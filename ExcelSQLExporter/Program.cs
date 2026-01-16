using Microsoft.Data.SqlClient;
using System;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;
using Microsoft.VisualBasic.FileIO;
using System.ComponentModel.DataAnnotations;
using System.Data;
using NPOI.OpenXmlFormats.Wordprocessing;
using NPOI.SS.UserModel;
using System.IO;
using MathNet.Numerics.Optimization;
using Microsoft.Extensions.Configuration;
using WinSCP;
using System.Reflection;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;
using System.Globalization;
using ExcelSQLExporter.Services;

namespace ExcelSQLExporter
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            bool? logToFile = true;
            bool? outputToScreen = true;

            string? toolName = Assembly.GetExecutingAssembly().GetName().Name;
            string logFileName = $"{toolName} - {DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.log";

            await LoggingService.Log(toolName, logFileName, logToFile, outputToScreen);
            await LoggingService.Log("Export SQL Table or View to Excel File", logFileName, logToFile, outputToScreen);
            await LoggingService.Log("=========================================", logFileName, logToFile, outputToScreen);

            string? productVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString();
            await LoggingService.Log($"\nVersion {productVersion}", logFileName, logToFile, outputToScreen);
            await LoggingService.Log($"\nCopyright Robin Wilson", logFileName, logToFile, outputToScreen);

            string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");
            string? customConfigFile = null;
            if (args.Length >= 1)
            {
                customConfigFile = args[0];
            }

            if (!string.IsNullOrEmpty(customConfigFile))
            {
                configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, customConfigFile);
            }

            await LoggingService.Log($"\nUsing Config File {configFile}", logFileName, logToFile, outputToScreen);

            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile(configFile, optional: false);

            IConfiguration config;
            try
            {
                config = builder.Build();
            }
            catch (Exception e)
            {
                await LoggingService.Log($"Error: {e}", logFileName, logToFile, outputToScreen);
                return 1;
            }

            await LoggingService.Log($"\nSetting Locale To {config["Locale"]}", logFileName, logToFile, outputToScreen);

            //Set locale to ensure dates and currency are correct
            CultureInfo culture = new CultureInfo(config["Locale"] ?? "en-GB");
            Thread.CurrentThread.CurrentCulture = culture;
            Thread.CurrentThread.CurrentUICulture = culture;
            CultureInfo.DefaultThreadCurrentCulture = culture;
            CultureInfo.DefaultThreadCurrentUICulture = culture;

            var databaseConnection = config.GetSection("DatabaseConnection");
            var databaseTable = config.GetSection("DatabaseTable");
            var excelFile = config.GetSection("ExcelFile");
            var ftpConnection = config.GetSection("FTPConnection");

            var sqlConnection = new SqlConnectionStringBuilder
            {
                DataSource = databaseConnection["Server"],
                UserID = databaseConnection["Username"],
                Password = databaseConnection["Password"],
                IntegratedSecurity = databaseConnection.GetValue<bool>("UseWindowsAuth", false),
                InitialCatalog = databaseConnection["Database"],
                TrustServerCertificate = true
            };

            //If not using windows auth then need username and password values too
            if (sqlConnection.IntegratedSecurity == false)
            {
                sqlConnection.UserID = databaseConnection["Username"];
                sqlConnection.Password = databaseConnection["Password"];
            }

            var connectionString = sqlConnection.ConnectionString;

            //Database Connection
            await LoggingService.Log("\nConnecting to Database", logFileName, logToFile, outputToScreen);
            await using var connection = new SqlConnection(connectionString);
            string excelFilePath = "";

            string? columnNameAsFileNameValue = null;
            int? columnNameAsFileNameIndex = null;

            try
            {
                await connection.OpenAsync();
                await LoggingService.Log($"\nConnected to {sqlConnection.DataSource}", logFileName, logToFile, outputToScreen);
                

                string sql = "";

                if (databaseTable["StoredProcedureCommand"]?.Length > 0)
                {
                    await LoggingService.Log($"Executing Stored Procedure {databaseTable["StoredProcedureCommand"]}", logFileName, logToFile, outputToScreen);

                    sql = $@"[{databaseTable["Database"]}].[{databaseTable["Schema"]}].[{databaseTable["StoredProcedureCommand"]}]";
                }
                else
                {
                    await LoggingService.Log($"Loading data from table {databaseTable["TableOrView"]}", logFileName, logToFile, outputToScreen);

                    sql =
                        $@"SELECT *
                        FROM [{databaseTable["Database"]}].[{databaseTable["Schema"]}].[{databaseTable["TableOrView"]}]";
                }

                await using var command = new SqlCommand(sql, connection);

                //If stored procedure specified with parameters then add these
                if (databaseTable["StoredProcedureCommand"]?.Length > 0)
                {
                    command.CommandType = CommandType.StoredProcedure;

                    if (databaseTable["StoredProcedureParam1IntegerName"]?.Length > 0)
                    {
                        command.Parameters.AddWithValue("@" + databaseTable["StoredProcedureParam1IntegerName"], SqlDbType.Int).Value = databaseTable["StoredProcedureParam1IntegerValue"];
                    }
                    if (databaseTable["StoredProcedureParam2IntegerName"]?.Length > 0)
                    {
                        command.Parameters.AddWithValue("@" + databaseTable["StoredProcedureParam2IntegerName"], SqlDbType.Int).Value = databaseTable["StoredProcedureParam2IntegerValue"];
                    }
                    if (databaseTable["StoredProcedureParam1StringName"]?.Length > 0)
                    {
                        command.Parameters.AddWithValue("@" + databaseTable["StoredProcedureParam1StringName"], SqlDbType.NVarChar).Value = databaseTable["StoredProcedureParam1StringValue"];
                    }
                    if (databaseTable["StoredProcedureParam2StringName"]?.Length > 0)
                    {
                        command.Parameters.AddWithValue("@" + databaseTable["StoredProcedureParam2StringName"], SqlDbType.NVarChar).Value = databaseTable["StoredProcedureParam2StringValue"];
                    }
                }

                await using var reader = await command.ExecuteReaderAsync();


                await LoggingService.Log("\nLoading Data into Excel", logFileName, logToFile, outputToScreen);
                //Excel File from NPOI
                XSSFWorkbook book = new XSSFWorkbook();
                
                //Get Sheet Name
                string? sheetName = "Sheet1";
                if (!String.IsNullOrEmpty(excelFile["SheetName"]))
                {
                    sheetName = excelFile["SheetName"];
                }

                ISheet sheet = book.CreateSheet(sheetName);

                //Cell Styles
                var cellStyleDate = book.CreateCellStyle();
                cellStyleDate.DataFormat = book.CreateDataFormat().GetFormat(excelFile["DateFormat"]);
                var cellStyleTime = book.CreateCellStyle();
                cellStyleTime.DataFormat = book.CreateDataFormat().GetFormat(excelFile["TimeFormat"]);
                var cellStyleCurrency = book.CreateCellStyle();
                cellStyleCurrency.DataFormat = book.CreateDataFormat().GetFormat(excelFile["CurrencyFormat"]);

                int line = 0;
                while (await reader.ReadAsync())
                {
                    //Add top row with column names
                    if (line == 0)
                    {
                        var topRow = sheet.CreateRow(line);

                        for (int cell = 0; cell < reader.FieldCount; cell++)
                        {
                            //Get file name from column if specified and found
                            if (reader.GetName(cell) == excelFile["ColumnNameAsFileName"])
                            {
                                columnNameAsFileNameValue = reader.GetString(cell);
                                columnNameAsFileNameIndex = cell;
                                await LoggingService.Log($"Using Custom File Name from Table Column '{excelFile["ColumnNameAsFileName"]}': {columnNameAsFileNameValue}", logFileName, logToFile, outputToScreen);
                            }
                            else
                            {
                                var headerCell = topRow.CreateCell(cell);
                                headerCell.SetCellValue(reader.GetName(cell));
                            } 
                        }

                        line++;
                    }

                    //Add data underneath
                    var row = sheet.CreateRow(line);

                    for (int cell = 0; cell < reader.FieldCount; cell++)
                    {
                        //Skip column used for file name
                        if (cell == columnNameAsFileNameIndex)
                        {
                            continue;
                        }

                        var bodyCell = row.CreateCell(cell);
                        if (reader.IsDBNull(cell) == true)
                        {
                            bodyCell.SetCellValue("");
                        }
                        else if (reader.GetFieldType(cell) == typeof(Int32))
                        {
                            bodyCell.SetCellValue(reader.GetInt32(cell));
                        }
                        else if (reader.GetFieldType(cell) == typeof(Decimal))
                        {
                            bodyCell.SetCellValue((double)reader.GetDecimal(cell));
                            bodyCell.CellStyle = cellStyleCurrency;
                        }
                        else if (reader.GetFieldType(cell) == typeof(string))
                        {
                            bodyCell.SetCellValue(reader.GetString(cell));
                        }
                        else if (reader.GetFieldType(cell) == typeof(DateTime))
                        {
                            DateTime dateTimeValue = reader.GetDateTime(cell);
                            bodyCell.SetCellValue(dateTimeValue);

                            //If Date is default then value only contains time element
                            if (dateTimeValue.ToString("yyyy-MM-dd") == "1900-01-01")
                            {
                                bodyCell.CellStyle = cellStyleTime;
                            }
                            else
                            {
                                bodyCell.CellStyle = cellStyleDate;
                            }

                        }
                        else
                        {
                            bodyCell.SetCellValue(reader.GetString(cell));
                        }
                    }

                    line++;
                }

                await LoggingService.Log("\nSaving Excel file", logFileName, logToFile, outputToScreen);

                string[]? filePaths = { @excelFile["Folder"] ?? "", excelFile["FileName"] ?? "" };

                //If column name specified then use this as the file name instead of the one in the config file
                if (columnNameAsFileNameValue?.Length > 0)
                {
                    if (columnNameAsFileNameValue.Substring(columnNameAsFileNameValue.Length - 5) != ".xlsx")
                    {
                        columnNameAsFileNameValue = columnNameAsFileNameValue + ".xlsx";
                    }
                    filePaths = [ @excelFile["Folder"] ?? "", columnNameAsFileNameValue ?? "" ];
                }

                excelFilePath = Path.Combine(filePaths);

                using (var fileStream = File.Create(excelFilePath ?? ""))
                {
                    book.Write(fileStream);
                    await LoggingService.Log($"File Saved to {fileStream.Name}", logFileName, logToFile, outputToScreen);
                }

                await connection.CloseAsync();
            }
            catch (Exception e)
            {
                await LoggingService.Log(e.ToString(), logFileName, logToFile, outputToScreen);

                if (connection != null)
                {
                    await connection.CloseAsync();
                }

                return 1;
            }

            if (System.IO.File.Exists(excelFilePath))
            {

                if (ftpConnection.GetValue<bool?>("UploadFile", false) == true)
                {
                    // Setup session options
                    SessionOptions sessionOptions = new SessionOptions
                    {
                        HostName = ftpConnection["Server"],
                        PortNumber = ftpConnection.GetValue<int>("Port", 21),
                        UserName = ftpConnection["Username"],
                        Password = ftpConnection["Password"]
                    };

                    switch (ftpConnection?["Type"])
                    {
                        case "FTP":
                            sessionOptions.Protocol = Protocol.Ftp;
                            break;
                        case "FTPS":
                            sessionOptions.Protocol = Protocol.Ftp;
                            sessionOptions.FtpSecure = FtpSecure.Explicit;
                            sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = true;
                            break;
                        case "SFTP":
                            sessionOptions.Protocol = Protocol.Sftp;
                            sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = true;
                            break;
                        case "SCP":
                            sessionOptions.Protocol = Protocol.Scp;
                            sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = true;
                            break;
                        default:
                            sessionOptions.Protocol = Protocol.Ftp;
                            break;
                    }

                    if (ftpConnection?["SSHHostKeyFingerprint"]?.Length > 0)
                    {
                        sessionOptions.SshHostKeyFingerprint = ftpConnection["SSHHostKeyFingerprint"];
                        sessionOptions.GiveUpSecurityAndAcceptAnyTlsHostCertificate = false;
                    }
                        

                    switch (ftpConnection?["Mode"])
                    {
                        case "Active":
                            sessionOptions.FtpMode = FtpMode.Active;
                            break;
                        case "Passive":
                            sessionOptions.FtpMode = FtpMode.Passive;
                            break;
                        default:
                            sessionOptions.FtpMode = FtpMode.Passive;
                            break;
                    }

                    await LoggingService.Log("\nUploading Excel File", logFileName, logToFile, outputToScreen);
                    await LoggingService.Log($"Uploading File to {sessionOptions.HostName}", logFileName, logToFile, outputToScreen);

                    string uploadPath = Path.Combine("/", ftpConnection?["FolderPath"] ?? "");

                    if (uploadPath.Substring(uploadPath.Length - 1) != "/")
                    {
                        uploadPath = uploadPath + "/";
                    }

                    try
                    {
                        using (Session session = new Session())
                        {
                            //When publishing to a self-contained exe file need to specify the location of WinSCP.exe
                            session.ExecutablePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WinSCP.exe");

                            // Connect
                            session.Open(sessionOptions);

                            // Upload files
                            TransferOptions transferOptions = new TransferOptions();
                            transferOptions.TransferMode = TransferMode.Binary;

                            TransferOperationResult transferResult;
                            transferResult =
                                session.PutFiles(excelFilePath, uploadPath, false, transferOptions);

                            // Throw on any error
                            transferResult.Check();

                            // Print results
                            foreach (TransferEventArgs transfer in transferResult.Transfers)
                            {
                                await LoggingService.Log($"Upload of {transfer.FileName} succeeded", logFileName, logToFile, outputToScreen);
                            }
                        }

                        await LoggingService.Log($"File Uploaded to {sessionOptions.HostName} to {uploadPath + columnNameAsFileNameValue ?? excelFile["FileName"] ?? ""}", logFileName, logToFile, outputToScreen);
                        return 0;
                    }
                    catch (Exception e)
                    {
                        await LoggingService.Log($"Error: {e}", logFileName, logToFile, outputToScreen);
                        return 1;
                    }
                }
                else
                {
                    await LoggingService.Log($"Not Uploading File to FTP as Option in Config is False", logFileName, logToFile, outputToScreen);
                    return 0;
                }
            }
            else
            {
                await LoggingService.Log($"The File at {excelFilePath} Could Not Be Found", logFileName, logToFile, outputToScreen);
                return 1;
            }
        }
    }
}