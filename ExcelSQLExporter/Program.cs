﻿using Microsoft.Data.SqlClient;
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

namespace ExcelSQLExporter
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            Console.WriteLine("\nExport SQL Table or View to Excel File");
            Console.WriteLine("=========================================\n");

            string? productVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString();
            Console.WriteLine($"Version {productVersion}");
            Console.WriteLine($"Copyright Robin Wilson");

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

            Console.WriteLine($"\nUsing Config File {configFile}");

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
                Console.WriteLine("Error: {0}", e);
                return 1;
            }

            Console.WriteLine($"\nSetting Locale To {config["Locale"]}");

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
                InitialCatalog = databaseConnection["Database"],
                TrustServerCertificate = true
            };

            var connectionString = sqlConnection.ConnectionString;

            //Database Connection
            Console.WriteLine("Connecting to Database\n");
            await using var connection = new SqlConnection(connectionString);
            string excelFilePath = "";

            string? columnNameAsFileNameValue = null;
            int? columnNameAsFileNameIndex = null;

            try
            {
                await connection.OpenAsync();
                Console.WriteLine($"\nConnected to {sqlConnection.DataSource}");
                

                string sql = "";

                if (databaseTable["StoredProcedureCommand"]?.Length > 0)
                {
                    Console.WriteLine($"Executing Stored Procedure {databaseTable["StoredProcedureCommand"]}");

                    sql = $@"[{databaseTable["Database"]}].[{databaseTable["Schema"]}].[{databaseTable["StoredProcedureCommand"]}]";
                }
                else
                {
                    Console.WriteLine($"Loading data from table {databaseTable["TableOrView"]}");

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


                Console.WriteLine("\nLoading Data into Excel");
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
                                Console.WriteLine($"Using Custom File Name from Table Column '{excelFile["ColumnNameAsFileName"]}': {columnNameAsFileNameValue}");
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

                Console.WriteLine("\nSaving Excel file");

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
                    Console.WriteLine($"File Saved to {fileStream.Name}");
                }

                await connection.CloseAsync();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());

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

                    Console.WriteLine("\nUploading Excel File");
                    Console.WriteLine($"Uploading File to {sessionOptions.HostName}");

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
                                Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                            }
                        }

                        Console.WriteLine($"File Uploaded to {sessionOptions.HostName} to {uploadPath + columnNameAsFileNameValue ?? excelFile["FileName"] ?? ""}");
                        return 0;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Error: {0}", e);
                        return 1;
                    }
                }
                else
                {
                    Console.WriteLine($"Not Uploading File to FTP as Option in Config is False");
                    return 0;
                }
            }
            else
            {
                Console.WriteLine($"The File at {excelFilePath} Could Not Be Found");
                return 1;
            }
        }
    }
}