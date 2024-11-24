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

namespace ExcelSQLExporter
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            Console.WriteLine("\nExport SQL Table or View to Excel File");
            Console.WriteLine("=========================================\n");
            Console.WriteLine("Copyright Robin Wilson");

            string configFile = "appsettings.json";
            string? customConfigFile = null;
            if (args.Length >= 1)
            {
                customConfigFile = args[0];
            }

            if (!string.IsNullOrEmpty(customConfigFile))
            {
                configFile = customConfigFile;
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

            var databaseConnection = config.GetSection("DatabaseConnection");
            var excelFile = config.GetSection("ExcelFile");
            var ftpConnection = config.GetSection("FTPConnection");
            string excelFilePath = excelFile["Folder"] + "\\" + excelFile["FileName"];

            var sqlConnection = new SqlConnectionStringBuilder
            {
                DataSource = databaseConnection["Server"],
                UserID = databaseConnection["Username"],
                Password = databaseConnection["Password"],
                InitialCatalog = databaseConnection["Database"],
                TrustServerCertificate = true
            };

            var connectionString = sqlConnection.ConnectionString;

            try
            {
                //Database Connection
                await using var connection = new SqlConnection(connectionString);
                Console.WriteLine("Connecting to Database\n");

                await connection.OpenAsync();
                Console.WriteLine($"\nConnected to {sqlConnection.DataSource}");

                var sql =
                    @"SELECT *
                    FROM ProSolutionReports.dbo.VW_WEB_CourseMarketingData CRS";

                await using var command = new SqlCommand(sql, connection);
                await using var reader = await command.ExecuteReaderAsync();


                Console.WriteLine("Loading Data into Excel\n");
                //Excel File from NPOI
                var book = new XSSFWorkbook();
                var sheet = book.CreateSheet("Sheet1");

                //Cell Styles
                var cellStyleDate = book.CreateCellStyle();
                cellStyleDate.DataFormat = book.CreateDataFormat().GetFormat("dd/MM/yyyy");
                var cellStyleTime = book.CreateCellStyle();
                cellStyleTime.DataFormat = book.CreateDataFormat().GetFormat("HH:mm");
                var cellStyleCurrency = book.CreateCellStyle();
                cellStyleCurrency.DataFormat = book.CreateDataFormat().GetFormat("£#,0");

                int line = 0;
                while (await reader.ReadAsync())
                {
                    //Add top row with column names
                    if (line == 0)
                    {
                        var topRow = sheet.CreateRow(line);

                        for (int cell = 0; cell < reader.FieldCount; cell++)
                        {
                            var headerCell = topRow.CreateCell(cell);
                            headerCell.SetCellValue(reader.GetName(cell));
                        }

                        line++;
                    }

                    //Add data underneath
                    var row = sheet.CreateRow(line);

                    for (int cell = 0; cell < reader.FieldCount; cell++)
                    {
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

                Console.WriteLine("Saving Excel file");
                using (var fileStream = File.Create(excelFilePath ?? ""))
                {
                    book.Write(fileStream);
                    Console.WriteLine($"File Saved to {fileStream.Name}");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
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

                    switch (ftpConnection["Type"])
                    {
                        case "FTP":
                            sessionOptions.Protocol = Protocol.Ftp;
                            break;
                        case "FTPS":
                            sessionOptions.Protocol = Protocol.Ftp;
                            sessionOptions.FtpSecure = FtpSecure.Explicit;
                            break;
                        case "SFTP":
                            sessionOptions.Protocol = Protocol.Sftp;
                            break;
                        default:
                            sessionOptions.Protocol = Protocol.Ftp;
                            break;
                    }

                    switch (ftpConnection["Mode"])
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

                    Console.WriteLine($"Uploading File to {sessionOptions.HostName}");

                    try
                    {
                        using (Session session = new Session())
                        {
                            // Connect
                            session.Open(sessionOptions);

                            // Upload files
                            TransferOptions transferOptions = new TransferOptions();
                            transferOptions.TransferMode = TransferMode.Binary;

                            TransferOperationResult transferResult;
                            transferResult =
                                session.PutFiles(excelFilePath, "/", false, transferOptions);

                            // Throw on any error
                            transferResult.Check();

                            // Print results
                            foreach (TransferEventArgs transfer in transferResult.Transfers)
                            {
                                Console.WriteLine("Upload of {0} succeeded", transfer.FileName);
                            }
                        }

                        Console.WriteLine($"File Uploaded to {sessionOptions.HostName}");
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