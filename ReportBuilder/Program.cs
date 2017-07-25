using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Mail;
using Ionic.Zip;
using Winnovative.ExcelLib;

namespace ReportBuilder
{
    public class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                int result;
                int.TryParse(args[0], out result);

                BuildReport(result, true);
            }
            else
            {
                StartManual();
            }
        }

        private static void StartManual()
        {
            Console.WriteLine(string.Empty);
            Console.WriteLine("Please enter the report number you wish to send");

            int result;
            int.TryParse(Console.ReadLine(), out result);

            if (result != 0)
            {
                goto confirmSelection;
            }

            selectReport:
            do
            {
                Console.WriteLine(string.Empty);
                Console.WriteLine("Please enter the report number you wish to send");
                int.TryParse(Console.ReadLine(), out result);
            }
            while (result == 0);

            confirmSelection:
            Console.WriteLine(string.Empty);
            Console.WriteLine("You have requested to send report id " + (object)result + ". Are you sure (Y/N)");

            var readLine = Console.ReadLine();
            if (readLine != null && readLine.ToLower() == "y")
            {
                BuildReport(result, false);
            }
            else
            {
                goto selectReport;
            }
        }

        private static void BuildReport(int reportNo, bool autoClose)
        {
            var resultsFound = false;
            var reportName = string.Empty;
            var emailList = new List<string>();
            var fileType = string.Empty;
            var query = string.Empty;
            var zipFile = false;
            var passwordProtect = false;
            var thePassword = string.Empty;
            var fileName = string.Empty;
            var gpgProtect = false;
            var gpgKey = string.Empty;
            var sendIfNoRecords = false;
            var reportLocation = string.Empty;
            var fromEmail = string.Empty;

            using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString))
            {
                using (var cmd = new SqlCommand("SELECT * FROM dbo.tbl_Report_Entries WHERE dbo.tbl_Report_Entries.report_id = " + reportNo, conn))
                {
                    conn.Open();
                    var dr = cmd.ExecuteReader();

                    if (dr.HasRows)
                    {
                        dr.Read();

                        resultsFound = true;
                        reportName = dr["report_name"].ToString();
                        emailList = dr["distrubution_list"].ToString().Split(',').ToList();
                        fileType = dr["file_type"].ToString();
                        query = dr["query"].ToString();
                        zipFile = bool.Parse(dr["zip_file"].ToString());
                        passwordProtect = bool.Parse(dr["password_protect"].ToString());
                        thePassword = dr["password"].ToString();
                        fileName = dr["report_filename"].ToString();
                        gpgProtect = bool.Parse(dr["gpg_protect"].ToString());
                        gpgKey = dr["gpg_key"].ToString();
                        sendIfNoRecords = bool.Parse(dr["send_if_no_records"].ToString());
                        reportLocation = dr["report_location"].ToString();
                        fromEmail = dr["from_email"].ToString();
                    }
                }
            }

            if (resultsFound)
            {
                var data = GetData(query);

                Console.WriteLine(string.Empty);
                Console.WriteLine("Report generated.");

                if (!Directory.Exists(reportLocation))
                {
                    Directory.CreateDirectory(reportLocation);
                }

                var path = reportLocation + reportName + "\\" + DateTime.Now.Year + "_" + DateTime.Now.Month + "_" +
                           DateTime.Now.Day + "_" + DateTime.Now.ToShortTimeString().Replace(":", string.Empty);
                var filePath = path + "\\" + fileName.Replace("{DD}", DateTime.Now.ToString("dd"))
                                   .Replace("{MM}", DateTime.Now.ToString("MM"))
                                   .Replace("{YY}", DateTime.Now.ToString("yy"));

                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                var fullFilePath = string.Empty;

                switch (fileType)
                {
                    case "CSV":
                        fullFilePath = filePath + ".csv";
                        SaveAsCsv(data, fullFilePath);
                        break;

                    case "XLS":
                        fullFilePath = filePath + ".xls";
                        SaveAsXls(data, fullFilePath);
                        break;

                    case "XLSX":
                        fullFilePath = filePath + ".xlsx";
                        SaveAsXlsx(data, fullFilePath);
                        break;
                }

                if (zipFile)
                {
                    Console.WriteLine(string.Empty);
                    Console.WriteLine("Zipping report...");

                    using (var zipFileObj = new ZipFile())
                    {
                        if (passwordProtect)
                        {
                            zipFileObj.Password = thePassword;
                            Console.WriteLine(string.Empty);
                            Console.WriteLine("Password protecting zip...");
                        }

                        zipFileObj.AddFile(fullFilePath, string.Empty);
                        zipFileObj.Save(filePath + ".zip");
                    }
                }

                if (gpgProtect)
                {
                    var fileInfo = new FileInfo(zipFile ? filePath + ".zip" : fullFilePath);
                    using (var process = Process.Start(new ProcessStartInfo("cmd.exe")
                    {
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardInput = true,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        WorkingDirectory = fileInfo.DirectoryName
                    }))
                    {
                        var str4 = "\"" + "C:\\Program Files (x86)\\GNU\\GnuPG\\pub\\gpg.exe" + "\"" +
                                   (" --recipient \"" + gpgKey + "\"") +
                                   (" --encrypt \"" + (zipFile ? filePath + ".zip" : fullFilePath) + "\"");
                        process.StandardInput.WriteLine(str4);
                        process.StandardInput.Flush();
                        process.StandardInput.Close();
                        process.WaitForExit(3500);
                        process.Close();
                    }
                }

                Console.WriteLine(string.Empty);
                Console.WriteLine("Sending report...");

                if (data.Rows.Count > 0 || data.Rows.Count == 0 && sendIfNoRecords)
                {
                    using (var smtpClient = new SmtpClient())
                    {
                        using (var message = new MailMessage())
                        {
                            foreach (var address in emailList)
                            {
                                message.To.Add(new MailAddress(address));
                            }

                            message.From = new MailAddress(fromEmail);
                            message.Subject = reportName;
                            message.Body = string.Empty;
                            message.IsBodyHtml = false;

                            if (gpgProtect)
                            {
                                message.Attachments.Add(zipFile
                                    ? new Attachment(filePath + ".zip.gpg")
                                    : new Attachment(filePath + ".gpg"));
                            }
                            else
                            {
                                message.Attachments.Add(zipFile
                                    ? new Attachment(filePath + ".zip")
                                    : new Attachment(fullFilePath));
                            }

                            smtpClient.Send(message);
                        }
                    }
                    Console.WriteLine(string.Empty);
                    Console.WriteLine("Report sent.");
                }
                else
                {
                    Console.WriteLine(string.Empty);
                    Console.WriteLine("Report not sent because no rows were found.");
                }
            }
            else
            {
                Console.WriteLine(string.Empty);
                Console.WriteLine("Report not found.");
            }

            if (autoClose)
            {
                return;
            }

            StartManual();
        }

        private static void SaveAsXlsx(DataTable data, string reportFileName)
        {
            var excelWorkbook2 = new ExcelWorkbook(ExcelWorkbookFormat.Xlsx_2007);
            var worksheet2 = excelWorkbook2.Worksheets[0];
            worksheet2.Name = "Sheet1";
            worksheet2.LoadDataTable(data, 1, 1, true);
            worksheet2.AutofitColumns();
            excelWorkbook2.Save(reportFileName);

            Console.WriteLine(string.Empty);
            Console.WriteLine("Saving report in XLSX format...");
        }

        private static void SaveAsXls(DataTable data, string reportFileName)
        {
            var excelWorkbook1 = new ExcelWorkbook(ExcelWorkbookFormat.Xls_2003);
            var worksheet1 = excelWorkbook1.Worksheets[0];
            worksheet1.Name = "Sheet1";
            worksheet1.LoadDataTable(data, 1, 1, true);
            worksheet1.AutofitColumns();
            excelWorkbook1.Save(reportFileName);

            Console.WriteLine(string.Empty);
            Console.WriteLine("Saving report in XLS format...");
        }

        private static void SaveAsCsv(DataTable data, string reportFileName)
        {
            CreateCsvFile(data, reportFileName);
            Console.WriteLine(string.Empty);
            Console.WriteLine("Saving report in CSV format...");
        }

        private static DataTable GetData(string query)
        {
            Console.WriteLine(string.Empty);
            Console.WriteLine("Building report...");

            var dataTable = new DataTable();
            if (query != string.Empty)
            {
                using (var connection = new SqlConnection(ConfigurationManager.ConnectionStrings["DbConnection"].ConnectionString))
                {
                    using (var sqlCommand = new SqlCommand(query, connection))
                    {
                        connection.Open();

                        sqlCommand.CommandTimeout = 600;

                        var sqlDataReader = sqlCommand.ExecuteReader();
                        dataTable.Load(sqlDataReader);
                        sqlDataReader.Dispose();
                    }
                }
            }

            return dataTable;
        }

        private static void CreateCsvFile(DataTable dt, string strFilePath)
        {
            var streamWriter = new StreamWriter(strFilePath, false);
            var count = dt.Columns.Count;

            for (var index = 0; index < count; ++index)
            {
                streamWriter.Write(((char)34) + dt.Columns[index].ToString() + '"');
                if (index < count - 1)
                {
                    streamWriter.Write(",");
                }
            }

            streamWriter.Write(streamWriter.NewLine);
            foreach (DataRow row in (InternalDataCollectionBase)dt.Rows)
            {
                for (var index = 0; index < count; ++index)
                {
                    if (!Convert.IsDBNull(row[index]))
                    {
                        streamWriter.Write(((char)34).ToString() + row[index] + '"');
                    }

                    if (index < count - 1)
                    {
                        streamWriter.Write(",");
                    }
                }

                streamWriter.Write(streamWriter.NewLine);
            }
            streamWriter.Close();
        }
    }
}