using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace ReadExcelProject
{
    /// <summary>
    /// // Requires EEPlus Nuget Package 
    /// </summary>
    class Program
    {        
        private const string _baseUrl = "https://octo.quickbase.com/db/";
        private const string _username = "rahulmisra2000@gmail.com";
        private const string _password = "123exceltoquickbase";
        private const string _appToken = "bwy3haxdh8fsi7v6p6k7dyexycn";
        private string _authTicket;

        private string _fileName = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) + @"\1.xlsx";
        private const string _tableId = "bkr5uxfek";

        private Stopwatch _stopWatch = new Stopwatch();    
        private DataTable dt = null;
        private string _finalMessage="";
        private const int FAIL = 100;
        private const int PASS = 101;        
        
        private const string _fromEmail = "ExcelToQuickBase@gmail.com"; // In gmail enable allow less secure apps
        private const string _fromPwd = "";                             // hidden because going to github
        private const string _toEmail = "rahulmisra2000@gmail.com";

        private enum AppFeatures
        {
            None = 0,
            Fiddler = 1,
            SendEmail = 2,
            EventViewer = 4,
            ConsoleOutput = 8,
            WriteDTtoScreen = 16,
            HttpPostPayload = 32,
            WriteToLogFile = 64
        }

        private AppFeatures _enabledFeatures;


        static void Main(string[] args)
        {            
            new Program().Run();                
        }
        
        private void Run()
        {
            BeginHouseKeeping();            
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            ProcessWorkFlow();
            EndHouseKeeping();                  
        }

        private void WriteToConsole(string v, bool wait=false)
        {
            if ((_enabledFeatures & AppFeatures.ConsoleOutput)== AppFeatures.ConsoleOutput)
            {
                Console.WriteLine(DateTime.Now + " : " + v);                
            }

            if (wait) {
                Console.WriteLine(DateTime.Now + " : Hit <Enter> to continue");
                Console.ReadLine();
            }
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            _finalMessage = e.ExceptionObject.ToString();
            LogToEventViewer(FAIL);
            WriteToLogFile();
            WriteToConsole("Severe Error: Check Event Viewer. Hit <Enter> to continue...",true);            
            Environment.Exit(1);
        }

        private void LogToEventViewer(int id)
        {
            if ((_enabledFeatures & AppFeatures.EventViewer) == AppFeatures.EventViewer)
            {
                using (EventLog eventLog = new EventLog("Application"))
                {
                    eventLog.Source = "Application";
                    eventLog.WriteEntry(DateTime.Now + " : " + _finalMessage, 
                                        id == FAIL ? EventLogEntryType.Error : EventLogEntryType.Information, 
                                        id, 
                                        1);
                }
            }
        }
        

        private void WriteToLogFile()
        {
            if ((_enabledFeatures & AppFeatures.WriteToLogFile) == AppFeatures.WriteToLogFile)
            {
                using (StreamWriter sw = File.AppendText("Log.txt"))
                {
                    sw.WriteLine(DateTime.Now + " : " + _finalMessage);
                }
            }
        }

        private void ProcessWorkFlow()
        {                                    
            if (!SuccessfullAuthenticated()) { _finalMessage="Failed Authentication"; return; }
            if (!SuccessfullyHydratedDataTable()) { _finalMessage = "Failed To Read Excel"; return; }
            WriteDataTableToScreen();
            if (!SuccessfullyWrittenToQuickBase()) { _finalMessage = "Failed To Write to QuickBase"; return; }
            // All OK
            _stopWatch.Stop();
            TimeSpan ts = _stopWatch.Elapsed;
            _finalMessage = String.Format("Successfully wrote {0} Records In {1:00}:{2:00}:{3:00}.{4:00}",
                                            dt.Rows.Count, ts.Hours, ts.Minutes, ts.Seconds, ts.Milliseconds / 10);
        }

        private void BeginHouseKeeping()
        {
            _enabledFeatures =  AppFeatures.SendEmail | 
                                AppFeatures.EventViewer | 
                                AppFeatures.ConsoleOutput | 
                                AppFeatures.WriteToLogFile;

            WriteToConsole("Working ...");
            _stopWatch.Start();
        }

        private void EndHouseKeeping()
        {            
            WriteToLogFile();
            SendEmail();
            LogToEventViewer(_finalMessage.Contains("Failed") ? FAIL : PASS);
            WriteToConsole(_finalMessage,true);
        }

        /// <summary>
        /// Remember to turn on "Access to less Secure apps in gmail
        /// </summary>
        private void SendEmail()
        {
            if ((_enabledFeatures & AppFeatures.SendEmail) == AppFeatures.SendEmail)
            {
                var client = new SmtpClient("smtp.gmail.com", 587)
                {
                    Credentials = new NetworkCredential(_fromEmail, _fromPwd),
                    EnableSsl = true
                };
                client.Send(_fromEmail, _toEmail, "Notification: Excel to QuickBase", DateTime.Now + " : " + _finalMessage);
            }
        }

        private bool SuccessfullyWrittenToQuickBase()
        {

            StringBuilder sb = new StringBuilder();
            foreach (DataRow r in dt.Rows)
            {
                sb.Append(String.Join(",", r.ItemArray) + "\r\n");
            }
            string csv_values = @"<![CDATA[" +
                                    sb.ToString() +
                                    @"]]>";
            if ((_enabledFeatures & AppFeatures.HttpPostPayload) == AppFeatures.HttpPostPayload) { WriteToConsole(csv_values); }
                
            string clist = "6.7.8.9.10.11.12.13.14.15.16.17.18.19.20.21.22.23.24.25.26.27.28.29.30.31.32.33.34.35.36.37";

            string result = CallApiParams(_baseUrl + _tableId, "API_ImportFromCSV", _authTicket, _appToken,
                        "records_csv", csv_values,
                        "clist", clist,
                        "skipfirst", 0);

            string errcode;
            ParseResult(result, "errcode", out errcode);
            return (Convert.ToInt32(errcode) == 0);
        }

        private bool SuccessfullAuthenticated()
        {            
            string result = CallApiParams(_baseUrl + "main", "API_Authenticate", null, null,"username", _username,"password", _password, "hours", 24);

            if (ParseResult(result, "ticket", out _authTicket))  { return false; }
            return true;            
        }


        private string CallApiParams(string url, string quickbase_action, string ticket, string appToken, params object[] name_value)
        {
            string values = "";
            for (int i = 0; i + 1 < name_value.Length; i += 2)
            {
                values += @"<" + name_value[i].ToString() + ">" + name_value[i + 1].ToString() + "</" + name_value[i].ToString() + ">";
            }
            return CallApi(url, quickbase_action, ticket, appToken, values);
        }

        private string CallApi(string url, string quickbase_action, string ticket, string appToken, string values)
        {
            WriteToConsole(string.Format("Step - {0} @ {1}",quickbase_action,url));
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/xml";
            request.Headers.Add("QUICKBASE-ACTION", quickbase_action);            
            if ((_enabledFeatures & AppFeatures.Fiddler) == AppFeatures.Fiddler) { request.Proxy = new WebProxy("127.0.0.1", 8888); }
            using (TextWriter writer = new StreamWriter(request.GetRequestStream()))
            {
                string temp = @"<qdbapi><udata>optional data</udata>" +
                                       (ticket != null ? @"<ticket>" + ticket + @"</ticket>" : "") +
                                       (appToken != null ? @"<apptoken>" + appToken + @"</apptoken>" : "") +
                                       (values != null ? values : "") +
                                 @"</qdbapi>";

                writer.Write(temp);
            }
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            using (TextReader reader = new StreamReader(response.GetResponseStream()))
                return reader.ReadToEnd();
        }


        private bool ParseResult(string result, string tag, out string data)
        {
            Match match = Regex.Match(result, @"<" + tag + @">([^\s<>]+)</" + tag + @">", RegexOptions.IgnoreCase);
            data = match.Success ? match.Groups[1].Value : "";
            return !match.Success;
        }


        private void WriteDataTableToScreen()
        {
            if ((_enabledFeatures & AppFeatures.WriteDTtoScreen) == AppFeatures.WriteDTtoScreen)
            {
                StringBuilder sb = new StringBuilder();
                foreach (DataRow r in dt.Rows)
                {
                    sb.Append(String.Join(",", r.ItemArray) + "\r\n");
                }
                string csv_values = @"<![CDATA[" +
                                        sb.ToString() +
                                        @"]]>";
                Console.WriteLine(csv_values);
            }
        }

        private bool SuccessfullyHydratedDataTable()
        {
            try
            {
                using (ExcelPackage pck = new ExcelPackage())
                {
                    using (var stream = File.OpenRead(_fileName))
                    {
                        pck.Load(stream);
                    }

                    ExcelWorksheet ws = pck.Workbook.Worksheets.First();
                    dt = WorksheetToDataTable(ws);
                }                
            }
            catch (Exception ex)
            {
                _finalMessage = ex.Message;
                WriteToConsole(DateTime.Now + " : " + "Reading of Excel failed. "  + ex.Message);
                LogToEventViewer(FAIL);
                return false;
            }
            return true;
        }

        private DataTable WorksheetToDataTable(ExcelWorksheet ws, bool hasHeader = true)
        {
            DataTable dt = new DataTable(ws.Name);
            int totalCols = ws.Dimension.End.Column;
            int totalRows = ws.Dimension.End.Row;
            int startRow = hasHeader ? 2 : 1;
            ExcelRange wsRow;
            DataRow dr;
            int unq = 0;
            foreach (var firstRowCell in ws.Cells[1, 1, 1, totalCols])
            {
              dt.Columns.Add(hasHeader ? firstRowCell.Text + unq++ : string.Format("Column {0}", firstRowCell.Start.Column));
            }

            for (int rowNum = startRow; rowNum <= totalRows; rowNum++)
            {
                if (rowNum % 100 == 0) {
                    WriteToConsole(string.Format("Step - {0} of {1} Records From Excel Processed", rowNum, totalRows));
                }
                wsRow = ws.Cells[rowNum, 1, rowNum, totalCols];
                dr = dt.NewRow();
                foreach (var cell in wsRow)
                {
                    dr[cell.Start.Column - 1] = cell.Text;
                }

                dt.Rows.Add(dr);
            }
            WriteToConsole(string.Format("Step - {0} of {1} Records From Excel Processed", totalRows, totalRows));
            return dt;
        }
    }
}

