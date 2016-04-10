using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
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
        private const string _fileName = "1.xlsx";
        private const string _tableId = "bkr5uxfek";
        private Stopwatch _stopWatch = new Stopwatch();    
        private DataTable dt = null;
        private string _finalMessage="";

        static void Main(string[] args)
        {
            Console.WriteLine("Working...");
            new Program().Run();                
        }

        private void Run()
        {
            ProcessWorkFlow();
            using (EventLog eventLog = new EventLog("Application"))
            {
                eventLog.Source = "Application";
                eventLog.WriteEntry(_finalMessage, EventLogEntryType.Information, 101, 1);
            }
        }

        private void ProcessWorkFlow()
        {
            
            BeginHouseKeeping();
            
            if (!SuccessfullAuthenticated()) { _finalMessage="Failed Authentication"; return; }
            if (!SuccessfullyHydratedDataTable()) { _finalMessage = "Failed To Read Excel"; return; }
            // DEBUG: WriteDataTableToScreen(dt);
            if (!SuccessfullyWrittenToQuickBase()) { _finalMessage = "Failed To Write to QuickBase"; return; }

            EndHouseKeeping();            
        }

        private void BeginHouseKeeping()
        {
            _stopWatch.Start();
        }

        private void EndHouseKeeping()
        {
            _stopWatch.Stop();
            TimeSpan ts = _stopWatch.Elapsed;
            _finalMessage = String.Format("Successfull wrote {0} Records In {1:00}:{2:00}:{3:00}.{4:00}", dt.Rows.Count, ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
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
            HttpWebRequest request = (HttpWebRequest)HttpWebRequest.Create(url);
            request.Method = "POST";
            request.ContentType = "application/xml";
            request.Headers.Add("QUICKBASE-ACTION", quickbase_action);
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
                Console.WriteLine("Reading of Excel failed. " + ex.Message);
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
                wsRow = ws.Cells[rowNum, 1, rowNum, totalCols];
                dr = dt.NewRow();
                foreach (var cell in wsRow)
                {
                    dr[cell.Start.Column - 1] = cell.Text;
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }
    }
}

