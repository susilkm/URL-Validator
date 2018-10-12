using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace URLVaidator
{
    class Program
    {
        static void Main()
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(@"C:\Personal\Tools\URLVaidator\URLVaidator\Dynamics365-Pages-Redirect.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)wb.ActiveSheet;

            IterateRows(worksheet);
            Console.Read();
        }

        public static void IterateRows(Excel.Worksheet worksheet)
        {
            //Get the used Range
            Excel.Range usedRange = worksheet.UsedRange;

            //Iterate the rows in the used range
            foreach (Excel.Range row in usedRange.Rows)
            {
                //while (row.Row != 0)
                {


                    //Do something with the row.

                    //Ex. Iterate through the row's data and put in a string array
                    String[] rowData = new String[row.Columns.Count];
                    for (int i = 0; i < row.Columns.Count; i++)
                    {
                        if (i == 1)
                        {
                            rowData[i] = row.Cells[1, i + 1].Value2.ToString();
                            Console.WriteLine(rowData[i]);

                            WebRequest _request;
                            string text;
                            string url = rowData[i];
                            try
                            {
                                _request = (HttpWebRequest)WebRequest.Create(url);
                                //using (WebResponse response = _request.GetResponse())
                                //{
                                //    text = response.ResponseUri.ToString();

                                //    //using (StreamReader reader = new StreamReader(response.GetResponseStream()))
                                //    //{
                                //    //    text = reader.ReadToEnd();
                                //    //}
                                //}

                                //HttpStatusCode status = GetHeaders(url);

                                //Console.WriteLine("Status---> " + status);


                                text = GetFinalRedirectedUrl(url);

                                Console.WriteLine(text);
                            }
                            catch (Exception ex)
                            { }
                            Console.WriteLine("----------- " + i + " ----------------");
                        }
                    }
                }
            }
        }

        public static string GetFinalRedirectedUrl(string url)
        {
            string result = string.Empty;

            Uri Uris = new Uri(url);

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(Uris);
            req.Method = "HEAD";
            req.AllowAutoRedirect = false;

            HttpWebResponse myResp = (HttpWebResponse)req.GetResponse();
            if (myResp.StatusCode == HttpStatusCode.Redirect)
            {
                string temp = myResp.GetResponseHeader("Location");
                //Recursive call
                //result = GetFinalRedirectedUrl(temp);
                result = temp;
            }
            else
            {
                result = url;
            }

            return result;
        }

        public static HttpStatusCode GetHeaders(string url)
        {
            HttpStatusCode result = default(HttpStatusCode);

            var request = HttpWebRequest.Create(url);
            request.Method = "HEAD";
            using (var response = request.GetResponse() as HttpWebResponse)
            {
                if (response != null)
                {
                    result = response.StatusCode;
                    response.Close();
                }
            }

            return result;
        }

        public static void CreateExcel()
        {
        }
    }
}
