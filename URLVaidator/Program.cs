using System;
using System.Collections.Generic;
using System.Data;
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
        static DataTable redirectsTable;

        static void Main()
        {
            Excel.Application excel = new Excel.Application();
            Excel.Workbook wb = excel.Workbooks.Open(@"C:\Personal\Tools\URLVaidator\URLVaidator\Dynamics365-Pages-Redirect-1.xlsx");
            Excel.Worksheet worksheet = (Excel.Worksheet)wb.ActiveSheet;

            IterateRows(worksheet);
            
            //Create a DataSet with the existing DataTables
            DataSet ds = new DataSet("Dynamics365");
            ds.Tables.Add(redirectsTable);

            try
            {
                ExportDataSetToExcel(ds);
            }
            catch (Exception ex)
            {
            }
            finally
            {
                Console.WriteLine("Excel sheet update successfull");
            }

        }

        public static void IterateRows(Excel.Worksheet worksheet)
        {
            redirectsTable = new DataTable();
            redirectsTable.Columns.Add("Locale", typeof(string));
            redirectsTable.Columns.Add("Page URL", typeof(string));
            redirectsTable.Columns.Add("Redirect URL", typeof(string));

            //Get the used Range
            Excel.Range usedRange = worksheet.UsedRange;

            //Iterate the rows in the used range
            int count = 0;
            
            foreach (Excel.Range row in usedRange.Rows)
            {
                string locale = string.Empty, 
                    url = string.Empty, 
                    redirectURL = string.Empty;

                //while (row.Row != 0)
                {


                    //Do something with the row.

                    //Ex. Iterate through the row's data and put in a string array
                    String[] rowData = new String[row.Columns.Count];
                    for (int i = 0; i < row.Columns.Count; i++)
                    {
                        //if (i == 1 || i == 2)
                        {
                            rowData[i] = row.Cells[1, i + 1].Value2.ToString();
                            Console.WriteLine(rowData[i]);
                            if (i == 2)
                            {
                                locale = rowData[i];
                            }
                            WebRequest _request;
                            if (i == 1)
                            {
                                url = rowData[i];
                                try
                                {
                                    _request = (HttpWebRequest)WebRequest.Create(url);

                                    redirectURL = GetFinalRedirect(url);

                                    Console.WriteLine(redirectURL);


                                }
                                catch (Exception ex)
                                { }
                            }

                            Console.WriteLine("----------- " + count++ + " ----------------");
                        }
                    }
                }
                redirectsTable.Rows.Add(locale, url, redirectURL);
            }
        }
                
        public static string GetFinalRedirect(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
                return url;

            int maxRedirCount = 8;  // prevent infinite loops
            string newUrl = url;
            do
            {
                HttpWebRequest req = null;
                HttpWebResponse resp = null;
                try
                {
                    req = (HttpWebRequest)HttpWebRequest.Create(url);
                    req.Method = "HEAD";
                    req.AllowAutoRedirect = false;
                    resp = (HttpWebResponse)req.GetResponse();
                    switch (resp.StatusCode)
                    {
                        case HttpStatusCode.OK:
                            return newUrl;
                        case HttpStatusCode.Redirect:
                        case HttpStatusCode.MovedPermanently:
                        case HttpStatusCode.RedirectKeepVerb:
                        case HttpStatusCode.RedirectMethod:
                            newUrl = resp.Headers["Location"];
                            if (newUrl == null)
                                return url;

                            if (newUrl.IndexOf("://", System.StringComparison.Ordinal) == -1)
                            {
                                // Doesn't have a URL Schema, meaning it's a relative or absolute URL
                                Uri u = new Uri(new Uri(url), newUrl);
                                newUrl = u.ToString();
                            }
                            break;
                        default:
                            return newUrl;
                    }
                    url = newUrl;
                }
                catch (WebException)
                {
                    // Return the last known good URL
                    return newUrl;
                }
                catch (Exception ex)
                {
                    return null;
                }
                finally
                {
                    if (resp != null)
                        resp.Close();
                }
            } while (maxRedirCount-- > 0);

            return newUrl;
        }

        public static System.Data.DataTable ExportToExcel()
        {
            DataTable table = new DataTable();
            table.Columns.Add("Page URL", typeof(string));
            table.Columns.Add("Redirect URL", typeof(string));

            table.Rows.Add("Amar", "M");
            table.Rows.Add("Mohit", "M");
            
            return table;
        }

        public static void ExportDataSetToExcel(DataSet ds)
        {
            //Creae an Excel application instance
            Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = excelApp.Workbooks.Open(@"C:\Personal\Tools\URLVaidator\Dynamics365-Sitemuse-Redirects.xlsx");

            foreach (DataTable table in ds.Tables)
            {
                //Add a new worksheet to workbook with the Datatable name
                Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                excelWorkSheet.Name = table.TableName;

                for (int i = 1; i < table.Columns.Count + 1; i++)
                {
                    excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }

                for (int j = 0; j < table.Rows.Count; j++)
                {
                    for (int k = 0; k < table.Columns.Count; k++)
                    {
                        excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                    }
                }
            }

            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApp.Quit();

        }


        public static void CreateExcel()
        {
            Excel.Application excel;
            Excel.Workbook worKbooK;
            Excel.Worksheet worKsheeT;
            Excel.Range celLrangE;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;
                worKbooK = excel.Workbooks.Add(Type.Missing);


                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "Redirects";

                worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[1, 8]].Merge();
                worKsheeT.Cells[1, 1] = "Redirects";
                worKsheeT.Cells.Font.Size = 15;


                int rowcount = 2;

                foreach (DataRow datarow in redirectsTable.Rows)
                {
                    rowcount += 1;
                    for (int i = 1; i <= redirectsTable.Columns.Count; i++)
                    {

                        if (rowcount == 3)
                        {
                            worKsheeT.Cells[2, i] = redirectsTable.Columns[i - 1].ColumnName;
                        }

                        worKsheeT.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == redirectsTable.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    celLrangE = worKsheeT.Range[worKsheeT.Cells[rowcount, 1], worKsheeT.Cells[rowcount, redirectsTable.Columns.Count]];
                                }

                            }
                        }

                    }

                }

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[rowcount, redirectsTable.Columns.Count]];
                celLrangE.EntireColumn.AutoFit();
                Microsoft.Office.Interop.Excel.Borders border = celLrangE.Borders;
                border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                celLrangE = worKsheeT.Range[worKsheeT.Cells[1, 1], worKsheeT.Cells[2, redirectsTable.Columns.Count]];

                worKbooK.SaveAs("Dynamics Page Redirects");
                worKbooK.Close();
                excel.Quit();

            }
            catch (Exception ex)
            {
                
            }
            finally
            {
                worKsheeT = null;
                celLrangE = null;
                worKbooK = null;
            }
        }
    }
}
