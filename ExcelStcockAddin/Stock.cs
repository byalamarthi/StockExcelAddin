using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using System.Drawing;

namespace ExcelStcockAddin
{
    public partial class Stock
    {
        private void Stock_Load(object sender, RibbonUIEventArgs e)
        {
            

        }

        private void BtnGetStock_Click(object sender, RibbonControlEventArgs e)
        {
        
            //FSOWAVY0NMELRMXN     StockAPI Key
        
       
            Worksheet CurrentWorksheet = Globals.ThisAddIn.GetActiveWorksheet();
            //get the active cell value
            Range activeCell = Globals.ThisAddIn.Application.ActiveCell;
            object cellValue = activeCell.Value;

            //get the row and column details
            int row = activeCell.Row;
            int column = activeCell.Column;
            
            string cellValues = "";
            string STRURL = "";
            if (!string.IsNullOrEmpty(activeCell.Value2))
            {
                CurrentWorksheet.Cells[row, column].Value2 = cellValue.ToString().ToUpper();
                cellValues = activeCell.Text;
                //API Call with selected stock ticker
                STRURL = string.Format("https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol=" + cellValues + "&apikey=FSOWAVY0NMELRMXN");
                WebRequest requestObject = WebRequest.Create(STRURL);
                requestObject.Method = "GET";
                HttpWebResponse webResponseObject = null;
                webResponseObject = (HttpWebResponse)requestObject.GetResponse();
                string strresulttest = null;
                using (Stream stream = webResponseObject.GetResponseStream())
                {
                    StreamReader sr = new StreamReader(stream);
                    strresulttest = sr.ReadToEnd();
                    // if the stock ticker is not valid or data is not available
                    if (strresulttest.Contains("Error"))
                    {
                        MessageBox.Show("Invalid Stock Ticker. Please retry");
                    }
                    else if(strresulttest.Contains("Note"))
                    {
                        MessageBox.Show("Only 5 API calls per minute is allowed.");

                    }
                    else
                    {
                        JObject jsonResult = JObject.Parse(strresulttest);
                        string toDate = DateTime.Now.ToString("yyyy-MM-dd");
                        DayOfWeek day = DateTime.Now.DayOfWeek;
                        string dayToday = day.ToString();
                        //check if the provided date is not weekend
                        if ((day == DayOfWeek.Saturday)) 
                        {
                            toDate = DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd");
                        }
                        else if ((day == DayOfWeek.Sunday))
                        {
                            toDate = DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd");
                        }
                        // get the open and close stock values from API json result and assign to string
                        string dailyOpenValue = jsonResult["Time Series (Daily)"][toDate]["1. open"].ToString();
                        string dailyHighValue= jsonResult["Time Series (Daily)"][toDate]["2. high"].ToString();
                        string dailyLowValue= jsonResult["Time Series (Daily)"][toDate]["3. low"].ToString();
                        string dailyCloseValue = jsonResult["Time Series (Daily)"][toDate]["4. close"].ToString();


                        CurrentWorksheet.Cells[row, column + 1].Value2 = "$" + dailyOpenValue;
                        
                        CurrentWorksheet.Cells[row, column + 2].Value2 = "$" + dailyHighValue;
                        CurrentWorksheet.Cells[row, column + 3].Value2 = "$" + dailyLowValue;
                        CurrentWorksheet.Cells[row, column + 4].Value2 = "$" + dailyCloseValue;
                    }



                    sr.Close();
                }
            }
            else
            {
                MessageBox.Show("Please make sure the selected cell contains valid stock ticker and the cell is not in edit mode.");
            }

        }

        private void BtnTemplate_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet CurrentWorksheet = Globals.ThisAddIn.GetActiveWorksheet();


            // Add header for the excel sheet.

            if (!string.IsNullOrEmpty(CurrentWorksheet.Range["A1"].Value2) && CurrentWorksheet.Range["A1"].Value2 != "Stock Ticker")
            {
                Range line = (Range)CurrentWorksheet.Rows[1];
                line.Insert();
                CurrentWorksheet.Range["A1"].Value2 = "Stock Ticker";
            }
            else if(string.IsNullOrEmpty(CurrentWorksheet.Range["A1"].Value2))
            {
                CurrentWorksheet.Range["A1"].Value2 = "Stock Ticker";
            }
            if (string.IsNullOrEmpty(CurrentWorksheet.Range["B1"].Value2))
                CurrentWorksheet.Range["B1"].Value2 = "Today open Value";
            if (string.IsNullOrEmpty(CurrentWorksheet.Range["C1"].Value2))
                CurrentWorksheet.Range["C1"].Value2 = "Today high Value";
            if (string.IsNullOrEmpty(CurrentWorksheet.Range["D1"].Value2))
                CurrentWorksheet.Range["D1"].Value2 = "Today low Value";
            if (string.IsNullOrEmpty(CurrentWorksheet.Range["E1"].Value2))
                CurrentWorksheet.Range["E1"].Value2 = "Today close Value";

            ((Excel.Range)CurrentWorksheet.Range["A1", "E1"]).Interior.Color = ColorTranslator.ToOle(Color.Yellow);
            ((Excel.Range)CurrentWorksheet.Range["A1", "E1"]).Font.Bold = true;
            CurrentWorksheet.Columns.AutoFit();
            //button disabled after load template
            btnTemplate.Enabled = false;

        }
    }
}
