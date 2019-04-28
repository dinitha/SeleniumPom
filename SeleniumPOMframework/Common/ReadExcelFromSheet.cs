using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumPOMframework.Common
{
    class ReadExcelFromSheet
    {
        public String getData(int row, int col, String var)
        {

            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {

                List<string> prop = new List<string>(
    var.Split(new string[] { "." }, StringSplitOptions.None));
                // open excel sheet
              //  string startupPath = System.AppDomain.CurrentDomain.BaseDirectory;
                String path = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\\..\\"));
                String dataDirectory = path+ "Data\\test.xlsx";

                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(dataDirectory, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                // open work sheet
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[prop[0]];
                excelWorksheet.Select(Type.Missing);
                 
                //method1- get value of given row and column
                Excel.Range range = (excelWorksheet.Cells[row, col] as Excel.Range);
                string cellValue = range.Value.ToString();

                // method 2-get value by passing key value
                String valueForKey=  getValueForSearchKey(excelWorksheet, prop[1]);



                excelWorkbook.Close();
                excelApp.Quit();
                return valueForKey;
            }
            return null;

        }

        public String getValueForSearchKey(Excel.Worksheet ws,String key) {


            int keyColum = 1;
            Excel.Range range1 = ws.Range["A:A"];
            Excel.Range range2 = range1.Find(key);
            string msg1 = range2.Count + " records found.";
            int row = range2.Row;
string msg2 = "1st in row " + range2.Row;
string msg3 = ws.Cells[row , keyColum + 1].value;
            return msg3;
        }

   }    
}
