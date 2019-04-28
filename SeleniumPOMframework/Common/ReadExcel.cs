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
    class ReadExcel
    {
        public String getData(int row,int col)
        {
            
            Excel.Application excelApp = new Excel.Application();
            if (excelApp != null)
            {

                Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(@"D:\\test.xls", 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets[1];

                Excel.Range excelRange = excelWorksheet.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;

              //for (int i = 1; i <= rowCount; i++)
               //
                 // for (int j = 1; j <= colCount; j++)
                  //{
                      Excel.Range range = (excelWorksheet.Cells[row,col] as Excel.Range);
                string cellValue = range.Value.ToString();


                // }
                //}
               
                excelWorkbook.Close();
                excelApp.Quit();
                return cellValue;
            }
            return null;
           
        }

    }
}
