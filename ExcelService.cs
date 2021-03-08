using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel; 

namespace AbbyOCRProcessOutputConversion
{
    public class ExcelService
    {

        public void ConvertExcelToCsv(string excelPath, string csvPath)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            //xlApp.DisplayAlerts = false;

            if (xlApp == null)
            {
                throw new Exception("Excel is not properly installed!!");
            }
            
            Excel.Workbook xlWorkBook = null;
            Excel.Worksheet xlWorkSheet = null;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Open(excelPath);
            xlWorkBook.SaveAs(csvPath, Excel.XlFileFormat.xlCSV, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            Thread.Sleep(2000);
            xlWorkBook.Close(true, misValue, misValue);
            xlWorkBook = null;
            Thread.Sleep(2000);
            xlApp.Quit();
            xlApp = null;


            if (xlWorkSheet != null)
                Marshal.ReleaseComObject(xlWorkSheet);

            if (xlWorkBook != null)
                Marshal.ReleaseComObject(xlWorkBook);

            if (xlApp != null)
                Marshal.ReleaseComObject(xlApp);

        }


    }
}
