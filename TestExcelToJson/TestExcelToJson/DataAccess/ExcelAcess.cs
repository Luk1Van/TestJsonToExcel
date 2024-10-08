using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestExcelToJson.DataAccess
{
    public  class ExcelAcess
    {
        public Excel.Application xlApp { get; set; }
        public Excel.Workbook xlWorkbook { get; set; }
        public Excel._Worksheet xlWorksheet { get; set; }
        public Excel.Range xlRange { get; set; }


        public  Excel.Range openExcel(string ExcelPath, int? SheetName = null)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(ExcelPath);
            xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[1];
                if (SheetName != null )
                {
                     xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets[SheetName];
                }
            xlRange = xlWorksheet.UsedRange;
            return xlRange;
        }

        public void closeExcel()
        {
            Marshal.ReleaseComObject(xlWorksheet);
            Marshal.ReleaseComObject(xlRange);
            xlWorkbook.Close();
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
