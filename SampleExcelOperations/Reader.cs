using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace SampleExcelOperations
{
    public class Reader
    {
        private static List<Excel.Range> ranges;
        private static Excel.Application xlApp;        
        protected static Excel.Workbook xlWorkbook;

        public Dictionary<string, Excel.Worksheet> book;

        public Reader()
        {
            ranges = new List<Excel.Range>();
        }

        public void open(String file)
        {
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(file, ReadOnly: false);
            this.book = this.toDictionary(xlWorkbook);
        }

        public void setWorkbook(Excel.Workbook workbook)
        {
            xlWorkbook = workbook;
            this.book = this.toDictionary(workbook);
        }

        public void close()
        {
            xlWorkbook.Save();
            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            foreach (KeyValuePair<string, Excel.Worksheet> entry in book)
                Marshal.ReleaseComObject(entry.Value);

            foreach (Excel.Range range in ranges)
                Marshal.ReleaseComObject(range);

            //close and release
            xlWorkbook.Close(false);
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }

        private Dictionary<string, Excel.Worksheet> toDictionary(Excel.Workbook workbook)
        {
            Dictionary<string, Excel.Worksheet> dict = new Dictionary<string, Excel.Worksheet>();
            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                dict.Add(worksheet.Name, worksheet);
            }
            return dict;
        }

        public Excel.Range getRange(String sheetName, String range)
        {
            Excel.Worksheet sheet = book[sheetName];
            Excel.Range excelRange = sheet.get_Range(range);
            ranges.Add(excelRange);
            return excelRange;
        }
    }
}
