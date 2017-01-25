using Excel = Microsoft.Office.Interop.Excel;

namespace SampleExcelOperations
{
    public class BookEditor:Reader
    {
        public BookEditor(): base()
        {

        }

        public void removeName(string name)
        {
            xlWorkbook.Names.Item(name).Delete();
        }

        public void changeRangeName(string range, string rangeName, string sheetName)
        {
            Excel.Range excelRange = this.getRange(sheetName, range);
            xlWorkbook.Names.Add(rangeName, excelRange);
        }
    }
}
