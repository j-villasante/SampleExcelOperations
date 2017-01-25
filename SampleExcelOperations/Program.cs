using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SampleExcelOperations
{
    class Program
    {
        static void Main(string[] args)
        {
            BookEditor editor = new BookEditor();
            editor.open(@"C:\Users\josue\Documents\Visual Studio 2015\Projects\SampleExcelOperations\nameTest.xlsx");
            editor.removeName("test2");
            editor.changeRangeName("A4:B10", "test2", "Hoja1");
            editor.close();
        }
    }
}
