using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace apptest
{
    class Program
    {
        static void Main(string[] args)
        {

            //Installer https://github.com/xceedsoftware/DocX (peut via nugget)
            // Installer NPOI sur nugget
            string chemin = "C:\\Users\\seymour\\Desktop\\testo.xls";
            ExcelRead.create_xls(chemin);
            ExcelRead.read(chemin);


            chemin = "C:\\Users\\seymour\\Desktop\\testi.docx";
            WordRead.CreateNewWordDocument(chemin);

            WordRead.ReadDocument(chemin);
           


        }
    }
}
