using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using ExcelDataReader;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace apptest
{
    class ExcelRead
    {
        //"C:\\Users\\seymour\\Desktop\\lelExcel.xls"
        public static void create_xls(string chemin)
        {
            using (var fs = new FileStream(chemin, FileMode.Create, FileAccess.Write))
            {
                
                IWorkbook workbook = new HSSFWorkbook();
                // new XSSFWorkbook(); pour xslx


                ISheet sheet1 = workbook.CreateSheet("Sheet1");

                //sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 0, 10));
                var rowIndex = 0;
                IRow row = sheet1.CreateRow(rowIndex);
                //row.Height = 30 * 80;
                row.CreateCell(0).SetCellValue("EXCEL");
                sheet1.AutoSizeColumn(0);
                rowIndex++;
                
                workbook.Write(fs);
            }



            


        }

        public static void read(string chemin)
        {
            using (var fs = new FileStream(chemin, FileMode.Open, FileAccess.Read))
            {

                IWorkbook workbook = new HSSFWorkbook(fs);
                // new XSSFWorkbook(); pour xslx


                ISheet sheet1 = workbook.GetSheet("Sheet1");

                
                
                IRow row = sheet1.GetRow(0);
                
                string titi;
                titi = row.GetCell(0).ToString();
                Console.WriteLine(row.GetCell(0).ToString());

                
            }

           
        }


    }
}
