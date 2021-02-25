using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Web;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Model;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Model;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace DataPrep
{
    class Program
    {
        static void Main(string[] args)
        {
            //string fileName = @"C:\Users\Master\Documents\C_sharp\Work\wall_test.output";
            //Console.WriteLine(ReadFile(fileName));
            CreateExcel();
            Console.Read();
        }

        static string ReadFile(string fileName)
        {
            return File.ReadAllText(fileName);
            //Console.WriteLine(File.ReadAllText(fileName));

        }

        static void CreateExcel()
        {
            //XSSFWorkbook wb1 = null;

            using (var stream = new FileStream(@"D:\test\Result.xlsx", FileMode.Create, FileAccess.ReadWrite))
            {
            //https://stackoverflow.com/questions/47793744/generate-excel-with-merged-header-using-npoi
                //wb1 = new XSSFWorkbook(file);
                var wb = new XSSFWorkbook();
                var sheet = wb.CreateSheet("Test wall");
                //creating cell style for header
                //var cellStyle =
                //var cellStyle = CreateCellStyleForHeader(wb);

                //filling the header
                var row = sheet.CreateRow(0);
                row.CreateCell(0, CellType.String).SetCellValue("x");
                row.CreateCell(1, CellType.String).SetCellValue("y");
                row.CreateCell(2, CellType.String).SetCellValue("z");
                row.CreateCell(3, CellType.String).SetCellValue("Hx");
                row.CreateCell(4, CellType.String).SetCellValue("Hy");
                row.CreateCell(5, CellType.String).SetCellValue("Hz");
                row.CreateCell(6, CellType.String).SetCellValue("Hsum");


                //wb1.GetSheetAt(0).GetRow(0).GetCell(0).SetCellValue("Sample");
                //file.
                wb.Write(stream);
                //file.Close();
            }


            //using (var file2 = new FileStream(@"C:\Users\Master\Documents\C_sharp\Work\Result.xlsx", FileMode.Create, FileAccess.ReadWrite))
            //{
            //    wb1.Write(file2);
            //    file2.Close();
            //}

        }
            
    }
}
