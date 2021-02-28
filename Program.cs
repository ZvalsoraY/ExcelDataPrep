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
using NPOI.HSSF.Util;
using System.Text.RegularExpressions;
using System.Globalization;

namespace DataPrep
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"C:\Users\Master\Documents\C_sharp\Work\wall_test.output";
            ReadFile(fileName);
            //FileArray(fileName);
           
            foreach (var item in FileArray(fileName))
            {
                //Console.Write(" ");
                foreach (var s in item)
                Console.Write(s);
                Console.WriteLine();
            }
                ////Console.WriteLine(ReadFile(fileName));
                //CreateExcel();
                Console.Read();
        }

        static Array ReadFile(string fileName)
        {
            //var res = File.ReadLines(fileName).Select(s => s.Split(' ')).ToArray();

            //foreach (string[] s in res)
            //{
            //    foreach (string r in s)
            //    {
            //        System.Console.Write("!! " + r);
            //        //Convert.ToDouble(r);
            //    }
            //    //System.Console.WriteLine("dddd");
            //    System.Console.WriteLine();
            //}


            //return res;
            //Console.WriteLine(File.ReadAllText(fileName));
            //return File.ReadAllText(fileName);
            return File.ReadLines(fileName).Select(s => s.Split(' ')).ToArray();

        }

        static double[][] FileArray(string fileName)
        {
            var lines = File.ReadAllLines(fileName);
            //double[][] resArray = new double[lines.Count()][];
            var resArray = new double[lines.Count()][];
            for (int i = 0; i < resArray.Length; i++)
            {
                string[] stringArray = lines[i].Split(' ','!').ToArray();
                //double[] doubleArray = stringArray.Select<string, double>(s => Double.Parse(s)).ToArray<double>();

                for (int j = 2; j < stringArray.Length; j++)
                {
                    resArray[i] = new double[7];
                    //var a = double.Parse(stringArray[j], CultureInfo.InvariantCulture);
                    resArray[i][j - 2] = double.Parse(stringArray[j], CultureInfo.InvariantCulture);
                    //Console.Write(stringArray[j] + "\t");
                }
            }
            return resArray;
        }
        
        static void CreateExcel()
        {
            //XSSFWorkbook wb1 = null;

            using (var stream = new FileStream(@"D:\test\Result.xlsx", FileMode.Create, FileAccess.ReadWrite))
            {
                //https://stackoverflow.com/questions/47793744/generate-excel-with-merged-header-using-npoi
                //https://stackoverflow.com/questions/32723483/adding-a-specific-autofilter-on-a-column
                //wb1 = new XSSFWorkbook(file);
                var wb = new XSSFWorkbook();
                var sheet = wb.CreateSheet("Test wall");
                //creating cell style for header
                var bStylehead = wb.CreateCellStyle();
                bStylehead.BorderBottom = BorderStyle.Thin;
                bStylehead.BorderLeft = BorderStyle.Thin;
                bStylehead.BorderRight = BorderStyle.Thin;
                bStylehead.BorderTop = BorderStyle.Thin;
                bStylehead.Alignment = HorizontalAlignment.Center;
                bStylehead.VerticalAlignment = VerticalAlignment.Center;
                bStylehead.FillBackgroundColor = HSSFColor.Green.Index;
                //var cellStyle =
                //var cellStyle = CreateCellStyleForHeader(wb);

                
                var row = sheet.CreateRow(0);
                row.CreateCell(0, CellType.String).SetCellValue("x");
                row.CreateCell(1, CellType.String).SetCellValue("y");
                row.CreateCell(2, CellType.String).SetCellValue("z");
                row.CreateCell(3, CellType.String).SetCellValue("Hx");
                row.CreateCell(4, CellType.String).SetCellValue("Hy");
                row.CreateCell(5, CellType.String).SetCellValue("Hz");
                row.CreateCell(6, CellType.String).SetCellValue("Hsum");
                row.Cells[0].CellStyle = bStylehead;

                //filling the data
                var rowsCounter = 1;

                string fileName = @"C:\Users\Master\Documents\C_sharp\Work\wall_test.output";
                var fileData = ReadFile(fileName);

                
                //var row = sheet.CreateRow(0);
                foreach (Array rowData in fileData)
                {
                    var rowD = sheet.CreateRow(rowsCounter++);
                    //rowD.CreateCell(0, CellType.String).SetCellValue(rowData.Length);
                    for (int i = 1; i < rowData.Length; i++)
                    {
                        rowD.CreateCell(i-1, CellType.Numeric).SetCellValue(Double.Parse(rowData.GetValue(i).ToString().Replace(@".", @",")));
                    }
                    
                }
                //sheet.SetAutoFilter(CellRangeAddress.ValueOf("A:C"));
                
                wb.Write(stream);
                //file.Close();
            }


        }
            
    }
}
