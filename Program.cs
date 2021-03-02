﻿using System;
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

            //FileDoubleArrayList(fileName);
            //FileDoubleArray(fileName);
            
            //selectCoordZ(FileDoubleArrayList(fileName));
            //sortByY(selectCoordZ(FileDoubleArrayList(fileName)));

            ////Console.WriteLine(ReadFile(fileName));
            CreateExcel(sortByY(selectCoordZ(FileDoubleArrayList(fileName))));
            Console.Read();
        }

        static void consoleWriteCheck(List<List<double>> writingArray)
        {
            foreach (var item in writingArray)
            {
                foreach (var s in item)
                    Console.Write(" " + s);
                //Console.Write(s);
                Console.WriteLine();
            }
        }
        static Array ReadFile(string fileName)
        {
            return File.ReadLines(fileName).Select(s => s.Split(' ')).ToArray();
        }

        static double[][] FileDoubleArray(string fileName)
        {
            var lines = File.ReadAllLines(fileName);
            var resArray = new double[lines.Count()][];
            for (int i = 0; i < resArray.Length; i++)
            {
                string[] stringArray = lines[i].Split(' ','!').ToArray();
                resArray[i] = new double[7];
                for (int j = 2; j < stringArray.Length; j++)
                {
                    resArray[i][j - 2] = double.Parse(stringArray[j], CultureInfo.InvariantCulture);
                }
            }
            //foreach (var item in resArray)
            //{
            //    //Console.Write(" ");
            //    foreach (var s in item)
            //        Console.Write(s);
            //    Console.WriteLine();
            //}

            return resArray;
        }
        static List<List<double>> FileDoubleArrayList(string fileName)
        {
            var resArray = new List<List<double>>();

            var lines = File.ReadAllLines(fileName);
            for (int i = 0; i < lines.Length; i++)
            {
                var resString = new List<double>();
                string[] stringArray = lines[i].Split(new[] { ' ', '!' }, StringSplitOptions.RemoveEmptyEntries).ToArray();
                foreach (var linPer in stringArray)
                {
                    resString.Add(Math.Round(double.Parse(linPer, CultureInfo.InvariantCulture),2));
                }                
                resArray.Add(resString);
            }
            //consoleWriteCheck(resArray);
            return resArray;
        }
              
        static List<List<double>> selectCoordZ(List<List<double>> sortedArray, double zCoord = 0.0)
        {
            var resArray = new List<List<double>>();
            foreach (var tag in sortedArray)
            {
                if (tag[2] == 0) resArray.Add(tag); 

            }
            //consoleWriteCheck(resArray);
            return resArray;
        }

        static List<List<double>> sortByY(List<List<double>> inputArray)
        {
            var resArray = new List<List<double>>();
            //resArray = inputArray.Sort()
            resArray = inputArray.OrderBy(l => l[1]).ToList();
            consoleWriteCheck(resArray);
            return resArray;
        }
        static void CreateExcel(List<List<double>> inputArray)
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

                foreach (var rowData in inputArray)
                {
                    var rowD = sheet.CreateRow(rowsCounter++);
                    var dCounter = 0;
                    foreach(var d in rowData)
                    {
                        rowD.CreateCell(dCounter++, CellType.Numeric).SetCellValue(Double.Parse(d.ToString().Replace(@".", @",")));
                    }
                }
                
                ////var row = sheet.CreateRow(0);
                //foreach (Array rowData in fileData)
                //{
                //    var rowD = sheet.CreateRow(rowsCounter++);
                //    //rowD.CreateCell(0, CellType.String).SetCellValue(rowData.Length);
                //    for (int i = 1; i < rowData.Length; i++)
                //    {
                //        rowD.CreateCell(i-1, CellType.Numeric).SetCellValue(Double.Parse(rowData.GetValue(i).ToString().Replace(@".", @",")));
                //    }
                    
                //}
                //sheet.SetAutoFilter(CellRangeAddress.ValueOf("A:C"));
                
                wb.Write(stream);
                //file.Close();
            }


        }
            
    }
}
