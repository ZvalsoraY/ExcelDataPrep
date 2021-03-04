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
using NPOI.SS.UserModel.Charts;

namespace DataPrep
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"C:\Users\Master\Documents\C_sharp\Work\wall_test.output";
            double axisCoordinateData = 12.0;

            //ReadFile(fileName);

            //FileDoubleArrayList(fileName);
            //FileDoubleArray(fileName);
            
            //selectCoordZ(FileDoubleArrayList(fileName));
            //sortByY(selectCoordZ(FileDoubleArrayList(fileName)));

            ////Console.WriteLine(ReadFile(fileName));
            CreateExcel(sortByY
                (selectCoordZ
                (FileDoubleArrayList(fileName), axisCoordinateData)), axisCoordinateData);
            //Console.Read();
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
            var lLdoubleArray = new List<List<double>>();
            //lLdoubleArray.Clear();
            var lines = File.ReadAllLines(fileName);
            for (int i = 0; i < lines.Length; i++)
            {
                var resString = new List<double>();

                string[] stringArray = lines[i].Split(new[] { ' ', '!' }, StringSplitOptions.RemoveEmptyEntries).ToArray();
                foreach (var linPer in stringArray)
                {
                    resString.Add(Math.Round(double.Parse(linPer, CultureInfo.InvariantCulture),2));
                }
                lLdoubleArray.Add(resString);
            }
            
            //consoleWriteCheck(resArray);
            return lLdoubleArray;
        }
              
        static List<List<double>> selectCoordZ(List<List<double>> arraySelectZ, double zCoord = 0.0)
        {
            var zArray = new List<List<double>>();
            zArray.Clear();
            foreach (var tag in arraySelectZ)
            {
                //if (tag[2] == 0) resArray.Add(tag);
                if (Math.Abs(tag[2] - zCoord) <= 0.01) zArray.Add(tag);
            }
            //consoleWriteCheck(resArray);
            return zArray;
        }

        static List<List<double>> sortByY(List<List<double>> arraySortY)
        {
            var arrayByY = new List<List<double>>();
            //resArray = inputArray.Sort()
            arrayByY = arraySortY.OrderBy(l => l[1]).ToList();
            //consoleWriteCheck(resArray);
            return arrayByY;
        }
        static void CreateExcel(List<List<double>> inputArray, double axisCoordinateData)
        {
            string parth = $"D:\\test\\Result{axisCoordinateData.ToString()}.xlsx";
            //XSSFWorkbook wb1 = null;
            using (var stream = new FileStream(parth, FileMode.Create, FileAccess.ReadWrite))
            //using (var stream = new FileStream(@"D:\test\Result.xlsx", FileMode.Create, FileAccess.ReadWrite))
            {
            //https://stackoverflow.com/questions/47793744/generate-excel-with-merged-header-using-npoi
            //https://stackoverflow.com/questions/32723483/adding-a-specific-autofilter-on-a-column
            //https://www.leniel.net/2009/10/npoi-with-excel-table-and-dynamic-chart.html
            //https://coderoad.ru/56089507/%D0%9A%D0%B0%D0%BA-%D1%81%D0%BE%D0%B7%D0%B4%D0%B0%D1%82%D1%8C-LineChart-%D0%BA%D0%BE%D1%82%D0%BE%D1%80%D1%8B%D0%B9-%D1%81%D0%BE%D0%B4%D0%B5%D1%80%D0%B6%D0%B8%D1%82-%D0%B4%D0%B2%D0%B0-CategoryAxis-%D0%B8%D1%81%D0%BF%D0%BE%D0%BB%D1%8C%D0%B7%D1%83%D1%8F-apache-POI
            //https://www.csharpcodi.com/vs2/1431/NPOI/src/NPOI.OOXML/XSSF/UserModel/Charts/XSSFLineChartData.cs/
            //https://itnan.ru/post.php?c=1&p=525492
            //https://overcoder.net/q/3405494/%D0%BA%D0%B0%D0%BA-%D1%81%D0%BE%D0%B7%D0%B4%D0%B0%D1%82%D1%8C-%D0%BB%D0%B8%D0%BD%D0%B5%D0%B9%D0%BD%D1%83%D1%8E-%D0%B4%D0%B8%D0%B0%D0%B3%D1%80%D0%B0%D0%BC%D0%BC%D1%83-%D0%B2%D0%BC%D0%B5%D1%81%D1%82%D0%B5-%D1%81-%D0%B4%D0%B0%D0%BD%D0%BD%D1%8B%D0%BC%D0%B8-%D0%B2-%D1%82%D0%B0%D0%B1%D0%BB%D0%B8%D1%86%D0%B5-excel-%D1%81-%D0%BF%D0%BE%D0%BC%D0%BE%D1%89%D1%8C%D1%8E
            //https://gist.github.com/Bykiev/2912494f5a3e4e6f91e02c12a6d6a82d
                //wb1 = new XSSFWorkbook(file);
                var wb = new XSSFWorkbook();
                //var wb = new ;
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



                var Drawing = sheet.CreateDrawingPatriarch();

                //IClientAnchor anchor = Drawing.CreateAnchor(0, 0, 0, 0, 8, 1, 18, 16);
                var anchor = Drawing.CreateAnchor(0, 0, 0, 0, 8, 1, 18, 16);
                var chart = Drawing.CreateChart(anchor);
                //var chart = 
                //IChart chart = Drawing.CreateChart(anchor);
                IChartAxis bottomAxis = chart.ChartAxisFactory.CreateCategoryAxis(AxisPosition.Bottom);
                IChartAxis leftAxis = chart.ChartAxisFactory.CreateValueAxis(AxisPosition.Left);

                var chartData =
                        chart.ChartDataFactory.CreateLineChartData<double, double>();
                var lenCellRange = inputArray.Count + 1;
                IChartDataSource<double> xs = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf($"B2:B{lenCellRange}"));
                IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf($"G2:G{lenCellRange}"));
                //IChartDataSource<double> ys = DataSources.FromNumericCellRange(sheet, CellRangeAddress.ValueOf("G2:G20"));
                var series = chartData.AddSeries(xs, ys);
                series.SetTitle("test");
                //chart.GetOrCreateLegend();
                
                chart.Plot(chartData, bottomAxis, leftAxis);





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
                    foreach (var d in rowData)
                    {
                        rowD.CreateCell(dCounter++, CellType.Numeric).SetCellValue(Double.Parse(d.ToString().Replace(@".", @",")));
                    }
                }


                
                


                wb.Write(stream);
                
                wb.Close();
                //file.Close();
            }


        }
            
    }
}
