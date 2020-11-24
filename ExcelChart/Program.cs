using System;

namespace ExcelChart
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelChartTest test = new ExcelChartTest();
            test.CreateExcelDoc(@"C:\Excel\test.xlsx");

            Console.WriteLine("Zrobione");
        }
    }
}
