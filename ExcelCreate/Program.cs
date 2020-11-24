using System;

namespace ExcelCreate
{
    class Program
    {
        static void Main(string[] args)
        {
            Test test = new Test();
            test.CreateExcelDoc(@"C:\Excel\test.xlsx");

            Console.WriteLine("Zrobione");
        }
    }
}
