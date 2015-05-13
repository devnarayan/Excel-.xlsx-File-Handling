using Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelWork
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Please enter 1 for read excel file and 0 for exit program.");
            bool bl = false;
           
            do
            {
                Console.WriteLine("Please enter option....");
                string st = Console.ReadLine();
                if (st == "1")
                    foreach (var worksheet in Workbook.Worksheets(@"E:\Projects\Excel Work\Excelwork.xlsx"))
                    {
                        foreach (var row in worksheet.Rows)
                        {
                            foreach (var cell in row.Cells)
                            {
                                Console.WriteLine(cell.Text);
                            }
                        }
                    }
                else if (st == "0")
                {
                    Console.WriteLine("Program terminated. Press any key to exit.");
                    bl = true;
                }
            }
            while (bl==false);
            Console.ReadKey();
        }
    }
}