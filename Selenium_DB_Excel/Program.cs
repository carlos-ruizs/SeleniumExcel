using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using SeleniumExcel;

namespace Selenium_DB_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Input the number of the data source for the program:");
                Console.WriteLine("1.- DB, 0.- Excel");
                int key = int.Parse(Console.ReadLine());
                if (key == 1)
                {
                    SupportSql supSql = new SupportSql();
                    supSql.DataFill();
                }
                else
                {
                    if (key == 0)
                    {
                        ProgramSE.Main();
                    }
                    else
                    {
                        Console.WriteLine("Invalid option");
                        Console.ReadLine();
                    }
                }
            }
            catch (FormatException fe)
            {
                Console.WriteLine(fe.Message);
                Console.WriteLine("\nPress the enter key to close");
                Console.ReadKey();
            }
        }
    }
}
