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
            Console.WriteLine("Elija desde dónde quiere correr el programa:");
            Console.WriteLine("1.- DB, 0.- Excel");
            int key = int.Parse(Console.ReadLine());
            if(key == 1)
            {
                SupportSql supSql = new SupportSql();
                supSql.DataFill();
                Console.ReadKey();
            }
            else
            {
                if (key == 0)
                {
                    ProgramSE.Main();
                }
                else
                {
                    Console.WriteLine("Opción inexistente");
                    Console.ReadLine();
                }
            }
        }
    }
}
