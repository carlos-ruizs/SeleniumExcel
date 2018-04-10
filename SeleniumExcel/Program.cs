using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using PruebaExcel_EPplus;
using OfficeOpenXml;
using System.IO;


namespace SeleniumExcel
{
    //TODO change this class so that we obtain which scenarios we will be executing before calling our methods and objects
    //TODO Add a for loop to make the program last for however long we want
    //TODO Add a case to iterate through every scenario available
    //TODO change this class to allow us to call every method at our disposition
    class Program
    {
        static void Main(string[] args)
        {
            FileInfo excelFile = new FileInfo(@"E:\WorkbookSelenium.xlsx");

            //checks if the file we have in excelFile exists and if it does, it instantiates the objects for the webdriver and everything else
            if (excelFile.Exists)
            {
                List<int> RunCases = new List<int>();
                IWebDriver driverFF = new FirefoxDriver(@"C:\geckodriver-v0.19.1-win64");
                LibExcel_epp objeto_Excel = new LibExcel_epp();
                Support objeto_Support = new Support("WorkbookSelenium", "Sheet1", driverFF, objeto_Excel);
                objeto_Excel.m_strFileName = "WorkbookSelenium";
                objeto_Excel.m_fileInfo = excelFile;
                objeto_Support.GetExcelElements();

                foreach (string value in objeto_Support.m_plRunElements)
                {
                    Console.WriteLine(value);
                }

                for (int listIndex = 0; listIndex <= objeto_Support.m_plNumberOfResultsToSave.Count - 1; listIndex++)
                {
                    if (objeto_Support.m_plRunElements[listIndex] == " ")
                    {
                        RunCases.Add(0);
                    }
                    else
                    {
                        int element = int.Parse(objeto_Support.m_plRunElements[listIndex]);
                        RunCases.Add(element);
                    }
                }

                Console.WriteLine();

                foreach(int val in RunCases)
                {
                    Console.WriteLine(val);
                }

                Console.ReadKey();

                //This gives me the name of a worksheet in a certain place in the file I want
                using(ExcelPackage excelPackage = new ExcelPackage(excelFile))
                {
                    Console.WriteLine("There are [{0}] worksheets in the file", objeto_Excel.GetWorksheetAmount(excelPackage));
                    Console.WriteLine(objeto_Excel.GetWorksheetName(excelPackage,1));
                }

                Console.ReadKey();

                objeto_Support.SearchGoogle();
            }
            else
            {
                Console.WriteLine("The file in " + excelFile.FullName + " wasn't found"); //Prints the full path of the file that we tried to use
                Console.ReadKey();
            }
        }
    }
}
