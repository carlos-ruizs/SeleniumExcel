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

                //This for-loop iterates through every element in the worksheet that has a number of results to save
                //TODO change "objeto_Support.m_plNumberOfResultsToSave.Count - 1" to something that better reflects how many actions we'll be checking 
                for (int listIndex = 0; listIndex <= objeto_Support.m_plNumberOfResultsToSave.Count - 1; listIndex++)
                {
                    //This if-else checks if there's a 0 or a blank space to avoid executing that element
                    if (objeto_Support.m_plRunElements[listIndex] == " ")
                    {
                        continue;
                    }
                    else
                    {
                        if (objeto_Support.m_plRunElements[listIndex] == "0")
                        {
                            continue;
                        }
                        //int element = int.Parse(objeto_Support.m_plRunElements[listIndex]);
                        //RunCases.Add(element);
                    }

                    //converts every element in the Actions column of the worksheet, so we can later check if an action in the column is valid or not
                    objeto_Support.m_plActions = objeto_Support.m_plActions.ConvertAll(d => d.ToUpper());

                    //This switch checks which actions are to be executed for the elements in the worksheet
                    switch (objeto_Support.m_plActions[listIndex])
                    {
                        case "SEARCH":
                            objeto_Support.SearchGoogle(listIndex);

                            break;
                        case "CREATE":
                            Console.WriteLine("You already created a .xlsx file");
                            break;

                        default:
                            Console.WriteLine("The case " + objeto_Support.m_plActions[listIndex] + " doesn't exist");
                            break;
                    }
                }

                objeto_Support.m_iwbWebDriver.Close();
            }
            else
            {
                Console.WriteLine("The file in " + excelFile.FullName + " wasn't found"); //Prints the full path of the file that we tried to use
                Console.ReadKey();
            }
        }
    }
}
