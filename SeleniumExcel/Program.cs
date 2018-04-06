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


namespace SeleniumExcel
{
    //TODO change this class so that we obtain which scenarios we will be executing before calling our methods and objects
    //TODO Add a way to check if the Excel file doesn't exist from this class and show an error if the path we need isn't there
    //TODO Add a for loop to make the program last for however long we want
    //TODO Add a case to iterate through every scenario available
    //TODO change this class to allow us to call every method at our disposition
    class Program
    {
        static void Main(string[] args)
        {
            IWebDriver driverFF = new FirefoxDriver(@"C:\geckodriver-v0.19.1-win64");
            LibExcel_epp objeto_Excel = new LibExcel_epp();
            Support objeto_Support = new Support("WorkbookSelenium", "Sheet1", driverFF, objeto_Excel);
            objeto_Support.SearchGoogle();

            //re
        }
    }
}
