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
    class Program
    {
        static void Main(string[] args)
        {
            IWebDriver driverFF = new FirefoxDriver(@"C:\geckodriver-v0.19.1-win64");
            libExcel_epp objeto_Excel = new libExcel_epp();
            Support objeto_Support = new Support("WorkbookSelenium", "Sheet1", driverFF, objeto_Excel);
            objeto_Support.SearchGoogle();
        }
    }
}
