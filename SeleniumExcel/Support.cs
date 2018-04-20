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
    class Support
    {
        //atributes
        public string m_strWorkbookName;
        public string m_strWorksheetName;
        public IWebDriver m_iwbWebDriver;
        public LibExcel_epp m_leeExcelObject;
        public List<string> m_plHeaderNames; //list of headers in the Excel file
        public List<string> m_plSearchTerms; //list of search strings in the Excel file
        public List<string> m_plNumberOfResultsToSave; //integers that we use to know how many results we will save inside the worksheet
        public List<string> m_plRunElements; //column that tells us if a search string is to be executed or not
        public List<string> m_plWorksheetNames; //names of all the worksheets in the file we're working with
        public FileInfo m_fiFilePath;
        public List<string> m_plActions; //Column that tells us what method is going to process
        public int m_listIndex;

        //Properties
        public string ProprWorkbookName {get => m_strWorkbookName; set => m_strWorkbookName = value;}
        public string ProprWorksheetName { get => m_strWorksheetName; set => m_strWorksheetName = value; }
        public IWebDriver ProprDriver { get => m_iwbWebDriver; set => m_iwbWebDriver = value; }
        public LibExcel_epp PropExcelObject { get => m_leeExcelObject; set => m_leeExcelObject = value; }
        public FileInfo ProprFilePath { get => m_fiFilePath; set => m_fiFilePath = value; }


        //constructors
        public Support()
        {

        }

        public Support(string pstrWorkbookName, string pstrWorksheetName, IWebDriver piwbDriver, LibExcel_epp pleeExcelObject)
        {
            m_iwbWebDriver = piwbDriver;
            m_leeExcelObject = pleeExcelObject;
            m_strWorkbookName = pstrWorkbookName;
            m_strWorksheetName = pstrWorksheetName;
            m_plHeaderNames = new List<string>();
            m_plSearchTerms = new List<string>();
            m_plNumberOfResultsToSave = new List<string>();
            m_plRunElements = new List<string>();
            m_plActions = new List<string>();
            m_fiFilePath = new FileInfo(@"E:\" + m_strWorkbookName + ".xlsx");
            m_plWorksheetNames = new List<string>();
            for (int worksheetsNumber = 1; worksheetsNumber <= m_leeExcelObject.GetWorksheetAmount(m_fiFilePath); worksheetsNumber++) //this has to start at 1, otherwise it will give an exception
            {
                m_plWorksheetNames.Add(m_leeExcelObject.GetWorksheetNameFI(m_fiFilePath,worksheetsNumber));
            }
        }

        public Support(FileInfo pfiExcelPath)
        {
            m_fiFilePath = pfiExcelPath;
            m_plHeaderNames = new List<string>();
            m_plSearchTerms = new List<string>();
            m_plNumberOfResultsToSave = new List<string>();
            m_plRunElements = new List<string>();
            m_leeExcelObject = new LibExcel_epp();
            m_strWorkbookName = m_leeExcelObject.GetWorkbookName(m_fiFilePath);
            m_strWorksheetName = m_leeExcelObject.FirstWorksheetName(m_fiFilePath);
            m_plWorksheetNames = new List<string>();
            for (int worksheetsNumber = 1; worksheetsNumber <= m_leeExcelObject.GetWorksheetAmount(m_fiFilePath); worksheetsNumber++) //this has to start at 1, otherwise it will give an exception
            {
                m_plWorksheetNames.Add(m_leeExcelObject.GetWorksheetNameFI(m_fiFilePath, worksheetsNumber));
            }
        }

        public void SearchGoogle(IWebDriver piwbDriver, LibExcel_epp pleeExcelObject, string pstrWorkbookName, string pstrWorksheetName)
        {
            piwbDriver.Navigate().GoToUrl("http://www.google.com/");

            List<string> headerNames = new List<string>();
            List<string> resultsToSave = new List<string>();
            List<string> searchTerms = new List<string>();

            //GetExcelElements(pleeExcelObject,pstrWorkbookName,pstrWorksheetName,headerNames,searchTerms,resultsToSave);
            //Results(piwbDriver,pleeExcelObject,pstrWorkbookName,pstrWorksheetName,headerNames,searchTerms,resultsToSave,searchString);
            
            piwbDriver.Close();
        }

        public void SearchGoogle(int Index)
        {
            Results(m_plHeaderNames, m_plSearchTerms, m_plNumberOfResultsToSave, m_plRunElements, Index);

            //m_iwbWebDriver.Close();
        }
        //TODO Change this too so that it complements what the outer for-loop in Program.cs when it comes to elements that have a 1 or 0 in it's Run column
        /// <summary>
        /// Enters a for loop that iterates through all the search strings we want to use, 
        /// as well as the number of saved results we want in the worksheet
        /// </summary>
        /// <param name="piwbDriver"></param>
        /// <param name="pleeExcelObject"></param>
        /// <param name="pstrWorkbookName"></param>
        /// <param name="pstrWorksheetName"></param>
        /// <param name="plHeaderNames"></param>
        /// <param name="plSearchStrings"></param>
        /// <param name="plResultNumbers"></param>
        /// <param name="pstrSearchString"></param>
        public void Results(List<string> plHeaderNames, List<string> plSearchStrings, List<string> plResultNumbers, List<string> plRunElements, int listIndex)
        {
            IfNIteration(listIndex, plRunElements, plHeaderNames);
            
            int elementsToSave = int.Parse(plResultNumbers[listIndex]); //converts the strings inside the resultsToSave list into integers we will use to determine how many results we will save for that particular search
            m_iwbWebDriver.FindElement(By.Id("lst-ib")).SendKeys(plSearchStrings[listIndex]); //finds the search bar and sends the string we want to search into it
            
            /*
            Checks if it's the first time it's searching on Google 
            If true, it looks for the btnK button and clicks it
            If false, it looks for the btnG button and clicks it
            The button changes names depending where you are. 
            */
            if (m_iwbWebDriver.Url == "https://www.google.com/?gws_rd=ssl" || m_iwbWebDriver.Url == "https://www.google.com/")
            {
                m_iwbWebDriver.FindElement(By.Name("btnK")).Click();
            }
            else
            {
                m_iwbWebDriver.FindElement(By.Name("btnG")).Click();
            }

            IList<IWebElement> h3Links = m_iwbWebDriver.FindElements(By.ClassName("g")); //saves all the links inside the webpage from the "g" class into an IList
            string totalSearchResults = m_iwbWebDriver.FindElement(By.Id("resultStats")).Text; //gets the total amount of results for that particular search
            IList<IWebElement> relatedResults = m_iwbWebDriver.FindElements(By.ClassName("nVcaUb")); //saves the links for all the related searches results into an IList


            //Sends the data we want to save into the worksheet for the corresponding column 
            //Also gets the lists from before and builds a big string with all the results that were saved in both their hyperlink and text forms
            m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, listIndex + 2, GetColumnIndex(plHeaderNames, "Saved Results Links"), GetResultsHref(h3Links, elementsToSave - 1));
            m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, listIndex + 2, GetColumnIndex(plHeaderNames, "Saved Results Text"), GetResultsTxt(h3Links, elementsToSave - 1));//debe tomar el nombre de la columna donde lo va a poner y el de la fila igual (el término de búsqueda)
            m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, listIndex + 2, GetColumnIndex(plHeaderNames, "Related Results Links"), GetResultsHref(relatedResults, relatedResults.Count - 1));
            m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, listIndex + 2, GetColumnIndex(plHeaderNames, "Related Results Text"), GetResultsTxt(relatedResults, relatedResults.Count - 1));
            m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, listIndex + 2, GetColumnIndex(plHeaderNames, "Total number of results"), GetTotalSearchResults(totalSearchResults));

            m_iwbWebDriver.FindElement(By.Id("lst-ib")).Clear(); //clears the search field when we finish with a search
        }

        /// <summary>
        /// This checks if it's the first iteration for the Search action and if the Run column associated to it has a 1, 0 or null value
        /// If it's a 1, then it simply goes to google.com
        /// If it's in any other iteration, it checks if the value in the Run column for the string before it was 0 or null so it knows
        /// it's going to google.com for the "first" time since it is the first string it will search.
        /// Right now it works only if the Search actions are the first things inside the worksheet
        /// </summary>
        /// <param name="pintIterationNumber"></param>
        /// <param name="plRunElements"></param>
        /// <param name="plHeaderNames"></param>
        //TODO Change this in the future so that it knows it's the first iteration of the "Search" action no matter where in the worksheet it's found
        private void IfNIteration(int pintIterationNumber, List<string> plRunElements, List<string> plHeaderNames)
        {
            if (pintIterationNumber == 0 && plRunElements[pintIterationNumber] == "1")
            {
                m_iwbWebDriver.Navigate().GoToUrl("http://www.google.com/");
            }
            else
            {
                if (pintIterationNumber > 0 && (plRunElements[pintIterationNumber - 1] == " " || plRunElements[pintIterationNumber - 1] == "0"))
                {
                    //Quiero ver cuál fue el último elemento que tuvo un uno para ubicar su índice
                    for (int i = pintIterationNumber - 1; i >= 0; i--)
                    {
                        if (plRunElements[i] == "1") //If there was at least one element before the current one that executed
                        {
                            break;
                        }
                        else
                        {
                            m_iwbWebDriver.Navigate().GoToUrl("http://www.google.com/");
                            break;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Fills some lists we created so we can use them to know how many search results we will
        /// save, what strings we will be searching for and the amount of searches in general we
        /// wil be doing
        /// </summary>
        /// <param name="pleeExcelObject"></param>
        /// <param name="pstrWorkbookName"></param>
        /// <param name="pstrWorksheetName"></param>
        /// <param name="plHeaderNames"></param>
        /// <param name="plSearchStrings"></param>
        /// <param name="plResultNumbers"></param>
        public void GetExcelElements(List<string> plHeaderNames, List<string> plSearchStrings, List<string> plResultNumbers, List<string> plRunElements, List<string> plActions)
        {
            try
            {
                FileStream stream = new FileStream(@"E:\" + m_strWorkbookName + ".xlsx", FileMode.Open); //creates a file stream to the file we want to manipulate
                ExcelPackage objExcel = new ExcelPackage();
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                objExcel.Load(stream);
                ExcelWorksheet worksheet = objExcel.Workbook.Worksheets[m_strWorksheetName];

                m_leeExcelObject.GetWorksheetHeader(worksheet, plHeaderNames);
                m_leeExcelObject.IterateByColumn(worksheet, GetColumnIndex(plHeaderNames, "Input Parameter"), plSearchStrings);
                m_leeExcelObject.IterateByColumn(worksheet, GetColumnIndex(plHeaderNames, "Number of results to save"), plResultNumbers);
                m_leeExcelObject.IterateByColumn(worksheet, GetColumnIndex(plHeaderNames, "Run"), plRunElements);
                m_leeExcelObject.IterateByColumn(worksheet, GetColumnIndex(plHeaderNames, "Actions"), plActions);

                stream.Close();
                stream.Dispose();
                objExcel.Save();
                objExcel.Dispose();
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception: ", e);
            }
        }

        public void GetExcelElements()
        {
            try
            {
                using (ExcelPackage objExcel = new ExcelPackage(m_fiFilePath))
                {
                    ExcelWorksheet worksheet = objExcel.Workbook.Worksheets[m_strWorksheetName];

                    m_leeExcelObject.GetWorksheetHeader(worksheet, m_plHeaderNames);
                    m_leeExcelObject.IterateByColumn(worksheet, GetColumnIndex(m_plHeaderNames, "Input Parameter"), m_plSearchTerms);
                    m_leeExcelObject.IterateByColumn(worksheet, GetColumnIndex(m_plHeaderNames, "Number of results to save"), m_plNumberOfResultsToSave);
                    m_leeExcelObject.IterateByColumn(worksheet, GetColumnIndex(m_plHeaderNames, "Run"), m_plRunElements);
                    m_leeExcelObject.IterateByColumn(worksheet, GetColumnIndex(m_plHeaderNames, "Actions"), m_plActions);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("{0} Exception: ", e);
            }
        }

        public int GetColumnIndex(List<string> plList, string pstrColumnName)
        {
            int colIndex = 0;

            try
            {
                if (plList.Contains(pstrColumnName))
                {
                    colIndex = plList.IndexOf(pstrColumnName);
                }
            }
            catch (Exception e)
            {

                throw e;
            }

            return colIndex + 1;
        }

        public int GetRowIndex(List<string> plList, string pstrRowName)
        {
            int rowIndex = 0;

            try
            {
                if (plList.Contains(pstrRowName))
                {
                    rowIndex = plList.IndexOf(pstrRowName);
                }
            }
            catch (Exception e)
            {

                throw e;
            }

            return rowIndex + 2;
        }

        /// <summary>
        /// Gets the text atribute for the search results in a google search and returns them as a 
        /// big string
        /// </summary>
        /// <param name="pilResultsList"></param>
        /// <param name="pintLimit"></param>
        /// <returns></returns>
        public string GetResultsTxt(IList<IWebElement> pilResultsList, int pintLimit)
        {
            string ResultString = null;

            for (int index = 0; index <= pintLimit; index++)
            {
                string myURL = pilResultsList[index].FindElement(By.TagName("a")).GetAttribute("text");//gets the text (name) of the link in a certain place in our list
                ResultString = ResultString + myURL + ", ";
            }

            return ResultString;
        }

        /// <summary>
        /// Gets the href atribute for the search results in a google search and returns them
        /// as a big string
        /// </summary>
        /// <param name="pilResultsList"></param>
        /// <param name="pintLimit"></param>
        /// <returns></returns>
        public string GetResultsHref(IList<IWebElement> pilResultsList, int pintLimit)
        {
            string ResultString = null;

            for (int index = 0; index <= pintLimit; index++)
            {
                string myURL = pilResultsList[index].FindElement(By.TagName("a")).GetAttribute("href");//gets the text (name) of the link in a certain place in our list
                ResultString = ResultString + myURL + ", ";
            }

            return ResultString;
        }

        /// <summary>
        /// Gets the total search results string and splits it so that it returns the 
        /// number in the middle as a string
        /// </summary>
        /// <param name="pstrTotalSearchResults"></param>
        /// <returns></returns>
        public string GetTotalSearchResults(string pstrTotalSearchResults)
        {
            string sub = pstrTotalSearchResults.Substring(9);
            string sub2 = sub.Substring(0, sub.Length - 28);

            return sub2;
        }
    }
}
