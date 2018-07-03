using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using PruebaExcel_EPplus;
using OfficeOpenXml;
using System.IO;
using NUnit.Framework;

namespace SeleniumExcel
{
    public class Support
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
        //public List<string> m_plSubmenuList; //Column that we use to know how many elements need to compare on submenu
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
            //m_plSubmenuList = new List<string>();
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
                if (pintIterationNumber > 0 && (plRunElements[pintIterationNumber - 1] == " " || plRunElements[pintIterationNumber - 1] != "1"))
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
        /// Function that gets parameters like username and password in order to validate what kind of results
        /// do we have by login in.
        /// </summary>
        /// <param name="pstrURL"></param>
        /// <param name="plHeaderNames"></param>
        public void Login(string pstrURL, int pintRowIndex)
        {
            string ResultsLogin, Validations = null;

            m_iwbWebDriver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "t");
            m_iwbWebDriver.Navigate().GoToUrl("http://opensource.demo.orangehrmlive.com");

            using (ExcelPackage excel = new ExcelPackage(m_fiFilePath))
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];

                string user = m_leeExcelObject.FindElement(m_strWorkbookName, worksheet.Name, pintRowIndex + 2, "Username");
                string pass = m_leeExcelObject.FindElement(m_strWorkbookName, worksheet.Name, pintRowIndex + 2, "Password");

                m_iwbWebDriver.FindElement(By.Id("txtUsername")).SendKeys(user);
                m_iwbWebDriver.FindElement(By.Id("txtPassword")).SendKeys(pass);
                m_iwbWebDriver.FindElement(By.Id("btnLogin")).Click();

                if (m_iwbWebDriver.Url == "http://opensource.demo.orangehrmlive.com/index.php/dashboard")
                {
                    ResultsLogin = "Succesful Login";

                    switch (m_leeExcelObject.FindElement(m_strWorkbookName, worksheet.Name, pintRowIndex + 2, "Test Case"))
                    {
                        case "1":
                        case "2":
                        case "3":
                            break;

                        case "4":
                            IList<IWebElement> bMenus = m_iwbWebDriver.FindElements(By.ClassName("firstLevelMenu"));
                            
                            for (int i = 0; i < bMenus.Count; i++)
                            {
                                string validatemenus = bMenus[i].FindElement(By.TagName("b")).Text;
                                Validations = Validations + validatemenus + " Exists" + ", ";
                                Console.WriteLine(Validations);
                            }
                            m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, pintRowIndex + 2, GetColumnIndex(m_plHeaderNames, "Validate Login"), Validations);

                            break;
                        case "5":
                            //Test case 5
                            //Validate the buttons, elements, and graphics in the screen
                            m_iwbWebDriver.FindElement(By.Id("menu_dashboard_index")).Click(); //Click the dashboard menu
                            var wait = new WebDriverWait(m_iwbWebDriver, TimeSpan.FromSeconds(10));
                            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@class='flot-base']")));
                            string resultString = null;
                            string graphDisplay = null;
                            string labels = null;
                            string percents = null;

                            //Validation of quick launch buttons, colors and labels for the legends of the graphic, the elements inside the graphic and the graphic itself
                            IList<IWebElement> quickLaunchButtons = m_iwbWebDriver.FindElements(By.XPath("//*[@class='quickLinkText']"));
                            IList<IWebElement> graphLegend = m_iwbWebDriver.FindElements(By.XPath("//*[@class='legendLabel']"));
                            IList<IWebElement> pieLabel = m_iwbWebDriver.FindElements(By.XPath("//*[@class='pieLabel']"));
                            IList<IWebElement> graphColor = m_iwbWebDriver.FindElements(By.XPath("//*[@class='legendColorBox']"));
                            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.XPath("//*[@class='pieLabel']")));

                            IWebElement graphic = m_iwbWebDriver.FindElement(By.XPath("//*[@class='flot-base']"));
                            if (graphic.Displayed)
                            {
                                graphDisplay = "The graphic is displayed";
                            }

                            if (graphColor.Count == graphLegend.Count)
                            {
                                graphDisplay = graphDisplay + " All the colors are displayed for their respective legends ";
                            }

                            resultString = GetWebElements(quickLaunchButtons,quickLaunchButtons.Count);
                            labels = GetWebElements(graphLegend, graphLegend.Count);
                            percents = GetWebElements(pieLabel, pieLabel.Count);
                            Console.WriteLine(resultString);
                            Console.WriteLine(labels);
                            Console.WriteLine(percents);
                            m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, pintRowIndex + 2, GetColumnIndex(m_plHeaderNames, "Validate Login"), resultString + labels + percents + graphDisplay);

                            break;

                        case "6":
                            //test case 6
                            //Goes to the Assign Leave submenu of the Leave menu and validates the elements required for the textboxes
                            int r = 0;//variable that keeps track of the labels inside the page
                            string ResultsLabels = null;

                            m_iwbWebDriver.Navigate().GoToUrl("http://opensource.demo.orangehrmlive.com/index.php/leave/assignLeave");
                            IList<IWebElement> labels1 = m_iwbWebDriver.FindElements(By.XPath("//*[@id='frmLeaveApply']/fieldset/ol/li/label"));
                            m_iwbWebDriver.FindElement(By.Id("assignBtn")).Click();
                            IList<IWebElement> validationslabels = m_iwbWebDriver.FindElements(By.XPath("//*[@id='frmLeaveApply']/fieldset/ol/li/span"));

                            //This for loop iterates through the elements that were saved inside our lists and gets their inner text
                            for (int i = 0; i < labels1.Count; i++)
                            {
                                if (i <= 4 || i == 9)
                                {
                                    Console.WriteLine("Label " + i + " " + labels1[i].Text);
                                    string labelsToCheck = labels1[i].Text;

                                    if (i <= 1 || i == 3 || i == 4)
                                    {
                                        Console.WriteLine("Validation " + r + " " + validationslabels[r].Text);
                                        string validation = validationslabels[r].Text;
                                        ResultsLabels += labelsToCheck + " " + validation + " " + ",";
                                        r++;
                                    }
                                }

                            }
                            
                            m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, pintRowIndex + 2, GetColumnIndex(m_plHeaderNames, "Validate Login"), ResultsLabels);
                            break;

                        case "7":
                            m_iwbWebDriver.Navigate().GoToUrl("http://opensource.demo.orangehrmlive.com/index.php/auth/logout");
                            if (m_iwbWebDriver.Url == "http://opensource.demo.orangehrmlive.com/index.php/auth/login")
                            {
                                m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, pintRowIndex + 2, GetColumnIndex(m_plHeaderNames, "Validate Login"), "The logout action was successful");
                            }
                            else
                            {
                                m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, pintRowIndex + 2, GetColumnIndex(m_plHeaderNames, "Validate Login"), "Logout action unsuccessful");
                            }

                            break;
                            
                        default:
                            Console.WriteLine("The test case doesn't exist");
                            break;
                    }
                    m_iwbWebDriver.Navigate().GoToUrl("http://opensource.demo.orangehrmlive.com/index.php/auth/logout");//Logout after each test case
                }
                else
                {
                    ResultsLogin = m_iwbWebDriver.FindElement(By.Id("spanMessage")).Text;
                }

                m_leeExcelObject.Excel_Mod_SingleWFI(m_strWorkbookName, m_strWorksheetName, pintRowIndex + 2, GetColumnIndex(m_plHeaderNames, "Results Login"), ResultsLogin);

            }

        }

        /// <summary>
        /// Gets the current row and checks what menu inside the web page you want to check
        /// then it goes to that menu and calls CompareItems to check if they exist inside the webpage
        /// </summary>
        /// <param name="pintRowIndex"></param>
        public void Hierarchy(int pintRowIndex) //RowIndex is the number of the current line we are executing
        {
            m_iwbWebDriver.FindElement(By.CssSelector("body")).SendKeys(Keys.Control + "t");
            m_iwbWebDriver.Navigate().GoToUrl("http://opensource.demo.orangehrmlive.com");
            using (ExcelPackage excel = new ExcelPackage(m_fiFilePath))
            {
                ExcelWorksheet worksheet = excel.Workbook.Worksheets["Sheet1"];
                ExcelWorksheet worksheet2 = excel.Workbook.Worksheets["Submenus"];

                string user = m_leeExcelObject.FindElement(m_strWorkbookName, worksheet.Name, pintRowIndex + 2, "Username");
                string pass = m_leeExcelObject.FindElement(m_strWorkbookName, worksheet.Name, pintRowIndex + 2, "Password");

                m_iwbWebDriver.FindElement(By.Id("txtUsername")).SendKeys(user);
                m_iwbWebDriver.FindElement(By.Id("txtPassword")).SendKeys(pass);
                m_iwbWebDriver.FindElement(By.Id("btnLogin")).Click();

                string firstLevelMenu = m_leeExcelObject.FindElement("WorkbookSelenium", worksheet.Name, pintRowIndex + 2, "Menu"); //The name of the menu inside the "Menu" column of the spreadsheet
                Console.WriteLine(firstLevelMenu);

                //This case gets the name of the menu we want to check and converts it to lowercase, if it's correct
                //it goes to the corresponding case and calls the CompareItems method with the respective Id of the menu we want to search
                switch (firstLevelMenu.ToLower())
                {
                    case "admin":
                        CompareItems(worksheet2, firstLevelMenu, "menu_admin_viewAdminModule");
                        break;

                    case "pim":
                        CompareItems(worksheet2, firstLevelMenu, "menu_pim_viewPimModule");
                        break;

                    case "leave":
                        CompareItems(worksheet2, firstLevelMenu, "menu_leave_viewLeaveModule");
                        break;

                    case "time":
                        CompareItems(worksheet2, firstLevelMenu, "menu_time_viewTimeModule");
                        break;

                    case "recruitment":
                        CompareItems(worksheet2, firstLevelMenu, "menu_recruitment_viewRecruitmentModule");
                        break;

                    case "performance":
                        CompareItems(worksheet2, firstLevelMenu, "menu__Performance");
                        break;
                            
                    case "dashboard":
                    case "directory":
                        Console.WriteLine("This menu doesn't contain any submenus");
                        break;

                    default:
                        Console.WriteLine("The menu doesn't exist");
                        break;
                }
                m_iwbWebDriver.Navigate().GoToUrl("http://opensource.demo.orangehrmlive.com/index.php/auth/logout");
            }
        }

        /// <summary>
        /// Gets the worksheet where all the menus and their respective submenus are, checks each of the submenus
        /// inside a specific menu and compares them with the webpage to see if they exist or so you can correctly
        /// write them in the spreadsheet. 
        /// </summary>
        /// <param name="pewWorksheet2"></param>
        /// <param name="pstFirstLevelMenu"></param>
        /// <param name="pstrByIdName"></param>
        public void CompareItems(ExcelWorksheet pewWorksheet2, string pstFirstLevelMenu, string pstrByIdName)
        {
            int r = 0; //A counter that checks each row of a certain column for the elements
            int ContSub = 0; //Counter that knows how many elements are in a certain column

            //This while checks the elements inside the Excel sheet that contains the menus and their submenus to know how much submenus a menu has
            //Menus are represented by columns and their submenus are the elements of that column in an ordered manner.
            while (m_leeExcelObject.FindElement(m_strWorkbookName,pewWorksheet2.Name,r+1,pstFirstLevelMenu) != null)
            {
                r++;
                ContSub++;
            }

            //This for loops through every submenu in a column, if it finds a null element (aka the element after the final 
            // element of that column, it ignores it. 
            for (int index2 = 0; index2 < ContSub; index2++)
            {
                if (m_leeExcelObject.FindElement(m_strWorkbookName,pewWorksheet2.Name,index2 + 2, pstFirstLevelMenu) == null)
                {
                    continue;
                }
                else
                {
                    try
                    {
                        string excelString = m_leeExcelObject.FindElement(m_strWorkbookName, pewWorksheet2.Name, index2 + 2, pstFirstLevelMenu);

                        By byId = By.Id(pstrByIdName);                         //This set of actions get the id of the menu we want to check
                        Actions action = new Actions(m_iwbWebDriver);      //And does a mouseover to check each one
                        IWebElement we = m_iwbWebDriver.FindElement(byId);
                        action.MoveToElement(we).Build().Perform();

                        string webString = m_iwbWebDriver.FindElement(By.LinkText(excelString)).GetAttribute("text");
                        Console.WriteLine(webString);
                        Assert.AreEqual(excelString, webString);
                    }
                    catch (Exception e) //catch ensures that the program continues even after it finds an incorrect submenu name
                    {
                        Console.WriteLine("Exception (Unable to locate): " + e);
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

        public string GetWebElements(IList<IWebElement> pilWebElements, int pintLimit)
        {
            string ResultString = null;
            for (int i = 0; i < pilWebElements.Count; i++)
            {
                string buttons = pilWebElements[i].Text;
                if (i != pilWebElements.Count - 1)
                {
                    ResultString = ResultString + buttons + ", ";
                }
                else
                {
                    ResultString = ResultString + buttons + " ";
                }
            }

            return ResultString;
        }
    }
}
