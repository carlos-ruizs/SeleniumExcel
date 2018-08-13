using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using SeleniumExcel;

namespace Selenium_DB_Excel
{
    class SupportSql
    {
        public SqlConnection connection;
        DataSet dataSet;
        DataTable masterTable;
        public IWebDriver m_iwbWebDriver;
        private Support support;
        SqlDataAdapter daAdapter;
        SqlCommandBuilder commandBuilder;

        public SupportSql()
        {
            try
            {
                connection = new SqlConnection("Data Source=.\\SQLEXPRESS;Initial Catalog = Selenium_DB;User ID=cruiz;Password=CR1194cr");
                m_iwbWebDriver = new FirefoxDriver(@"C:\geckodriver-v0.19.1-win64");
                support = new Support();
            }
            catch (WebDriverException wdException)
            {
                Console.WriteLine(wdException.Message);
            }
        }

        /// <summary>
        /// Uses the DataAdapter object to fill a DataSet with the rows from the Master table where the Run column is 
        /// set to true. Then, uses a CommandBuilder to automatically have all the Insert, Update and Delete operations
        /// for the database. Finally the DataSet is set to include a new table called Master. 
        /// </summary>
        public void DataFill()
        {
            daAdapter = new SqlDataAdapter("SELECT * FROM Master WHERE Run = 1", connection); //Query that takes the elements with a 1 in the Run column from the database and uses that as the SELECT command of the adapter 
            dataSet = new DataSet();
            commandBuilder = new SqlCommandBuilder(daAdapter); //We use the CommandBuilder to generate the INSERT, UPDATE and DELETE commands after we've already set the SELECT command
            daAdapter.Fill(dataSet, "Master"); //Fills the DataSet with the data from the DataAdapter in a table we call "Master"
            masterTable = dataSet.Tables["Master"];
            ExecuteCases();
        }

        /// <summary>
        /// Method that calls all the other methods that use the data from the database. If the method checks and 
        /// there's no 1 on the Run column for that particular Action, it does nothing.
        /// </summary>
        private void ExecuteCases()
        {
            Search();
            Reservation();
            m_iwbWebDriver.Close();
        }

        /// <summary>
        /// Gets the rows with the "Search" action and calls the webdriver to do a Google Search and return the results.
        /// It then uses the desired results and saves them to the DataTable created at the beginning of the execution.
        /// Finally, the DataAdapter updates the database with the updated DataSet.
        /// </summary>
        public void Search()
        {
            m_iwbWebDriver.Navigate().GoToUrl("http://www.google.com/");
            DataRow[] searchRows = masterTable.Select("Actions = 'Search'"); //Gets the rows in the DataTable with the Search action to work exclusively on them

            for (int i = 0; i < searchRows.Length; i++)
            {
                int elementsToSave = int.Parse(searchRows[i]["NoResultsToSave"].ToString()); //Gets the number of links to save from the current search and converts it to an integer
                m_iwbWebDriver.FindElement(By.Id("lst-ib")).SendKeys(searchRows[i]["InputParameter"].ToString()); //Gets the string in the InputParameter column and sends it to the search bar

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

                //Inserts the results of the search into the Master DataTable
                searchRows[i]["TotalResults"] = support.GetTotalSearchResults(totalSearchResults);
                searchRows[i]["SavedResultsLinks"] = support.GetResultsHref(h3Links, elementsToSave - 1);
                searchRows[i]["SavedResultsText"] = support.GetResultsTxt(h3Links, elementsToSave - 1);
                searchRows[i]["RelatedResultsLinks"] = support.GetResultsHref(relatedResults, relatedResults.Count - 1);
                searchRows[i]["RelatedResultsText"] = support.GetResultsTxt(relatedResults, relatedResults.Count - 1);

                m_iwbWebDriver.FindElement(By.Id("lst-ib")).Clear(); //Clears the search bar for the next word
            }
            daAdapter.Update(dataSet.Tables["Master"]); //Uses the updated Master table to update the database
        }

        //@TODO Make it so that whenever something is validated, the program takes a screenshot of the webpage for future reference
        public void Reservation()
        {
            DateTime local = DateTime.Today; //Used to get the date for the test cases that need it

            DataRow[] loginRows = masterTable.Select("Actions = 'Reservation'"); //Gets all the "Reservation" rows to run them

            SqlCommand updateLogin = new SqlCommand("UpValidateResults", connection); //Changes the Update method of the adapter for a stored procedure inside the database
            updateLogin.CommandType = CommandType.StoredProcedure;
            daAdapter.UpdateCommand = updateLogin;

            //Sets the parameters used for the stored procedure
            SqlParameter param1 = new SqlParameter("@ResultsLogin", SqlDbType.VarChar);
            SqlParameter param2 = new SqlParameter("@ValidateLogin", SqlDbType.VarChar);
            SqlParameter param3 = new SqlParameter("@TestCase", SqlDbType.TinyInt);

            for (int i = 0; i < loginRows.Length; i++)
            {
                m_iwbWebDriver.Navigate().GoToUrl("https://www.phptravels.net/");
                m_iwbWebDriver.FindElement(By.XPath("//*[@data-title='HOTELS']")).Click(); //Clicks the Hotels button by finding it's Xpath

                string resultsLoginString = null; //strings used to tell the database if the tests where succesful in specific things
                string validateLoginString = null;

                int switchOption = int.Parse(loginRows[i]["TestCase"].ToString()); //Gets the number of test case to execute and converts it to an integer to be used in the switch
                switch (switchOption)
                {
                    case 1:
                        //Checks if the default text for the search field is displayed
                        var element = m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-chosen']"));
                        string elementText = element.Text;
                        resultsLoginString += IsValid(element.Displayed);

                        //Checks if the check in field doesn't have anything inside of it, same thing applies to the check out field
                        var elementCkin = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']"));
                        string elementCkinText = elementCkin.GetAttribute("value");
                        resultsLoginString += IsValid(elementCkinText == "");

                        var elementCkout = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']"));
                        string elementCkoutText = elementCkout.GetAttribute("value");
                        resultsLoginString += IsValid(elementCkoutText == "");

                        //checks if the travellers field has the default value    
                        var elementTvlrs = m_iwbWebDriver.FindElement(By.XPath("//*[@name='travellers']"));
                        string elementTvlrsText = elementTvlrs.GetAttribute("value");
                        resultsLoginString += IsValid(elementTvlrsText == "2 Adult 0 Child");

                        validateLoginString += ValidCount(resultsLoginString);

                        //Updates the database and clears the parameters 
                        loginRows[i]["ResultsLogin"] = " ";

                        param1.Value = resultsLoginString;
                        updateLogin.Parameters.Add(param1);

                        param2.Value = validateLoginString;
                        updateLogin.Parameters.Add(param2);

                        param3.Value = switchOption;
                        updateLogin.Parameters.Add(param3);

                        daAdapter.Update(dataSet.Tables["Master"]);

                        updateLogin.Parameters.Clear();
                        break;

                    case 2:
                        var wait1 = new WebDriverWait(m_iwbWebDriver, TimeSpan.FromSeconds(30));
                        Console.WriteLine(local.Date.ToString("dd/MM/yyyy"));
                        DateTime dbDate = DateTime.Parse(loginRows[i]["Date"].ToString());

                        int monthDB = int.Parse(dbDate.Month.ToString());
                        Console.WriteLine(monthDB);
                        int dayDB = int.Parse(dbDate.Day.ToString());
                        Console.WriteLine(dayDB);
                        int yearDB = int.Parse(dbDate.Year.ToString());
                        Console.WriteLine(yearDB);

                        int monthNow = int.Parse(local.Date.Month.ToString());
                        int today = int.Parse(local.Date.Day.ToString());
                        int thisYear = int.Parse(local.Date.Year.ToString());

                        //This checks if the date inside the database is equal to today's or a future date so that the program can work correctly
                        if (yearDB < thisYear)
                        {
                            Console.WriteLine("Date inside the database is before today");
                            break;
                        }
                        else
                        {
                            if (yearDB > thisYear)
                            {

                            }
                            else
                            {
                                if (monthDB < monthNow)
                                {
                                    Console.WriteLine("Date inside the database is before today");
                                    break;
                                }
                                else
                                {
                                    if (monthDB > monthNow)
                                    {

                                    }
                                    else
                                    {
                                        if (dayDB < today)
                                        {
                                            Console.WriteLine("Date is before today");
                                            break;
                                        }
                                    }
                                }
                            }
                        }

                        //This part is what gets the check in and check out buttons and sends a regular date to the fields and then clears them
                        IWebElement checkInBtn = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']"));
                        IWebElement checkOutBtn = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']"));
                        checkInBtn.SendKeys(dbDate.Date.ToString("dd/MM/yyyy"));
                        checkOutBtn.SendKeys(dbDate.Date.AddDays(5).ToString("dd/MM/yyyy"));
                        checkInBtn.Clear();
                        checkOutBtn.Clear();

                        //This parts clicks the fields again to select the date from the tables
                        //First it clicks the check in button, gets the table elements from the website and uses the corresponding table body to select the day with the SelectDay method
                        //The check out button is not pressed since it automatically is pressed when the check in date is selected.
                        checkInBtn.Click();
                        IJavaScriptExecutor js = (IJavaScriptExecutor)m_iwbWebDriver;
                        IList<IWebElement> daysTable = m_iwbWebDriver.FindElements(By.ClassName("datepicker-days"));
                        Console.WriteLine(daysTable.Count);
                        IWebElement daysBody = daysTable[0].FindElement(By.TagName("tbody"));
                        SelectDay(daysBody, dbDate.Day.ToString());
                        IWebElement daysBod2 = daysTable[1].FindElement(By.TagName("tbody"));
                        SelectDay(daysBod2, dbDate.AddDays(5).Day.ToString());

                        string checkInValue = checkInBtn.GetAttribute("value");
                        Console.WriteLine(checkInValue);
                        resultsLoginString += IsValid(checkInValue == dbDate.Date.ToString("dd/MM/yyyy"));
                        string checkOutValue = checkOutBtn.GetAttribute("value");
                        Console.WriteLine(checkOutValue);
                        resultsLoginString += IsValid(checkOutValue == dbDate.AddDays(5).Date.ToString("dd/MM/yyyy"));

                        validateLoginString += ValidCount(resultsLoginString);

                        //Updates the database and clears the parameters
                        loginRows[i]["ResultsLogin"] = " ";

                        param1.Value = resultsLoginString;
                        updateLogin.Parameters.Add(param1);

                        param2.Value = validateLoginString;
                        updateLogin.Parameters.Add(param2);

                        param3.Value = switchOption;
                        updateLogin.Parameters.Add(param3);

                        daAdapter.Update(dataSet.Tables["Master"]);

                        updateLogin.Parameters.Clear();

                        break;

                    case 3:
                        //Clicks the travellers field and adds 2 children since 2 adults is the default text, then clicks the field again to hide it
                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='childPlusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='childPlusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();

                        //This gets the value inside the travellers field to check if it changed
                        var travellers = m_iwbWebDriver.FindElement(By.XPath("//*[@name='travellers']"));
                        string travellersText = travellers.GetAttribute("value");

                        resultsLoginString += IsValid(travellersText != "2 Adult 0 Child"); //Checks if the default text changed

                        //clicks the traveller field again and removes one adult and one child, then hides the field
                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='adultMinusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='childMinusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();

                        var travellersChange = m_iwbWebDriver.FindElement(By.XPath("//*[@name='travellers']"));
                        string tcText = travellersChange.GetAttribute("value");

                        //Checks if the value inside the travellers input changed from the last time it had something by comparing it to it's last value
                        resultsLoginString += IsValid(tcText != travellersText);

                        validateLoginString += ValidCount(resultsLoginString);

                        //Updates the database and clears the parameters
                        loginRows[i]["ResultsLogin"] = " ";

                        param1.Value = resultsLoginString;
                        updateLogin.Parameters.Add(param1);

                        param2.Value = validateLoginString;
                        updateLogin.Parameters.Add(param2);

                        param3.Value = switchOption;
                        updateLogin.Parameters.Add(param3);

                        daAdapter.Update(dataSet.Tables["Master"]);

                        updateLogin.Parameters.Clear();

                        break;

                    case 4:
                        var wait = new WebDriverWait(m_iwbWebDriver, TimeSpan.FromSeconds(30)); //Wait used for when we want to check if the results drop list exists
                        var elemnt = m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']"));

                        Actions action = new Actions(m_iwbWebDriver); //Webdriver action used to move the cursor to click the search field
                        action.MoveToElement(elemnt).Click().Build().Perform();
                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("select2-no-results"))); //waits for the textbox that appears when you click the search field

                        //If the textbox is displayed, it adds it to the validations string
                        var txtBox = m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-no-results']"));

                        resultsLoginString += IsValid(txtBox.Displayed);

                        action.MoveToElement(elemnt).Click().Build().Perform();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-input select2-focused']")).SendKeys(loginRows[i]["InputParameter"].ToString()); //This sends the word "Hotel" to the input field
                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("select2-results"))); //If this class exists, it's because there's 1 or more search results
                        var resultsList = m_iwbWebDriver.FindElement(By.ClassName("select2-results"));

                        //if this class is enabled, it means there was one or more results found
                        resultsLoginString += IsValid(resultsList.Enabled);

                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.ClassName("select2-result-label"))); //This waits until an element from the list is available 

                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-results-dept-1 select2-result select2-result-selectable select2-highlighted']")).Click(); //selects the first result of the search

                        var firstResult = m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']")); //gets the value inside the element to check if it matches the search result
                        string hotelTxt = firstResult.GetAttribute("value");

                        DateTime date = DateTime.Parse(loginRows[i]["Date"].ToString()); //gets the date value inside the database and casts into the DateTime type to better manipulate it

                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).SendKeys((date.Date.ToString("dd/MM/yyyy"))); //sends today's date to the check in field
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).SendKeys(date.Date.AddDays(3).ToString("dd/MM/yyyy")); //sends the date three days from today to the check out field

                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='btn btn-lg btn-block btn-danger pfb0 loader']")).Click(); //Clicks the search button

                        //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.CssSelector(".ellipsis ttu")));//Waits for the webpage to load. Check further

                        resultsLoginString += IsValid(m_iwbWebDriver.Url != "https://www.phptravels.net/"); //This is to see if the webpage changes after the search button is clicked

                        resultsLoginString += IsValid(m_iwbWebDriver.Url.Contains(hotelTxt)); //This is to see if the webpage has the info for the chosen hotel

                        validateLoginString += ValidCount(resultsLoginString);

                        //Updates the database and clears the parameters
                        loginRows[i]["ResultsLogin"] = " ";

                        param1.Value = resultsLoginString;
                        updateLogin.Parameters.Add(param1);

                        param2.Value = validateLoginString;
                        updateLogin.Parameters.Add(param2);

                        param3.Value = switchOption;
                        updateLogin.Parameters.Add(param3);

                        daAdapter.Update(dataSet.Tables["Master"]);

                        updateLogin.Parameters.Clear();
                        break;

                    default:
                        Console.WriteLine("The Test Case doesn't exist");
                        break;
                }
            }
        }

        private string IsValid(bool pb_expression)
        {
            string validation = null;

            if (pb_expression)
            {
                validation += "Validation successful ";
                TakeScreenshot(@"C:\Users\cruiz\Documents\TestCases_Screenshots\", "Valid");
            }
            else
            {
                validation += "Validation unsuccessful ";
                TakeScreenshot(@"C:\Users\cruiz\Documents\TestCases_Screenshots\", "Invalid");
            }

            return validation;
        }

        private string ValidCount(string validationString)
        {
            int noValidations = Regex.Matches(validationString, "Validation successful").Count;
            return noValidations + " validations succesful";
        }

        private void SelectDay(IWebElement tableBody, string daySelected)
        {
            IList<IWebElement> tableRows = tableBody.FindElements(By.TagName("tr"));

            foreach (var rows in tableRows)
            {
                try
                {
                    var tds = rows.FindElements(By.ClassName("day"));
                    string dataRows = rows.GetAttribute("innerText");
                    Console.WriteLine(dataRows);
                    foreach (var td in tds)
                    {
                        string data = td.GetAttribute("innerText");
                        Console.WriteLine(data);
                        if (data == daySelected)
                        {
                            td.Click();
                            break;
                        }
                    }
                }
                catch (OpenQA.Selenium.StaleElementReferenceException SERE)
                {
                    break;
                }
            }
        }

        private void SelectYear(string yearText, IWebElement prevButton, IWebElement nextButton, IWebElement headerTitle)
        {
            int yearInt = int.Parse(yearText);
            Console.WriteLine(yearText);
            Console.WriteLine(yearInt);

            if (yearInt > 2018)
            {
                while (yearInt != 2018)
                {
                    prevButton.Click();
                    //monthsYear = monthsHeader.FindElement(By.ClassName("switch"));
                    yearText = headerTitle.Text;
                    yearInt = int.Parse(yearText);
                }
            }
            else
            {
                if (yearInt < 2018)
                {
                    while (yearInt != 2018)
                    {
                        nextButton.Click();
                        //monthsYear = monthsHeader.FindElement(By.ClassName("switch"));
                        yearText = headerTitle.Text;
                        yearInt = int.Parse(yearText);
                    }
                }
            }
        }

        private void CheckYear(DateTime date, IWebElement webElement)
        {
            int monthSelct = int.Parse(date.Month.ToString()); //mes actual

            string daySelct = date.Day.ToString(); //día actual

            int dayCheck = int.Parse(daySelct); //día actual convertido en int
            DateTime local = DateTime.Today;
            int monthNow = int.Parse(local.Date.Month.ToString());
            int today = int.Parse(local.Date.Day.ToString());

            //var year = m_iwbWebDriver.FindElement(By.XPath("//*[@class='switch']"));
            //string yearText = year.GetAttribute("textContent"); //innerText da "Mes Año" como resultado
            //Console.WriteLine(yearText);
            //m_iwbWebDriver.FindElement(By.XPath("//*[@class='switch']")).Click();
            //yearText = year.GetAttribute("innerText");
            //Console.WriteLine(yearText);

            IWebElement tableBody = m_iwbWebDriver.FindElement(By.XPath("//table/tbody"));
            IWebElement tableHeader = m_iwbWebDriver.FindElement(By.XPath("//table/thead"));

            IList<IWebElement> tableRows = tableBody.FindElements(By.TagName("tr"));
            int rows_count = tableRows.Count();
            Console.WriteLine(rows_count); //Da igual a 6, que es el número de filas del calendario como tal para un mes

            IList<IWebElement> tableCells = tableBody.FindElements(By.TagName("td"));
            int cell_count = tableCells.Count();
            Console.WriteLine(cell_count); //Da igual a 42 que es el número de celdas que tiene el calendario para un mes

            IList<IWebElement> row = tableHeader.FindElements(By.TagName("tr"));
            IList<IWebElement> heads = tableHeader.FindElements(By.TagName("th"));
            int rc = row.Count;
            int hc = heads.Count;
            Console.WriteLine(rc); //Da como resultado 2 filas porque son la fila del nombre del mes y el año junto con prev y next y además la fila de los días de la semana
            Console.WriteLine(hc); //Da como resultado 10 porque son los nombres de los botones (3) más los días de la semana (7)
            IWebElement switchBtn = tableHeader.FindElement(By.ClassName("switch"));
            switchBtn.Click(); //Tenemos que darle click para poder hacer que se vea la tabla de los meses

            //-------------There are two tables, one for months, one for days

            IWebElement months = m_iwbWebDriver.FindElement(By.ClassName("datepicker-months"));
            IWebElement monthTable = months.FindElement(By.TagName("table"));

            IWebElement monthsHeader = months.FindElement(By.TagName("thead"));
            IList<IWebElement> monthsHeaders = monthsHeader.FindElements(By.TagName("th"));
            int hcount = monthsHeaders.Count;
            Console.WriteLine(hcount);
            IWebElement monthsYear = monthsHeader.FindElement(By.ClassName("switch"));
            IWebElement yearPrev = monthsHeader.FindElement(By.ClassName("prev"));
            IWebElement yearNext = monthsHeader.FindElement(By.ClassName("next"));
            string mytxt = monthsYear.Text;
            Console.WriteLine(mytxt);
            int yearInt = int.Parse(mytxt);
            Console.WriteLine(mytxt);

            if (yearInt > 2018)
            {
                while (yearInt != 2018)
                {
                    yearPrev.Click();
                    monthsYear = monthsHeader.FindElement(By.ClassName("switch"));
                    mytxt = monthsYear.Text;
                    yearInt = int.Parse(mytxt);
                }
            }
            else
            {
                if (yearInt < 2018)
                {
                    while (yearInt != 2018)
                    {
                        yearNext.Click();
                        monthsYear = monthsHeader.FindElement(By.ClassName("switch"));
                        mytxt = monthsYear.Text;
                        yearInt = int.Parse(mytxt);
                    }
                }
            }

            IWebElement monthsBody = months.FindElement(By.TagName("tbody"));
            IList<IWebElement> monthsSelect = monthsBody.FindElements(By.TagName("span"));
            int selectCount = monthsSelect.Count;
            Console.WriteLine(selectCount);

            for (int i = 0; i < monthsSelect.Count; i++) //Los meses están en orden de manera 0 A n-1.
            {
                string monthsTxt = monthsSelect[i].Text;
                Console.WriteLine(monthsTxt);
                if (monthSelct == i + 1)
                {
                    monthsSelect[i].Click(); //Cuando le da click, inmediatamente se va al siguiente campo, no espera a que elijas el día. 
                    break;
                }
            }

            webElement.Click();
            IList<IWebElement> dayList = tableBody.FindElements(By.ClassName("day"));
            Console.WriteLine(dayList.Count);

            for (int i = 0; i < dayList.Count; i++)
            {
                if (daySelct == dayList[i].Text)
                {
                    dayList[i].Click();
                    break;
                }
            }
        }

        private void TakeScreenshot(string Path, string fileName)
        {
            StringBuilder TimeAndDate = new StringBuilder(DateTime.Now.ToString());
            TimeAndDate.Replace("/", "_");
            TimeAndDate.Replace(":", "_");

            Screenshot file = ((ITakesScreenshot)m_iwbWebDriver).GetScreenshot();

            file.SaveAsFile(Path + fileName + TimeAndDate.ToString() + ".jpeg", ScreenshotImageFormat.Jpeg);
        }
    }
}
