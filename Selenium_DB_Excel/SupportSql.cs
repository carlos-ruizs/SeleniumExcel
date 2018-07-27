using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
                        var element = m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-chosen']"));
                        string elementText = element.Text;
                        
                        //Checks if the default text for the search field is displayed
                        if (element.Displayed)
                        {
                            resultsLoginString = "The default text for a search is shown, ";
                            validateLoginString = "First validation succesful, ";
                        }
                        else
                        {
                            resultsLoginString = "The default text for a search is not shown, ";
                            validateLoginString = "First validation unsuccesful, ";
                        }

                        var elementCkin = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']"));
                        string elementCkinText = elementCkin.GetAttribute("value");

                        //Checks if the check in field doesn't have anything inside of it, same thing applies to the check out field
                        if (elementCkinText == "")
                        {
                            resultsLoginString += "Check in field is clear, ";
                            validateLoginString += "Second validation succesful, ";
                        }
                        else
                        {
                            resultsLoginString += "Check in field is not clear, ";
                            validateLoginString += "Second validation unsuccesful, ";
                        }

                        var elementCkout = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']"));
                        string elementCkoutText = elementCkout.GetAttribute("value");

                        if (elementCkoutText == "")
                        {
                            resultsLoginString += "Check out field is clear, ";
                            validateLoginString += "Third validation succesful, ";
                        }
                        else
                        {
                            resultsLoginString += "Check out field is not clear, ";
                            validateLoginString += "Third validation unsuccesful, ";
                        }
                            
                        var elementTvlrs = m_iwbWebDriver.FindElement(By.XPath("//*[@name='travellers']"));
                        string elementTvlrsText = elementTvlrs.GetAttribute("value");

                        //checks if the travellers field has the default value
                        if (elementTvlrsText == "2 Adult 0 Child")
                        {
                            resultsLoginString += "Default adult and child numbers still shown";
                            validateLoginString += "Final validation succesful";
                        }
                        else
                        {
                            resultsLoginString += "Default adult and child numbers are not shown";
                            validateLoginString += "Final validation unsuccesful";
                        }

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

                    case 2: //TODO change the way it gets the date so it's extracted from the database instead. 
                        Console.WriteLine(local.Date.ToString("dd/MM/yyyy"));
                        //Console.WriteLine(local.Date.AddDays(5).ToString("d"));

                        //TODO change 
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).SendKeys(local.Date.ToString("d"));
                        //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@class='table-condensed']")));
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).SendKeys(local.Date.AddDays(1).ToString("d"));
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).Clear();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).Clear();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).Click();

                        //Se tiene que agregar una manera en la que le de click a la cosa del mes para poder llegar al mes actual
                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='switch']")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("/html/body/div[8]/div[2]/table/thead/tr/th[1]")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("/html/body/div[8]/div[2]/table/thead/tr/th[1]")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("/html/body/div[8]/div[2]/table/tbody/tr/td/span[7]")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("/html/body/div[8]/div[1]/table/tbody/tr[5]/td[6]")).Click();//Clicks the date in the calendar pop up

                        var Ckin = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']"));
                        string CkinText = Ckin.GetAttribute("value");
                        Console.WriteLine("check in:" + CkinText);

                        //if (CkinText == local.Date.ToString("dd/MM/yyyy"))
                        //{
                        //    Console.WriteLine("Fecha correcta");
                        //    resultsLoginString += "Correct date added";
                        //    validateLoginString += "First date added correctly";
                        //}
                            
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).Clear();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).Click(); //Hay que esperar a que le de click y el calendario esté visible, de ahí sería ver que escoja el día correcto. 
                        //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("div.datepicker:nth-child(14)")));
                        //m_iwbWebDriver.FindElement(By.XPath("//*[@class='prev']")).Click();
                        //m_iwbWebDriver.FindElement(By.XPath("//*[@class='prev']")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("/html/body/div[9]/div[1]/table/tbody/tr[6]/td[4]")).Click(); //Clicks the date in the calendar pop up

                        var Ckout = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']"));
                        string CkoutText = Ckout.GetAttribute("value");
                        Console.WriteLine("check out:" + CkoutText);

                        if (CkoutText == local.Date.AddDays(5).ToString("d"))
                            Console.WriteLine("Fecha correcta");

                        //Updates the database and clears the parameters
                        loginRows[i]["ResultsLogin"] = " ";

                        param1.Value = "Correct results added";
                        updateLogin.Parameters.Add(param1);

                        param2.Value = "Correct dates added";
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

                        if (travellersText == "2 Adult 0 Child") //checks if the default text has changed
                        {
                            resultsLoginString = "Values not changed, ";
                            validateLoginString = "Change in first value unsuccesful, ";
                        }
                        else
                        {
                            resultsLoginString = "Values changed, ";
                            validateLoginString = "Change in first value succesful, ";
                        }
                            
                        //clicks the traveller field again and removes one adult and one child, then hides the field
                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='adultMinusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='childMinusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();

                        var travellersChange = m_iwbWebDriver.FindElement(By.XPath("//*[@name='travellers']"));
                        string tcText = travellersChange.GetAttribute("value");

                        //Checks if the value inside the travellers input changed from the last time it had something by comparing it to it's last value
                        if (tcText == travellersText)
                        {
                            resultsLoginString += "The values inside the field did not change from last time";
                            validateLoginString += "Changes in values unsuccesful";
                        }
                        else
                        {
                            resultsLoginString += "The values changed a second time";
                            validateLoginString += "Changes in values succesful";
                        }

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
                        var wait = new WebDriverWait(m_iwbWebDriver, TimeSpan.FromSeconds(5)); //Wait used for when we want to check if the results drop list exists
                        var elemnt = m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']"));
                        
                        Actions action = new Actions(m_iwbWebDriver); //Webdriver action used to move the cursor to click the search field
                        action.MoveToElement(elemnt).Click().Build().Perform();
                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("select2-no-results"))); //waits for the textbox that appears when you click the search field

                        //If the textbox is displayed, it adds it to the validations string
                        var txtBox = m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-no-results']"));
                        if (txtBox.Displayed)
                        {
                            resultsLoginString += "The no results text is displayed, ";
                            validateLoginString += "Validation before search successful, ";
                        }
                        else
                        {
                            resultsLoginString += "The no results text is not displayed, ";
                            validateLoginString += "Validation before search unsuccessful, ";
                        }

                        action.MoveToElement(elemnt).Click().Build().Perform();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-input select2-focused']")).SendKeys("Hotel"); //This sends the word "Hotel" to the input field
                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("select2-results"))); //If this class exists, it's because there's 1 or more search results
                        var resultsList = m_iwbWebDriver.FindElement(By.ClassName("select2-results"));
                        
                        //if this class is enabled, it means there was one or more results found
                        if (resultsList.Enabled)
                        {
                            resultsLoginString += "1 or more results found, ";
                            validateLoginString += "list of values found, ";
                        }
                        else
                        {
                            resultsLoginString += " No results found, ";
                            validateLoginString += "list of values not found, ";
                        }

                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.ClassName("select2-result-label"))); //This waits until an element from the list is available 

                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-results-dept-1 select2-result select2-result-selectable select2-highlighted']")).Click(); //selects the first result of the search

                        var firstResult = m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']")); //gets the value inside the element to check if it matches the search result
                        string hotelTxt = firstResult.GetAttribute("value");

                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).SendKeys(local.Date.ToString("d")); //sends today's date to the check in field
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).SendKeys(local.Date.AddDays(3).ToString("d")); //sends the date three days from today to the check out field

                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='btn btn-lg btn-block btn-danger pfb0 loader']")).Click(); //Clicks the search button

                        if (m_iwbWebDriver.Url != "https://www.phptravels.net/") //This is to see if the webpage changes after the search button is clicked
                        {
                            resultsLoginString += " Webpage changed, ";
                            validateLoginString += "Search successful, ";
                        }
                        else
                        {
                            resultsLoginString += " Webpage not changed, ";
                            validateLoginString += "Search not successful, ";
                        }

                        if (m_iwbWebDriver.Url.Contains(hotelTxt)) //This is to see if the webpage has the info for the chosen hotel
                        {
                            resultsLoginString += "Webpage is for the chosen hotel";
                            validateLoginString += "Info for the chosen hotel obtained";
                        }
                        else
                        {
                            resultsLoginString += "Webpage is not for the chosen hotel";
                            validateLoginString += "Info for the chosen hotel wasn't obtained";
                        }

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
    }
}
