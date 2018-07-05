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
            connection = new SqlConnection("Data Source=.\\SQLEXPRESS;Initial Catalog = Selenium_DB;User ID=cruiz;Password=CR2018cr");
            m_iwbWebDriver = new FirefoxDriver(@"C:\geckodriver-v0.19.1-win64");
            support = new Support();
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
            Login();
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

        public void Login()
        {
            DataRow[] loginRows = masterTable.Select("Actions = 'Login'");

            for (int i = 0; i < loginRows.Length; i++)
            {
                Console.WriteLine(loginRows[i]["Actions"]+" "+loginRows[i]["TestCase"]);
            }
        }
    }
}
