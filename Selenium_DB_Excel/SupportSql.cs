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

namespace Selenium_DB_Excel
{
    class SupportSql
    {
        SqlConnection connection;
        DataSet dataSet;
        DataTable masterTable;
        public IWebDriver m_iwbWebDriver;

        public SupportSql()
        {
            connection = new SqlConnection("Data Source=.\\SQLEXPRESS;Initial Catalog = Selenium_DB;User ID=cruiz;Password=CR2018cr");
        }

        public void DataFill()
        {
            SqlDataAdapter daAdapter = new SqlDataAdapter("SELECT * FROM Master WHERE Run = 1", connection); //Esta query mete los elementos con un 1 en su columna Run al programa, de tal forma que sólo aquellos que se vayan a ejecutar, entren. 
            dataSet = new DataSet();
            SqlCommandBuilder commandBuilder = new SqlCommandBuilder(daAdapter);
            daAdapter.Fill(dataSet, "Master");
            masterTable = dataSet.Tables["Master"];
            ExecuteCases();
            
        /*
        foreach (DataColumn col in masterTable.Columns)
        {
            Console.WriteLine(col.ColumnName);
        }

        foreach (DataRow row in masterTable.Rows)
        {
            foreach (var item in row.ItemArray)
            {
                Console.Write(item);
            }
        }
        */

        Console.ReadKey();
        }

        public void ExecuteCases()
        {
            Search();
            Login();
        }

        /// <summary>
        /// Gets the rows with the "Search" action and calls the webdriver to do a Google Search and return the results
        /// </summary>
        public void Search()
        {
            DataRow[] searchRows = masterTable.Select("Actions = 'Search'");

            for (int i = 0; i < searchRows.Length; i++)
            {
                Console.WriteLine(searchRows[i]["Actions"]+" "+searchRows[i]["InputParameter"] + " " + searchRows[i]["NoResultsToSave"]);
            }
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
