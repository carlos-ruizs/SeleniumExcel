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

        public void Reservation()
        {
            DateTime local = DateTime.Today;

            //El atributo "value" de las textboxes es lo que me dice efectivamente si cambió algo o si está vacío
            //Y el texto placeholder sigue visible. Eso es lo que debo checar para el task 1. Continuar el lunes. 


            //Entonces necesito ver si está intacta la caja de texto antes de hacer algo. 



            //This two pieces of code will help me for the third test case
            //Estos 2 pedacitos de código me van a servir para el caso 3 para cuando tenga que ver que cambie y luego volver a cambiarlo
            //También podría servirme probablemente para el segundo test case porque igual hay que darle click a la text box
            //Y luego darle click al elemento del calendario.




            //TestCase 3 y 4
            //DateTime local = DateTime.Today;

            DataRow[] loginRows = masterTable.Select("Actions = 'Reservation'");

            SqlCommand updateLogin = new SqlCommand("UpValidateResults", connection);
            updateLogin.CommandType = CommandType.StoredProcedure;

            daAdapter.UpdateCommand = updateLogin;

            SqlParameter param1 = new SqlParameter("@ResultsLogin", SqlDbType.VarChar);
            SqlParameter param2 = new SqlParameter("@ValidateLogin", SqlDbType.VarChar);
            SqlParameter param3 = new SqlParameter("@TestCase", SqlDbType.TinyInt);
            //Revisar la parte de los stored procedures mañana


            for (int i = 0; i < loginRows.Length; i++)
            {
                m_iwbWebDriver.Navigate().GoToUrl("https://www.phptravels.net/");
                m_iwbWebDriver.FindElement(By.XPath("//*[@data-title='HOTELS']")).Click(); //Clicks the Hotels button by finding it's Xpath

                Console.WriteLine(loginRows[i]["Actions"] + " " + loginRows[i]["TestCase"]);
                int switchOption = int.Parse(loginRows[i]["TestCase"].ToString());
                switch (switchOption)
                {
                    case 1:
                        //TestCase 1: verify that all the text boxes have their default values
                        var element = m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-chosen']"));
                        string elementText = element.Text;
                        var elementForm = m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']"));
                        string elementFormTxt = element.GetAttribute("value");
                        //element.Text sería para obtener el inner text del elemento
                        //El GetAttribute sería para saber si hay algo dentro de la caja de texto
                        if (elementFormTxt == null)
                            Console.WriteLine("El texto por defecto sigue ahí");
                        Console.WriteLine(elementText);

                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']")).SendKeys("Hola");
                        var elmnt = m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']"));
                        string elmntText = elmnt.GetAttribute("value");
                        Console.WriteLine(elmntText);

                        var elementCkin = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']"));
                        string elementCkinText = elementCkin.GetAttribute("value"); //value se actualiza si se le pone algo antes,
                        Console.WriteLine(elementCkinText);
                        if (elementCkinText == "")
                            Console.WriteLine("Check in field is clear");

                        var elementCkout = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']"));
                        string elementCkoutText = elementCkout.GetAttribute("value");
                        Console.WriteLine(elementCkoutText);
                        if (elementCkoutText == "")
                            Console.WriteLine("Check out field is clear");

                        var elementTvlrs = m_iwbWebDriver.FindElement(By.XPath("//*[@name='travellers']"));
                        string elementTvlrsText = elementTvlrs.GetAttribute("value");
                        Console.WriteLine(elementTvlrsText);
                        if (elementTvlrsText == "2 Adult 0 Child")
                            Console.WriteLine("Default text is still shown");

                        loginRows[i]["ResultsLogin"] = "Everything's fine";

                       
                        //param1.ParameterName = "@ResultsLogin";
                        param1.Value = "default text still shows" + " Check in field is clear" + " Check out field is clear" + " Default text is still shown";

                        updateLogin.Parameters.Add(param1);

                        
                        //param2.ParameterName = "@ValidateLogin";
                        param2.Value = "Login succesful";

                        updateLogin.Parameters.Add(param2);

                        
                        //param3.ParameterName = "@TestCase";
                        param3.Value = switchOption;

                        updateLogin.Parameters.Add(param3);

                        daAdapter.Update(dataSet.Tables["Master"]);
                        break;

                    case 2:
                        Console.WriteLine(local.Date.ToString("d"));
                        Console.WriteLine(local.Date.AddDays(5).ToString("d"));

                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).SendKeys(local.Date.ToString("d"));
                        //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.XPath("//*[@class='table-condensed']")));
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).SendKeys(local.Date.AddDays(1).ToString("d"));
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).Clear();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).Clear();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).Click();

                        //Se tiene que agregar una manera en la que le de click a la cosa del mes para poder llegar al mes actual
                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='switch']")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("/html/body/div[8]/div[2]/table/tbody/tr/td/span[7]")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).Click();
                        m_iwbWebDriver.FindElement(By.CssSelector("div.datepicker:nth-child(13) > div:nth-child(1) > table:nth-child(1) > tbody:nth-child(2) > tr:nth-child(3) > td:nth-child(4)")).Click(); //Clicks the date in the calendar pop up

                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).Clear();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).Click(); //Hay que esperar a que le de click y el calendario esté visible, de ahí sería ver que escoja el día correcto. 
                        //wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.CssSelector("div.datepicker:nth-child(14)")));
                        //m_iwbWebDriver.FindElement(By.XPath("//*[@class='prev']")).Click();
                        //m_iwbWebDriver.FindElement(By.XPath("//*[@class='prev']")).Click();
                        m_iwbWebDriver.FindElement(By.CssSelector("div.datepicker:nth-child(14) > div:nth-child(1) > table:nth-child(1) > tbody:nth-child(2) > tr:nth-child(4) > td:nth-child(2)")).Click(); //Clicks the date in the calendar pop up
                        var Ckout = m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']"));
                        string CkoutText = Ckout.GetAttribute("value");
                        Console.WriteLine(CkoutText);

                        break;

                    case 3:
                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='childPlusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='childPlusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();

                        var travellers = m_iwbWebDriver.FindElement(By.XPath("//*[@name='travellers']"));
                        string travellersText = travellers.GetAttribute("value");
                        Console.WriteLine(travellersText);

                        if (travellersText == "2 Adult 0 Child")
                            Console.WriteLine("Default text is still shown");
                        else
                            Console.WriteLine("The values changed");

                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='adultMinusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@id='childMinusBtn']")).Click();
                        m_iwbWebDriver.FindElement(By.Id("travellersInput")).Click();

                        var travellersChange = m_iwbWebDriver.FindElement(By.XPath("//*[@name='travellers']"));
                        string tcText = travellersChange.GetAttribute("value");
                        Console.WriteLine(tcText);

                        if (tcText == travellersText)
                            Console.WriteLine("The value is the same");
                        else
                            Console.WriteLine("The value changed");

                        break;

                    case 4:
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']")).SendKeys(" ");
                        var wait = new WebDriverWait(m_iwbWebDriver, TimeSpan.FromSeconds(5));
                        var elemnt = m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']"));
                        string elemntText = elemnt.GetAttribute("value");
                        Console.WriteLine(elemntText);
                        
                        Actions action = new Actions(m_iwbWebDriver);
                        action.MoveToElement(elemnt).Click().Build().Perform();
                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("select2-no-results")));

                        var txtBox = m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-no-results']"));
                        if (txtBox.Displayed)
                        {
                            Console.WriteLine("The results text is displayed");
                        }
                        else
                        {
                            Console.WriteLine("The results text is not displayed");
                        }

                        action.MoveToElement(elemnt).Click().Build().Perform();
                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-input select2-focused']")).SendKeys("Hotel");
                        wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementExists(By.ClassName("select2-results"))); //Si existe results, es que hay resultados ahí. Uno o más
                        IList<IWebElement> results = m_iwbWebDriver.FindElements(By.ClassName("select2-result-label"));
                        var val = m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']"));
                        string valText = val.GetAttribute("value");

                        Console.WriteLine(valText);

                       wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(By.ClassName("select2-result-label")));

                        for (int j = 0; j < results.Count; j++)
                        {
                            string name = results[j].GetAttribute("value");
                            Console.WriteLine();
                        }

                        var firstResult = m_iwbWebDriver.FindElement(By.ClassName("select2-result-label"));
                        //action.MoveToElement(firstResult).Click().Build().Perform();
                        var hotelName = m_iwbWebDriver.FindElement(By.XPath("//*[@name='hotel_s2_text']"));
                        string hotelTxt = hotelName.GetAttribute("value");

                        Console.WriteLine(hotelTxt);

                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='select2-results-dept-1 select2-result select2-result-selectable select2-highlighted']")).Click();

                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkin']")).SendKeys(local.Date.ToString("d"));
                        m_iwbWebDriver.FindElement(By.XPath("//*[@name='checkout']")).SendKeys(local.Date.AddDays(3).ToString("d"));

                        m_iwbWebDriver.FindElement(By.XPath("//*[@class='btn btn-lg btn-block btn-danger pfb0 loader']")).Click();

                        if (m_iwbWebDriver.Url != "https://www.phptravels.net/")
                            Console.WriteLine("The web page changed");

                        var hotelInfo = m_iwbWebDriver.FindElement(By.XPath("//*[@class='ellipsis ttu']"));
                        string infoTxt = hotelInfo.Text; //Sí lo consigue, pero está en mayúsculas, habría que volverlo minúsculas.

                        Console.WriteLine(infoTxt);

                        //Falta verificar que sea el hotel que elegí, por lo que debo de ver la forma de agarrar el value dentro de la caja de texto
                        //del search antes de apretar el botón de buscar.


                        //var resultsBox = m_iwbWebDriver.FindElement(By.ClassName(""));
                        //var defText = m_iwbWebDriver.FindElement(By.ClassName("select2-no-results"));
                        //string dfValue = defText.Text;
                        //Console.WriteLine(elemntText);

                        break;

                    default:
                        Console.WriteLine("The Test Case doesn't exist");
                        break;
                }
            }
        }
    }
}
