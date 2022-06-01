using System;
using System.Collections.Generic;
using System.Threading;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace AutoTest
{
    public class TestRequest
    {
        private IWebDriver sdriver;
        private readonly By errorTag = By.XPath("//div[@class='error_number']");
        //private string elemResult;
        public List<string> results = new List<string>();

        public void Setup()
        {
            sdriver = new ChromeDriver();
        }

        public void Quit()
        {
            sdriver.Quit();
        }

        // https://bus.gov.ru/agency/1437
        public void FindElementError(string num)
        {
            try
            {
                sdriver.Navigate().GoToUrl("https://bus.gov.ru/agency/" + num);
                Thread.Sleep(5000);
                IWebElement webElem = sdriver.FindElement(errorTag);
                //elemResult = webElem.Text;
            }
            catch(Exception ex)
            {
                //elemResult = "OK";
                results.Add(num);
            }
            finally
            {
                //results.Add(num, elemResult);
            }
            //foreach (var el in results)
            //{
            //    Console.WriteLine("{0} : {1}", el.Key, el.Value);
            //}
        }
    }
}
