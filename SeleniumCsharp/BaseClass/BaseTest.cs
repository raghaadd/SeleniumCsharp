using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumCsharp.BaseClass
{
    public class BaseTest
    {
        public IWebDriver driver;
        [OneTimeSetUp]
        public void Open()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://borsheims.com/");


        }

        [OneTimeTearDown]
        public void Close()
        {
            Thread.Sleep(5000);
            driver.Dispose();
        }
    }


}
       
       
