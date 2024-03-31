using OpenQA.Selenium;
using SeleniumCsharp.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumCsharp
{
    [TestFixture]
    public class ParallelTesting
    {
        IWebDriver driver;
        [Test, Category("UAT Testing"),Category("Module1")]
        public void TestMethod1()
        {
            var Driver=new BrowserUtility().Init(driver);
            driver.FindElement(By.Id("l-Customer_LoginEmail")).SendKeys("testuser29@testqa.com");
            driver.FindElement(By.Id("l-Customer_Password")).SendKeys("wtx@91187");
            Driver.Close();


        }
    }
}
