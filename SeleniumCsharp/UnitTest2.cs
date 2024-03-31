using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumCsharp.BaseClass;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SeleniumCsharp
{
   
    [TestFixture]
    public class UnitTest2 : BaseTest2
    {
        string expectedUsername = null;
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://borsheims.com/");
        }

        [Test, Order(0)]
        public void NavigateWebsite()
        {
            try
            {
                //Navigate to Account page:
                driver.FindElement(By.XPath("//*[@id=\"js-header\"]/div[1]/div/div[3]")).Click();
                //Navigate to register:
                driver.FindElement(By.XPath("//*[@id=\"js-main-content\"]/div/div[1]/div/div/div[3]/div[2]/p[2]/a")).Click();
                WriteStatusToExcel(2, "Passed");
            }
            catch(Exception e)
            {
                WriteStatusToExcel(2, "Failed");
                Console.WriteLine(e);
                throw;
            }

        }

     
        [Test, Order(1), TestCaseSource(nameof(ReadExcel))]
        public void registerInformation(String email, String password, String confirmPassword, String firstName, String lastName, String emailAdd, String phoneNo,
                String faxNo, String company, String address, String city, String state, String otherState, String zipCode, String country, String done)
        { 

            expectedUsername = firstName;
            try
            {
                if ("X".Equals(done, StringComparison.InvariantCultureIgnoreCase))
                {
                    Console.WriteLine("User already registered.");
                    WriteStatusToExcel(3,"Failed");
                    return; 
                }
                driver.FindElement(By.Id("Customer_LoginEmail")).SendKeys(email);
                driver.FindElement(By.Id("l-Customer_Password")).SendKeys(password);
                driver.FindElement(By.Id("l-Customer_VerifyPassword")).SendKeys(confirmPassword);

                driver.FindElement(By.Id("l-Customer_ShipFirstName")).SendKeys(firstName);
                driver.FindElement(By.Id("l-Customer_ShipLastName")).SendKeys(lastName);
                driver.FindElement(By.Id("Customer_ShipEmail")).SendKeys(emailAdd);
                driver.FindElement(By.Id("l-Customer_ShipPhone")).SendKeys(phoneNo);
                driver.FindElement(By.Id("l-Customer_ShipFax")).SendKeys(faxNo);
                driver.FindElement(By.Id("l-Customer_ShipCompany")).SendKeys(company);
                driver.FindElement(By.Id("l-Customer_ShipAddress1")).SendKeys(address);
                driver.FindElement(By.Id("l-Customer_ShipCity")).SendKeys(city);
                // Find the dropdown element
                IWebElement dropdownState = driver.FindElement(By.Id("Customer_ShipStateSelect"));
                SelectElement selectState = new SelectElement(dropdownState);
                selectState.SelectByText(state);
                driver.FindElement(By.Id("l-Customer_ShipState")).SendKeys(otherState);
                driver.FindElement(By.Id("l-Customer_ShipZip")).SendKeys(zipCode);
                // Find the dropdown element
                IWebElement dropdownCountry = driver.FindElement(By.Id("Customer_ShipCountry"));
                SelectElement selectCountry = new SelectElement(dropdownCountry);
                selectCountry.SelectByText(country);

                //click on save:
                driver.FindElement(By.XPath("//*[@id=\"shipping_fields\"]/div[14]/div/input")).Click();

                WriteStatusToExcel(3, "Passed");
            }
            catch (Exception e)
            {
                WriteStatusToExcel(3,"Failed");
                Console.WriteLine(e);
                throw;
            }
        }

        [Test, Order(2)]
        public void usernameDisplayed()
        {
            try
            {
                //write done 
                WriteDoneToExcel(2);
                IWebElement messageElement = driver.FindElement(By.ClassName("message-success"));
                String messageText = messageElement.Text;
                // Extract the username from the message
                string[] words = messageText.Split(' ');
                string username = words[1]; // Username is the second word
                                            // Remove any unwanted characters from the username
                username = Regex.Replace(username, @"[^a-zA-Z0-9]", ""); // Remove any non-alphanumeric characters
                username = username.TrimEnd('.'); // Remove any trailing period

                // Expected username
                Console.WriteLine("Expected username: " + expectedUsername); ;

                // Compare the extracted username with the expected username
                Assert.AreEqual(username, expectedUsername, "Actual username does not match expected username");
                WriteStatusToExcel(4, "Passed");
            }
            catch(Exception e)
            {
                WriteStatusToExcel(4, "Failed");
                Console.WriteLine(e);
                throw;
            }

        }

        [Test, Order(3)]
        public void logoutFunction()
        {
            try
            {
                //Navigate to MyAccount page:
                driver.FindElement(By.XPath("//*[@id=\"js-header\"]/div[1]/div/div[3]")).Click();
                Thread.Sleep(2000);
                //logout from the account:
                driver.FindElement(By.XPath("//*[@id=\"js-main-content\"]/div/div[1]/div/div/p/a")).Click();
                //wait for 5 sec.
                Thread.Sleep(5000);
                WriteStatusToExcel(5, "Passed");
            }
            catch(Exception e)
            {
                WriteStatusToExcel(5, "Failed");
                Console.WriteLine(e);
                throw;
            }
        }

        [Test, Order(4), TestCaseSource(nameof(ReadExcel2))]
        public void loginInformation(String email, String password)
        {
            try
            {
                ////Navigate to Account page:
                driver.FindElement(By.XPath("//*[@id=\"js-header\"]/div[1]/div/div[3]")).Click();
                driver.FindElement(By.Id("l-Customer_LoginEmail")).SendKeys(email);
                driver.FindElement(By.Id("l-Customer_Password")).SendKeys(password);
                driver.FindElement(By.Id("l-Customer_Password")).SendKeys(Keys.Enter);
                WriteStatusToExcel(6, "Passed");
            }
            catch(Exception e)
            {
                WriteStatusToExcel(6, "Failed");
                Console.WriteLine(e);
                throw;
            }
        }

         // Close the browser
        private void CloseBrowser()
        {
            driver.Quit();
        }

            public static IEnumerable<object[]> ReadExcel()
        {
            //create worksheet object
            using (ExcelPackage package = new ExcelPackage(new FileInfo("dataa.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int rowCount = worksheet.Dimension.End.Row;
                Console.WriteLine("rowCount: " + rowCount);
                for(int i = 2; i < rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),
                        worksheet.Cells[i,3].Value?.ToString().Trim(),
                        worksheet.Cells[i,4].Value?.ToString().Trim(),
                        worksheet.Cells[i,5].Value?.ToString().Trim(),
                        worksheet.Cells[i,6].Value?.ToString().Trim(),
                        worksheet.Cells[i,7].Value?.ToString().Trim(),
                        worksheet.Cells[i,8].Value?.ToString().Trim(),
                        worksheet.Cells[i,9].Value?.ToString().Trim(),
                        worksheet.Cells[i,10].Value?.ToString().Trim(),
                        worksheet.Cells[i,11].Value?.ToString().Trim(),
                        worksheet.Cells[i,12].Value?.ToString().Trim(),
                        worksheet.Cells[i,13].Value?.ToString().Trim(),
                        worksheet.Cells[i,14].Value?.ToString().Trim(),
                        worksheet.Cells[i,15].Value?.ToString().Trim(),
                        worksheet.Cells[i,16].Value?.ToString().Trim(),

                    };
                }
            }
           
        }
       
        public static IEnumerable<object[]> ReadExcel2()
        {
            //create worksheet object
            using (ExcelPackage package = new ExcelPackage(new FileInfo("dataa.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int rowCount = worksheet.Dimension.End.Row;
                Console.WriteLine("rowCount: " + rowCount);
                for (int i = 2; i < rowCount; i++)
                {
                    yield return new object[]
                    {
                        worksheet.Cells[i,1].Value?.ToString().Trim(),
                        worksheet.Cells[i,2].Value?.ToString().Trim(),

                    };
                }
            }

        }
        public void WriteDoneToExcel(int rowIndex)
        {
            // Load existing workbook
            FileInfo file = new FileInfo(@"C:\Users\Msys\source\repos\SeleniumCsharp\SeleniumCsharp\dataa.xlsx");
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets["Sheet1"];

                // Write "X" to the "done" column of the specified row index
                ExcelRange cell = sheet.Cells[rowIndex, 16]; // 16th column corresponds to "P" column in Excel
                cell.Value = "X";

                // Save changes to the workbook
                package.Save();
            }
        }

        public void WriteStatusToExcel(int rowIndex,String status)
        {
            // Load existing workbook
            FileInfo file = new FileInfo(@"C:\Users\Msys\source\repos\SeleniumCsharp\SeleniumCsharp\dataa.xlsx");
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets["testcase2"];

                // Write "Status" to the "Status" column of the specified row index
                ExcelRange cell = sheet.Cells[rowIndex, 2]; 
                cell.Value = status;

                // Save changes to the workbook
                package.Save();
            }
        }

        [OneTimeTearDown]
        public void Close()
        {
            Thread.Sleep(5000);
            driver.Dispose();
        }
    }
}
