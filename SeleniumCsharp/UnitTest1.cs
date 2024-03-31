using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using SeleniumCsharp.BaseClass;
using System;

namespace SeleniumCsharp
{
    [TestFixture]
    public class Tests: BaseTest
    {
        [SetUp]
        public void Setup()
        {
            // Initialize your driver here
            // Example:
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
                //enter valid email and password:
                driver.FindElement(By.Id("l-Customer_LoginEmail")).SendKeys("test@testqa2024.com");
                driver.FindElement(By.Id("l-Customer_Password")).SendKeys("wtx@91187");
                driver.FindElement(By.Id("l-Customer_Password")).SendKeys(Keys.Enter);
                WriteStatusToExcel(2, "Passed");
            }
            catch(Exception e)
            {
                WriteStatusToExcel(2, "Failed");
                Console.WriteLine(e);
                throw;
            }
        }
        private readonly Random random = new Random();
        [DatapointSource]
        String[] brandList = { "Hobo", "Vhernier", "Baccarat" };

        [Test, Order(1)]
        public void searchRandomBrand()
        {
            try
            {
               // Choose a random brand from brandList array
                int randomIndex = random.Next(brandList.Length);
                string randomBrand = brandList[randomIndex];
                Console.WriteLine(randomBrand);

                driver.FindElement(By.Id("l-desktop-search")).SendKeys(randomBrand);
                driver.FindElement(By.Id("l-desktop-search")).SendKeys(Keys.Enter);
                Thread.Sleep(1000);

                IWebElement actualTextElement = driver.FindElement(By.XPath("//*[@id=\"js-category-header-title\"]"));
                String text = actualTextElement.Text;
                Console.WriteLine(text);

                // Extract the first word
                string[] words = text.Split(' ');
                string firstWordBrand = words[0];
                Console.WriteLine("First word: " + firstWordBrand);
                Assert.AreEqual(firstWordBrand, randomBrand, "Actual text doesn't match expected text");


                Thread.Sleep(2000);
                if (randomBrand.Equals("vhernier", StringComparison.InvariantCultureIgnoreCase))
                {
                    var element = driver.FindElement(By.XPath("//*[@id=\"js-main-content\"]/div[2]/div[1]/div/div[4]/x-search-app"));

                    var shadowroot = element.GetShadowRoot();
                    var xcardelement = shadowroot.FindElement(By.CssSelector("x-card"));
                    xcardelement.Click();
                }
                else
                {
                    var element = driver.FindElement(By.XPath("//*[@id=\"js-main-content\"]/div[2]/div[1]/div/div[3]/x-search-app"));

                    var shadowroot = element.GetShadowRoot();
                    var xcardelement = shadowroot.FindElement(By.CssSelector("x-card"));
                    xcardelement.Click();
                }
                Thread.Sleep(1000);
                WriteStatusToExcel(3, "Passed");

            }
            catch(Exception e)
            {
                WriteStatusToExcel(3, "Failed");
                Console.WriteLine(e);
                throw;
            }
        }

        [Test, Order(2)]
        public void addItemBag()
        {
            try
            {
                driver.FindElement(By.Id("js-add-to-cart")).SendKeys(Keys.Enter);
                Thread.Sleep(4000);
                driver.Navigate().GoToUrl("https://www.borsheims.com/vhernier-mon-jeu-titanium-link");
                Thread.Sleep(1000);
                driver.FindElement(By.Id("js-add-to-cart")).SendKeys(Keys.Enter);


                driver.FindElement(By.XPath("//*[@id=\"js-header-contents\"]/a/img")).Click();
                driver.FindElement(By.XPath("//*[@id=\"js-header\"]/div[1]/div/div[4]")).Click();
                driver.FindElement(By.XPath("//*[@id=\"js-mini-basket-container\"]/div[2]/div/a")).SendKeys(Keys.Enter);

                //increase the QT.:
                driver.FindElement(By.XPath("//*[@id=\"js-main-content\"]/div/div[1]/div[2]/div/div[2]/div[3]/div[1]/div[3]/div[4]/form/div/span[2]")).Click();

                // Find all parent divs containing the product details
                IList<IWebElement> productDivs = driver.FindElements(By.CssSelector("div.flex.rigid.row.ai-center.basket-product-row"));

                // Create a 2D array to store item details including a header row
                String[][] valueToWrite = new String[productDivs.Count + 1][];

                // Add header row
                valueToWrite[0] = new String[] { "Product Name", "Quantity", "Price", "Total" };

                // Iterate through each product div
                for (int i = 0; i < productDivs.Count; i++)
                {
                    // Find the element containing the item name
                    IWebElement itemNameElement = productDivs[i].FindElement(By.CssSelector("h4.lineitem-name a"));

                    // Get the text of the item name
                    String itemName = itemNameElement.Text.Trim();
                    Console.WriteLine("item name: " + itemName);

                    // Find the element containing the price
                    IWebElement priceElement = productDivs[i].FindElement(By.CssSelector("div.lineitem-has-price p"));

                    // Get the text of the price
                    String price = priceElement.Text.Trim();
                    Console.WriteLine("price: " + price);

                    // Find the element containing the total price
                    IWebElement totalPriceElement = productDivs[i].FindElement(By.CssSelector("div.lineitem-has-subtotal p"));

                    // Get the text of the total price
                    String totalPrice = totalPriceElement.Text.Trim();
                    Console.WriteLine("total price: " + totalPrice);

                    // Find the element containing the quantity input field
                    IWebElement quantityInputElement = productDivs[i].FindElement(By.CssSelector("input#l-quantity"));

                    // Get the value attribute of the input element to get the quantity
                    String quantity = quantityInputElement.GetAttribute("value").Trim();
                    Console.WriteLine("Quantity: " + quantity);

                    // Write the item details to the array
                    valueToWrite[i + 1] = new String[] { itemName, quantity, price, totalPrice };
                }
                WriteStatusToExcel(4, "Passed");
            }
            catch(Exception e)
            {
                WriteStatusToExcel(4, "Failed");
                Console.WriteLine(e);
                throw;
            }

        }

        public void WriteStatusToExcel(int rowIndex, String status)
        {
            // Load existing workbook
            FileInfo file = new FileInfo(@"C:\Users\Msys\source\repos\SeleniumCsharp\SeleniumCsharp\dataa.xlsx");
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets["testcase1"];

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