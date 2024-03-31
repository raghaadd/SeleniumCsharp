using NUnit.Framework.Interfaces;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Execution;
using OfficeOpenXml;
using SeleniumCsharp.BaseClass;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NUnit.Framework;
using NUnit.Framework.Internal;
using NUnit.Framework.Internal.Commands;

namespace SeleniumCsharp
{
    [TestFixture]
    class UnitTestMain
    {
       // int rowIndex = 0;
        [Test, Order(0)]
        public void readExecutionColumn()
        {
            int executionColumnIndex = 2;
            using (ExcelPackage package = new ExcelPackage(new FileInfo("dataa.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["testCases"];
                int rowCount = worksheet.Dimension.End.Row;
                Console.WriteLine("rowCount: " + rowCount);
                for (int i = 2; i <= rowCount; i++)
                {
                    String executionStatus = worksheet.Cells[i, executionColumnIndex].Value?.ToString().Trim();
                    Console.WriteLine("Test Case #" + i + " Execution Status: " );
                    if("X".Equals(executionStatus, StringComparison.InvariantCultureIgnoreCase))
                    {
                        executeTestCase(i);
                    }
                }
            }
        }
      
        public void executeTestCase(int rowIndex)
        {
            switch(rowIndex)
            {
                case 2:
                    Tests tests = new Tests();
                    tests.Setup();
                    tests.NavigateWebsite();
                    tests.searchRandomBrand();
                    tests.addItemBag();
                    tests.Close();
                    break;
                case 3:
                    UnitTest2 unitTest2 = new UnitTest2();
                    unitTest2.Setup();
                    unitTest2.NavigateWebsite();
                    String email="";
                    String password="";
                    foreach (object[] data in UnitTest2.ReadExcel())
                    {
                        // Extract data from the object[] array
                        email = data[0]?.ToString();
                        password = data[1]?.ToString();
                        string confirmPassword = data[2]?.ToString();
                        string firstName = data[3]?.ToString();
                        string lastName = data[4]?.ToString();
                        string emailAdd = data[5]?.ToString();
                        string phoneNo = data[6]?.ToString();
                        string faxNo = data[7]?.ToString();
                        string company = data[8]?.ToString();
                        string address = data[9]?.ToString();
                        string city = data[10]?.ToString();
                        string state = data[11]?.ToString();
                        string otherState = data[12]?.ToString();
                        string zipCode = data[13]?.ToString();
                        string country = data[14]?.ToString();
                        string done = data[15]?.ToString();

                        
                        unitTest2.registerInformation(email, password, confirmPassword, firstName, lastName, emailAdd, phoneNo,
                            faxNo, company, address, city, state, otherState, zipCode, country, done);
                        
                    }
                    unitTest2.usernameDisplayed();
                    unitTest2.logoutFunction();
                    unitTest2.loginInformation(email, password);
                    unitTest2.Close();

                    break;

                default:
                    Console.WriteLine("Invalid test case number: " + rowIndex);
                    break;

            }
        }
    }
}
