using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using NUnit.Framework.Interfaces;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OrderStatusTest
{
    public class QLDH_Intergrated_Test
    {
        private IWebDriver driver;
        private static readonly string excelPath = @"D:\\BDCLPM\\QLDH_Test\\QLDH_DataTest.xlsx";
        private string currentOrderId = "";

        [SetUp]
        public void Setup()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            options.AddArgument("force-device-scale-factor=0.8"); 
            driver = new ChromeDriver(options);
        }

        public void Login(string url, string email = "", string password = "")
        {
        

            driver.Navigate().GoToUrl(url);
            Thread.Sleep(2000);

            if (url.Contains("4300")) 
            {
                driver.FindElement(By.Id("email")).SendKeys(email);
                driver.FindElement(By.Id("password")).SendKeys(password);
                driver.FindElement(By.CssSelector("button[type='submit']")).Click();
            }
            else
            {
                driver.FindElement(By.Id("email")).SendKeys(email);
                driver.FindElement(By.Id("password")).SendKeys(password);
                driver.FindElement(By.CssSelector("button")).Click();
            }
            Thread.Sleep(3000);
        }

        public static IEnumerable<TestCaseData> GetOrderData()
        {
            var workbook = new XLWorkbook(excelPath);
            var worksheet = workbook.Worksheet(7);
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

            foreach (var row in rows)
            {
                string testCaseId = row.Cell(1).GetValue<string>();
                string orderId = row.Cell(2).GetValue<string>();
                string currentStatus = row.Cell(3).GetValue<string>();
                string expectedStatus = row.Cell(4).GetValue<string>();

                yield return new TestCaseData(testCaseId, orderId, currentStatus, expectedStatus)
                    .SetName($"TestCase_{testCaseId}_Order_{orderId}");
            }
        }

        [Test, TestCaseSource(nameof(GetOrderData))]
        public void ChangeOrderStatus(string testCaseId, string orderId, string currentStatus, string expectedStatus)
        {
           
            Login("http://localhost:4200/login", "minhtam39@gmail.com", "YourSecurePassword");
            driver.Navigate().GoToUrl("http://localhost:4200/manages/order/order-list");
            Thread.Sleep(2000);
            currentOrderId = orderId;
            string xpath = $"//tbody/tr[td[2][normalize-space()='{orderId}']]";
            var orderRow = driver.FindElements(By.XPath(xpath));

            if (orderRow.Count == 0)
            {
                Console.WriteLine($"❌ TestCase {testCaseId}: Không tìm thấy đơn hàng {orderId}");
                WriteResultToExcel(testCaseId, orderId, "Fail");
                throw new AssertionException("Order not found");
            }

            var detailButton = orderRow[0].FindElement(By.XPath(".//td[last()]//button"));
            new Actions(driver).MoveToElement(detailButton).Click().Perform();
            Thread.Sleep(2000);

            var approveButton = driver.FindElement(By.XPath("//button[contains(@class, 'btn') and contains(@class, 'mx-1')]"));
            approveButton.Click();

            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//button[text()='Thực hiện!']")).Click();
            Thread.Sleep(2000);

            Login("http://localhost:4300/dangnhap", "taolatui@gmail.com", "123456");
            driver.Navigate().GoToUrl("http://localhost:4300/lichsudonhang");
            Thread.Sleep(2000);

            //var userOrderRow = driver.FindElements(By.XPath($"//tbody/tr[td[1][normalize-space()='{orderId}']]"));
            //if (userOrderRow.Count == 0)
            //{
            //    WriteResultToExcel(testCaseId, orderId, "Fail");
            //    throw new AssertionException("Order not found in user view");
            //}

            //string userOrderStatus = userOrderRow[0].FindElement(By.XPath("./td[3]")).Text.Trim();
            //Assert.AreEqual(expectedStatus, userOrderStatus, "Trạng thái không khớp");
            //WriteResultToExcel(testCaseId, orderId, "Pass");
        }

        [TearDown]
        public void Cleanup()
        {
            driver.Quit();
        }

        public void WriteResultToExcel(string testCaseId, string orderId, string result)
        {
            try
            {
                using (var stream = new FileStream(excelPath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                using (var workbook = new XLWorkbook(stream))
                {
                    var worksheet = workbook.Worksheet(7);
                    var row = worksheet.RangeUsed().RowsUsed().Skip(1)
                        .FirstOrDefault(r => r.Cell(1).GetValue<string>() == testCaseId && r.Cell(2).GetValue<string>() == orderId);

                    if (row != null)
                    {
                        row.Cell(5).Value = result;
                        workbook.SaveAs(stream);
                    }
                }
            }
            catch (IOException)
            {
                Console.WriteLine("❌ File Excel đang mở hoặc bị khóa");
            }
        }
    }
}