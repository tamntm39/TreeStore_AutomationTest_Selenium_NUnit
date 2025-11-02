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
namespace OrderDetailTest
{
    public class ChiTietDonHang
    {
        public IWebDriver driver;
        public static readonly string excelPath = @"D:\BDCLPM\QLDH_Test\QLDH_DataTest.xlsx";

        [SetUp]
        public void Setup()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            driver = new ChromeDriver(options);
        }

        public void Login()
        {
            driver.Navigate().GoToUrl("http://localhost:4200/login");


            driver.FindElement(By.CssSelector("input[type='email']")).SendKeys("minhtam39@gmail.com");

            driver.FindElement(By.CssSelector("input[type='password']")).SendKeys("YourSecurePassword");

            driver.FindElement(By.CssSelector("button")).Click();

           
        }



        public static IEnumerable<TestCaseData> GetOrderData()
        {
            var workbook = new XLWorkbook(excelPath);
            var worksheet = workbook.Worksheet(3);
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Bỏ qua tiêu đề

            foreach (var row in rows)
            {
                string orderId = row.Cell(2).GetValue<string>();
                string customerName = row.Cell(3).GetValue<string>();
                string address = row.Cell(4).GetValue<string>();
                string totalAmount = row.Cell(5).GetValue<string>();
                string status = row.Cell(6).GetValue<string>();
                yield return new TestCaseData(orderId, customerName, address, totalAmount, status);
            }
        }

        [Test, TestCaseSource(nameof(GetOrderData))]
        public void VerifyOrderDetails(string orderId, string customerName, string address, string totalAmount, string status)
        {
            Login();
            Thread.Sleep(2000);

            driver.Navigate().GoToUrl("http://localhost:4200/manages/order/order-list");
            Thread.Sleep(2000);

            bool orderFound = false;


            string xpath = $"//tbody/tr[td[2][normalize-space()='{orderId}']]";
            var orderRow = driver.FindElements(By.XPath(xpath));

            if (orderRow.Count == 0)
            {
                Console.WriteLine($"Không tìm thấy đơn hàng có ID: {orderId}");
                throw new AssertionException("Order not found");
            }

            var detailButton = orderRow[0].FindElement(By.XPath(".//td[last()]//button"));
            Actions actions = new Actions(driver);
            actions.MoveToElement(detailButton).Click().Perform();
            Thread.Sleep(2000);

            string displayedCustomerName = driver.FindElement(By.CssSelector(".order-details-container h5:nth-of-type(1)")).Text;
            string displayedAddress = driver.FindElement(By.CssSelector(".order-details-container h5:nth-of-type(2)")).Text;
            string displayedStatus = driver.FindElement(By.CssSelector(".order-details-container h5:nth-of-type(3)")).Text;

            bool isMatch = displayedCustomerName.Contains(customerName) &&
                           displayedAddress.Contains(address) &&
                           displayedStatus.Contains(status);

            WriteResultToExcel(orderId, isMatch ? "Pass" : "Fail");

            Assert.IsTrue(isMatch, "Thông tin đơn hàng không khớp với dữ liệu trong file Excel!");
        }

        public void WriteResultToExcel(string orderId, string result)
        {
            var workbook = new XLWorkbook(excelPath);
            var worksheet = workbook.Worksheet(3);  
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

            foreach (var row in rows)
            {
                if (row.Cell(2).GetValue<string>() == orderId)
                {
                    row.Cell(7).Value = result; 
                    break;
                }
            }

            workbook.Save();
            workbook.Dispose();
        }

        [TearDown]
        public void Cleanup()
        {
            driver.Quit();
        }
    }
}
