using System;
using System.IO;
using System.Threading;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using ClosedXML.Excel;

namespace QLDH_Test
{
    [TestFixture]
    public class RevenuePageTests
    {
        private IWebDriver driver;
        private string testFilePath = @"D:\BDCLPM\QLDH_Test\\QLDH_DataTest.xlsx";

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://localhost:4200/manages/revenue");
        }

        public void Login()
        {
            driver.Navigate().GoToUrl("http://localhost:4200/login");
            driver.FindElement(By.CssSelector("input[type='email']")).SendKeys("minhtam39@gmail.com");

            driver.FindElement(By.CssSelector("input[type='password']")).SendKeys("YourSecurePassword");

            driver.FindElement(By.CssSelector("button")).Click();
        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
        }

        private string ExtractNumber(string input)
        {
            return new string(input.Where(char.IsDigit).ToArray());
        }

        [Test]
        public void VerifyRevenuePageData()
        {
            Login();

            Thread.Sleep(2000);
            driver.Navigate().GoToUrl("http://localhost:4200/manages/revenue");

            Thread.Sleep(2000);

            var expectedData = ReadTestDataFromExcel(testFilePath);

            string totalOrders = ExtractNumber(driver.FindElement(By.XPath("(//h6[contains(@class, 'card-number')])[1]")).Text);
            string totalProducts = ExtractNumber(driver.FindElement(By.XPath("(//h6[contains(@class, 'card-number')])[2]")).Text);
            string totalCustomers = ExtractNumber(driver.FindElement(By.XPath("(//h6[contains(@class, 'card-number')])[3]")).Text);
            string totalReviews = ExtractNumber(driver.FindElement(By.XPath("(//h6[contains(@class, 'card-number')])[4]")).Text);

            bool isPass = expectedData.TotalOrders == totalOrders &&
                          expectedData.TotalProducts == totalProducts &&
                          expectedData.TotalCustomers == totalCustomers &&
                          expectedData.TotalReviews == totalReviews;


            string result = isPass ? "Pass" : "Fail";
            WriteTestResultToExcel(testFilePath, result);

            if (!isPass)
            {
                string errorMessage = "Dữ liệu trên trang không khớp với dữ liệu kiểm thử:\n";

               
                if (expectedData.TotalOrders != totalOrders)
                    errorMessage += $"❌ Tổng đơn hàng không khớp: {totalOrders} (thực tế) != {expectedData.TotalOrders} (mong đợi)\n";
                if (expectedData.TotalProducts != totalProducts)
                    errorMessage += $"❌ Tổng sản phẩm không khớp: {totalProducts} (thực tế) != {expectedData.TotalProducts} (mong đợi)\n";
                if (expectedData.TotalCustomers != totalCustomers)
                    errorMessage += $"❌ Tổng khách hàng không khớp: {totalCustomers} (thực tế) != {expectedData.TotalCustomers} (mong đợi)\n";
                if (expectedData.TotalReviews != totalReviews)
                    errorMessage += $"❌ Tổng đánh giá không khớp: {totalReviews} (thực tế) != {expectedData.TotalReviews} (mong đợi)\n";

                Assert.Fail(errorMessage);
            }
            else
            {
                Assert.Pass("Tất cả dữ liệu trên trang khớp với dữ liệu kiểm thử.");
            }
        }

        private (string TotalOrders, string TotalProducts, string TotalCustomers, string TotalReviews) ReadTestDataFromExcel(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(4);
                return (
                    worksheet.Cell(2, 1).GetString(),
                    worksheet.Cell(2, 2).GetString(),
                    worksheet.Cell(2, 3).GetString(),
                    worksheet.Cell(2, 4).GetString()
                );
            }
        }

        private void WriteTestResultToExcel(string filePath, string result)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(4);
                worksheet.Cell(2, 5).Value = result;
                workbook.Save();
            }
        }
    }
}