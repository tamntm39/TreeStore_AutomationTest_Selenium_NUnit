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

namespace OrderStatusTest
{
    public class ThayDoiTrangThai_Test
    {
        private IWebDriver driver;
        private static readonly string excelPath = @"D:\BDCLPM\QLDH_Test\\QLDH_DataTest.xlsx";
        private string currentOrderId = "";

        [SetUp]
        public void Setup()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            options.AddArgument("force-device-scale-factor=0.8"); 
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
            var worksheet = workbook.Worksheet(5);
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
            currentOrderId = orderId;
            Login();
            Thread.Sleep(2000);
            driver.Navigate().GoToUrl("http://localhost:4200/manages/order/order-list");
            Thread.Sleep(2000);

            string xpath = $"//tbody/tr[td[2][normalize-space()='{orderId}']]";
            var orderRow = driver.FindElements(By.XPath(xpath));

            if (orderRow.Count == 0)
            {
                Console.WriteLine($"❌ TestCase {testCaseId}: Không tìm thấy đơn hàng có ID: {orderId}");
                WriteResultToExcel(testCaseId, orderId, "Fail");
                throw new AssertionException("Order not found");
            }

            var detailButton = orderRow[0].FindElement(By.XPath(".//td[last()]//button"));
            Actions actions = new Actions(driver);
            actions.MoveToElement(detailButton).Click().Perform();
            Thread.Sleep(2000);

            string currentStatusText = driver.FindElement(By.XPath("//h5[contains(text(), 'Trạng thái đơn hàng')]")).Text.Replace("Trạng thái đơn hàng :", "").Trim();
            Assert.AreEqual(currentStatus, currentStatusText, "Trạng thái hiện tại không đúng");

            if (expectedStatus.Trim() == "Hủy")
            {
                var cancelButton = driver.FindElement(By.XPath("//button[contains(text(), 'Hủy')]"));
                cancelButton.Click();
                Thread.Sleep(2000);
                driver.FindElement(By.XPath("//button[text()='Thực hiện!']")).Click();
                Thread.Sleep(2000);

                try
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    IWebElement successPopup = wait.Until(drv => drv.FindElement(By.ClassName("swal2-popup")));

                    if (successPopup.Displayed)
                    {
                        Console.WriteLine($"✅ TestCase {testCaseId}: Hủy đơn hàng {orderId} thành công!");
                        WriteResultToExcel(testCaseId, orderId, "Pass");
                        return;
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine($"❌ TestCase {testCaseId}: Hủy đơn hàng {orderId} thất bại!");
                    WriteResultToExcel(testCaseId, orderId, "Fail");
                    throw new AssertionException("Cancel order failed");
                }
            }
            else
            {
                var approveButton = driver.FindElement(By.XPath("//button[contains(@class, 'btn') and contains(@class, 'mx-1')]"));
                approveButton.Click();
                Thread.Sleep(2000);
                driver.FindElement(By.XPath("//button[text()='Thực hiện!']")).Click();
                Thread.Sleep(2000);

                try
                {
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
                    IWebElement successPopup = wait.Until(drv => drv.FindElement(By.ClassName("swal2-popup")));

                    if (successPopup.Displayed)
                    {
                        Console.WriteLine($"✅ TestCase {testCaseId}: Đổi trạng thái đơn hàng {orderId} thành công!");
                        IWebElement okButton = driver.FindElement(By.XPath("//button[text()='OK']"));
                        okButton.Click();
                        WriteResultToExcel(testCaseId, orderId, "Pass");
                    }
                }
                catch (WebDriverTimeoutException)
                {
                    Console.WriteLine($"❌ TestCase {testCaseId}: Đổi trạng thái đơn hàng {orderId} thất bại!");
                    WriteResultToExcel(testCaseId, orderId, "Fail");
                    throw new AssertionException("Change order status failed");
                }
            }
            Thread.Sleep(2000);
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
                    var worksheet = workbook.Worksheet(5);
                    var rowsToUpdate = worksheet.RangeUsed().RowsUsed().Skip(1) 
                        .Where(row => row.Cell(1).GetValue<string>() == testCaseId
                                   && row.Cell(2).GetValue<string>() == orderId)
                        .ToList(); 

                    if (rowsToUpdate.Count > 0)
                    {
                        foreach (var row in rowsToUpdate)
                        {
                            row.Cell(5).Value = result; 
                        }

                        workbook.SaveAs(stream); 
                        Console.WriteLine($"📄 Đã ghi kết quả '{result}' cho tất cả TestCase {testCaseId}, Order {orderId}.");
                    }
                    else
                    {
                        Console.WriteLine($"⚠ Không tìm thấy TestCase {testCaseId}, Order {orderId} trong file Excel.");
                    }
                }
            }
            catch (IOException ioEx)
            {
                Console.WriteLine($"❌ File Excel có thể đang bị khóa hoặc mở: {ioEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Lỗi khi ghi Excel: {ex.Message}");
            }
        }




    }
}