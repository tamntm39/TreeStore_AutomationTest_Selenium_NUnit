using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using ClosedXML.Excel;
using System.IO.Packaging;
using OpenQA.Selenium.Support.UI;

namespace Login_Test
{
    public class Login_Test
    {
        private IWebDriver driver;
        private static readonly string excelPath = @"D:\\BDCLPM\\QLDH_Test\\QLDH_DataTest.xlsx";

        [SetUp]
        public void Setup()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddExcludedArgument("enable-automation");
            options.AddAdditionalOption("useAutomationExtension", false);
            options.AddArgument("--disable-blink-features=AutomationControlled");

            driver = new ChromeDriver(options);
            driver.Manage().Window.Maximize();
        }

        public static IEnumerable<TestCaseData> GetLoginData()
        {
            using (var workbook = new XLWorkbook(excelPath))
            {
                var worksheet = workbook.Worksheet(6);
                var rows = worksheet.RangeUsed().RowsUsed();

                foreach (var row in rows.Skip(1)) // Bỏ qua dòng tiêu đề
                {
                    string username = row.Cell(3).GetValue<string>();
                    string password = row.Cell(4).GetValue<string>();
                    string expected = row.Cell(5).GetValue<string>();
                    yield return new TestCaseData(username, password, expected);
                }
            } 
        }



        [Test, TestCaseSource(nameof(GetLoginData))]
        public void Login_With_Excel_Data(string username, string password, string expected)
        {
            driver.Navigate().GoToUrl("http://localhost:4200/login");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            var emailInput = driver.FindElement(By.CssSelector("input[type='email']"));
            emailInput.Clear();
            emailInput.SendKeys(" ");
            emailInput.SendKeys(Keys.Backspace);
            emailInput.SendKeys(username);

            // Nhập mật khẩu
            var passwordInput = driver.FindElement(By.CssSelector("input[type='password']"));
            passwordInput.Clear();
            passwordInput.SendKeys(" ");
            passwordInput.SendKeys(Keys.Backspace);
            passwordInput.SendKeys(password);

            // Click nút Login
            driver.FindElement(By.CssSelector("button")).Click();

            // Chờ hộp thoại xuất hiện (tối đa 5 giây)
            bool isSuccess = false;
            for (int i = 0; i < 10; i++) // Kiểm tra 10 lần, mỗi lần 500ms (tổng 5 giây)
            {
                Thread.Sleep(500);
                var popups = driver.FindElements(By.ClassName("swal2-popup"));
                if (popups.Count > 0) // Hộp thoại xuất hiện
                {
                    isSuccess = true;
                    break;
                }
            }

            if (isSuccess)
            {
                // Click nút OK để đóng hộp thoại nếu có
                var okButtons = driver.FindElements(By.ClassName("swal2-confirm"));
                if (okButtons.Count > 0)
                {
                    okButtons[0].Click();
                }
            }

            // Ghi kết quả vào Excel
            WriteResultToExcel(username, isSuccess ? "Pass" : "Fail");

            // So sánh với kết quả mong đợi
            Assert.AreEqual(expected, isSuccess ? "Pass" : "Fail", $"Đăng nhập với {username} thất bại!");
        }

        public void WriteResultToExcel(string username, string result)
        {
            var workbook = new XLWorkbook(excelPath);
            var worksheet = workbook.Worksheet(4);
            var rows = worksheet.RangeUsed().RowsUsed().Skip(1);

            foreach (var row in rows)
            {
                if (row.Cell(3).GetValue<string>() == username)
                {
                    row.Cell(6).Value = result; 
                    break;
                }
            }

            workbook.Save(); // Lưu file Excel
            workbook.Dispose();
        }

        [TearDown]
        public void Cleanup()
        {
            driver.Quit();
        }
    }
}
