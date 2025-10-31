using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using ClosedXML.Excel;
using System.IO.Packaging;
using OpenQA.Selenium.Support.UI;

namespace QLDH_Test
{
    public class HienThiDonHang_Test
    {
        private IWebDriver driver;

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);
            driver.Navigate().GoToUrl("http://localhost:4200/manages/order/order-list");
        }

        public void Login()
        {
            driver.Navigate().GoToUrl("http://localhost:4200/login");

            //Nhập email
            driver.FindElement(By.CssSelector("input[type='email']")).SendKeys("minhtam39@gmail.com");

            // Nhập mật khẩu
            driver.FindElement(By.CssSelector("input[type='password']")).SendKeys("YourSecurePassword");

            // Click nút Login
            driver.FindElement(By.CssSelector("button")).Click();

            // Chờ chuyển hướng sang trang admin
            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            //wait.Until(d => d.Url.Contains("http://localhost:4200/manages/order/order-list"));
        }

        [Test]
        public void QLDH_HienThienDonHang_Test()
        {
            Login();
            Thread.Sleep(2000);

            // 2. Mở danh sách đơn hàng
            driver.Navigate().GoToUrl("http://localhost:4200/manages/order/order-list");
            Thread.Sleep(2000);
            // Kiểm tra danh sách có đơn hàng không
            var orderRows = driver.FindElements(By.CssSelector("table tbody tr"));
            Assert.IsTrue(orderRows.Count > 0, "Không có đơn hàng nào hiển thị!");
        }

        
        [TearDown]
        public void TearDown()
        {
            driver.Quit();
        }
    }
}