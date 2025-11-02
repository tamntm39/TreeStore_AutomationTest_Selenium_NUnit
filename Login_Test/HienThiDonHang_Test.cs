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

            driver.FindElement(By.CssSelector("input[type='email']")).SendKeys("minhtam39@gmail.com");

            driver.FindElement(By.CssSelector("input[type='password']")).SendKeys("YourSecurePassword");

            driver.FindElement(By.CssSelector("button")).Click();

            //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(5));
            //wait.Until(d => d.Url.Contains("http://localhost:4200/manages/order/order-list"));
        }

        [Test]
        public void QLDH_HienThienDonHang_Test()
        {
            Login();
            Thread.Sleep(2000);

            driver.Navigate().GoToUrl("http://localhost:4200/manages/order/order-list");
            Thread.Sleep(2000);
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