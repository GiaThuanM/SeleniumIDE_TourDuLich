using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Engine.ClientProtocol;
using NUnit.Framework.Interfaces;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Data.SqlClient;
using System.IO;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using static System.Runtime.InteropServices.JavaScript.JSType;
namespace SeleniumIDE
{
    public class LoginTest
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        private int counterTestcase = 0;

        [SetUp]
        public void Setup()
        {
            new WebDriverManager.DriverManager().SetUpDriver(new ChromeConfig()); // Tự động tải phiên bản đúng
            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        }

        [TearDown]
        public void TearDown()
        {
            driver.Quit();
            driver.Dispose();
        }
        [Test]
        [TestCaseSource(typeof(ExcelHelper), nameof(ExcelHelper.GetTestDataFromExcel), new object[] { "Login" })]
        public void LoginTestMain(string email, string password)
        {
            Login(email, password);

            IWebElement? logged = null;
            IWebElement? loginMessage = null;
            string validationMessageEmail = string.Empty;
            string validationMessagePassword = string.Empty;
            if (driver.Url.Contains("LoginAndRegister"))
            {
                IWebElement validationEmailElement = driver.FindElement(By.Name("Mail_KH"));
                bool isInputEmailEmpty = string.IsNullOrEmpty(validationEmailElement.GetAttribute("value"));
                if (isInputEmailEmpty)
                {
                    validationMessageEmail = validationEmailElement.GetAttribute("validationMessage");
                }
                else
                {
                    validationMessagePassword = driver.FindElement(By.Name("MatKhau")).GetAttribute("validationMessage");
                }

            }
            else if (driver.Url.Contains("Login"))
            {
                loginMessage = driver.FindElement(By.CssSelector(".field-validation-error"));
            }
            else
            {
                logged = driver.FindElement(By.XPath("//*[@id=\"myNavbar\"]/ul[2]/li[1]/a"));
            }
            string actualResult;
            if (logged != null)
            {
                actualResult = logged.Text;
            }
            else if (loginMessage != null)
            {
                actualResult = loginMessage.Text;
            }
            else if (validationMessageEmail != null)
            {
                actualResult = validationMessageEmail;
            }
            else if (validationMessagePassword != null)
            {
                actualResult = validationMessagePassword;
            }
            else
            {
                actualResult = "Xảy ra kết quả ngoài ý muốn";
            }
            counterTestcase++;
            ExcelHelper.WriteResultToExcel(actualResult, "Login", counterTestcase);
            Assert.Pass(actualResult);
        }
        private void Login(string email, string password)
        {
            driver.Navigate().GoToUrl("https://localhost:44385/");

            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector(".loader")));

            IWebElement btnLogin = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".ms-auto .nav-link")));
            btnLogin.Click();

            IWebElement emailElement = driver.FindElement(By.CssSelector(".infield:nth-child(3) > #Mail_KH"));
            emailElement.SendKeys(email);
            IWebElement passwordElement = driver.FindElement(By.CssSelector(".infield:nth-child(4) > #MatKhau"));
            passwordElement.SendKeys(password);
            driver.FindElement(By.CssSelector("button:nth-child(6)")).Click();
        }
    }
}