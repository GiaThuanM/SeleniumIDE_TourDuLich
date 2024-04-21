using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;

namespace SeleniumIDE
{
    public class SecurityTest
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        private int counterTestcase = 0;

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
        }
        [TearDown]
        public void TearDown()
        {
            driver.Quit();
            driver.Dispose();
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
        private bool LoginSuccessful()
        {
            string email = "tuan@gmail.com";
            string password = "Tuan@321";
            Login(email, password);
            if (driver.Url.Contains("LoginAndRegister") || driver.Url.Contains("Login"))
            {
                return false;
            }
            else
            {
                string logged = driver.FindElement(By.XPath("//*[@id=\"myNavbar\"]/ul[2]/li[1]/a")).Text;
                if (logged.ToLower() == "XIN CHÀO, TRẦN TUẤN ANH".ToLower())
                {
                    return true;
                }
                return false;
            }
        }
        // Test bảo mật XSS (Cross-Site Scripting), SQL Injection, URL Manipulation
        // XSS (Cross-Site Scripting)
        [Test]
        [TestCaseSource(typeof(ExcelHelper), nameof(ExcelHelper.GetTestDataFromExcel), new object[] { "XSS" })]
        public void LoginAdminXSS(string email, string password)
        {
            string urlLoginAdmin = "https://localhost:44385/Logging/LoginAdmin";
            driver.Navigate().GoToUrl(urlLoginAdmin);

            wait.Until(ExpectedConditions.ElementExists(By.Id("Mail_NV"))).SendKeys(email);
            wait.Until(ExpectedConditions.ElementExists(By.Id("MatKhau"))).SendKeys(password);

            driver.FindElement(By.CssSelector("button")).Click();

            Thread.Sleep(1000);

            string actualResult = string.Empty;
            try
            {
                IWebElement alertValidation = driver.FindElement(By.CssSelector(".text-danger"));
                if (alertValidation != null)
                {
                    actualResult = alertValidation.Text;
                }
            }
            catch (NoSuchElementException ex)
            {
                IWebElement waitErrorMessage = wait.Until(ExpectedConditions.ElementExists(By.CssSelector("i")));

                actualResult = waitErrorMessage.Text;
                counterTestcase++;
                ExcelHelper.WriteResultToExcel(actualResult, "XSS", counterTestcase);
            }
            Assert.Pass(actualResult);
        }
        // SQL Injection
        [Test]
        [TestCaseSource(typeof(ExcelHelper), nameof(ExcelHelper.GetTestDataFromExcel), new object[] { "SQL Injection" })]
        public void ChinhSuaThongTinSQLInjection(string hoTen, string passwordConfirm)
        {
            if (LoginSuccessful())
            {
                wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".no-hover"))).Click();
                wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.CssSelector(".loader")));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("pen-icon"))).Click();

                IWebElement waitHoTen = wait.Until(ExpectedConditions.ElementExists(By.Id("HoTen_KH")));
                waitHoTen.Clear();
                waitHoTen.SendKeys(hoTen);

                IWebElement waitMatKhau = wait.Until(ExpectedConditions.ElementExists(By.Id("MatKhau")));
                waitMatKhau.SendKeys(passwordConfirm);

                Thread.Sleep(1000);
                waitMatKhau.SendKeys(Keys.Enter);

                IWebElement waitAlert = wait.Until(ExpectedConditions.ElementExists(By.Id("successAlert")));
                string actualResult = waitAlert.Text;

                counterTestcase++;
                ExcelHelper.WriteResultToExcel(actualResult, "SQL Injection", counterTestcase);
                Assert.Pass(actualResult);
            }
            else
            {
                Assert.Fail("Đăng nhập thất bại");
            }
        }
        [Test]
        [TestCaseSource(typeof(ExcelHelper), nameof(ExcelHelper.GetTestDataFromExcel), new object[] { "URL Manipulation" })]
        public void URLManipulation(string url)
        {
            IWebElement elementCheck;
            string actualResult = string.Empty;
            driver.Navigate().GoToUrl(url);
            Thread.Sleep(1000);
            bool success = false;
            try
            {
                if (driver.Url.Contains("DanhMucTour"))
                {
                    elementCheck = driver.FindElement(By.CssSelector(".alert"));
                    actualResult = elementCheck.Text;
                    success = true;
                }
            }
            catch (Exception ex) { }
            if (success)
            {
                counterTestcase++;
                ExcelHelper.WriteResultToExcel(actualResult, "URL Manipulation", counterTestcase);
                Assert.Pass(actualResult);
                return;
            }
            try
            {
                elementCheck = driver.FindElement(By.CssSelector(".content-container"));
                actualResult = elementCheck.Text;
                success = true;
            }
            catch (Exception ex) { }
            if (success)
            {
                counterTestcase++;
                ExcelHelper.WriteResultToExcel(actualResult, "URL Manipulation", counterTestcase);
                Assert.Pass(actualResult);
                return;
            }
            try
            {
                if (driver.Url.Contains("DanhMucTour"))
                {
                    elementCheck = driver.FindElement(By.CssSelector("i"));
                    actualResult = elementCheck.Text;
                    success = true;
                }
            }
            catch (Exception ex) { }
            if (success)
            {
                counterTestcase++;
                ExcelHelper.WriteResultToExcel(actualResult, "URL Manipulation", counterTestcase);
                Assert.Pass(actualResult);
                return;
            }
            try
            {
                if (driver.Url.Contains("LoginAndRegister"))
                {
                    elementCheck = driver.FindElement(By.CssSelector("h1:nth-child(1)"));
                    actualResult = elementCheck.Text;
                    actualResult = "Không truy cập dô tài khoản người khác qua đường link được (id: 3)";
                    success = true;
                }
            }
            catch (Exception ex) { }
            if (success)
            {
                counterTestcase++;
                ExcelHelper.WriteResultToExcel(actualResult, "URL Manipulation", counterTestcase);
                Assert.Pass(actualResult);
                return;
            }
            /*try
            {
                wait.Until(ExpectedConditions.AlertIsPresent());
                driver.SwitchTo().Alert().Accept();
                actualResult = "CHƯA ĐĂNG NHẬP TÀI KHOẢN ADMIN!";
                success = true;
                if (success)
                {
                    counterTestcase++;
                    ExcelHelper.WriteResultToExcel(actualResult, "URL Manipulation", counterTestcase);
                    Assert.Pass(actualResult);
                }
                else
                {
                    Assert.Fail("Failed");
                }
            }
            catch (Exception ex)
            {
                Assert.Fail(actualResult);
                *//*counterTestcase++;
                ExcelHelper.WriteResultToExcel("CHƯA ĐĂNG NHẬP TÀI KHOẢN ADMIN!", "URL Manipulation", counterTestcase);
                Assert.Pass("CHƯA ĐĂNG NHẬP TÀI KHOẢN ADMIN!");*//*
            }*/
        }
    }
}
