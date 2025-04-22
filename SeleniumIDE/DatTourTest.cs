using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;

namespace SeleniumIDE
{
    public class DatTourTest
    {
        private IWebDriver driver;
        private WebDriverWait wait;
        private int counterTestcase = 0;
        private string connectString = "data source=.;initial catalog=QLTOUR;integrated security=True;trustservercertificate=True;MultipleActiveResultSets=True;App=EntityFramework";
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
            string email = "thuanle394@gmail.com";
            string password = "112233@Th";
            Login(email, password);
            if (driver.Url.Contains("LoginAndRegister") || driver.Url.Contains("Login"))
            {
                return false;
            }
            else
            {
                string logged = driver.FindElement(By.XPath("//*[@id=\"myNavbar\"]/ul[2]/li[1]/a")).Text;
                if (logged.ToLower() == "XIN CHÀO, LEE THUAN".ToLower())
                {
                    return true;
                }
                return false;
            }
        }
        private void MoveToDatTour()
        {
            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".me-auto > .nav-item:nth-child(2) > .nav-link"))).Click();

            Random random = new Random();
            int randomOneToNine = random.Next(1, 6);
            int randomOneToTwentyThree = random.Next(26, 30);
            string randomString = randomOneToTwentyThree.ToString("00000");
            Thread.Sleep(500);
            try
            {
                IWebElement detailTourBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath($"/html/body/div[2]/div[3]/div/div[{randomOneToNine}]/div/div/a")));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", detailTourBtn);
                Thread.Sleep(500);
                detailTourBtn.Click();
            }
            catch (ElementClickInterceptedException ex)
            {
                driver.Navigate().GoToUrl($"https://localhost:44385/Home/ChiTietTour/{randomString}");
                //throw new ElementClickInterceptedException("Error click card view item tour: ", ex);
            }
            try
            {
                IWebElement moveToDatTour = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".btn")));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", moveToDatTour);
                Thread.Sleep(500);
                moveToDatTour.Click();
            }
            catch (ElementClickInterceptedException )
            {
                driver.Navigate().GoToUrl($"https://localhost:44385/Home/DatTour/{randomString}");
                //throw new ElementClickInterceptedException("Error click button book tour: ", ex);
            }
        }
        [Test]
        [TestCaseSource(typeof(ExcelHelper), nameof(ExcelHelper.GetTestDataFromExcel), new object[] { "Đặt tour" })]
        public void DatTour(string soLuongNguoiLon, string soLuongTreEm, string methodPayment)
        {
            if (LoginSuccessful())
            {
                MoveToDatTour();
                IWebElement inputNguoiLon = driver.FindElement(By.Id("nguoilondat"));
                inputNguoiLon.Clear();
                inputNguoiLon.SendKeys(soLuongNguoiLon);

                IWebElement inputTreEm = driver.FindElement(By.Id("treemdat"));
                inputTreEm.Clear();
                inputTreEm.SendKeys(soLuongTreEm);

                var dropdown = driver.FindElement(By.Id("payment_method"));
                var selectDropdown = new SelectElement(dropdown);

                if (methodPayment == "VNPAY")
                {
                    selectDropdown.SelectByText(methodPayment);
                    IWebElement vnpayBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#vnpayBtn")));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", vnpayBtn);
                    Thread.Sleep(1000);
                    vnpayBtn.Click();
                    Thread.Sleep(1000);

                    if (driver.Url.Contains("DatTour"))
                    {
                        ValidationDatTour();
                    }
                    else
                    {
                        string cardNumber = "9704198526191432198";
                        string cardName = "NGUYEN VAN A";
                        DateTime releaseDate = new DateTime(2015, 07, 20);
                        DatTourVNPAY(cardNumber, cardName, releaseDate);
                    }
                }
                else if (methodPayment == "Tiền Mặt")
                {
                    selectDropdown.SelectByText(methodPayment);
                    IWebElement cashBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#cashBtn")));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cashBtn);
                    Thread.Sleep(1000);
                    cashBtn.Click();
                    Thread.Sleep(1000);

                    if (driver.Url.Contains("DatTour"))
                    {
                        ValidationDatTour();
                    }
                    else
                    {
                        DatTourTienMat();
                    }
                }
                else
                {
                    Assert.Fail("Phương thức thanh toán không hợp lệ");
                }
            }
            else
            {
                Assert.Fail("Đăng nhập thất bại");
            }
        }
        private void ValidationDatTour()
        {
            string messageValidationNguoiLon = string.Empty;
            string messageValidationTreEm = string.Empty;
            string messageValidationSoLuong = string.Empty;
            try
            {
                messageValidationNguoiLon = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("nguoilondat"))).GetAttribute("validationMessage");
            }
            catch (NoSuchElementException ) { }
            try
            {
                messageValidationTreEm = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("treemdat"))).GetAttribute("validationMessage");
            }
            catch (NoSuchElementException) { }
            try
            {
                IWebElement alertWarning = driver.FindElement(By.CssSelector(".alert"));
                if (alertWarning != null)
                {
                    messageValidationSoLuong = alertWarning.Text;
                }
            }
            catch (NoSuchElementException ) { }

            string actualResult;
            if (!string.IsNullOrEmpty(messageValidationNguoiLon))
            {
                actualResult = messageValidationNguoiLon;
            }
            else if (!string.IsNullOrEmpty(messageValidationTreEm))
            {
                actualResult = messageValidationTreEm;
            }
            else if (!string.IsNullOrEmpty(messageValidationSoLuong))
            {
                actualResult = messageValidationSoLuong;
            }
            else
            {
                actualResult = "Kết quả xảy ra ngoài ý muốn";
            }

            counterTestcase++;
            ExcelHelper.WriteResultToExcel(actualResult, "Đặt tour", counterTestcase);
            Assert.Pass(actualResult);
        }
        private void DatTourVNPAY(string cardNumber, string cardName, DateTime releaseDate)
        {

            string month = releaseDate.Month.ToString();
            string year = releaseDate.Year.ToString().Substring(2);

            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".list-method-item:nth-child(2) .title"))).Click();

            Thread.Sleep(1000);

            driver.FindElement(By.Id("searchPayMethod2")).SendKeys("NCB");
            Thread.Sleep(1000);
            driver.FindElement(By.Id("NCB")).Click();

            Thread.Sleep(1000);

            driver.FindElement(By.Id("card_number_mask")).SendKeys(cardNumber);
            driver.FindElement(By.Id("cardHolder")).SendKeys(cardName);
            driver.FindElement(By.Id("cardDate")).SendKeys(month + year);
            driver.FindElement(By.Id("btnContinue")).Click();

            Thread.Sleep(1000);

            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("btnAgree"))).Click();

            Thread.Sleep(1000);

            string otpNumber = "123456";
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("otpvalue"))).SendKeys(otpNumber);
            driver.FindElement(By.Id("btnConfirm")).Click();

            Thread.Sleep(1000);

            string txtIdHoaDon = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("p:nth-child(3)"))).Text;

            txtIdHoaDon = txtIdHoaDon.Substring(30, 3);

            string txtStatusHoaDon = driver.FindElement(By.CssSelector("h2")).Text;

            driver.FindElement(By.LinkText("Quay về lịch sử đặt tour")).Click();

            Thread.Sleep(1000);

            string txtIdHoaDonMoi = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".container-md:nth-child(8) .fw-semibold:nth-child(2)"))).Text;
            txtIdHoaDonMoi = txtIdHoaDonMoi.Substring(txtIdHoaDonMoi.Length - txtIdHoaDon.Length);

            string txtStatusThanhToan = driver.FindElement(By.CssSelector(".container-md:nth-child(8) .btn-done")).Text;

            int idHoaDon = int.Parse(txtIdHoaDonMoi);
            string actualResult;
            if (txtIdHoaDon == txtIdHoaDonMoi && txtStatusHoaDon == "Thanh toán thành công" && txtStatusThanhToan == "Đã Thanh Toán" && QueryDatabaseCompareDatTour(idHoaDon))
            {
                actualResult = "Đặt tour thành công và chuyển đến trang lịch sử đặt tour với thông tin tour vừa đặt";
            }
            else
            {
                actualResult = "Kết quả xảy ra ngoài ý muốn";
            }

            counterTestcase++;
            ExcelHelper.WriteResultToExcel(actualResult, "Đặt tour", counterTestcase);
            Assert.Pass(actualResult);
        }

        private void DatTourTienMat()
        {
            string txtIdHoaDon = driver.FindElement(By.Id("idHoaDon")).GetAttribute("placeholder");

            driver.FindElement(By.LinkText("Quay về trang lịch sử đặt tour")).Click();

            string txtIdHoaDonMoi = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".container-md:nth-child(8) .fw-semibold:nth-child(2)"))).Text;

            txtIdHoaDonMoi = txtIdHoaDonMoi.Substring(txtIdHoaDonMoi.Length - txtIdHoaDon.Length);

            int idHoaDon = int.Parse(txtIdHoaDonMoi);
            string actualResult;
            if (txtIdHoaDon == txtIdHoaDonMoi && QueryDatabaseCompareDatTour(idHoaDon))
            {
                actualResult = "Đặt tour thành công và chuyển đến trang lịch sử đặt tour với thông tin tour vừa đặt";
            }
            else
            {
                actualResult = "Kết quả xảy ra ngoài ý muốn";
            }

            counterTestcase++;
            ExcelHelper.WriteResultToExcel(actualResult, "Đặt tour", counterTestcase);
            Assert.Pass(actualResult);
        }
        private bool QueryDatabaseCompareDatTour(int idHoaDon)
        {
            using (SqlConnection sqlConnection = new SqlConnection(connectString))
            {
                sqlConnection.Open();
                string query = "SELECT COUNT(*) FROM HOADON WHERE ID_HoaDon = @idHoaDon";
                SqlCommand command = new SqlCommand(query, sqlConnection);
                command.Parameters.AddWithValue("@idHoaDon", idHoaDon);

                int count = (int)command.ExecuteScalar();
                if (count > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
    }
}
