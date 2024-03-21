using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Engine.ClientProtocol;
using NUnit.Framework.Interfaces;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using static System.Runtime.InteropServices.JavaScript.JSType;
namespace SeleniumIDE
{
    public class Tests
    {
        private IWebDriver driver;
        private int counterTestcase = 0;
        
        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
        }
        [TearDown]
        public void TearDown()
        {
            // Clean up WebDriver (called once after all tests are executed)
            driver.Quit();
            driver.Dispose();
        }
        private static IEnumerable<object[]> GetTestDataFromExcel(string worksheetName)
        {
            List<object[]> testCases = new List<object[]>();
            string filePath = @"C:\Users\Admin\Desktop\HK2-N3\Software Quality Assurance\LT\testscript.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];
                int startRow = worksheet.Dimension.Start.Row + 1;//Bỏ dòng đầu tiêu đề
                int endRow = worksheet.Dimension.End.Row;
                for (int row = startRow; row <= endRow; row++)
                {
                    string email = worksheet.Cells[row, 1].Text;
                    string password = worksheet.Cells[row, 2].Text;

                    testCases.Add(new object[] { email, password });
                }
            }

            return testCases;
        }
        private void WriteResultToExcel(string actualResult, string worksheetName, int rowIndex)
        {
            string filePath = @"C:\Users\Admin\Desktop\HK2-N3\Software Quality Assurance\LT\testscript.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];
                int currentRow = rowIndex + 1;//Bỏ dòng đầu tiêu đề
                int column = 0;
                if (worksheetName == "Login")
                {
                    column = 3;
                }
                worksheet.Cells[currentRow, column + 1].Value = actualResult;
                string actualResultExcel = worksheet.Cells[currentRow, column + 1].Text;
                string expectedResultExcel = worksheet.Cells[currentRow, column].Text;
                string status = (expectedResultExcel.ToLower() == actualResultExcel.ToLower()) ? "Passed" : "Failed";
                worksheet.Cells[currentRow, column + 2].Value = status;
                package.Save();
            }
        }

        [Test]
        [TestCaseSource(nameof(GetTestDataFromExcel), new object[] { "Login" })]
        public void Test1Login(string email, string password)
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

            } else if (driver.Url.Contains("Login"))
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
            else if(loginMessage != null)
            {
                actualResult = loginMessage.Text;
            }else if (validationMessageEmail != null)
            {
                actualResult = validationMessageEmail;
            }else if(validationMessagePassword != null)
            {
                actualResult = validationMessagePassword;
            }
            else
            {
                actualResult = "Xảy ra kết quả ngoài ý muốn";
            }
            counterTestcase++;
            WriteResultToExcel(actualResult, "Login" , counterTestcase);
            Assert.Pass(actualResult);
        }
        private void Login(string email,string password)
        {
            driver.Navigate().GoToUrl("https://localhost:44385/");

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            wait.Until(ExpectedConditions.InvisibilityOfElementLocated(By.ClassName("loader")));

            IWebElement btnLogin = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".ms-auto .nav-link")));
            btnLogin.Click();

            IWebElement emailElement = driver.FindElement(By.CssSelector(".infield:nth-child(3) > #Mail_KH"));
            emailElement.SendKeys(email);
            IWebElement passwordElement = driver.FindElement(By.CssSelector(".infield:nth-child(4) > #MatKhau"));
            passwordElement.SendKeys(password);
            driver.FindElement(By.CssSelector("button:nth-child(6)")).Click();
        }
        private void ThreadSleepOneSecond()
        {
            Thread.Sleep(1000);
        }
        private void ThreadSleepTwoSecond()
        {
            Thread.Sleep(2000);
        }
        private void ThreadSleepThreeSecond()
        {
            Thread.Sleep(3000);
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
                if(logged.ToLower() == "XIN CHÀO, TRẦN TUẤN ANH".ToLower())
                {
                    return true;
                }
                return false;
            }
        }
        private void MoveToDatTour()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".me-auto > .nav-item:nth-child(2) > .nav-link"))).Click();

            Random random = new Random();
            int randomOneToTwentyThree = 0;
            do
            {
                randomOneToTwentyThree = random.Next(1, 24);
            } while (randomOneToTwentyThree == 14);
            int randomOneToNine = random.Next(1, 10);
            string randomString = randomOneToTwentyThree.ToString("00000");

            try
            {
                IWebElement detailTourBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector($".col:nth-child({randomOneToNine}) .btn-more")));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", detailTourBtn);
                Thread.Sleep(500);
                detailTourBtn.Click();
            }
            catch(ElementClickInterceptedException ex)
            {
                driver.Navigate().GoToUrl($"https://localhost:44385/Home/ChiTietTour/{randomString}");
            }
            try
            {
                IWebElement moveToDatTour = wait.Until(ExpectedConditions.ElementToBeClickable(By.LinkText("Đặt tour")));
                ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", moveToDatTour);
                Thread.Sleep(500);
                moveToDatTour.Click();
            }
            catch (ElementClickInterceptedException ex)
            {
                driver.Navigate().GoToUrl($"https://localhost:44385/Home/DatTour/{randomString}");
            }
        }
        [Test]
        [TestCase(1,0, "VNPAY")]
        public void DatTour(int soLuongNguoiLon, int soLuongTreEm, string methodPayment)
        {
            if(LoginSuccessful())
            {
                MoveToDatTour();
                IWebElement inputNguoiLon = driver.FindElement(By.Id("nguoilondat"));
                inputNguoiLon.Clear();
                inputNguoiLon.SendKeys(soLuongNguoiLon.ToString());

                IWebElement inputTreEm = driver.FindElement(By.Id("treemdat"));
                inputTreEm.Clear();
                inputTreEm.SendKeys(soLuongTreEm.ToString());

                var dropdown = driver.FindElement(By.Id("payment_method"));
                var selectDropdown = new SelectElement(dropdown);

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

                if(methodPayment == "VNPAY")
                {
                    selectDropdown.SelectByText(methodPayment);
                    IWebElement vnpayBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#vnpayBtn")));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", vnpayBtn);
                    ThreadSleepOneSecond();
                    vnpayBtn.Click();
                    ThreadSleepOneSecond();
                    string cardNumber = "9704198526191432198";
                    string cardName = "NGUYEN VAN A";
                    DateTime releaseDate = new DateTime(2015,07,20);
                    DatTourVNPAY(cardNumber, cardName, releaseDate);
                }
                else if(methodPayment == "Tiền Mặt")
                {
                    selectDropdown.SelectByText(methodPayment);
                    IWebElement cashBtn = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#cashBtn")));
                    ((IJavaScriptExecutor)driver).ExecuteScript("arguments[0].scrollIntoView(true);", cashBtn);
                    ThreadSleepOneSecond();
                    cashBtn.Click();
                    ThreadSleepOneSecond();
                    DatTourTienMat();
                }
                else
                {
                    Assert.Fail("Phương thức thanh toán không hợp lệ");
                }
                // Done
                // Còn VNPAY
            }
            else
            {
                Assert.Fail("Đăng nhập thất bại");
            }
        }

        private void DatTourVNPAY(string cardNumber,string cardName,DateTime releaseDate)
        {
            string month = releaseDate.Month.ToString();
            string year = releaseDate.Year.ToString().Substring(2);

            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".list-method-item:nth-child(2) .title"))).Click();

            ThreadSleepOneSecond();

            driver.FindElement(By.Id("searchPayMethod2")).SendKeys("NCB");
            ThreadSleepOneSecond();
            driver.FindElement(By.Id("NCB")).Click();

            ThreadSleepOneSecond();

            driver.FindElement(By.Id("card_number_mask")).SendKeys(cardNumber);
            driver.FindElement(By.Id("cardHolder")).SendKeys(cardName);
            driver.FindElement(By.Id("cardDate")).SendKeys(month+year);
            driver.FindElement(By.Id("btnContinue")).Click();

            ThreadSleepOneSecond();

            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("btnAgree"))).Click();

            ThreadSleepOneSecond();

            string otpNumber = "123456";
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("otpvalue"))).SendKeys(otpNumber);
            driver.FindElement(By.Id("btnConfirm")).Click();

            ThreadSleepOneSecond();

            string txtIdHoaDon = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("p:nth-child(3)"))).Text;
                
            txtIdHoaDon = txtIdHoaDon.Substring(30, 2);

            string txtStatusHoaDon = driver.FindElement(By.CssSelector("h2")).Text;

            driver.FindElement(By.LinkText("Quay về lịch sử đặt tour")).Click();

            ThreadSleepOneSecond();

            string txtIdHoaDonMoi = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".container-md:nth-child(8) .fw-semibold:nth-child(2)"))).Text;
            txtIdHoaDonMoi = txtIdHoaDonMoi.Substring(txtIdHoaDonMoi.Length - 2);

            string txtStatusThanhToan = driver.FindElement(By.CssSelector(".container-md:nth-child(8) .btn-done")).Text;

            ThreadSleepThreeSecond();

            if (txtIdHoaDon == txtIdHoaDonMoi && txtStatusHoaDon == "Thanh toán thành công" && txtStatusThanhToan == "Đã Thanh Toán")
            {
                Assert.Pass($"{txtIdHoaDon} = {txtIdHoaDonMoi}, trạng thái thanh toán VNPAY: {txtStatusHoaDon}, xác nhận thanh toán: {txtStatusThanhToan}");
            }
            else
            {
                Assert.Fail($"{txtIdHoaDon}, {txtIdHoaDonMoi}, {txtStatusHoaDon}, {txtStatusThanhToan}");
            }
        }

        private void DatTourTienMat()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));

            string txtIdHoaDon = driver.FindElement(By.Id("idHoaDon")).GetAttribute("placeholder");

            driver.FindElement(By.LinkText("Quay về trang lịch sử đặt tour")).Click();

            string txtIdHoaDonMoi = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector(".container-md:nth-child(8) .fw-semibold:nth-child(2)"))).Text;
            txtIdHoaDonMoi = txtIdHoaDonMoi.Substring(txtIdHoaDonMoi.Length - 2);

            ThreadSleepThreeSecond();

            if (txtIdHoaDon == txtIdHoaDonMoi)
            {
                Assert.Pass($"{txtIdHoaDon} = {txtIdHoaDonMoi}");
            }
            else
            {
                Assert.Fail(txtIdHoaDon);
            }
        }
    }
}