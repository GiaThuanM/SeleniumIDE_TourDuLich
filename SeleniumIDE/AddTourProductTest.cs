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
using OpenQA.Selenium.Interactions;

namespace SeleniumIDE
{
    public class AddTourProductTest
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
            driver.Dispose();
            driver.Quit();
        }
        private bool LoginAdminSuccess()
        {
            driver.Navigate().GoToUrl("https://localhost:44385/Logging/LoginAdmin");
            wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("Mail_NV"))).SendKeys("admin");
            IWebElement password = driver.FindElement(By.Id("MatKhau"));
            password.SendKeys("admin");
            password.SendKeys(Keys.Enter);
            if (driver.Url.Contains("/NHANVIENs/GetData"))
            {
                return true;
            }
            else if (driver.Url.Contains("/Logging/LoginAdmin"))
            {
                return false;
            }
            else
            {
                throw new Exception("LoginAdminSuccess lỗi ngoài hệ thống");
            }
        }
        [Test]
        [TestCaseSource(typeof(ExcelHelper), nameof(ExcelHelper.GetTestDataFromExcel), new object[] { "Thêm sản phẩm tour" })]
        public void AddTourProduct(string idSPTour, string tenSPTour, string giaNguoiLon, string ngayKhoiHanh, string ngayKetThuc, string moTa, string diemTapTrung, string diemDen, string soNguoi, string hinhAnh, string giaTreEm, string tenNV, string tenTour)
        {
            if (LoginAdminSuccess())
            {
                Actions action = new Actions(driver);
                IWebElement navbar = driver.FindElement(By.XPath("/html/body/div[1]/div[2]/ul/li[4]/a/span"));
                action.MoveToElement(navbar).MoveToElement(driver.FindElement(By.CssSelector("li:nth-child(4) span"))).Click().Build().Perform();

                wait.Until(ExpectedConditions.ElementExists(By.LinkText("Thêm mới sản phẩm tour"))).Click();

                wait.Until(ExpectedConditions.ElementExists(By.Id("ID_SPTour"))).SendKeys(idSPTour);
                wait.Until(ExpectedConditions.ElementExists(By.Id("TenSPTour"))).SendKeys(tenSPTour);
                wait.Until(ExpectedConditions.ElementExists(By.Id("GiaNguoiLon"))).SendKeys(giaNguoiLon);
                wait.Until(ExpectedConditions.ElementExists(By.Id("NgayKhoiHanh"))).SendKeys(ngayKhoiHanh);
                wait.Until(ExpectedConditions.ElementExists(By.Id("NgayKetThuc"))).SendKeys(ngayKetThuc);
                wait.Until(ExpectedConditions.ElementExists(By.Id("MoTa"))).SendKeys(moTa);
                wait.Until(ExpectedConditions.ElementExists(By.Id("DiemTapTrung"))).SendKeys(diemTapTrung);
                wait.Until(ExpectedConditions.ElementExists(By.Id("DiemDen"))).SendKeys(diemDen);
                wait.Until(ExpectedConditions.ElementExists(By.Id("SoNguoi"))).SendKeys(soNguoi);
                wait.Until(ExpectedConditions.ElementExists(By.Id("HinhAnh"))).SendKeys(hinhAnh);
                wait.Until(ExpectedConditions.ElementExists(By.Id("GiaTreEm"))).SendKeys(giaTreEm);

                SelectElement selectedIdNV = new SelectElement(driver.FindElement(By.Id("ID_NV")));
                selectedIdNV.SelectByText(tenNV);

                SelectElement selectedIdTour = new SelectElement(driver.FindElement(By.Id("ID_TOUR")));
                selectedIdTour.SelectByText(tenTour);
                wait.Until(ExpectedConditions.ElementExists(By.CssSelector(".btn"))).SendKeys(Keys.Enter);
                if (QueryDatabaseCompareAddProductTour(idSPTour))
                {
                    counterTestcase++;
                    string actualResult = "Thêm sản phẩm tour thành công";
                    ExcelHelper.WriteResultToExcel(actualResult, "Thêm sản phẩm tour", counterTestcase);
                    Assert.Pass("Thêm thành công");
                }
                else
                {
                    Assert.Fail("Thêm thất bại");
                }
            }
            else
            {
                Assert.Fail("Đăng nhập admin thất bại");
            }
        }
        private bool QueryDatabaseCompareAddProductTour(string idSPTour)
        {
            using (SqlConnection sqlConnection = new SqlConnection(connectString))
            {
                sqlConnection.Open();
                string query = "SELECT COUNT(*) FROM SPTOUR WHERE ID_SPTour = @idSPTour";
                SqlCommand command = new SqlCommand(query, sqlConnection);
                command.Parameters.AddWithValue("@idSPTour", idSPTour);

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
