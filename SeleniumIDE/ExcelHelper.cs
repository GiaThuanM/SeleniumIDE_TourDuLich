using Microsoft.VisualStudio.TestPlatform.ObjectModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumIDE
{
    public class ExcelHelper
    {
        public static IEnumerable<object[]> GetTestDataFromExcel(string worksheetName)
        {
            List<object[]> testCases = new List<object[]>();
            string filePath = @"C:\Users\Admin\Desktop\HK2-N3\Software Quality Assurance\LT\testscript.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];
                int startRow = worksheet.Dimension.Start.Row + 1;//Bỏ dòng đầu tiêu đề
                int endRow = worksheet.Dimension.End.Row;
                for (int row = startRow; row <= endRow; row++)
                {
                    if (worksheetName.Equals("Login") || worksheetName.Equals("XSS"))
                    {
                        string email = worksheet.Cells[row, 1].Text;
                        string password = worksheet.Cells[row, 2].Text;

                        testCases.Add(new object[] { email, password });
                    }
                    else if (worksheetName.Equals("Đặt tour"))
                    {
                        string soNguoiLon = worksheet.Cells[row, 1].Text;
                        string soTreEm = worksheet.Cells[row, 2].Text;
                        string phuongThucThanhToan = worksheet.Cells[row, 3].Text;

                        testCases.Add(new object[] { soNguoiLon, soTreEm, phuongThucThanhToan });
                    }
                    else if (worksheetName.Equals("SQL Injection"))
                    {
                        string hoTen = worksheet.Cells[row, 1].Text;
                        string matKhauXacNhan = worksheet.Cells[row, 2].Text;

                        testCases.Add(new object[] { hoTen, matKhauXacNhan });
                    }
                    else if (worksheetName.Equals("URL Manipulation"))
                    {
                        string url = worksheet.Cells[row, 1].Text;

                        testCases.Add(new object[] { url });
                    }
                    else if (worksheetName.Equals("Thêm sản phẩm tour"))
                    {
                        string idSPTour = worksheet.Cells[row, 1].Text;
                        string tenSPTour = worksheet.Cells[row, 2].Text;
                        string giaNguoiLon = worksheet.Cells[row, 3].Text;
                        string ngayKhoiHanh = worksheet.Cells[row, 4].Text;
                        string ngayKetThuc = worksheet.Cells[row, 5].Text;
                        string moTa = worksheet.Cells[row, 6].Text;
                        string diemTapTrung = worksheet.Cells[row, 7].Text;
                        string diemDen = worksheet.Cells[row, 8].Text;
                        string soNguoi = worksheet.Cells[row, 9].Text;
                        string hinhAnh = worksheet.Cells[row, 10].Text;
                        string giaTreEm = worksheet.Cells[row, 11].Text;
                        string tenNV = worksheet.Cells[row, 12].Text;
                        string tenTour = worksheet.Cells[row, 13].Text;

                        testCases.Add(new object[] { idSPTour, tenSPTour, giaNguoiLon, ngayKhoiHanh, ngayKetThuc, moTa, diemTapTrung, diemDen, soNguoi, hinhAnh, giaTreEm, tenNV, tenTour });
                    }
                }
            }
            return testCases;
        }
        public static void WriteResultToExcel(string actualResult, string worksheetName, int rowIndex)
        {
            string filePath = @"C:\Users\Admin\Desktop\HK2-N3\Software Quality Assurance\LT\testscript.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];
                int currentRow = rowIndex + 1;//Bỏ dòng đầu tiêu đề
                int column = 0; // expected result
                if (worksheetName.Equals("Login") || worksheetName.Equals("XSS") || worksheetName.Equals("SQL Injection"))
                {
                    column = 3;
                }
                else if (worksheetName.Equals("Đặt tour"))
                {
                    column = 4;
                }
                else if (worksheetName.Equals("URL Manipulation"))
                {
                    column = 2;
                }
                else if (worksheetName.Equals("Thêm sản phẩm tour"))
                {
                    column = 14;
                }
                worksheet.Cells[currentRow, column + 1].Value = actualResult;
                string actualResultExcel = worksheet.Cells[currentRow, column + 1].Text.ToLower();
                string expectedResultExcel = worksheet.Cells[currentRow, column].Text.ToLower();
                string status = expectedResultExcel.Equals(actualResultExcel) ? "Passed" : "Failed";
                worksheet.Cells[currentRow, column + 2].Value = status;
                package.Save();
            }
        }
    }
}
