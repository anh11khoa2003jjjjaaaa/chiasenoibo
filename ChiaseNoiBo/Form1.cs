using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Drive.v3;
using Guna.UI2.WinForms;

namespace ChiaseNoiBo
{
    public partial class Form1 : Form
    {
        private readonly string fileId;
        private readonly string fileName;

        public Form1(string fileId, string fileName)
        {
            InitializeComponent();
            this.fileId = fileId;
            this.fileName = fileName;
            this.Text = $"Xem File: {fileName}"; // Hiển thị tên file trên tiêu đề
        }

        private async void Form1_Load(object sender, EventArgs e)
        {
            await LoadExcelDataAsync(); // Gọi hàm tải dữ liệu khi form mở
        }

        private async Task LoadExcelDataAsync()
        {
            if (string.IsNullOrEmpty(fileId))
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show("Không có file nào được chọn!", "Lỗi");
                return;
            }

            try
            {
                var stream = await DownloadFileFromGoogleDriveAsync(fileId);

                if (stream == null || stream.Length < 100) // Kiểm tra file có hợp lệ không
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("File tải về không hợp lệ hoặc có kích thước nhỏ bất thường!", "Lỗi");
                    return;
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(stream))
                {
                    if (package.Workbook.Worksheets.Count == 0)
                    {
                        guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                        guna2MessageDialog1.Show("File Excel không chứa dữ liệu hợp lệ.", "Lỗi");
                        return;
                    }

                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    DataTable dt = ConvertWorksheetToDataTable(worksheet);

                    guna2DataGridView1.DataSource = dt; // Gán dữ liệu lên DataGridView
                }
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show($"Lỗi khi tải file Excel: {ex.Message}", "Lỗi");
            }
        }

        /// <summary>
        /// Tải file từ Google Drive và trả về MemoryStream
        /// </summary>
        private async Task<MemoryStream> DownloadFileFromGoogleDriveAsync(string fileId)
        {
            try
            {
                var service = GoogleDriveHelper.GetDriveService();
                var request = service.Files.Get(fileId);

                MemoryStream stream = new MemoryStream();
                await request.DownloadAsync(stream);
                stream.Position = 0; // Đưa stream về đầu để đọc dữ liệu
                return stream;
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show($"Lỗi khi tải file từ Google Drive: {ex.Message}", "Lỗi");
                return null;
            }
        }

        /// <summary>
        /// Chuyển đổi Worksheet thành DataTable
        /// </summary>
        private DataTable ConvertWorksheetToDataTable(ExcelWorksheet worksheet)
        {
            DataTable dt = new DataTable();

            // Đọc header (hàng đầu tiên)
            for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
            {
                dt.Columns.Add(worksheet.Cells[1, col].Text);
            }

            // Đọc dữ liệu
            for (int row = worksheet.Dimension.Start.Row + 1; row <= worksheet.Dimension.End.Row; row++)
            {
                DataRow newRow = dt.NewRow();
                for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                {
                    newRow[col - 1] = worksheet.Cells[row, col].Text;
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }
    }
}
