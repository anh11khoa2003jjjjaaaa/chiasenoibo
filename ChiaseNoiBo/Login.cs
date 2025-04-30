using OfficeOpenXml;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChiaseNoiBo
{
    public partial class Login : Form
    {


        public Login()
        {
            InitializeComponent();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            this.Close(); // Đóng form đăng nhập
        }

        private async void guna2Button1_Click(object sender, EventArgs e)
        {
            await PerformLogin();
        }

        private async void txt_password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = true;
                e.SuppressKeyPress = true; // Ngăn tiếng "bíp"
                await PerformLogin();
            }
        }

        private async Task PerformLogin()
        {
            string email = txt_email.Text.Trim();
            string password = txt_password.Text.Trim();
            string username = await AuthenticateUser(email, password);

            if (username != null)
            {
                ShowMessage("✅ Đăng nhập thành công!", "Thông báo", Guna.UI2.WinForms.MessageDialogIcon.Information);
                this.Hide();
                new Home1(username).Show();
            }
            else
            {
                ShowMessage("❌ Tài khoản hoặc mật khẩu không chính xác!", "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
            }
        }

        private async Task<string> FindFileDanhSachTaiKhoanfromGGdrive()
        {
            try
            {
                var service = GoogleDriveHelper.GetDriveService();
                var request = service.Files.List();
                request.Q = $"'{GoogleDriveHelper.FolderId}' in parents and name = 'DanhSachTaiKhoan.xlsx'";
                request.Fields = "files(id, name)";

                var result = await request.ExecuteAsync();
                if (result.Files.Count == 0) return null;

                var fileId = result.Files[0].Id;
                string localFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "DanhSachTaiKhoan.xlsx");

                using (var stream = new MemoryStream())
                {
                    await service.Files.Get(fileId).DownloadAsync(stream);
                    File.WriteAllBytes(localFilePath, stream.ToArray());
                }
                return localFilePath;
            }
            catch (Exception ex)
            {
                ShowMessage($"Lỗi khi tải file từ Google Drive: {ex.Message}", "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
                return null;
            }
        }

        private async Task<string> AuthenticateUser(string email, string password)
        {
            string excelPath = await FindFileDanhSachTaiKhoanfromGGdrive();
            if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath)) return null;
            return CheckCredentialsFromExcel(excelPath, email, password);
        }

        private string CheckCredentialsFromExcel(string filePath, string email, string password)
        {
            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        string excelEmail = worksheet.Cells[row, 3].Text.Trim();
                        string username = worksheet.Cells[row, 2].Text.Trim();

                        if (email == excelEmail && password == "phcn@123")
                        {
                            return username;
                        }
                    }
                }
                return null;
            }
            catch (Exception ex)
            {
                ShowMessage($"Lỗi khi đọc file Excel: {ex.Message}", "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
                return null;
            }
        }

        private void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            guna2MessageDialog1.Icon = icon;
            guna2MessageDialog1.Show(message, title);
        }
    }
}