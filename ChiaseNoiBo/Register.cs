using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Upload;
using Google.Apis.Download;
using Guna.UI2.WinForms;
using System.Security.Cryptography;

namespace ChiaseNoiBo
{
    public partial class Register : Form
    {
        private const string EXCEL_FILENAME = "DanhSachTaiKhoan.xlsx";
        private readonly Color ERROR_COLOR = Color.Red;
        private readonly Color PLACEHOLDER_COLOR = Color.Gray;
        public Register()
        {
            InitializeComponent();
            SetupPlaceholders();
        }
        private void SetupPlaceholders()
        {
            txt_name.PlaceholderText = "Nhập họ và tên";
            txt_email.PlaceholderText = "Nhập email";
            txt_password.PlaceholderText = "Nhập mật khẩu";
            txt_password.PlaceholderText = "Nhập lại mật khẩu";
            txt_name.PlaceholderForeColor = PLACEHOLDER_COLOR;
            txt_email.PlaceholderForeColor = PLACEHOLDER_COLOR;
            txt_password.PlaceholderForeColor = PLACEHOLDER_COLOR;
            confirm_password.PlaceholderForeColor = PLACEHOLDER_COLOR;
        }
        private bool AddUserToLocalExcel(string filePath, string name, string email, string password)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var fileInfo = new FileInfo(filePath);

            using (var package = new ExcelPackage(fileInfo))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault() ??
                               package.Workbook.Worksheets.Add("Accounts");

                // Tạo headers nếu chưa có
                if (worksheet.Dimension == null)
                {
                    worksheet.Cells["A1"].Value = "STT";
                    worksheet.Cells["B1"].Value = "Họ và tên";
                    worksheet.Cells["C1"].Value = "Email";
                    worksheet.Cells["D1"].Value = "Password";
                    worksheet.Cells["E1"].Value = "Xác thực";
                    worksheet.Cells["F1"].Value = "Vai trò";
                }

                // Kiểm tra email đã tồn tại chưa
                for (int row = 2; row <= worksheet.Dimension?.End.Row; row++)
                {
                    if (worksheet.Cells[row, 3]?.Value?.ToString() == email)
                        return false;
                }

                // Thêm bản ghi mới
                int newRow = (worksheet.Dimension?.End.Row ?? 1) + 1;
                worksheet.Cells[newRow, 1].Value = GetNextId(worksheet);
                worksheet.Cells[newRow, 2].Value = name;
                worksheet.Cells[newRow, 3].Value = email;
                worksheet.Cells[newRow, 4].Value = password;
                worksheet.Cells[newRow, 5].Value = false;
                worksheet.Cells[newRow, 6].Value = "User";

                package.Save();
                return true;
            }
        }
        private async Task SyncToGoogleDrive(string localFilePath)
        {
           
                var service = GoogleDriveHelper.GetDriveService();
                string fileName = Path.GetFileName(localFilePath);

                // Upload file mới
                await UploadOrUpdateFile(service, localFilePath, fileName);
        }

        private async Task UploadOrUpdateFile(DriveService service, string filePath, string fileName)
        {
            // 1. Tìm file cũ trên Drive
            var listRequest = service.Files.List();
            listRequest.Q = $"name='{fileName}' and '{GoogleDriveHelper.FolderId}' in parents and trashed=false";
            listRequest.Fields = "files(id)";
            var fileList = await listRequest.ExecuteAsync();

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                if (fileList.Files.Any())
                {
                    // 2. Ghi đè lên file cũ
                    var fileId = fileList.Files[0].Id;
                    var updateRequest = service.Files.Update(null, fileId, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    await updateRequest.UploadAsync();
                }
                else
                {
                    // 3. Nếu chưa có → tạo mới
                    var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                    {
                        Name = fileName,
                        Parents = new List<string> { GoogleDriveHelper.FolderId }
                    };

                    var createRequest = service.Files.Create(fileMetadata, stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                    createRequest.Fields = "id";
                    await createRequest.UploadAsync();
                }
            }
        }


        // Helper methods
        private int GetNextId(ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null) return 1;

            int maxId = 0;
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                if (int.TryParse(worksheet.Cells[row, 1]?.Value?.ToString(), out int currentId))
                {
                    maxId = Math.Max(maxId, currentId);
                }
            }
            return maxId + 1;
        }

     
        private void ResetForm()
        {
            txt_name.Text = "";
            txt_email.Text = "";
            txt_password.Text = "";
            confirm_password.Text = "";
            SetupPlaceholders();
        }

        private void ShowSuccess(string message)
        {
            MessageBox.Show(message, "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowError(string message)
        {
            MessageBox.Show(message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

    
       
        private void linkLogin_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new Login().Show();
            this.Hide();
        }




        private bool isValidate()
        {
            bool isValid = true;

            // Validate họ tên
            if (string.IsNullOrWhiteSpace(txt_name.Text))
            {
                SetError(txt_name, "Họ tên không được để trống!");
                isValid = false;
            }

            // Validate email
            string email = txt_email.Text.Trim();
            if (string.IsNullOrWhiteSpace(email))
            {
                SetError(txt_email, "Email không được để trống!");
                isValid = false;
            }
            else
            {
                try
                {
                    var addr = new MailAddress(email);
                }
                catch
                {
                    SetError(txt_email, "Email không hợp lệ!");
                    isValid = false;
                }
            }

            // Validate mật khẩu
            string password = txt_password.Text.Trim();
            if (string.IsNullOrWhiteSpace(password)||string.IsNullOrWhiteSpace(confirm_password.Text))
            {
                SetError(txt_password, "Mật khẩu không được để trống!");
                SetError(confirm_password, "Mật khẩu không được để trống!");
                isValid = false;
            }
            else if (password.EndsWith("_"))
            {
                SetError(txt_password, "Mật khẩu không được kết thúc bằng dấu gạch dưới (_)");
                isValid = false;
            }

            // Validate xác nhận mật khẩu
            if (confirm_password.Text.Trim() != password)
            {
                SetError(confirm_password, "Mật khẩu không đúng");
                isValid = false;
            }

            return isValid;
        }

        private void SetError(Guna2TextBox textBox, string errorMessage)
        {
            textBox.Text = "";
            textBox.PlaceholderText = errorMessage;
            textBox.PlaceholderForeColor = Color.Red;
        }

        private async Task<bool> CheckEmailExistsInExcelAsync(string emailToCheck)
        {
            var service = GoogleDriveHelper.GetDriveService();

            // Tìm file trên Google Drive
            var listRequest = service.Files.List();
            listRequest.Q = $"name='{EXCEL_FILENAME}' and '{GoogleDriveHelper.FolderId}' in parents and trashed=false";
            listRequest.Fields = "files(id)";
            var result = await listRequest.ExecuteAsync();

            if (result.Files == null || result.Files.Count == 0)
                throw new Exception("Không tìm thấy file DanhSachTaiKhoan.xlsx trên Google Drive");

            string fileId = result.Files[0].Id;
            string localTempPath = Path.Combine(Path.GetTempPath(), EXCEL_FILENAME);

            // Tải file về tạm thời
            using (var stream = new FileStream(localTempPath, FileMode.Create, FileAccess.Write))
            {
                await service.Files.Get(fileId).DownloadAsync(stream);
            }

            // Đọc và kiểm tra email
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(localTempPath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null) return false;

                int rowCount = worksheet.Dimension?.End.Row ?? 0;

                for (int row = 2; row <= rowCount; row++) // bắt đầu từ dòng 2 vì dòng 1 là header
                {
                    var emailInCell = worksheet.Cells[row, 3]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(emailInCell) && emailInCell.Equals(emailToCheck, StringComparison.OrdinalIgnoreCase))
                    {
                        return true; // email đã tồn tại
                    }
                }
            }

            return false; // không tìm thấy email
        }
        public string HashPasswordSHA256(string password)
        {
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] inputBytes = Encoding.UTF8.GetBytes(password);
                byte[] hashBytes = sha256.ComputeHash(inputBytes);

                // Convert bytes to hex string
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                    sb.Append(b.ToString("x2")); // "x2" to get hex format

                return sb.ToString();
            }
        }
        private async void guna2Button1_Click(object sender, EventArgs e)
        {
            if (!isValidate()) return;

            string name = txt_name.Text.Trim();
            string email = txt_email.Text.Trim();
            string password = txt_password.Text.Trim() + "_phcn";
            string hashedPassword = HashPasswordSHA256(password);
            if (await CheckEmailExistsInExcelAsync(email))
            {
                ShowMessage("Email đã tồn tại trong hệ thống!", "Cảnh báo",MessageDialogIcon.Warning);
                return;
            }

            string localPath = Path.Combine(Application.StartupPath, EXCEL_FILENAME);
                bool isNewRecordAdded = AddUserToLocalExcel(localPath, name, email, hashedPassword);

                if (isNewRecordAdded)
                {
                    await SyncToGoogleDrive(localPath); // Xóa file cũ và upload mới

                    ShowMessage("Đăng ký thành công!", "Thông báo", MessageDialogIcon.Information);
                    ResetForm();
               Login login=new Login();
                login.Show();
                this.Hide();
                }
                else
                {
                ShowMessage("Đăng ký thất bại! Vui lòng nhập lại", "Thông báo lỗi", MessageDialogIcon.Error);
                }
            
          
        }


        private void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            
                var dialog = new Guna.UI2.WinForms.Guna2MessageDialog
                {
                    Buttons = Guna.UI2.WinForms.MessageDialogButtons.OK,
                    Icon = icon,
                    Style = Guna.UI2.WinForms.MessageDialogStyle.Dark,
                    Caption = title,
                    Text = message,
                    Parent = this
                };
                dialog.Show();
            
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Login login = new Login();
            login.Show();
            this.Hide();
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}