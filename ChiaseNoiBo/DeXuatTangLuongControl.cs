using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChiaseNoiBo
{
    public partial class DeXuatTangLuongControl : UserControl
    {
        public DeXuatTangLuongControl()
        {
            InitializeComponent();
        }
        public string Lanhdao => txt_lanhdao.Text;
        public string HoTen => txt_name.Text;
        public string Phongban => txt_bophan.Text;
        public string Vitri => txt_chucvu.Text;
        public string LyDo => richTextBox1_lydo.Text;
        public double Salary_current => (double)guna2NumericUpDown1_hientai.Value;
        public double Salary_expected => (double)guna2NumericUpDown2_mongmuon.Value;
        public string TuNgay => guna2DateTimePicker1.Value.ToShortDateString();
        public DateTime DenNgay => DateTime.Now;
        public static async Task<string> DownloadFileFromGoogleDriveAsync(string fileId, string savePath)
        {
            var url = $"https://drive.google.com/uc?export=download&id={fileId}";
            using (var client = new HttpClient())
            {
                var response = await client.GetAsync(url);
                if (response.IsSuccessStatusCode)
                {
                    using (var fs = new FileStream(savePath, FileMode.Create, FileAccess.Write))
                    {
                        await response.Content.CopyToAsync(fs);
                    }
                    return savePath;
                }
                else
                {
                    throw new Exception("Tải file từ Google Drive thất bại.");
                }
            }
        }
        public static void ReplacePlaceholders(string filePath, Dictionary<string, string> replacements)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                var body = wordDoc.MainDocumentPart.Document.Body;

                foreach (var text in body.Descendants<Text>())
                {
                    foreach (var kvp in replacements)
                    {
                        if (text.Text.Contains(kvp.Key))
                        {
                            text.Text = text.Text.Replace(kvp.Key, kvp.Value);
                        }
                    }
                }

                wordDoc.MainDocumentPart.Document.Save();
            }
        }

        private void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            var mainForm = this.FindForm() as Home1;

            if (mainForm != null)
            {
                mainForm.ShowMessage(message, title, icon);
            }
            else
            {
                // Nếu không phải là Home1 hoặc không tìm thấy form
                MessageBox.Show(message, title, MessageBoxButtons.OK, GetMessageBoxIcon(icon));
            }
        }


        private MessageBoxIcon GetMessageBoxIcon(Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            switch (icon)
            {
                case Guna.UI2.WinForms.MessageDialogIcon.Error:
                    return MessageBoxIcon.Error;
                case Guna.UI2.WinForms.MessageDialogIcon.Information:
                    return MessageBoxIcon.Information;
                case Guna.UI2.WinForms.MessageDialogIcon.Warning:
                    return MessageBoxIcon.Warning;
                case Guna.UI2.WinForms.MessageDialogIcon.Question:
                    return MessageBoxIcon.Question;
                default:
                    return MessageBoxIcon.None;
            }
        }
   
    
        private bool ValidateFormInputsV2(out string errorMessage)
        {
            StringBuilder sb = new StringBuilder();

            if (string.IsNullOrWhiteSpace(HoTen))
                sb.AppendLine("⚠️ Vui lòng nhập Họ tên.");

            if (string.IsNullOrWhiteSpace(Phongban))
                sb.AppendLine("⚠️ Vui lòng nhập Phòng ban.");

            if (string.IsNullOrWhiteSpace(Vitri))
                sb.AppendLine("⚠️ Vui lòng nhập Vị trí công việc.");

            if (string.IsNullOrWhiteSpace(Lanhdao))
                sb.AppendLine("⚠️ Vui lòng nhập tên Lãnh đạo.");

            if (string.IsNullOrWhiteSpace(LyDo))
                sb.AppendLine("⚠️ Vui lòng nhập Lý do.");

            if (Salary_current <= 0)
                sb.AppendLine("⚠️ Lương hiện tại phải lớn hơn 0.");

            if (Salary_expected <= 0)
                sb.AppendLine("⚠️ Lương mong muốn phải lớn hơn 0.");

            if (Salary_expected < Salary_current)
                sb.AppendLine("⚠️ Lương mong muốn không được nhỏ hơn lương hiện tại.");

            if (guna2DateTimePicker1.Value.Date < DateTime.Today)
                sb.AppendLine("⚠️ Ngày bắt đầu không được nhỏ hơn ngày hôm nay.");

            errorMessage = sb.ToString();
            return string.IsNullOrEmpty(errorMessage);
        
    }

        private void DeXuatTangLuongControl_Load(object sender, EventArgs e)
        {
            CultureInfo viCulture = new CultureInfo("vi-VN");

            // Đặt mặc định cho toàn app (nên dùng)
            CultureInfo.DefaultThreadCurrentCulture = viCulture;
            CultureInfo.DefaultThreadCurrentUICulture = viCulture;

            guna2DateTimePicker1.Format = DateTimePickerFormat.Custom;
            guna2DateTimePicker1.CustomFormat = "'Ngày' dd 'tháng' MM 'năm' yyyy";
        }

        private async void guna2Button2_luu_Click(object sender, EventArgs e)
        {
            if (!ValidateFormInputsV2(out string error))
            {
                ShowMessage(error, "Lỗi nhập liệu", Guna.UI2.WinForms.MessageDialogIcon.Warning);
                return;
            }
            try
            {
                string fileId = "171SJndW98HPjyNzy4--8L3ELyBPo556N";
                string tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DonDeXuatTangLuong.docx");

                await DownloadFileFromGoogleDriveAsync(fileId, tempPath);

                // Tạo dictionary chứa các giá trị cần thay thế
                var replacements = new Dictionary<string, string>
{
    { "{{name}}", HoTen },
    { "{{department}}", Phongban },
    { "{{position}}", Vitri },
    { "{{current_salary}}", Salary_current.ToString("N0") + " VND" },
    { "{{expected_salary}}", Salary_expected.ToString("N0") + " VND" },
    { "{{effective_date}}", TuNgay },
    { "{{reason}}", LyDo },
    { "{{recipient}}", Lanhdao },
    { "{{current_date}}", DateTime.Now.ToString("dd/MM/yyyy") }
};

                // Thay thế nội dung trong file
                ReplacePlaceholders(tempPath, replacements);

                // Chọn nơi lưu file
                SaveFileDialog dialog = new SaveFileDialog
                {
                    Filter = "Word Documents (*.docx)|*.docx",
                    FileName = "DonDeXuatTangLuong_Filled.docx"
                };

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    File.Copy(tempPath, dialog.FileName, true);
                    ShowMessage("Lưu thành công!", "Thông báo", Guna.UI2.WinForms.MessageDialogIcon.Information);
                    ResetData();
                }
            }
            catch (Exception ex)
            {
                ShowMessage("Lỗi: " + ex.Message, "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
            }
        }
        public void ResetData()
        {
            // Xóa nội dung các TextBox
            txt_lanhdao.Text = string.Empty;
            txt_name.Text = string.Empty;
            txt_bophan.Text = string.Empty;
            txt_chucvu.Text = string.Empty;

            // Xóa nội dung RichTextBox
            richTextBox1_lydo.Text = string.Empty;

            // Reset giá trị lương về 0 hoặc giá trị mặc định tùy bạn
            guna2NumericUpDown1_hientai.Value = 0;
            guna2NumericUpDown2_mongmuon.Value = 0;

            // Reset ngày về ngày hiện tại
            guna2DateTimePicker1.Value = DateTime.Now;
        }

       
            private void guna2Button1_huy_Click(object sender, EventArgs e)
            {
                // Tìm form chính
                var mainForm = this.FindForm() as Home1;

                // Tạo dialog xác nhận
                Guna.UI2.WinForms.Guna2MessageDialog dialog = new Guna.UI2.WinForms.Guna2MessageDialog
                {
                    Parent = mainForm, // đảm bảo dialog nằm giữa form cha
                    Buttons = Guna.UI2.WinForms.MessageDialogButtons.YesNo,
                    Caption = "Xác nhận",
                    Text = "Bạn có chắc muốn hủy thao tác không?",
                    Icon = Guna.UI2.WinForms.MessageDialogIcon.Question,
                    Style = Guna.UI2.WinForms.MessageDialogStyle.Light
                };

                // Hiển thị dialog và lấy kết quả
                var result = dialog.Show();

                if (result == DialogResult.Yes)
                {
                    // Gọi Reset nếu người dùng đồng ý hủy
                    ResetData();
                }
            }


            
        }

    }

