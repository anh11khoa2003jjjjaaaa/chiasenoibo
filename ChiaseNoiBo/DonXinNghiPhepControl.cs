using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Guna.UI2.WinForms;

namespace ChiaseNoiBo
{
    public partial class DonXinNghiPhepControl : UserControl
    {
        

        public DonXinNghiPhepControl()
        {
            InitializeComponent();
        }

        public string Lanhdao => txt_lanhdao.Text;
        public string HoTen => txt_name.Text;
        public string Phongban => txt_bophan.Text;
        public string Vitri => txt_chucvu.Text;
        public string LyDo => richTextBox1.Text;
        public string NguoiThayThe => txt_nguoithaythe.Text;
        public string BoPhanThayThe => txt_bophanthaythe.Text;
        public string TuNgay => guna2DateTime_batdau.Value.ToShortDateString();
        public string DenNgay => guna2DateTimePicker_ketthuc.Value.ToShortDateString();

        private void label1_Click(object sender, EventArgs e)
        {
            // Xử lý sự kiện click nếu cần
        }

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

        private async void guna2Button2_luu_Click(object sender, EventArgs e)
        {
            if (!ValidateFormInputs(out string errorMessage))
            {
                ShowMessage(errorMessage, "Thông báo lỗi", Guna.UI2.WinForms.MessageDialogIcon.Warning);
                return;
            }
            try
            {
                string fileId = "1I728PpgBMMdBiTBO5hWDTyEIea3RiWyr";
                string tempPath = Path.Combine(Path.GetTempPath(), "DonXinNghiPhep.docx");

                await DownloadFileFromGoogleDriveAsync(fileId, tempPath);

                // Tạo dictionary chứa các giá trị cần thay thế
                var replacements = new Dictionary<string, string>
                {
                    { "{{name}}", HoTen },
                    { "{{department}}", Phongban },
                    { "{{position}}", Vitri },
                    { "{{from_date}}", TuNgay },
                    { "{{to_date}}", DenNgay },
                    { "{{reason}}", LyDo },
                    { "{{recipient}}", Lanhdao },
                    { "{{handover_person}}", NguoiThayThe },
                    { "{{handover_department}}", BoPhanThayThe },
                    { "{{current_date}}", DateTime.Now.ToString("dd/MM/yyyy") }
                };

                // Thay thế nội dung trong file
                ReplacePlaceholders(tempPath, replacements);

                // Chọn nơi lưu file
                SaveFileDialog dialog = new SaveFileDialog
                {
                    Filter = "Word Documents (*.docx)|*.docx",
                    FileName = "DonXinNghiPhep_Filled.docx"
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
        public void ResetData()
        {
            // Xóa dữ liệu trong các TextBox
            txt_lanhdao.Text = string.Empty;
            txt_name.Text = string.Empty;
            txt_bophan.Text = string.Empty;
            txt_chucvu.Text = string.Empty;
            txt_nguoithaythe.Text = string.Empty;
            txt_bophanthaythe.Text = string.Empty;

            // Xóa nội dung RichTextBox
            richTextBox1.Text = string.Empty;

            // Reset ngày về ngày hiện tại
            guna2DateTime_batdau.Value = DateTime.Now;
            guna2DateTimePicker_ketthuc.Value = DateTime.Now;
        }
        private void guna2Button1_huy_Click(object sender, EventArgs e)
        {
            ResetData();
        }
        private bool ValidateFormInputs(out string errorMessage)
        {
            StringBuilder sb = new StringBuilder();

            if (string.IsNullOrWhiteSpace(HoTen))
                sb.AppendLine("Vui lòng nhập Họ tên.");

            if (string.IsNullOrWhiteSpace(Phongban))
                sb.AppendLine("Vui lòng nhập Phòng ban.");

            if (string.IsNullOrWhiteSpace(Vitri))
                sb.AppendLine("Vui lòng nhập Vị trí công việc.");

            if (string.IsNullOrWhiteSpace(Lanhdao))
                sb.AppendLine("Vui lòng nhập tên Lãnh đạo nhận đơn.");

            if (string.IsNullOrWhiteSpace(LyDo))
                sb.AppendLine("Vui lòng nhập Lý do xin nghỉ.");

            if (string.IsNullOrWhiteSpace(NguoiThayThe))
                sb.AppendLine("Vui lòng nhập Người thay thế.");

            if (string.IsNullOrWhiteSpace(BoPhanThayThe))
                sb.AppendLine("Vui lòng nhập Bộ phận của người thay thế.");

            if (guna2DateTime_batdau.Value.Date > guna2DateTimePicker_ketthuc.Value.Date)
                sb.AppendLine("Ngày bắt đầu không được lớn hơn ngày kết thúc.");

            if (guna2DateTimePicker_ketthuc.Value.Date < DateTime.Today)
                sb.AppendLine("Ngày kết thúc không được nhỏ hơn ngày hiện tại.");

            errorMessage = sb.ToString();
            return string.IsNullOrEmpty(errorMessage);
        }

    }
}