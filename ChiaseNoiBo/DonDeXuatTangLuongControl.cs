using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Guna.UI2.WinForms;
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

namespace ChiaseNoiBo
{
    public partial class DonDeXuatTangLuongControl : UserControl
    {
        public DonDeXuatTangLuongControl()
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
        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }
        //171SJndW98HPjyNzy4--8L3ELyBPo556N
        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

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

        private async void guna2Button2_Click_1(object sender, EventArgs e)
        {
            if (!ValidateFormInputsV2(out string error))
            {
                ShowMessage(error, "Lỗi nhập liệu", Guna.UI2.WinForms.MessageDialogIcon.Warning);
                return;
            }
            try
            {
                string fileId = "171SJndW98HPjyNzy4--8L3ELyBPo556N";
                string tempPath = Path.Combine(Path.GetTempPath(), "DonDeXuatTangLuong.docx");

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
                }
            }
            catch (Exception ex)
            {
                ShowMessage("Lỗi: " + ex.Message, "Lỗi", Guna.UI2.WinForms.MessageDialogIcon.Error);
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

    }

}
