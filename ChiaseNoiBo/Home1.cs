using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Guna.UI2.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChiaseNoiBo
{
   
    public partial class Home1 : Form
    {
        private Guna.UI2.WinForms.Guna2MessageDialog guna2MessageDialog12;
        private string excelusername;
        private readonly UserControl_LoadFile userControl_LoadFile;
        private readonly HuongDanSuDungControl huongDanSuDungControl;
       
        public Home1()
        {
            InitializeComponent();
            guna2MessageDialog12 = new Guna.UI2.WinForms.Guna2MessageDialog();
            guna2MessageDialog12.Parent = this;

            userControl_LoadFile = new UserControl_LoadFile();
          

            // Khởi tạo label hiển thị trạng thái


        }
        public Home1(string excelusername)
        {
            InitializeComponent();
            this.excelusername = excelusername;
            guna2MessageDialog12 = new Guna.UI2.WinForms.Guna2MessageDialog();
            guna2MessageDialog12.Parent = this;
            userControl_LoadFile = new UserControl_LoadFile();

            // Khởi tạo label hiển thị trạng thái

        }
        public void LoadUserControl(UserControl uc)
        {
            panel2.Controls.Clear();
            uc.Dock = DockStyle.Fill;
            panel2.Controls.Add(uc);
        }

        private void Home1_Load(object sender, EventArgs e)
        {

            label2.Text = excelusername;
           LoadUserControl(new HuongDanSuDungControl());
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Login login = new Login();
            login.ShowDialog();
        }
        //Nut Excel
        private async void guna2Button3_Click(object sender, EventArgs e)
        {

            var uc = new UserControl_LoadFile();
            LoadUserControl(uc);
           await uc.LoadExcelFilesAsync();

        }
        //Nut van ban
        private async void guna2Button4_Click(object sender, EventArgs e)
        {
            var uc = new UserControl_LoadFile();
            LoadUserControl(uc);
            await uc.LoadWordAndPdfFilesAsync();


        }
        // tao bieu mau
        private void guna2Button6_Click(object sender, EventArgs e)
        {

        }

        private void guna2ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            panel2.Controls.Clear();

            if (guna2ComboBox1.SelectedItem.ToString() == "Đơn xin nghỉ phép")
            {
                var control = new DonXinNghiPhepControl();
                control.Dock = DockStyle.Fill;
                panel2.Controls.Add(control);
            }
            else if (guna2ComboBox1.SelectedItem.ToString() == "Đề xuất tăng lương")
            {
                var control = new DeXuatTangLuongControl();
                control.Dock = DockStyle.Fill;
                panel2.Controls.Add(control);
            }
        }

        public void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            guna2MessageDialog12.Icon = icon;
            guna2MessageDialog12.Show(message, title);
        }


        //Nút cập nhật phiên bản
        private async void guna2Button6_Click_1(object sender, EventArgs e)
        {
            try
            {
                // Khởi tạo UI hiển thị tiến trình
                //InitializeUpdateStatusUI();
                //_updateCancellationTokenSource = new CancellationTokenSource();

                // Bắt đầu đếm thời gian cập nhật
                //var countdownTask = ShowUpdateCountdownAsync(_lblUpdateStatus, _updateCancellationTokenSource.Token);

                var credential = GoogleCredential.FromFile(GoogleDriveUpdater.CredentialPath)
                    .CreateScoped(GoogleDriveUpdater.Scopes);
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = GoogleDriveUpdater.ApplicationName,
                });

                string localVersion = btn_version.Text.Replace("Phiên bản ", "").Trim();
                string driveVersion = await userControl_LoadFile.CheckOnlineVersion(service);

                if (string.IsNullOrEmpty(driveVersion))
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không thể lấy phiên bản từ Google Drive!", "Lỗi");
                    //_updateCancellationTokenSource.Cancel();
                    //CleanupUpdateStatusUI();
                    return;
                }

                if (localVersion == driveVersion)
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                    guna2MessageDialog1.Show($"Bạn đang sử dụng phiên bản mới nhất: {localVersion}", "Thông báo");
                    //_updateCancellationTokenSource.Cancel();
                    //CleanupUpdateStatusUI();
                    return;
                }

                // Thông báo có phiên bản mới
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Warning;
                guna2MessageDialog1.Show($"Hệ thống đã có phiên bản mới! Vui lòng cập nhật để trải nghiệm chức năng mới nhất!");

                // Bắt đầu tải file
                string latestMsiFileName = await userControl_LoadFile.GetMsiFileName(service, GoogleDriveUpdater.FolderId);
                if (string.IsNullOrEmpty(latestMsiFileName))
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không tìm thấy file MSI trên Google Drive!", "Lỗi");
                    //_updateCancellationTokenSource.Cancel();
                    //CleanupUpdateStatusUI();
                    return;
                }

                string downloadPath = Path.Combine(userControl_LoadFile.SharedDirectory, latestMsiFileName);
                bool downloadSuccess = await userControl_LoadFile.DownloadFileFromDrive(service, GoogleDriveUpdater.FolderId, latestMsiFileName, downloadPath);
                if (!downloadSuccess)
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không thể tải file cập nhật từ Google Drive!", "Lỗi");
                    //_updateCancellationTokenSource.Cancel();
                    //CleanupUpdateStatusUI();
                    return;
                }

                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                guna2MessageDialog1.Show("Cập nhật hoàn tất! Ứng dụng sẽ tự động cài đặt phiên bản mới.", "Cập nhật thành công");

                // Cài đặt file MSI
                userControl_LoadFile.InstallMsi(downloadPath);

                //// Dừng đếm thời gian và dọn dẹp
                //_updateCancellationTokenSource.Cancel();
                //CleanupUpdateStatusUI();
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show($"Lỗi khi kiểm tra/cập nhật phiên bản: {ex.Message}", "Lỗi");
                //_updateCancellationTokenSource?.Cancel();
                //CleanupUpdateStatusUI();
            }
        }




        private void guna2Button5_Click(object sender, EventArgs e)
        {
            LoadUserControl(new HuongDanSuDungControl());
        }
    }
    }

