using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Guna.UI2.WinForms;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using Org.BouncyCastle.Asn1.Cmp;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using File = System.IO.File;

namespace ChiaseNoiBo
{
    public partial class UserControl_LoadFile : UserControl
    {
        private string SharedDirectory = AppDomain.CurrentDomain.BaseDirectory;
        private GoogleDriveHelper gg=new GoogleDriveHelper();
      
        private static string FolderId = "15viUYINHRFLMIuCNVI4khVOHZgMf13jN"; // Thư mục chứa file Excel
        private System.Windows.Forms.Label _lblUpdateStatus;
        private CancellationTokenSource _updateCancellationTokenSource;
       
        public UserControl_LoadFile()
        {

            InitializeComponent();

        }

        private async void UserControl_LoadFile_Load(object sender, EventArgs e)
        {
            if (this.ParentForm is Home homeForm)
            {
                guna2MessageDialog1.Parent = homeForm;
            }
            await LoadExcelFilesAsync();
        }

        /// <summary>
        /// Lấy danh sách file Excel trong thư mục Google Drive và hiển thị lên panel1
        /// </summary>
        private async Task LoadExcelFilesAsync()
        {
            try
            {
                var service = GoogleDriveHelper.GetDriveService();
                var request = service.Files.List();
                request.Q = $"'{FolderId}' in parents and (mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or mimeType='application/vnd.ms-excel' or mimeType='text/csv')";

                request.Fields = "files(id, name)";

                var result = await request.ExecuteAsync();

                if (result.Files == null || result.Files.Count == 0)
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không tìm thấy file Excel nào trong thư mục!", "Thông báo");
                    return;
                }

                flowLayoutPanel1.Controls.Clear(); // Xóa danh sách cũ

                // Sử dụng FlowLayoutPanel để tự động sắp xếp

                foreach (var file in result.Files.Where(f => f.Name != "DanhSachTaiKhoan.xlsx"))
                {
                    Button fileButton = new Button
                    {
                        Text = file.Name,
                        Tag = file.Id,
                        Width = panel1.Width - 20,
                        Height = 50, // Tăng chiều cao để tránh bị cắt chữ
                        BackColor = System.Drawing.Color.WhiteSmoke,
                        FlatStyle = FlatStyle.Flat,
                        TextAlign = System.Drawing.ContentAlignment.MiddleLeft,
                        Padding = new Padding(10, 5, 10, 5),
                        Font = new Font("Segoe UI", 12, FontStyle.Regular), // Tăng kích thước chữ
                        Margin = new Padding(5, 5, 5, 5) // Tạo khoảng cách giữa các nút
                    };

                    fileButton.FlatAppearance.BorderSize = 0;
                    fileButton.Click += FileButton_Click;
                    flowLayoutPanel1.Controls.Add(fileButton);
                }

                flowLayoutPanel1.AutoScroll = true; // Bật cuộn nếu danh sách dài


            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show($"Lỗi khi tải danh sách file: {ex.Message}", "Lỗi");
            }
        }


        /// <summary>
        /// Xử lý khi người dùng nhấn vào file để mở
        /// </summary>
        private void FileButton_Click(object sender, EventArgs e)
        {
            Button btn = sender as Button;
            string fileId = btn.Tag.ToString();
            string fileName = btn.Text;

            // Mở Form1 để hiển thị nội dung file
            Form1 form = new Form1(fileId, fileName);
            form.Show();
        }


        private async Task<string> CheckOnlineVersion(DriveService service)
        {
            try
            {
                var fileMetadata = await service.Files.Get(GoogleDriveUpdater.VersionFileId).ExecuteAsync();

                if (fileMetadata.MimeType == "application/vnd.google-apps.document")
                {
                    var request = service.Files.Export(GoogleDriveUpdater.VersionFileId, "text/plain");
                    using (var stream = new MemoryStream())
                    {
                        await request.DownloadAsync(stream);
                        stream.Position = 0;
                        using (var reader = new StreamReader(stream))
                        {
                            string firstLine = reader.ReadLine();
                            return firstLine?.Trim() ?? "Error";
                        }
                    }
                }
                else if (fileMetadata.MimeType == "text/plain")
                {
                    var request = service.Files.Get(GoogleDriveUpdater.VersionFileId);
                    using (var stream = new MemoryStream())
                    {
                        await request.DownloadAsync(stream);
                        stream.Position = 0;
                        using (var reader = new StreamReader(stream))
                        {
                            string firstLine = reader.ReadLine();
                            return firstLine?.Trim() ?? "Error";
                        }
                    }
                }
                else
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("File không đúng định dạng TXT hoặc Google Docs!", "Lỗi");
                    return "Error";
                }
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show("Lỗi khi tải file: " + ex.Message, "Lỗi");
                return "Error";
            }
        }

        private async Task<bool> DownloadFileFromDrive(DriveService service, string folderId, string fileName, string downloadPath)
        {
            try
            {
                var listRequest = service.Files.List();
                listRequest.Q = $"'{FolderId}' in parents and name = \"{fileName}\"";
                listRequest.Fields = "files(id, name)";

                var files = await listRequest.ExecuteAsync();

                if (files.Files.Count == 0)
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show($"Không tìm thấy file {fileName} trên Google Drive!", "Lỗi");
                    return false;
                }

                var file = files.Files[0];
                var request = service.Files.Get(file.Id);

                using (var fileStream = new FileStream(downloadPath, FileMode.Create, FileAccess.Write))
                {
                    await request.DownloadAsync(fileStream);
                }
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                guna2MessageDialog1.Show($"File {fileName} đã được tải về thành công!", "Thành công");
                return true;
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show("Lỗi khi tải file từ Google Drive: " + ex.Message, "Lỗi");
                return false;
            }
        }


        private async Task<string> GetLatestMsiFileName(DriveService service, string folderId)
        {
            try
            {
                var listRequest = service.Files.List();
                listRequest.Q = $"'{folderId}' in parents and name contains '.msi'";
                listRequest.Fields = "files(id, name, createdTime)";
                listRequest.OrderBy = "createdTime desc"; // Lấy file mới nhất trước

                var files = await listRequest.ExecuteAsync();

                if (files.Files.Count == 0)
                {
                    return null; // Không tìm thấy file MSI nào
                }

                return files.Files[0].Name; // Trả về tên file MSI mới nhất
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show("Lỗi khi lấy file MSI mới nhất: " + ex.Message, "Lỗi");
                return null;
            }
        }

        //private async Task InstallMsi(string msiFilePath)
        //{
        //    try
        //    {
        //        guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
        //        guna2MessageDialog1.Show($"Bắt đầu cài đặt: {Path.GetFileName(msiFilePath)}", "Cài đặt");

        //        using (Process process = new Process())
        //        {
        //            process.StartInfo.FileName = "msiexec";
        //            process.StartInfo.Arguments = $"/i \"{msiFilePath}\" /qn /norestart";
        //            process.StartInfo.UseShellExecute = false;

        //            CancellationTokenSource cts = new CancellationTokenSource();
        //            System.Windows.Forms.Label lblStatus = CreateStatusLabel();
        //            this.Controls.Add(lblStatus);
        //            var updatingTask = ShowUpdateCountdownAsync(lblStatus, cts.Token); // Bắt đầu vòng lặp UI

        //            process.Start();
        //            await Task.Run(() => process.WaitForExit());
        //            // Chờ cài đặt xong

        //            cts.Cancel(); // Dừng hiển thị thông báo cập nhật
        //            await updatingTask; // Đợi task UI dừng hẳn
        //            this.Controls.Remove(lblStatus);

        //            if (process.ExitCode == 0)
        //            {
        //                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
        //                guna2MessageDialog1.Show("Cài đặt hoàn tất! Ứng dụng sẽ khởi động lại.", "Thành công");

        //                string filedelete = msiFilePath.Replace("\\", @"\");

        //                // Xóa file MSI sau khi cài đặt thành công
        //                if (File.Exists(filedelete))
        //                {
        //                    File.Delete(Path.Combine(filedelete));
        //                }

        //                // Khởi động lại ứng dụng
        //                RestartApplication();
        //            }
        //            else
        //            {
        //                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
        //                guna2MessageDialog1.Show($"Cài đặt thất bại, mã lỗi: {process.ExitCode}", "Lỗi");
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
        //        guna2MessageDialog1.Show($"Lỗi khi cài đặt MSI:\n{ex.Message}", "Lỗi");
        //    }
        //}

        

        

        private void InstallMsi(string msiFilePath)
        {
            try
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                guna2MessageDialog1.Show($"Bắt đầu cài đặt: {Path.GetFileName(msiFilePath)}", "Cài đặt");

                using (Process process = new Process())
                {
                    process.StartInfo.FileName = "msiexec";
                    process.StartInfo.Arguments = $"/i \"{msiFilePath}\" /qn /norestart";
                    process.StartInfo.UseShellExecute = false;


                    process.Start();
                    process.WaitForExit();

                    if (process.ExitCode == 0)
                    {
                        guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                        guna2MessageDialog1.Show("Cài đặt hoàn tất! Ứng dụng sẽ khởi động lại.", "Thành công");

                        string filedelete = msiFilePath.Replace("\\", @"\");


                        // Xóa file MSI sau khi cài đặt thành công
                        if (File.Exists(filedelete))
                        {
                            File.Delete(Path.Combine(filedelete));
                            //guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Information;
                            //guna2MessageDialog1.Show($"Đã xóa file: {Path.GetFileName(msiFilePath)}", "Dọn dẹp");
                        }

                        // Khởi động lại ứng dụng
                        RestartApplication();
                    }
                    else
                    {
                        guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                        guna2MessageDialog1.Show($"Cài đặt thất bại, mã lỗi: {process.ExitCode}", "Lỗi");
                    }
                }
            }
            catch (Exception ex)
            {
                guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                guna2MessageDialog1.Show($"Lỗi khi cài đặt MSI:\n{ex.Message}", "Lỗi");
            }
        }

        private void RestartApplication()
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = Application.ExecutablePath,
                UseShellExecute = true
            });

            Application.Exit();
        }

        private async void guna2Button2_Click(object sender, EventArgs e)
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

                string localVersion = guna2Button2.Text.Replace("Phiên bản ", "").Trim();
                string driveVersion = await CheckOnlineVersion(service);

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
                string latestMsiFileName = await GetLatestMsiFileName(service, GoogleDriveUpdater.FolderId);
                if (string.IsNullOrEmpty(latestMsiFileName))
                {
                    guna2MessageDialog1.Icon = Guna.UI2.WinForms.MessageDialogIcon.Error;
                    guna2MessageDialog1.Show("Không tìm thấy file MSI trên Google Drive!", "Lỗi");
                    //_updateCancellationTokenSource.Cancel();
                    //CleanupUpdateStatusUI();
                    return;
                }

                string downloadPath = Path.Combine(SharedDirectory, latestMsiFileName);
                bool downloadSuccess = await DownloadFileFromDrive(service, GoogleDriveUpdater.FolderId, latestMsiFileName, downloadPath);
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
                InstallMsi(downloadPath);

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

        //private async Task ShowUpdateCountdownAsync(System.Windows.Forms.Label lblStatus, CancellationToken token)
        //{
        //    int seconds = 0;

        //    while (!token.IsCancellationRequested)
        //    {
        //        try
        //        {
        //            // Cập nhật UI an toàn
        //            if (lblStatus.InvokeRequired)
        //            {
        //                lblStatus.Invoke((MethodInvoker)delegate
        //                {
        //                    lblStatus.Text = $"Đang cập nhật ứng dụng... {seconds} giây";
        //                });
        //            }
        //            else
        //            {
        //                lblStatus.Text = $"Đang cập nhật ứng dụng... {seconds} giây";
        //            }

        //            await Task.Delay(1000, token);
        //            seconds++;
        //        }
        //        catch (TaskCanceledException)
        //        {
        //            // Bỏ qua khi task bị hủy
        //        }
        //        catch (Exception ex)
        //        {
        //            Debug.WriteLine($"Lỗi khi cập nhật timer: {ex.Message}");
        //            break;
        //        }
        //    }

        //    // Xóa label sau khi kết thúc
        //    if (lblStatus != null && lblStatus.IsHandleCreated)
        //    {
        //        if (lblStatus.InvokeRequired)
        //        {
        //            lblStatus.Invoke((MethodInvoker)delegate
        //            {
        //                if (this.Controls.Contains(lblStatus))
        //                    this.Controls.Remove(lblStatus);
        //            });
        //        }
        //        else
        //        {
        //            if (this.Controls.Contains(lblStatus))
        //                this.Controls.Remove(lblStatus);
        //        }
        //    }
        //}
        //private void InitializeUpdateStatusUI()
        //{
        //    // Tạo label hiển thị trạng thái cập nhật
        //    _lblUpdateStatus = new System.Windows.Forms.Label
        //    {
        //        AutoSize = false,
        //        Size = new Size(350, 60),
        //        Font = new Font("Arial", 12, FontStyle.Bold),
        //        ForeColor = Color.White,
        //        BackColor = Color.FromArgb(50, 0, 0, 0),
        //        TextAlign = ContentAlignment.MiddleCenter,
        //        Location = new Point((this.Width - 350) / 2, (this.Height - 60) / 2),
        //        Text = "Đang chuẩn bị cập nhật... 0 giây"
        //    };

        //    this.Controls.Add(_lblUpdateStatus);
        //    _lblUpdateStatus.BringToFront();
        //}

        //private void CleanupUpdateStatusUI()
        //{
        //    // Dừng đếm thời gian nếu đang chạy
        //    _updateCancellationTokenSource?.Cancel();

        //    // Xóa label hiển thị trạng thái
        //    if (_lblUpdateStatus != null && this.Controls.Contains(_lblUpdateStatus))
        //    {
        //        this.Controls.Remove(_lblUpdateStatus);
        //        _lblUpdateStatus.Dispose();
        //        _lblUpdateStatus = null;
        //    }
        //}





    }
}
