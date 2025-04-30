using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
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
        public Home1()
        {
            InitializeComponent();
            guna2MessageDialog12 = new Guna.UI2.WinForms.Guna2MessageDialog();
            guna2MessageDialog12.Parent = this;

        }
        public Home1(string excelusername)
        {
            InitializeComponent();
            this.excelusername = excelusername;

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
                var control = new DonDeXuatTangLuongControl();
                control.Dock = DockStyle.Fill;
                panel2.Controls.Add(control);
            }
        }

        public void ShowMessage(string message, string title, Guna.UI2.WinForms.MessageDialogIcon icon)
        {
            guna2MessageDialog12.Icon = icon;
            guna2MessageDialog12.Show(message, title);
        }

    }
}
