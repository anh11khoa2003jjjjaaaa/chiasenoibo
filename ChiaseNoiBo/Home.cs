using Guna.UI2.WinForms;
using OfficeOpenXml;
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
    public partial class Home : Form
    {
        private string excelusername;
        public Home()
        {
            InitializeComponent();
        }
        public Home(string excelusername)
        {
            InitializeComponent();
            this.excelusername = excelusername;
           
        }
        public void LoadUserControl(UserControl uc)
        {
            panel1.Controls.Clear();
            uc.Dock = DockStyle.Fill;
            panel1.Controls.Add(uc);
        }

        private void Home_Load(object sender, EventArgs e)
        {
            LoadUserControl(new UserControl_LoadFile());
            label2.Text = excelusername;

        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            this.Hide();
            Login login = new Login();
            login.ShowDialog();
        }
    }
}

