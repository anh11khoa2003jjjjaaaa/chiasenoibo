using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Guna.UI2.Native.WinApi;

namespace ChiaseNoiBo
{
    public partial class ForgotPassword : Form
    {
        public ForgotPassword()
        {
            InitializeComponent();
        }

        private bool IsValidate()
        {
            string email = txt_email.Text.Trim();

            // Kiểm tra rỗng
            if (string.IsNullOrEmpty(email))
            {
                txt_email.PlaceholderText = "Email không được để trống!";
                txt_email.PlaceholderForeColor = Color.Red;
                txt_email.Text = ""; // Xoá nội dung để hiển thị placeholder
                return false;
            }

            // Kiểm tra định dạng email
            try
            {
                var addr = new System.Net.Mail.MailAddress(email);
                return true;
            }
            catch
            {
                txt_email.Text = "";
                txt_email.PlaceholderText = "Email không hợp lệ!";
                txt_email.PlaceholderForeColor = Color.Red;
                return false;
            }
        }


        private void guna2Button1_Click(object sender, EventArgs e)
        {
            if(IsValidate())
            {
                return;
            }
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
    }
