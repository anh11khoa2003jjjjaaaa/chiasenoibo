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
    public partial class UpdateNotificationForm : Form
    {
        private int _countdown = 10; // Thời gian đếm ngược (giây)
        private Timer _timer;

        public UpdateNotificationForm()
        {
            InitializeComponent();
            InitializeTimer();
        }

        private void InitializeTimer()
        {
            _timer = new Timer();
            _timer.Interval = 1000; // 1 giây
            _timer.Tick += Timer_Tick;
            _timer.Start();
        }

        private void Timer_Tick(object sender, EventArgs e)
        {
            _countdown--;
            lblMessage.Text = $"Vui lòng chờ... đang cập nhật hệ thống ({_countdown} giây)";

            if (_countdown <= 0)
            {
                _timer.Stop();
                this.Close(); // Đóng form sau khi đếm ngược kết thúc
            }
        }

        private void UpdateNotificationForm_Load(object sender, EventArgs e)
        {

        }
    }

}
