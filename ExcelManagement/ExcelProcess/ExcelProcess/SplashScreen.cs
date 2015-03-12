using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProcess
{
    public partial class SplashScreen : Form
    {
        public SplashScreen()
        {
            InitializeComponent();
        }

        private void SplashScreen_Load(object sender, EventArgs e)
        {
            timer1.Start();
        }
        private void openMainForm()
        {
            Application.Run(new Form2());
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            System.Threading.Thread tMain = new System.Threading.Thread(new System.Threading.ThreadStart(openMainForm));
            tMain.Start();
            this.Close();
        }
    }
}
