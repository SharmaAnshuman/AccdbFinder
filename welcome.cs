using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Finder
{
    public partial class welcome : Form
    {
        public welcome()
        {
            InitializeComponent();
        }

        private void welcome_Load(object sender, EventArgs e)
        {
            
        }
        Timer tmr;
        private void welcome_Shown(object sender, EventArgs e)
        {
            tmr = new Timer();
            tmr.Interval = 3000;
            tmr.Start();
            tmr.Tick += tmr_Tick;
            
        }

        void tmr_Tick(object sender, EventArgs e)
        {
            tmr.Stop();
            Form1 mf = new Form1();
            mf.Show();
            this.Hide();

        }


        private void welcome_Paint(object sender, PaintEventArgs e)
        {
        }

        private void welcome_ControlAdded(object sender, ControlEventArgs e)
        {

        }

        private void pictureBox1_Paint(object sender, PaintEventArgs e)
        {
        }

        private void pictureBox1_LoadCompleted(object sender, AsyncCompletedEventArgs e)
        {

        }
    }
}
