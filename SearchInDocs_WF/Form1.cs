using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SearchInDocs_WF
{
    public partial class Form1 : Form
    {

        Point lastPoint;

        public Form1()
        {
            InitializeComponent();
            FormBorderStyle = FormBorderStyle.FixedDialog;

            this.FormBorderStyle = FormBorderStyle.None;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        

        private void Main_menu_panel_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Left += e.X - lastPoint.X;
                this.Top += e.Y - lastPoint.Y;
            }
        }

        private void Main_menu_panel_MouseDown(object sender, MouseEventArgs e) => lastPoint = new Point(e.X, e.Y);

        private void CloseTheApp(object sender, EventArgs e) => this.Close();

        private void MinimizeTheApp(object sender, EventArgs e) => this.WindowState = FormWindowState.Minimized;

    }
}
