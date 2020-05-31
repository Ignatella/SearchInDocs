using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordLibrary;

namespace SearchInDocs_WF
{
    public partial class Form1 : Form
    {
        private Point lastPoint;
        private readonly SynchronizationContext syncContext;
        public Form1()
        {
            InitializeComponent();
            FormBorderStyle = FormBorderStyle.FixedDialog;

            this.FormBorderStyle = FormBorderStyle.None;

            syncContext = SynchronizationContext.Current;
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

        private void select_btn_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folder = new FolderBrowserDialog())
            {
                if (folder.ShowDialog() == DialogResult.OK)
                    dir_txtBox.Text = folder.SelectedPath;
            }
            dir_txtBox.SelectAll();
            dir_txtBox.ScrollToCaret();
        }

        private void start_btn_Click(object sender, EventArgs e)
        {
            if (dir_txtBox.Text.Length > 0 && Directory.Exists(dir_txtBox.Text) 
                && word_txtBox.Text.Length > 0 && agree_checkBox.Checked)
            {

                progress_progressBar.Maximum = Search.GetFileCount(dir_txtBox.Text);

                SearchOptions options = new SearchOptions(word_txtBox.Text, dir_txtBox.Text);

                Thread thread = new Thread(new ParameterizedThreadStart(SearchInFilesAndConvertPagesToJpg));
                
                thread.IsBackground = true;
                thread.Start(options);
            }
        }

        private void SearchInFilesAndConvertPagesToJpg(object data)
        {
            if(data is SearchOptions)
            {
                SearchOptions options = (SearchOptions)data;
                Search.SearchInFilesAndConvertPagesToJpg(options.StrToSearchFor, options.Path, (() =>
                {
                    syncContext.Post(UpdateProgressBar, null);
                }), (() =>
                {
                    syncContext.Post((actionObject) => {
                        progress_label.Text = "Done.";
                    }, null);
                }));
            }
        }

        private void UpdateProgressBar(object actionText) =>
            progress_progressBar.Increment(1);
    }
}
