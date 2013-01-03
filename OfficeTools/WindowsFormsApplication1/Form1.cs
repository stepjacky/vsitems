using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            string homeurl = @"http://www.iteye.com";
           
            InitializeComponent();
            webBrowser1.Url = new Uri(homeurl);
        }

        private void webBrowser1_NewWindow(object sender, CancelEventArgs e)
        {
            webBrowser1.Navigate(webBrowser1.StatusText);
            e.Cancel = true;   

        }

        private void webBrowser1_Navigating(object sender, WebBrowserNavigatingEventArgs e)
        {
            if (e.Url.Equals("http://www.microsoft.com"))
            {
                e.Cancel = true;
            }   

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
              //DialogResult dr  = MessageBox.Show("网页加载完毕!", "系统提示");
              
        }

        private void webBrowser1_ProgressChanged(object sender, WebBrowserProgressChangedEventArgs e)
        {
           // toolStripProgressBar1.Minimum = 0;
           
           // toolStripProgressBar1.Maximum = Convert.ToInt32(e.MaximumProgress);
           // toolStripProgressBar1.Value = Convert.ToInt32(e.CurrentProgress<0?0:e.CurrentProgress);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            webBrowser1.ShowPageSetupDialog();
        }

        private void 打印预览ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.ShowPrintPreviewDialog();
        }

        private void 打印ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.ShowPrintDialog();
        }

        private void 关于InternetExplorerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            webBrowser1.Print();
        }

        private void 版本ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(webBrowser1.ProductName);
        }

       

        

      
    }
}
