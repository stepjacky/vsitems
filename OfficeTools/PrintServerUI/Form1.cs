using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using myutils;
using System.Threading;

namespace PrintServerUI
{
    public partial class Form1 : Form
    {
        static OfficeTools tools = new OfficeTools();
        string ipaddr = "localhost";
        int port = 10086;
        Thread uithread=null;  
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            IPHostEntry ipHost = Dns.GetHostEntry(Dns.GetHostName());
            foreach (IPAddress ip in ipHost.AddressList)
            {
                listBox1.Items.Add(ip.ToString());
            }
            ComboxItem ci = new ComboxItem();
            ci.text = "文件服务";
            ci.value = "file";
            comboBox1.Items.Add(ci);
            ci = new ComboxItem();
            ci.text = "文件数据服务";
            ci.value = "fileData";
            comboBox1.Items.Add(ci);
            comboBox1.SelectedIndex = 0;
            button2.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ipaddr == "localhost")
            {
                MessageBox.Show(this, "请选择一个IPV4的地址");
                return;
            }
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show(this, "请选择服务类型");
                return;
            }
            ComboxItem si = comboBox1.SelectedItem as ComboxItem;
            
            if ("file"==si.value)
            {
                uithread = new Thread(new ThreadStart(this.startListener));
            }

            if ("fileData"==si.value)
            {
                uithread = new Thread(new ThreadStart(this.startDataListener));
            }
             
            uithread.IsBackground = true;

            uithread.Start(); 
            
            button1.Enabled = false;
            button1.Text = "正在监听..";
            button2.Enabled = true;
                
        }

        private void listBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            ipaddr = listBox1.SelectedItem.ToString();
        }

        private void startListener()
        {
            tools.startServer(ipaddr, port);
        }

        private void startDataListener()
        {
            tools.startJsonServer(ipaddr, port);
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

            closeThread();            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            closeThread();
        }

        void closeThread()
        {
            try
            {
                if (uithread.IsAlive) uithread.Abort();

                uithread.Join(1000);
                button1.Enabled = true;
                button1.Text = "●开始监听";
                button2.Enabled = false;
                
            }
            catch (ThreadAbortException)
            {

            }
            catch (Exception)
            {

            }
        }
    }
}
