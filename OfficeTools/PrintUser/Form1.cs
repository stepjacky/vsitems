using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using myutils;
namespace PrintUser
{
    public partial class Form1 : Form
    {
        OfficeTools tools = new OfficeTools();
        OpenFileDialog od = new OpenFileDialog();
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string ip = textBox1.Text;
            od.Filter = "Word2003(*.doc)|*.doc";
            DialogResult rst = od.ShowDialog(this);
            //MessageBox.Show(rst.ToString());
            if (rst == DialogResult.OK)
            {
                tools.Connect(ip, od.FileName);
            }
        }
    }
}
