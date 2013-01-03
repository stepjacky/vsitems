using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ChatClient
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Process currProcess = Process.GetCurrentProcess();
            TreeNode tn = new TreeNode();
            tn.Name = "pdtdoc";
            tn.ImageIndex = 0;
            tn.Text = "生产管理处";
            treeView1.Nodes.Add(tn);
            tn.Nodes.Add("qujiakang", "屈甲康", 1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Process currProcess = Process.GetCurrentProcess();
            MessageBox.Show("结束当前进程" + currProcess.ProcessName);
            KillProcess(currProcess.ProcessName);
           
            
        }


        private void KillProcess(string processName)
        {
            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcessesByName(processName);

            foreach (System.Diagnostics.Process procCur in procs)
            {
                procCur.Kill();
                procCur.Close();
            }
        }
    }
}
