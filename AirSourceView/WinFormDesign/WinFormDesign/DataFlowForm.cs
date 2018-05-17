using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace WinFormDesign
{
    public partial class DataFlowForm : Form
    {
        private MainForm frmMain = new MainForm();
       
        public DataFlowForm(WinFormDesign.MainForm parentt)
        {

            InitializeComponent();
            frmMain = parentt;
            this.StartPosition = FormStartPosition.CenterScreen;

        }

        protected override void OnClosing(CancelEventArgs e)
        {
            // 这里写关闭窗体要执行的代码
            frmMain.SHOW = 0;
            frmMain.Childflag = 0;
            base.OnClosing(e);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            if (button1.Text == "暂停")
            {

                frmMain.SHOW = 0;
                button1.Text = "开始";
                button1.BackColor = Color.CadetBlue;

            }
            else
            {

                frmMain.SHOW = 1;
                button1.Text = "暂停";
                button1.BackColor = Color.PaleVioletRed;
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {

            saveFileDialog1.Filter = "*.txt|*.txt";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamWriter sw = new StreamWriter(saveFileDialog1.FileName);

                if (richTextBox1.Text != null)
                {
                    foreach (string str in richTextBox1.Lines)
                    {
                        sw.WriteLine(str);
                    }
                }
                else
                {
                    MessageBox.Show("文本框为空！");
                }
                sw.Close();
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void DataFlow_Load(object sender, EventArgs e)
        {
            button1.Text = "暂停";
            button2.Text = "清空";
            button3.Text = "保存";
            button1.BackColor = Color.PaleVioletRed;
            frmMain.SHOW = 1;
        }

        
    }
}
