using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;

namespace WinFormDesign
{
    public partial class SerialSetForm : Form
    {
        private WinFormDesign.MainForm frmMain = new WinFormDesign.MainForm();
        public SerialSetForm(WinFormDesign.MainForm parent)
        {
            InitializeComponent();
            frmMain = parent;
            this.StartPosition = FormStartPosition.CenterScreen;
        }
        protected override void OnClosing(CancelEventArgs e)
        {
            // 这里写关闭窗体要执行的代码
            frmMain.Childflag = 0;
            base.OnClosing(e);
        }
        public void getComPart(ComboBox Cname)
        {
            Microsoft.VisualBasic.Devices.Computer pc = new Microsoft.VisualBasic.Devices.Computer(); ;

            foreach (string s in pc.Ports.SerialPortNames)
            {
                Cname.Items.Add(s);
            }

        }


        private void ComboxEnableSet(bool Set)
        {
            comboBox1.Enabled = Set;
            comboBox2.Enabled = Set;
            comboBox3.Enabled = Set;
            comboBox4.Enabled = Set;
            comboBox5.Enabled = Set;
            textBox1.Enabled = Set;

        }

        private void SerialSetForm_Load(object sender, EventArgs e)
        {
            if (frmMain.serialPort1.IsOpen)
            {
                button2.BackColor = Color.Green;
                comboBox1.Items.Add(frmMain.serialPort1.PortName);
                comboBox1.Text = frmMain.serialPort1.PortName;
                comboBox2.Text = Convert.ToString(frmMain.serialPort1.BaudRate);
                comboBox3.Text = Convert.ToString(frmMain.serialPort1.DataBits); ;
                comboBox4.Text = (frmMain.serialPort1.Parity == Parity.None) ? "None" : ((frmMain.serialPort1.Parity == Parity.Odd) ? "Odd" : "Even");
                comboBox5.Text = (frmMain.serialPort1.StopBits == StopBits.One) ? "1" : "2";
                button1.Text = "关闭串口";

                ComboxEnableSet(false);

            }
            else
            {

                getComPart(comboBox1);
                comboBox2.Text = "9600";
                comboBox3.Text = "8";
                comboBox4.Text = "None";
                comboBox5.Text = "1";
                button1.Text = "打开串口";

            }
            textBox1.Text = Convert.ToString(frmMain.MailAddr);

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

       

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (comboBox1.Text != "" && button1.Text == "打开串口")
            {
               
                frmMain.serialPort1.PortName = comboBox1.Text;
                frmMain.serialPort1.BaudRate = Convert.ToInt32(comboBox2.Text);
                frmMain.serialPort1.DataBits = Convert.ToInt32(comboBox3.Text);
                frmMain.serialPort1.Parity = (comboBox4.Text == "None") ? Parity.None : ((comboBox4.Text == "Odd") ? Parity.Odd : Parity.Even);
                frmMain.serialPort1.StopBits = (comboBox5.Text == "1") ? StopBits.One : StopBits.Two;
                if (!frmMain.serialPort1.IsOpen)
                {
                    try
                    {

                        frmMain.serialPort1.Open();
                      
                        button1.Text = "关闭串口";
                        button2.BackColor = Color.Green;
                        ComboxEnableSet(false);

                    }
                    catch
                    {
                        MessageBox.Show("端口错误！", "提示");
                    }

                }

            }
            else if (button1.Text == "关闭串口")
            {

                if (frmMain.serialPort1.IsOpen)
                {
                    string NowPart = comboBox1.Text;
                    comboBox1.Items.Clear();
                    frmMain.serialPort1.Close();
                    getComPart(comboBox1);
                    comboBox1.Text = NowPart;
                    button2.BackColor = Color.White;
                    button1.Text = "打开串口";
                    ComboxEnableSet(true);
                 

                }
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            frmMain.MailAddr = Convert.ToByte(textBox1.Text);
        }
    }
}
