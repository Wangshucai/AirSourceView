using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WinFormDesign
{
    public partial class SoftWareForm : Form
    {
        private WinFormDesign.MainForm frmMain = new WinFormDesign.MainForm();
        public SoftWareForm(WinFormDesign.MainForm parent)
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

        private void SoftWareForm_Load(object sender, EventArgs e)
        {

        }
    }
}
