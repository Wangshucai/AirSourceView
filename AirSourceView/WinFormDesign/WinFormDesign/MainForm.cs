using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace WinFormDesign
{
    public partial class MainForm : Form
    {
        public byte MailAddr = 1;  //默认通信地址1，可在串口设置界面更改
        Cparameter cparameter = new Cparameter();
        Ccontrols ccontrols = new Ccontrols();
        Modbus modbus = new Modbus();
        CparameterControl SetFlag = new CparameterControl();
        public int Labe10ShowTime = 0;
        public int Childflag = 0;
        public int SHOW;

        public long ErrorDataCount = 0;
        public long ErrorDataFlag = 0;

        private static string str = "";
        private string StartTime;
        private DataFlowForm dataflowform;


        //创建Excel 对象
        Excel.Application excelApp;//创建实例
        Excel.Workbook workBook;
        Excel.Worksheet workSheet;
        int  ExcelOpen = 0;
        /****************************************************/


        public MainForm()
        {

            ExcelCreat();
            this.StartPosition = FormStartPosition.CenterScreen;
            InitializeComponent();
        }

        private void ExcelCreat()
        {
           

            if (ExcelOpen == 0)
            {
                ExcelOpen = 1;
                excelApp = new Excel.Application();//创建实例
                StartTime = DateTime.Now.ToString("yyyy")+"年"+ DateTime.Now.ToString("MM")+"月"+DateTime.Now.ToString("dd")+"日"+DateTime.Now.ToString("HH")+"-"+ DateTime.Now.ToString("mm")+"-"+ DateTime.Now.ToString("ss");
                workBook = excelApp.Workbooks.Add(true);
                workSheet = workBook.ActiveSheet as Excel.Worksheet;
                Excel.Worksheet workSheet1 = workBook.Worksheets.Add(Missing.Value, workSheet, 3, Missing.Value);
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);//获得第i个sheet，准备写入  
                workSheet.Name = "状态显示";

                for (int i = 1; i < cparameter.OPER_STA_NUM + 2; i++)
                {
                    if (i == 1)
                    {
                        workSheet.Cells[1, i] = "时间";
                    }
                    else
                    {
                        workSheet.Cells[1, i] = cparameter.OperNameConst[i - 2];
                    }

                }

                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(2);//获得第i个sheet，准备写入  
                workSheet.Name = "故障显示";
                for (int i = 1; i < cparameter.ERR_STA_NUM + 2; i++)
                {
                    if (i == 1)
                    {
                        workSheet.Cells[1, i] = "时间";
                    }
                    else
                    {
                        workSheet.Cells[1, i] = cparameter.ERRNameConst[i - 2];
                    }

                }
                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(3);//获得第i个sheet，准备写入  
                workSheet.Name = "DO输出显示";
                for (int i = 1; i < cparameter.DO_STA_NUM + 2; i++)
                {
                    if (i == 1)
                    {
                        workSheet.Cells[1, i] = "时间";
                    }
                    else
                    {
                        workSheet.Cells[1, i] = cparameter.DONameConst[i - 2];
                    }

                }

                workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(4);//获得第i个sheet，准备写入  
                workSheet.Name = "参数二显示";
                for (int i = 1; i < cparameter.OPER2_STA_NUM + 2; i++)
                {
                    if (i == 1)
                    {
                        workSheet.Cells[1, i] = "时间";
                    }
                    else
                    {
                        workSheet.Cells[1, i] = cparameter.OPER2NameConst[i - 2];
                    }

                }


            }

           


        }


        private void ExcelExit()
        {

            if (ExcelOpen == 1)
            {
                ExcelOpen = 0;
                string currentPath = Directory.GetCurrentDirectory();
                string filename = currentPath + "\\" + StartTime +"——"+ DateTime.Now.ToString("dd")+"日"+ DateTime.Now.ToString("HH") + "-" + DateTime.Now.ToString("mm") + "-" + DateTime.Now.ToString("ss") + ".xlsx";
                workBook.SaveAs(filename);

                workBook.Close(false, Missing.Value, Missing.Value);



                excelApp.Quit();
                workSheet = null;
                workBook = null;
                excelApp = null;
                GC.Collect();
               

            }
           

        }


        protected override void OnClosing(CancelEventArgs e)
        {
            ExcelExit();
            base.OnClosing(e);
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            CreatOperStaUI();
            CreatOperStaUII();
            CreatOperStaUIII();
            CreatOperStaUIV();
            timer1.Start();
            timer2.Start();
            SetFlag.CONTROL_FLAG = SetFlag.GET_OPER_REG_STA;
            numericUpDown1.Value = 12;  //制冷设定默认值
            numericUpDown2.Value = 41;  //制热设定默认值
            label12.ForeColor = Color.DarkBlue;
            label12.Text = "软件版本：";


        }

        public void CreatOperStaUI()
        {

            short creatCount = cparameter.OPER_STA_NUM;
            ccontrols.creatLabUI(ref cparameter.OperName, ref cparameter.OperValue,
                                      120, 20, 80, 20, groupBox1, creatCount, 3,40, 250);
            for (int i = 0; i < cparameter.OPER_STA_NUM; i++)
            {
                if (cparameter.OperName[i] != null)
                    cparameter.OperName[i].Text = cparameter.OperNameConst[i];

                if (cparameter.OperValue[i] != null)
                    cparameter.OperValue[i].Text = "未检测";
            }
            label7.Text = "通讯状态：断开";
            label7.ForeColor = Color.Red;
            label7.Font = new System.Drawing.Font("宋体", 9);
            label10.Text = "";

        }
        //故障界面
        public void CreatOperStaUII()
        {
            short creatCount = cparameter.ERR_STA_NUM;
            ccontrols.creatLabUI(ref cparameter.ERRName, ref cparameter.ERRValue,
                                     160, 20, 80, 20, groupBox3, creatCount, 3, 26, 240);
            for (int i = 0; i < creatCount; i++)
            {
                if (cparameter.ERRName[i] != null)
                    cparameter.ERRName[i].Text = cparameter.ERRNameConst[i];

                if (cparameter.ERRValue[i] != null)
                    cparameter.ERRValue[i].Text = "未检测";
            }
        }
        //DO界面
        public void CreatOperStaUIII()
        {
            short creatCount = cparameter.DO_STA_NUM;
            ccontrols.creatLabUI(ref cparameter.DOName, ref cparameter.DOValue,
                                     120,20, 80, 20, groupBox4, creatCount, 1, 30, 350);
            for (int i = 0; i < creatCount; i++)
            {
                if (cparameter.DOName[i] != null)
                    cparameter.DOName[i].Text = cparameter.DONameConst[i];

                if (cparameter.DOValue[i] != null)
                    cparameter.DOValue[i].Text = "未检测";
            }
        }

        //参数二界面
        public void CreatOperStaUIV()
        {
            short creatCount = cparameter.OPER2_STA_NUM;
            ccontrols.creatLabUI(ref cparameter.OPER2Name, ref cparameter.OPER2Value,
                                     160, 20, 70, 20, groupBox5, creatCount, 3, 30, 240);
            for (int i = 0; i < creatCount; i++)
            {
                if (cparameter.OPER2Name[i] != null)
                    cparameter.OPER2Name[i].Text = cparameter.OPER2NameConst[i];

                if (cparameter.OPER2Value[i] != null)
                    cparameter.OPER2Value[i].Text = "未检测";
            }
        }

        /*发送请求，并设置接收标志*/
        public void RequestStep()
        {
            if (SetFlag.CONTROL_FLAG == SetFlag.SET_SET_REG_STA)
            {

                Send_Set_Procedure(0x10, 30001, cparameter.SetValue, cparameter.SET_STA_NUM);
                SetFlag.RECEIVE_FLAG = SetFlag.SET_SET_REG_STA;

            }
            else if (SetFlag.CONTROL_FLAG == SetFlag.GET_SET_REG_STA)
            {
                Send_Get_Procedure(0x03, 30001, cparameter.SET_STA_NUM);
                SetFlag.RECEIVE_FLAG = SetFlag.GET_SET_REG_STA;
              
            }

            else if (SetFlag.CONTROL_FLAG == SetFlag.GET_OPER_REG_STA)
            {
                Send_Get_Procedure(03, 50001, cparameter.OPER_STA_NUM);
               
                SetFlag.RECEIVE_FLAG = SetFlag.GET_OPER_REG_STA;
            }
            else if (SetFlag.CONTROL_FLAG == SetFlag.GET_ERR_REG_STA)
            {
                Send_Get_Procedure(01, 7001, cparameter.ERR_STA_NUM);
                SetFlag.RECEIVE_FLAG = SetFlag.GET_ERR_REG_STA;
              
            }
            else if (SetFlag.CONTROL_FLAG == SetFlag.GET_DO_REG_STA)
            {
                Send_Get_Procedure(01, 8001, cparameter.DO_STA_NUM);
                SetFlag.RECEIVE_FLAG = SetFlag.GET_DO_REG_STA;
               
            }
            else if (SetFlag.CONTROL_FLAG == SetFlag.GET_SOFTWARE_REG_STA)
            {
                Send_Get_Procedure(0x03, 62001, 15);
                SetFlag.RECEIVE_FLAG = SetFlag.GET_SOFTWARE_REG_STA;
                
            }
            else if (SetFlag.CONTROL_FLAG == SetFlag.GET_OPER2_REG_STA)
            {
                Send_Get_Procedure(03, 60001, cparameter.OPER2_STA_NUM);
                SetFlag.RECEIVE_FLAG = SetFlag.GET_OPER2_REG_STA;
                
            }
        }

       //发送请求数据帧
        public void Send_Get_Procedure(byte funCode, ushort startAddr, ushort Num)
        {
            byte[] buf = new byte[8];
            ushort checksum, ctr = 0;
            buf[0] = (byte)(MailAddr);
            buf[1] = funCode;
            modbus.WordToByte(startAddr, ref buf[2], ref buf[3]);
            modbus.WordToByte(Num, ref buf[4], ref buf[5]);
            ctr = 6;
            checksum = modbus.Crc16_override(buf, 0, (ushort)(ctr - 1));
            modbus.WordToByte(checksum, ref buf[7], ref buf[6]);
            ctr += 2;

            if (serialPort1.IsOpen)
            {
                serialPort1.Write(buf, 0, ctr);
                if (SHOW == 1)
                {
                    str = ToHexString(buf);
                    dataflowform.richTextBox1.ForeColor = Color.DarkBlue;
                    dataflowform.richTextBox1.Text += DateTime.Now.ToString("[TX_HH:mm:ss]") + str + "\r\n";
                    
                    // dataflowform.richTextBox1.SelectionStart = dataflowform.richTextBox1.TextLength;//将滑动条移动到底部
                    // dataflowform.richTextBox1.ScrollToCaret();

                }
                else
                {
                    str = "";
                }
            }


        }
        //发送设定数据帧
        public void Send_Set_Procedure(byte funCode, ushort startAddr, ushort[] SetValue, byte RegNum)
        {
            byte[] buf = new byte[9 + 2 * RegNum];
            ushort checksum, ctr = 0;
            short index = 0;

            buf[index++] = (byte)(MailAddr);
            buf[index++] = funCode;
            modbus.WordToByte(startAddr, ref buf[index++], ref buf[index++]);
            modbus.WordToByte(RegNum, ref buf[index++], ref buf[index++]);
            buf[index++] = (byte)((2) * (int)RegNum);

            ctr += 7;

            for (int i = 0; i < RegNum; i++)
            {
                modbus.WordToByte(SetValue[i], ref buf[index++], ref buf[index++]);
                ctr += 2;
            }

            checksum = modbus.Crc16_override(buf, 0, (ushort)(ctr - 1));

            modbus.WordToByte(checksum, ref buf[index + 1], ref buf[index]);
            ctr += 2;
            if (SHOW == 1)
            {
                str = ToHexString(buf);
                dataflowform.richTextBox1.ForeColor = Color.DarkBlue;
                dataflowform.richTextBox1.Text += DateTime.Now.ToString("[TX_HH:mm:ss]") + str + "\r\n";

                //dataflowform.richTextBox1.SelectionStart = dataflowform.richTextBox1.TextLength;//将滑动条移动到底部
               // dataflowform.richTextBox1.ScrollToCaret();


            }
            else
            {
                str = "";
            }
            if (!serialPort1.IsOpen)
            {
                serialPort1.Open();
                serialPort1.Write(buf, 0, ctr);
            }
            else
            {
                serialPort1.Write(buf, 0, ctr);
            }

        }

        //接收回复数据帧解析，并显示
        private void DataAnalyze()
        {
            switch (SetFlag.RECEIVE_FLAG)
            {

                /*状态读取 */
                case 0x01:
                    { 
                        OperDataShow();
                    } break;

                /*设定读 */
                case 0x02:
                    {
                        SetDataShow();
                    } break;

                /*设定写*/
                case 0x03:
                    {
                        SetDataCheck();
                    }
                    break;

                /*读故障*/
                case 0x04:
                    {
                        ErrDataShow();
                    }
                    break;
                /*读开关量*/
                case 0x05:
                    {
                        DoDataShow();
                    }
                    break;
                case 0x06:
                    {
                        Oper2DataShow();
                    }
                    break;
                case 0x07:
                    {
                        SoftWareDataShow();
                    }
                    break;
                default:
                    {
                        SetFlag.RECEIVE_FLAG = SetFlag.GET_OPER_REG_STA;
                        SetFlag.CONTROL_FLAG = SetFlag.GET_OPER_REG_STA;
                    }
                    break;

            }
           
        }
        
        //机组状态参数显示
        private void OperDataShow()
        {

            if (SetFlag.RECEIVE_FLAG == SetFlag.GET_OPER_REG_STA)
            {
                if (cparameter.ShowInformation[1] == 0x03 && cparameter.ShowInformation[2] == cparameter.OPER_STA_NUM*2)
                {
                    if (SetFlag.CONTROL_FLAG != SetFlag.SET_SET_REG_STA && SetFlag.CONTROL_FLAG != SetFlag.GET_SET_REG_STA)
                    {
                        SetFlag.CONTROL_FLAG = SetFlag.GET_ERR_REG_STA;
                    }
                   
                    short Count = 3;
                    short[] Data = new short[cparameter.OPER_STA_NUM];
                    for (int i = 0; i < cparameter.OPER_STA_NUM; i++)
                    {
                        Data[i] = modbus.TwoToWord(cparameter.ShowInformation[Count++], cparameter.ShowInformation[Count++]);
                    }

                    cparameter.ShowInformation.Clear();                              //清空接收缓存区
                    cparameter.ShowInformationCrt = 0;

                    if (tabControl1.SelectedIndex == 0)
                    {
                        
                        if (Data[0] == 0) { cparameter.OperValue[0].Text = "关机"; } else { cparameter.OperValue[0].Text = "开机"; }
                        if (Data[1] == 1) { cparameter.OperValue[1].Text = "制冷"; } else { cparameter.OperValue[1].Text = "制热"; }
                        cparameter.OperValue[2].Text = Convert.ToString(Data[2]) + " Hz";

                        ccontrols.RefreshAnalogEEV(cparameter.OperValue[3], Data[3], " B");

                        if (Data[4] == 1) { cparameter.OperValue[4].Text = "告警"; cparameter.OperValue[4].ForeColor = Color.Red; } else { cparameter.OperValue[4].Text = "正常"; cparameter.OperValue[4].ForeColor = Color.DarkBlue; }
                        ccontrols.RefreshAnalog(cparameter.OperValue[5], Data[5], "℃");
                        ccontrols.RefreshAnalog(cparameter.OperValue[6], Data[6], "℃");
                        cparameter.OperValue[7].Text = Convert.ToString((float)(Data[7]) / 10) + " Kpa";
                        ccontrols.RefreshAnalog(cparameter.OperValue[8], Data[8], "℃");
                        ccontrols.RefreshAnalog(cparameter.OperValue[9], Data[9], "℃");
                        ccontrols.RefreshAnalog(cparameter.OperValue[10], Data[10], "℃");
                        ccontrols.RefreshAnalog(cparameter.OperValue[11], Data[11], "℃");
                        ccontrols.RefreshAnalog(cparameter.OperValue[12], Data[12], "℃");
                        ccontrols.RefreshAnalog(cparameter.OperValue[13], Data[13], "℃");
       
                    }


                    if (ExcelOpen == 1)
                    {
                        workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(1);//获得第i个sheet，准备写入  
                        Int64 num = workSheet.UsedRange.Rows.Count;

                        workSheet.Cells[num + 1, 1] = Convert.ToString(DateTime.Now.ToString("HH:mm:ss"));
                        for (int i = 0; i < cparameter.OPER_STA_NUM; i++)
                        {
                            if (i < 5)
                            {
                                workSheet.Cells[num + 1, i + 2] = Convert.ToString(Data[i]);
                            }
                            else
                            {

                                workSheet.Cells[num + 1, i + 2] = Convert.ToString((float)Data[i] / 10);
                            }

                        }

                    }


                }



               

               
            }

        }

        //设置参数显示
        private void SetDataShow()
        {

            if (SetFlag.RECEIVE_FLAG == SetFlag.GET_SET_REG_STA)
            {
                if (cparameter.ShowInformation[1] == 0x03 && cparameter.ShowInformation[2] == cparameter.SET_STA_NUM * 2)
                {
                    if (SetFlag.CONTROL_FLAG != SetFlag.SET_SET_REG_STA)
                    {
                        SetFlag.CONTROL_FLAG = SetFlag.GET_OPER_REG_STA;
                    }
                   
                   
                    short Count = 3;

                    short[] Data = new short[cparameter.SET_STA_NUM];

                    for (int i = 0; i < cparameter.SET_STA_NUM; i++)
                    {
                        Data[i] = modbus.TwoToWord(cparameter.ShowInformation[Count++], cparameter.ShowInformation[Count++]);
                    }
                    cparameter.ShowInformation.Clear();                              //清空接收缓存区
                    cparameter.ShowInformationCrt = 0;

                    if (Data[0] == 1) { comboBox2.Text = "制冷"; } else { comboBox2.Text = "制热"; }
                     if (Data[1] == 1) { comboBox1.Text = "开机"; } else { comboBox1.Text = "关机"; }
                     numericUpDown1.Value = Data[2] / 10;
                     numericUpDown2.Value = Data[3] / 10;
                     if (Data[4] == 1) { comboBox3.Text = "禁止"; } else { comboBox3.Text = "允许"; }

                    label10.Text = "机组参数读取成功！";
                    label10.ForeColor = Color.DarkBlue;

                    label10.Font = new System.Drawing.Font("宋体", 9);

                }

            }
        }

        //设置参数响应分析
        private void SetDataCheck()
        {
            if (SetFlag.RECEIVE_FLAG == SetFlag.SET_SET_REG_STA)
            {
                if (modbus.TwoToWord(cparameter.ShowInformation[2], cparameter.ShowInformation[3]) == 30001)
                {
                    if (SetFlag.CONTROL_FLAG != SetFlag.GET_SET_REG_STA)
                    {
                        SetFlag.CONTROL_FLAG = SetFlag.GET_OPER_REG_STA;

                    }



                    cparameter.ShowInformation.Clear();                              //清空接收缓存区
                    cparameter.ShowInformationCrt = 0;
                    label10.Text = "机组参数设定成功！";
                    label10.ForeColor = Color.DarkBlue;

                    label10.Font = new System.Drawing.Font("宋体", 9);
                    
                }

            }

        }
        //故障显示
        private void ErrDataShow()
        {
            if (SetFlag.RECEIVE_FLAG == SetFlag.GET_ERR_REG_STA)
            {
                int Num = cparameter.ERR_STA_NUM;
                int NumCount = (Num%8 == 0)?(Num/8):(Num/8)+1;

                if (cparameter.ShowInformation[1] == 0x01 && cparameter.ShowInformation[2] == NumCount)
                {
                    if (SetFlag.CONTROL_FLAG != SetFlag.SET_SET_REG_STA && SetFlag.CONTROL_FLAG != SetFlag.GET_SET_REG_STA)
                    {
                        SetFlag.CONTROL_FLAG = SetFlag.GET_DO_REG_STA;
                    }
                  
                    short Count = 3;
                    short ShowNum = 0;
                    byte NumData = cparameter.ShowInformation[2];
                    byte[] Data = new byte[NumData];

                    for (int i = 0; i < NumData; i++)
                    {
                        Data[i] = cparameter.ShowInformation[Count++];
                    }

                    cparameter.ShowInformation.Clear();                              //清空接收缓存区
                    cparameter.ShowInformationCrt = 0;

                    if (tabControl1.SelectedIndex == 1)
                    {
                       
                       for (int j = 0; j < NumData; j++)
                        {
                          for (int i = 0; i < 8; i++)
                            {
                             if (ShowNum < cparameter.ERR_STA_NUM)
                             {
                                    if (((Data[j] >> i) & 0x01) == 0x01)
                                    {
                                       cparameter.ERRValue[ShowNum].Text = "故障";
                                       cparameter.ERRValue[ShowNum].ForeColor = Color.Red;
                                     }
                                    else
                                    {
                                     cparameter.ERRValue[ShowNum].Text = "正常";
                                     cparameter.ERRValue[ShowNum].ForeColor = Color.DarkBlue;
                                    }
                                        ShowNum++;
                             } 

                          }

                       }
       
                    }

                    if (ExcelOpen == 1)
                    {
                        ShowNum = 0;
                        workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(2);//获得第i个sheet，准备写入  
                        Int64 num = workSheet.UsedRange.Rows.Count;

                        workSheet.Cells[num + 1, 1] = Convert.ToString(DateTime.Now.ToString("HH:mm:ss"));

                        for (int j = 0; j < NumData; j++)
                        {
                            for (int i = 1; i < 9; i++)
                            {
                                if (ShowNum < cparameter.ERR_STA_NUM)
                                {
                                    workSheet.Cells[num + 1, ShowNum + 2] = Convert.ToString((Data[j] >> (i - 1)) & 0x01);

                                }
                                ShowNum++;
                            }
                        }

                    }



                }

            }
        }
        //DO显示
        private void DoDataShow()
        {
            if (SetFlag.RECEIVE_FLAG == SetFlag.GET_DO_REG_STA)
            {
                int Num = cparameter.DO_STA_NUM;
                int NumCount = (Num % 8 == 0) ? (Num / 8) : (Num / 8) + 1;
                byte NumData = cparameter.ShowInformation[2];

                if (cparameter.ShowInformation[1] == 0x01 && cparameter.ShowInformation[2] == NumCount)
                {
                    if (SetFlag.CONTROL_FLAG != SetFlag.SET_SET_REG_STA && SetFlag.CONTROL_FLAG != SetFlag.GET_SET_REG_STA)
                    {
                        SetFlag.CONTROL_FLAG = SetFlag.GET_SOFTWARE_REG_STA;
                    }
                  
                    short Count = 3;
                    short ShowNum = 0;
                    byte[] Data = new byte[NumData];

                    for (int i = 0; i < NumData; i++)
                    {
                        Data[i] = cparameter.ShowInformation[Count++];
                    }

                    cparameter.ShowInformation.Clear();                              //清空接收缓存区
                    cparameter.ShowInformationCrt = 0;

                    if (tabControl1.SelectedIndex == 2)
                    {
                       
                                 for (int j = 0; j < NumData; j++)
                                 {
                                     for (int i = 0; i < 8; i++)
                                     {
                                         if (ShowNum < cparameter.DO_STA_NUM)
                                         {
                                             if (((Data[j] >> i) & 0x01) == 0x01)
                                             {
                                                 cparameter.DOValue[ShowNum].Text = "开";
                                                 cparameter.DOValue[ShowNum].ForeColor = Color.Red;
                                             }
                                             else
                                             {
                                                 cparameter.DOValue[ShowNum].Text = "关";
                                                 cparameter.DOValue[ShowNum].ForeColor = Color.DarkBlue;
                                             }
                                             ShowNum++;
                                         }

                                     }

                                 }

                    }


                    if (ExcelOpen == 1)
                    {
                        ShowNum = 0;
                        workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(3);//获得第i个sheet，准备写入  
                        Int64 num = workSheet.UsedRange.Rows.Count;

                        workSheet.Cells[num + 1, 1] = Convert.ToString(DateTime.Now.ToString("HH:mm:ss"));

                        for (int j = 0; j < NumData; j++)
                        {
                            for (int i = 1; i < 9; i++)
                            {
                                if (ShowNum < cparameter.DO_STA_NUM)
                                {
                                    workSheet.Cells[num + 1, ShowNum + 2] = Convert.ToString((Data[j] >> (i - 1)) & 0x01);

                                }
                                ShowNum++;
                            }
                        }

                    }



                }

            }

        }
        //参数二显示
        private void Oper2DataShow()
        {
            if (SetFlag.RECEIVE_FLAG == SetFlag.GET_OPER2_REG_STA)
            {
                if (cparameter.ShowInformation[1] == 0x03 && cparameter.ShowInformation[2] == cparameter.OPER2_STA_NUM * 2)
                {
                    if (SetFlag.CONTROL_FLAG != SetFlag.SET_SET_REG_STA && SetFlag.CONTROL_FLAG != SetFlag.GET_SET_REG_STA)
                    {
                        SetFlag.CONTROL_FLAG = SetFlag.GET_OPER_REG_STA;
                    }
                   
                    short Count = 3;
                    short ShowNum = 0;
                    short[] Data = new short[cparameter.OPER2_STA_NUM];
                    for (int i = 0; i < cparameter.OPER2_STA_NUM; i++)
                    {
                        Data[i] = modbus.TwoToWord(cparameter.ShowInformation[Count++], cparameter.ShowInformation[Count++]);
                    }

                    cparameter.ShowInformation.Clear();                              //清空接收缓存区
                    cparameter.ShowInformationCrt = 0;

                    if (tabControl1.SelectedIndex == 3)
                    {
                       
                                ccontrols.RefreshAnalog(cparameter.OPER2Value[ShowNum], Data[ShowNum++], "℃");
                                ccontrols.RefreshAnalog(cparameter.OPER2Value[ShowNum], Data[ShowNum++], "℃");
                                ccontrols.RefreshAnalog(cparameter.OPER2Value[ShowNum], Data[ShowNum++], "℃");
                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum++]);
                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum++]);
                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum++]);
                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum++]);
                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum++]);

                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum++]) + "Hz";
                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum++]) + "Hz";

                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum]) + "%";
                                ShowNum++;


                                ccontrols.RefreshAnalog(cparameter.OPER2Value[ShowNum], Data[ShowNum++], "℃");
                                ccontrols.RefreshAnalog(cparameter.OPER2Value[ShowNum], Data[ShowNum++], "℃");
                                ccontrols.RefreshAnalogEEV(cparameter.OPER2Value[ShowNum], Data[ShowNum++], " B");
                                ccontrols.RefreshAnalogEEV(cparameter.OPER2Value[ShowNum], Data[ShowNum++], " B");

                                // cparameter.OPER2Value[ShowNum].Text = (Data[ShowNum] == 0x01) ? "本地允许" : ((Data[ShowNum] == 0x02) ? "远程" : ((Data[ShowNum] == 0x03) ? "本地允许+远程" : "未检测"));
                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum]);
                                ShowNum++;
                                ccontrols.RefreshAnalog(cparameter.OPER2Value[ShowNum], Data[ShowNum++], "A");
                                ccontrols.RefreshAnalog(cparameter.OPER2Value[ShowNum], Data[ShowNum++], "A");
                                ccontrols.RefreshAnalog(cparameter.OPER2Value[ShowNum], Data[ShowNum++], "℃");
                                ccontrols.RefreshAnalog(cparameter.OPER2Value[ShowNum], Data[ShowNum++], "℃");

                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum++]) + "V";


                                cparameter.OPER2Value[ShowNum].Text = (Data[ShowNum] == 0x00) ? "无" : ((Data[ShowNum] == 0x01) ? "一级" : "二级");
                                ShowNum++;
                                for (int i = 0; i < 6; i++)
                                {
                                    cparameter.OPER2Value[ShowNum].Text = (Data[ShowNum] == 0x01) ? "不升频" : ((Data[ShowNum] == 0x02) ? "降频" : ((Data[ShowNum] == 0x03) ? "停机" : "正常"));
                                    ShowNum++;
                                }

                                cparameter.OPER2Value[ShowNum].Text = (Data[ShowNum] == 0x01) ? "三花" : ((Data[ShowNum] == 0x02) ? "视源" : ((Data[ShowNum] == 0x03) ? "海悟" : "未检测"));
                                ShowNum++;

                                cparameter.OPER2Value[ShowNum].Text = Convert.ToString(Data[ShowNum]);
                                // cparameter.OPER2Value[ShowNum].Text = (Data[ShowNum] == 0x01) ? "所有故障" : ((Data[ShowNum] == 0x02) ? "停水泵" : ((Data[ShowNum] == 0x04) ? "停压缩机": (Data[ShowNum] == 0x08) ?"停电加热":"正常"));

                      
                    }


                    if (ExcelOpen == 1)
                    {
                        workSheet = (Excel.Worksheet)workBook.Worksheets.get_Item(4);//获得第i个sheet，准备写入  
                        Int64 num = workSheet.UsedRange.Rows.Count;

                        workSheet.Cells[num + 1, 1] = Convert.ToString(DateTime.Now.ToString("HH:mm:ss"));
                        for (int i = 0; i < cparameter.OPER2_STA_NUM; i++)
                        {
                            if (i < 3)
                            {
                                workSheet.Cells[num + 1, i + 2] = Convert.ToString((float)Data[i] / 10);
                            }
                            else if (i >= 3 && i <= 10)
                            {
                                workSheet.Cells[num + 1, i + 2] = Convert.ToString(Data[i]);
                            }
                            else if (i >= 11 && i <= 12)
                            {
                                workSheet.Cells[num + 1, i + 2] = Convert.ToString(((float)Data[i]) / 10);
                            }
                            else if (i >= 13 && i <= 15)
                            {
                                workSheet.Cells[num + 1, i + 2] = Convert.ToString(Data[i]);
                            }
                            else if (i >= 16 && i <= 19)
                            {
                                workSheet.Cells[num + 1, i + 2] = Convert.ToString((float)Data[i] / 10);
                            }
                            else if (i >= 20 && i < cparameter.OPER2_STA_NUM)
                            {
                                workSheet.Cells[num + 1, i + 2] = Convert.ToString(Data[i]);
                            }

                        }

                    }

                }




            }
        }

        //版本号显示
        private void SoftWareDataShow()
        {
            if (SetFlag.RECEIVE_FLAG == SetFlag.GET_SOFTWARE_REG_STA)
            {
                if (cparameter.ShowInformation[1] == 0x03 && cparameter.ShowInformation[2] == 15 * 2)
                {
                    if (SetFlag.CONTROL_FLAG != SetFlag.SET_SET_REG_STA && SetFlag.CONTROL_FLAG != SetFlag.GET_SET_REG_STA)
                    {
                        SetFlag.CONTROL_FLAG = SetFlag.GET_OPER2_REG_STA;
                    }
                   
                    int num = 0;

                    byte[] byteArray = new byte[30];


                    for (int i = 3; i < 33; i += 2)
                    {

                        byteArray[num] = cparameter.ShowInformation[i + 1];
                        num++;
                        byteArray[num] = cparameter.ShowInformation[i];
                        num++;
                    }

                    cparameter.ShowInformation.Clear();                              //清空接收缓存区
                    cparameter.ShowInformationCrt = 0;

                    if (tabControl1.SelectedIndex == 0)
                    {
                       
                                 label12.Text = "软件版本：";
                                 label12.Text += "内机：";
                                 for (num = 0; num < 10; num++)
                                 {
                                     label12.Text += Chr(byteArray[num]);
                                 }
                                 label12.Text += "外机：";
                                 for (num = 10; num < 20; num++)
                                 {
                                     label12.Text += Chr(byteArray[num]);
                                 }
                                 label12.Text += "驱动：";
                                 for (num = 20; num < 30; num++)
                                 {
                                     label12.Text += Chr(byteArray[num]);
                                 }

                       
                    }
                }
            }
           
        }

       
        private string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
            }
        }

        //数据接收
        private void DataReceive()
        {



            //if (ExcelOpen == 1 && DateTime.Now.ToString("ss") == "00")
            //{
            //    ExcelExit();
            //}
            //else if (ExcelOpen == 0 && DateTime.Now.ToString("ss") == "01")
            //{
            //    ExcelCreat();
            //}

            if (ExcelOpen == 1 && Convert.ToByte(DateTime.Now.ToString("HH")) % 6 == 0 && DateTime.Now.ToString("mmss") == "0000")
            {
                ExcelExit();
            }
            else if (ExcelOpen == 0 && Convert.ToByte(DateTime.Now.ToString("HH")) % 6 == 0 && DateTime.Now.ToString("mmss") == "0030")
            {
                ExcelCreat();
            }


            if (serialPort1.IsOpen)
            {
                int rxCnt = 0;
                rxCnt = serialPort1.BytesToRead;
                if (rxCnt != 0)
                {
                    byte[] TempBuffer = new byte[rxCnt];
                    serialPort1.Read(TempBuffer, 0, rxCnt);


                    if (SHOW == 1)
                    {
                        str = ToHexString(TempBuffer);

                        dataflowform.richTextBox1.ForeColor = Color.DarkBlue;
                        dataflowform.richTextBox1.Text += DateTime.Now.ToString("[RX_HH:mm:ss] ") + str + "\r\n";
              
                    }
                    else
                    {
                        str = "";
                    }

                    for (int i = 0; i < rxCnt; i++)
                    {
                        cparameter.ShowInformation.Insert(cparameter.ShowInformationCrt, TempBuffer[i]);
                        cparameter.ShowInformationCrt++;
                    }
                    if (cparameter.ShowInformation[0] == MailAddr)
                    {
                        ushort Local_CHKSUM;
                        Local_CHKSUM = (ushort)(modbus.TwoToWord(cparameter.ShowInformation[cparameter.ShowInformationCrt - 1], cparameter.ShowInformation[cparameter.ShowInformationCrt - 2]));
                        modbus.checksum = modbus.Crc16(cparameter.ShowInformation, 0, (ushort)(cparameter.ShowInformationCrt - 3));

                        if (Local_CHKSUM == modbus.checksum)
                        {
                            Invoke(new MethodInvoker(
                                () =>
                                {
                                    ErrorDataFlag = 0;
                                    label7.Text = "通讯状态：正常";
                                    label7.ForeColor = Color.Black;
                                    DataAnalyze();
                                }));
                        }
                        else
                        {
                           // ErrorDataCount ++;   //校验错误！
                           
                        }
                        cparameter.ShowInformation.Clear();                              //清空接收缓存区
                        cparameter.ShowInformationCrt = 0;
                    }
                    else
                    {
                         
                          cparameter.ShowInformation.Clear();                              //清空接收缓存区
                          cparameter.ShowInformationCrt = 0;
                    }
                }
                else
                {
                     ErrorDataFlag ++;
                    if (ErrorDataFlag == 20)//拔线滤波
                    {
                        ErrorDataFlag = 0;
                        label7.Text = "通讯状态：断开";
                        label7.ForeColor = Color.Red;

                        for (int i = 0; i < cparameter.OPER_STA_NUM; i++)
                        {
                            cparameter.OperValue[i].Text = "未检测";
                        }
                        for (int i = 0; i < cparameter.ERR_STA_NUM; i++)
                        {
                            cparameter.ERRValue[i].Text = "未检测";
                        }
                        for (int i = 0; i < cparameter.DO_STA_NUM; i++)
                        {
                            cparameter.DOValue[i].Text = "未检测";
                        }
                        for (int i = 0; i < cparameter.OPER2_STA_NUM; i++)
                        {
                            cparameter.OPER2Value[i].Text = "未检测";
                        }
                    }

                   

                }


            }
            else
            {
                Invoke(new MethodInvoker(
                    () =>
                    {

                        label7.Text = "通讯状态：断开";
                        label7.ForeColor = Color.Red;
                        for (int i = 0; i < cparameter.OPER_STA_NUM; i++)
                        {
                            cparameter.OperValue[i].Text = "未检测";
                        }
                        for (int i = 0; i < cparameter.ERR_STA_NUM; i++)
                        {
                            cparameter.ERRValue[i].Text = "未检测";
                        }
                        for (int i = 0; i < cparameter.DO_STA_NUM; i++)
                        {
                            cparameter.DOValue[i].Text = "未检测";
                        }
                        for (int i = 0; i < cparameter.OPER2_STA_NUM; i++)
                        {
                            cparameter.OPER2Value[i].Text = "未检测";
                        }

                    }));
            }
        }

        //设置参数获取
        private void SetValueGet()
        {
            if (comboBox1.Text == "开机")
            {
                cparameter.SetValue[1] = 1;
            }
            else
            {
                cparameter.SetValue[1] = 0;
            }

            if (comboBox2.Text == "制冷")
            {
                cparameter.SetValue[0] = 1;
            }
            else
            {
                cparameter.SetValue[0] = 2;
            }

            cparameter.SetValue[2] = (ushort)(numericUpDown1.Value * 10);
            cparameter.SetValue[3] = (ushort)(numericUpDown2.Value * 10);
           

            if (comboBox3.Text == "允许")
            {
                cparameter.SetValue[4] = 0;
            }
            else
            {
                cparameter.SetValue[4] = 1;
            }

        }


        private void 串口设置ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (Childflag == 0)
            {
                Childflag = 1;
                SerialSetForm serialset = new SerialSetForm(this);
                serialset.Text = "串口设置";
                serialset.Show();
            }
        }
        private void 通信数据流ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Childflag == 0)
            {
                Childflag = 1;
                DataFlowForm dataflowset = new DataFlowForm(this);

                dataflowform = dataflowset;
                dataflowform.Text = "通信数据流";
                dataflowform.Show();
                SHOW = 1;
            }
        }
        private void 关于软件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Childflag == 0)
            {
                Childflag = 1;
                SoftWareForm softwareshow = new SoftWareForm(this);
                softwareshow.Text = "关于软件";
                softwareshow.Show();
            }
        }

        public string ToHexString(byte[] bytes)
        {
            string hexString = string.Empty;
            if (bytes != null)
            {
                StringBuilder strB = new StringBuilder();
                for (int i = 0; i < bytes.Length; i++)
                {
                    for (int j = 0; j < 1; j++)
                    {
                        strB.Append(bytes[i].ToString("X2") + " ");
                    }
                }
                hexString = strB.ToString();
            }
            return hexString;
        }



        private void timer1_Tick(object sender, EventArgs e)
        {
            DataReceive();
            RequestStep();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
           
            

            DateTimePicker NowTime = new DateTimePicker();

            NowTime.Format = DateTimePickerFormat.Long;
            NowTime.Format = DateTimePickerFormat.Time;
            label1.Text = NowTime.Value.ToString();
            label1.Font = new System.Drawing.Font("宋体", 10);

            label11.Text = "解析出错：" + Convert.ToString(ErrorDataCount) + "次";

            if (label10.Text != "")
            {
                Labe10ShowTime++;

                if (Labe10ShowTime == 10)
                {
                    Labe10ShowTime = 0;
                    label10.Text = "";
                }
            }
           
            
        }
        
        private void numericUpDown1_KeyPress(object sender, KeyPressEventArgs e)
        {
          
        }

        private void numericUpDown2_KeyPress(object sender, KeyPressEventArgs e)
        {
           
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SetValueGet();
          
            SetFlag.CONTROL_FLAG = SetFlag.SET_SET_REG_STA;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SetFlag.CONTROL_FLAG = SetFlag.GET_SET_REG_STA;
           
        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
