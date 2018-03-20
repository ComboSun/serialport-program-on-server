using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;

using System.ServiceProcess;
using System.Text;

using System.Data.SqlClient;
using System.Timers;
using System.Threading;
using Microsoft.Win32;
using System.IO;

namespace SerialPortService
{
    public partial class Service1 : ServiceBase
    {        
        System.Timers.Timer aTimer = new System.Timers.Timer();
        System.Timers.Timer bTimer = new System.Timers.Timer();
        RegistryKey rsg = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\CZDevice", true);        
        SqlConnection conn = new SqlConnection("server=192.168.1.6;database=scz;uid=devicemanage;pwd=scz640318!");
        string regInfo;
        string adr;
        string AddressId;
        int DeviceCount;

        public Service1()
        {
            InitializeComponent();
            aTimer.Interval = 1000;
            bTimer.Interval = 5000;//定时5s
            aTimer.Enabled = true;
            aTimer.Elapsed += new ElapsedEventHandler(aTimer_Elapsed);
            bTimer.Enabled = true;
            bTimer.Elapsed += new ElapsedEventHandler(bTimer_Elapsed);
            serialPort1.WriteTimeout = 1000;
            serialPort1.ReadTimeout = 1000;
        }

        protected override void OnStart(string[] args)
        {
            open();
        }

        private void open()
        {
            try
            {
                if (rsg.GetValue("PortName") != null)
                {
                    string Portname = rsg.GetValue("PortName").ToString();
                    serialPort1.PortName = Portname;
                }
                else
                {
                    rsg.Close();
                    WriteLog("无串口号信息，需创建注册表为PortName赋值。");
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            try
            {
                if (rsg.GetValue("BaudRate") != null)
                {
                    string Baudrate = rsg.GetValue("BaudRate").ToString();
                    serialPort1.BaudRate = Convert.ToInt32(Baudrate);
                }
                else
                {
                    rsg.Close();
                    WriteLog("无波特率信息，需创建注册表为BaudRate赋值。");
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            try
            {
                if (rsg.GetValue("DeviceCount") != null)
                {
                    DeviceCount = Convert.ToInt32(rsg.GetValue("DeviceCount"));
                }
                else
                {
                    rsg.Close();
                    WriteLog("无设备数量信息，需创建注册表为DeviceCount赋值。");
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
            try
            {
                serialPort1.Open();
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
                bTimer.Enabled = false;
            }
            if (serialPort1.IsOpen)
            {
                WriteLog("串口已打开！");
            }
            serialPort1.DiscardInBuffer();
        }

        //定时5s发送一次数据
        private void bTimer_Elapsed(object sender, ElapsedEventArgs e)
        {          
            for (int i = 0; i < DeviceCount; i++)
            {
                regInfo = @"SOFTWARE\CZDevice\Device"+i.ToString();
                RegistryKey rsg1 = Registry.LocalMachine.OpenSubKey(regInfo, true);                
                try
                {
                    if (rsg1.GetValue("DeviceAddress") != null)
                    {
                       adr = "0" + rsg1.GetValue("DeviceAddress").ToString();
                    }
                    else
                    {
                       rsg1.Close();
                       WriteLog("无设备地址信息，需创建注册表为DeviceAddress赋值。");
                    }
                }
                catch (Exception ex)
                {
                    WriteLog(ex.Message);
                }
                try
                {
                    if (rsg1.GetValue("DeviceName") != null)
                    {
                       AddressId = rsg1.GetValue("DeviceName").ToString();
                    }
                    else
                    {
                       rsg1.Close();
                       WriteLog("无设备名称信息，需创建注册表为DeviceName赋值。");
                    }
                }
                catch (Exception ex)
                {
                    WriteLog(ex.Message);
                }
                int buffersize = 13; //十六进制数的大小
                byte[] buffer = new Byte[buffersize]; //创建缓冲区
                byte[] dizhi = strToToHexByte(adr);
                byte[] CRC = new byte[2];
                byte[] data = new byte[8]; //例：{ 0x01,0x03,0x00,0x00,0x00,0x04,0x44,0x09 };
                data[0] = dizhi[0];
                data[1] = 0x03;
                data[2] = 0x00;
                data[3] = 0x00;
                data[4] = 0x00;
                data[5] = 0x04;
                //生成CRC码
                GetCRC(data, ref CRC);
                data[6] = CRC[0];
                data[7] = CRC[1];
                
                try
                {
                    serialPort1.Write(data, 0, data.Length);
                }
                catch (Exception ex)
                {
                    WriteLog(ex.Message + AddressId + "数据发送异常！");
                    serialPort1.Close();
                    open();
                    bTimer_Elapsed(sender, e);
                }
                Thread.Sleep(100);
                //从缓冲区接收数据，获取电流、电压、功率值
                try
                {
                    serialPort1.Read(buffer, 0, buffersize);
                }
                catch (Exception ex)
                {
                    WriteLog(ex.Message + AddressId + "数据接收异常！");                                   
                    if (ex == null)
                    {
                        open();
                        bTimer_Elapsed(sender, e);
                    }
                }
                serialPort1.DiscardInBuffer();
                string ss = null;
                ss = ByteToHexStr(buffer);//字节数组转换为16进制字符串
                string zjs = ss.Trim().Substring(4, 2);
                if (zjs.Equals("08"))//判断接收数据字节数是否为08
                {
                    //获取电压
                    string dy = ss.Trim().Substring(8, 2);
                    int dec1 = Convert.ToInt32(dy, 16);
                    //获取电流
                    string dl = ss.Trim().Substring(12, 2);
                    int dec2 = Convert.ToInt32(dl, 16);
                    Decimal dd1 = Math.Round((decimal)dec2 / 100, 1);
                    //获取功率
                    string gl = ss.Trim().Substring(16, 2);
                    int dec3 = Convert.ToInt32(gl, 16);
                    Decimal dd2 = Math.Round((decimal)dec3 / 1000, 2);
                    string sql = "insert into DeviceData (AddressId,OperateDate,DianYa,DianLiu,GongLv) values('" + AddressId + "','" + DateTime.Now + "','" + dec1 + "','" + dd1 + "','" + dd2 + "')";
                    SqlDataAdapter sda = new SqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    sda.Fill(ds);
                 }                               
            }
        }

        private void aTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                int intHour = e.SignalTime.Hour;
                int intMinute = e.SignalTime.Minute;
                int intSecond = e.SignalTime.Second;
                int iHour = 00;
                int iMinute = 00;
                int iSecond = 00;
                if (intHour == iHour && intMinute == iMinute && intSecond == iSecond)
                {
                    string sql = @"insert into DeviceGh 
select addressid,convert(varchar(12),DATEADD(DAY,-1,GETDATE()),111) as fdate,sum(gonglv/3600*5) as ydl from DeviceData where datediff(day,operatedate,getdate())=1 group by addressid";
                    SqlDataAdapter sda = new SqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    sda.Fill(ds);
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
                aTimer.Enabled = false;
            }
        }

        //函数,字节数组转16进制字符串
        public static string ByteToHexStr(byte[] bytes)
        {
            string returnStr = "";
            if (bytes != null)
            {
                for (int i = 0; i < bytes.Length; i++)
                {
                    returnStr += bytes[i].ToString("X2");
                }
            }
            return returnStr;
        }

        //字符串转16进制字符数组
        private static byte[] strToToHexByte(string hexString)
        {
            hexString = hexString.Replace(" ", "");
            if ((hexString.Length % 2) != 0)
                hexString += " ";
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            return returnBytes;
        }

        //生成CRC码
        private void GetCRC(byte[] message, ref byte[] CRC)
        {
            ushort CRCFull = 0xFFFF;
            byte CRCHigh = 0xFF, CRCLow = 0xFF;
            char CRCLSB;
            for (int i = 0; i < message.Length - 2; i++)
            {
                CRCFull = (ushort)(CRCFull ^ message[i]);
                for (int j = 0; j < 8; j++)
                {
                    CRCLSB = (char)(CRCFull & 0x0001);
                    //下面两句所得结果一样
                    //CRCFull = (ushort)(CRCFull >> 1);
                    CRCFull = (ushort)((CRCFull >> 1) & 0x7FFF);
                    if (CRCLSB == 1)
                        CRCFull = (ushort)(CRCFull ^ 0xA001);
                }
            }
            CRC[1] = CRCHigh = (byte)((CRCFull >> 8) & 0xFF);
            CRC[0] = CRCLow = (byte)(CRCFull & 0xFF);
        }

        protected override void OnStop()
        {
            if (serialPort1.IsOpen)
            {
                serialPort1.DiscardOutBuffer();
                serialPort1.Close();
            }
            WriteLog("串口已关闭，Windows已停止服务！");
        }

        //写日志；
        public static void WriteLog(string strLog)
        {
            string LogAddress = "";
            string path = Environment.CurrentDirectory.ToString()+'\\'+"CZsp_Log";
            //如果目录下无 CZsp_log 的文件夹，那么新建此文件夹
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            //如果日志文件为空，则默认在Debug目录下新建 YYYY-mm-dd_Log.log文件
            if (LogAddress == "")
            {
                LogAddress = path +  '\\' +
                    DateTime.Now.Year + '-' +
                    DateTime.Now.Month + '-' +
                    DateTime.Now.Day + "_Log.log";
            }
            StreamWriter fs = new StreamWriter(LogAddress, true);
            fs.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + "   ---   " + strLog);
            fs.Close();
        }
    }
}
