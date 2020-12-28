using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CQ
{
    delegate void ReadNewLine(int nPort, string str);
    class SerialPortService
    {
        private static SerialPortService _intance = null;
        private static readonly object _lock = new object();
        public static SerialPortService Instance
        {
            get
            {
                if (_intance == null)
                {
                    lock (_lock)
                    {
                        if (_intance == null)
                        {
                            _intance = new SerialPortService();
                        }
                    }
                }
                return _intance;
            }
        }

        private SerialPort serialPort1 = null;
        public ReadNewLine SerialPort_ReadNewLine = null;

        private SerialPort Open(string szSection)
        {
            SerialPort serialPort = new SerialPort();
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            string PortName = IniService.Instance.ReadIniData(szSection, "串口号", "COM1", str + "Config.ini");
            string BaudRate = IniService.Instance.ReadIniData(szSection, "波特率", "115200", str + "Config.ini");
            string Parity = IniService.Instance.ReadIniData(szSection, "校验位", "0", str + "Config.ini");
            string DataBits = IniService.Instance.ReadIniData(szSection, "数据位", "8", str + "Config.ini");
            string StopBits = IniService.Instance.ReadIniData(szSection, "停止位", "1", str + "Config.ini");

            serialPort.PortName = PortName;
            if (int.TryParse(BaudRate, out int nBaudRate) == false)
            {
                MessageBox.Show("波特率设置错误!");
                return null;
            }
            serialPort.BaudRate = nBaudRate;
            if (int.TryParse(Parity, out int nParity) == false)
            {
                MessageBox.Show("校验位设置错误!");
                return null;
            }
            serialPort.Parity = (System.IO.Ports.Parity)nParity;
            if (int.TryParse(DataBits, out int nDataBits) == false)
            {
                MessageBox.Show("数据位设置错误!");
                return null;
            }
            serialPort.DataBits = nDataBits;
            if (int.TryParse(StopBits, out int nStopBits) == false)
            {
                MessageBox.Show("停止位设置错误!");
                return null;
            }
            serialPort.StopBits = (System.IO.Ports.StopBits)nStopBits;
            serialPort.NewLine = "\r";
            serialPort.DataReceived += SerialPort_DataReceived;
            try
            {
                serialPort.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            return serialPort;
        }
        private SerialPortService()
        {
            serialPort1 = Open("扫码枪1");
        }

        private void SerialPort_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort serialPort = sender as SerialPort;
            if (serialPort == serialPort1)
            {
                SerialPort_ReadNewLine?.Invoke(1, serialPort.ReadLine());
            }
        }

        public void OpenSerialPort()
        {
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

        }
    }
}
