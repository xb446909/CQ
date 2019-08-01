using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CQ
{
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

        SerialPort serialPort = null;

        private SerialPortService()
        {
            serialPort = new SerialPort();
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            string PortName = IniService.Instance.ReadIniData("扫码枪", "串口号", "COM1", str + "Config.ini");
            string BaudRate = IniService.Instance.ReadIniData("扫码枪", "波特率", "115200", str + "Config.ini");
            string Parity = IniService.Instance.ReadIniData("扫码枪", "校验位", "0", str + "Config.ini");
            string DataBits = IniService.Instance.ReadIniData("扫码枪", "数据位", "8", str + "Config.ini");
            string StopBits = IniService.Instance.ReadIniData("扫码枪", "停止位", "1", str + "Config.ini");

            serialPort.PortName = PortName;
            if (int.TryParse(BaudRate, out int nBaudRate) == false)
            {
                MessageBox.Show("波特率设置错误!");
                return;
            }
            serialPort.BaudRate = nBaudRate;
            if (int.TryParse(Parity, out int nParity) == false)
            {
                MessageBox.Show("校验位设置错误!");
                return;
            }
            serialPort.Parity = (System.IO.Ports.Parity)nParity;
            if (int.TryParse(DataBits, out int nDataBits) == false)
            {
                MessageBox.Show("数据位设置错误!");
                return;
            }
            serialPort.DataBits = nDataBits;

        }

        public void OpenSerialPort()
        {
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

        }
    }
}
