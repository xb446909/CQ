using System;
using System.Collections.Generic;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            string BuadRate = IniService.Instance.ReadIniData("扫码枪", "", "DB2.0", str + "Config.ini");
        }

        public void OpenSerialPort()
        {
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

        }
    }
}
