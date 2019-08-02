using HslCommunication;
using HslCommunication.Profinet.Siemens;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CQ
{
    public class PLCService
    {
        private static PLCService _intance = null;
        private static readonly object _lock = new object();
        public static PLCService Instance
        {
            get
            {
                if (_intance == null)
                {
                    lock (_lock)
                    {
                        if (_intance == null)
                        {
                            _intance = new PLCService();
                        }
                    }
                }
                return _intance;
            }
        }

        private PLCService()
        {
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            string type = IniService.Instance.ReadIniData("PLC", "Type", "S1200", str + "Config.ini");

            SiemensPLCS siemensPLCS = 0;
            switch (type)
            {
                case "S1200":
                    siemensPLCS = SiemensPLCS.S1200;
                    break;
                case "S300":
                    siemensPLCS = SiemensPLCS.S300;
                    break;
                case "S1500":
                    siemensPLCS = SiemensPLCS.S1500;
                    break;
                case "S200Smart":
                    siemensPLCS = SiemensPLCS.S200Smart;
                    break;
                case "S200":
                    siemensPLCS = SiemensPLCS.S200;
                    break;
                default:
                    break;
            }

            siemensTcpNet = new SiemensS7Net(siemensPLCS);
        }

        private SiemensS7Net siemensTcpNet = null;

        public bool Connect()
        {
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            string Ip = IniService.Instance.ReadIniData("PLC", "Ip", "127.0.0.1", str + "Config.ini");
            string port = IniService.Instance.ReadIniData("PLC", "Port", "102", str + "Config.ini");

            if (!System.Net.IPAddress.TryParse(Ip, out System.Net.IPAddress address))
            {
                return false;
            }

            if (!int.TryParse(port, out int nPort))
            {
                return false;
            }

            siemensTcpNet.IpAddress = Ip;
            siemensTcpNet.Port = nPort;

            try
            {
                string type = IniService.Instance.ReadIniData("PLC", "Type", "S1200", str + "Config.ini");
                if (type != "S200Smart")
                {
                    string Rack = IniService.Instance.ReadIniData("PLC", "Rack", "0", str + "Config.ini");
                    string Slot = IniService.Instance.ReadIniData("PLC", "Slot", "0", str + "Config.ini");
                    siemensTcpNet.Rack = byte.Parse(Rack);
                    siemensTcpNet.Slot = byte.Parse(Slot);
                }


                OperateResult connect = siemensTcpNet.ConnectServer();
                if (connect.IsSuccess)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return true;
        }

        public UInt32 ReadUInt32(string addr)
        {
            var val = siemensTcpNet.ReadUInt32(addr);
            if (val.IsSuccess)
            {
                return val.Content;
            }
            return 0;
        }

        public double ReadDouble(string addr)
        {
            var val = siemensTcpNet.ReadDouble(addr);
            if (val.IsSuccess)
            {
                return val.Content;
            }
            return 0.0;
        }

        public bool ReadBool(string addr)
        {
            var val = siemensTcpNet.ReadBool(addr);
            if (val.IsSuccess)
            {
                return val.Content;
            }
            return false;
        }

        public float ReadFloat(string addr)
        {
            var val = siemensTcpNet.ReadFloat(addr);
            if (val.IsSuccess)
            {
                return (val.Content / 100.0f);
            }
            return 0.0f;
        }

        public void WriteUInt32(string addr, UInt32 val)
        {
            var ret = siemensTcpNet.Write(addr, val);
            if (!(ret.IsSuccess))
            {
                MessageBox.Show("写入地址:" + addr + "失败!");
            }
        }
    }
}
