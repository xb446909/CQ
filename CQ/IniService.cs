using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CQ
{
    public class IniService
    {
        private static IniService _intance = null;
        private static readonly object _lock = new object();
        public static IniService Instance
        {
            get
            {
                if (_intance == null)
                {
                    lock(_lock)
                    {
                        if (_intance == null)
                        {
                            _intance = new IniService();
                        }
                    }
                }
                return _intance;
            }
        }

        private IniService()
        {
        }


        [DllImport("kernel32")]//返回0表示失败，非0为成功
        private static extern long WritePrivateProfileString(string section, string key,
            string val, string filePath);

        [DllImport("kernel32")]//返回取得字符串缓冲区的长度
        private static extern long GetPrivateProfileString(string section, string key,
            string def, StringBuilder retVal, int size, string filePath);

        public string ReadIniData(string Section, string Key, string defText, string iniFilePath)
        {
            if (File.Exists(iniFilePath))
            {
                StringBuilder temp = new StringBuilder(1024);
                GetPrivateProfileString(Section, Key, defText, temp, 1024, iniFilePath);
                WriteIniData(Section, Key, temp.ToString(), iniFilePath);
                return temp.ToString();
            }
            else
            {
                WriteIniData(Section, Key, defText, iniFilePath);
                return defText;
            }
        }

        public bool WriteIniData(string Section, string Key, string Value, string iniFilePath)
        {
            long OpStation = WritePrivateProfileString(Section, Key, Value, iniFilePath);
            if (OpStation == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
    }
}
