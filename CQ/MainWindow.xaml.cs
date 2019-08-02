using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CQ
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        private ManualResetEvent closeEvent = new ManualResetEvent(false);
        private Thread thread1 = null;
        private Thread thread2 = null;
        ObservableCollection<Model> models1 = new ObservableCollection<Model>();
        ObservableCollection<Model> models2 = new ObservableCollection<Model>();
        public MainWindow()
        {
            InitializeComponent();
            DataList1.DataContext = models1;
            DataList2.DataContext = models2;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SerialPortService.Instance.SerialPort_ReadNewLine = SerialPort_NewLine;
            if (PLCService.Instance.Connect() == false)
            {
                MessageBox.Show("连接PLC失败!");
                return;
            }
            thread1 = new Thread(new ThreadStart(ThreadProc1));
            thread1.Start();
            thread2 = new Thread(new ThreadStart(ThreadProc2));
            thread2.Start();
        }

        private void SerialPort_NewLine(int nSerialPort, string str)
        {
            if (nSerialPort == 1)
            {
                Dispatcher.BeginInvoke(new Action(() => { Barcode1.Text = str; }));
            }
            else
            {
                Dispatcher.BeginInvoke(new Action(() => { Barcode2.Text = str; }));
            }
        }

        private void ThreadProc1()
        {
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            string ExcelFilePath = IniService.Instance.ReadIniData("保存", "Excel地址", "D:\\", str + "Config.ini");
            if (!ExcelFilePath.EndsWith("\\"))
            {
                ExcelFilePath += "\\";
            }

            string ID = IniService.Instance.ReadIniData("序号1", "地址", "DB2.0", str + "Config.ini");
            string Flow = IniService.Instance.ReadIniData("流量1", "地址", "DB2.4", str + "Config.ini");
            string APressure = IniService.Instance.ReadIniData("A股压力1", "地址", "DB2.8", str + "Config.ini");
            string BPressure = IniService.Instance.ReadIniData("B股压力1", "地址", "DB2.12", str + "Config.ini");

            string Start = IniService.Instance.ReadIniData("启动信号1", "地址", "DB2.28", str + "Config.ini");

            bool bLastStart = false;
            bool bStart = false;

            while (true)
            {
                if (closeEvent.WaitOne(100) == true)
                {
                    break;
                }

                bStart = PLCService.Instance.ReadBool(Start);
                if ((bLastStart == false) && (bStart == true))
                {
                    UInt32 nID = PLCService.Instance.ReadUInt32(ID);
                    UInt32 nFlow = PLCService.Instance.ReadUInt32(Flow);
                    UInt32 nAPressure = PLCService.Instance.ReadUInt32(APressure);
                    UInt32 nBPressure = PLCService.Instance.ReadUInt32(BPressure);

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                    if ((nID == 1) && (models1.Count > 0))
                    {
                        string FileName = ExcelFilePath + DateTime.Now.ToString("yyyy-MM-dd") + "_1.xlsx";
                        ExcelService.Instance.Save(FileName, models1);
                        models1.Clear();
                    }

                        models1.Add(new Model()
                        {
                            Date = DateTime.Now.ToString("d"),
                            Time = DateTime.Now.ToString("T"),
                            Id = nID.ToString(),
                            QRCode = Barcode1.Text,
                            Flow = nFlow.ToString(),
                            APressure = nAPressure.ToString(),
                            BPressure = nBPressure.ToString(),
                        });
                    }));

                    
                }
                else
                {
                    bLastStart = bStart;
                }
                
            }
        }

        private void ThreadProc2()
        {
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            string ExcelFilePath = IniService.Instance.ReadIniData("保存", "Excel地址", "D:\\", str + "Config.ini");
            if (!ExcelFilePath.EndsWith("\\"))
            {
                ExcelFilePath += "\\";
            }

            string ID = IniService.Instance.ReadIniData("序号2", "地址", "DB2.0", str + "Config.ini");
            string Flow = IniService.Instance.ReadIniData("流量2", "地址", "DB2.4", str + "Config.ini");
            string APressure = IniService.Instance.ReadIniData("A股压力2", "地址", "DB2.8", str + "Config.ini");
            string BPressure = IniService.Instance.ReadIniData("B股压力2", "地址", "DB2.12", str + "Config.ini");

            string Start = IniService.Instance.ReadIniData("启动信号2", "地址", "DB2.28", str + "Config.ini");

            bool bLastStart = false;
            bool bStart = false;

            while (true)
            {
                if (closeEvent.WaitOne(100) == true)
                {
                    break;
                }

                bStart = PLCService.Instance.ReadBool(Start);
                if ((bLastStart == false) && (bStart == true))
                {
                    UInt32 nID = PLCService.Instance.ReadUInt32(ID);
                    UInt32 nFlow = PLCService.Instance.ReadUInt32(Flow);
                    UInt32 nAPressure = PLCService.Instance.ReadUInt32(APressure);
                    UInt32 nBPressure = PLCService.Instance.ReadUInt32(BPressure);

                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        if ((nID == 1) && (models2.Count > 0))
                        {
                            string FileName = ExcelFilePath + DateTime.Now.ToString("yyyy-MM-dd") + "_2.xlsx";
                            ExcelService.Instance.Save(FileName, models2);
                            models2.Clear();
                        }

                        models2.Add(new Model()
                        {
                            Date = DateTime.Now.ToString("d"),
                            Time = DateTime.Now.ToString("T"),
                            Id = nID.ToString(),
                            QRCode = Barcode2.Text,
                            Flow = nFlow.ToString(),
                            APressure = nAPressure.ToString(),
                            BPressure = nBPressure.ToString(),
                        });
                    }));


                }
                else
                {
                    bLastStart = bStart;
                }

            }
        }


        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            closeEvent.Set();
            thread1?.Join();
            thread2?.Join();
        }
    }
}
