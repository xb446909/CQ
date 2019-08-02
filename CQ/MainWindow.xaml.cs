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
        private Thread thread = null;
        ObservableCollection<Model> models = new ObservableCollection<Model>();
        public MainWindow()
        {
            InitializeComponent();
            DataList.DataContext = models;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            SerialPortService.Instance.SerialPort_ReadNewLine = SerialPort_NewLine;
            if (PLCService.Instance.Connect() == false)
            {
                MessageBox.Show("连接PLC失败!");
                return;
            }
            thread = new Thread(new ThreadStart(ThreadProc));
            thread.Start();
        }

        private void SerialPort_NewLine(string str)
        {
            Dispatcher.BeginInvoke(new Action(() => { Barcode.Text = str; }));
        }

        private void ThreadProc()
        {
            string str = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase;
            string ExcelFilePath = IniService.Instance.ReadIniData("保存", "Excel地址", "D:\\", str + "Config.ini");
            if (!ExcelFilePath.EndsWith("\\"))
            {
                ExcelFilePath += "\\";
            }

            string ID = IniService.Instance.ReadIniData("序号", "地址", "DB2.0", str + "Config.ini");
            string Flow = IniService.Instance.ReadIniData("流量", "地址", "DB2.4", str + "Config.ini");
            string APressure = IniService.Instance.ReadIniData("A股压力", "地址", "DB2.8", str + "Config.ini");
            string BPressure = IniService.Instance.ReadIniData("B股压力", "地址", "DB2.12", str + "Config.ini");

            string Start = IniService.Instance.ReadIniData("启动信号", "地址", "DB2.28", str + "Config.ini");

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
                    if ((nID == 0) && (models.Count > 0))
                    {
                        string FileName = ExcelFilePath + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";
                        ExcelService.Instance.Save(FileName, models);
                        models.Clear();
                    }

                        models.Add(new Model()
                        {
                            Date = DateTime.Now.ToString("d"),
                            Time = DateTime.Now.ToString("T"),
                            Id = nID.ToString(),
                            QRCode = Barcode.Text,
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
            thread?.Join();
        }
    }
}
