using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CQ
{
    class ExcelService
    {
        private static ExcelService _intance = null;
        private static readonly object _lock = new object();
        public static ExcelService Instance
        {
            get
            {
                if (_intance == null)
                {
                    lock (_lock)
                    {
                        if (_intance == null)
                        {
                            _intance = new ExcelService();
                        }
                    }
                }
                return _intance;
            }
        }

        private ExcelService()
        {
        }

        public void Save(string FileName, ObservableCollection<Model> models)
        {
            try
            {
                FileStream fs = new FileStream(FileName, FileMode.Create);

                int idx = FileName.LastIndexOf('.');
                string ext = FileName.Substring(idx + 1);
                NPOI.SS.UserModel.IWorkbook workbook = null;

                if (ext == "xls")
                {
                    workbook = new NPOI.HSSF.UserModel.HSSFWorkbook();
                }
                else if (ext == "xlsx")
                {
                    workbook = new NPOI.XSSF.UserModel.XSSFWorkbook();
                }
                else
                {
                    return;
                }

                var sheet = workbook.CreateSheet();
                var row0 = sheet.CreateRow(0);
                row0.CreateCell(0).SetCellValue("序号");
                row0.CreateCell(1).SetCellValue("时间");
                row0.CreateCell(1).SetCellValue("条码");
                row0.CreateCell(2).SetCellValue("流量/ml");
                row0.CreateCell(3).SetCellValue("A胶压力");
                row0.CreateCell(4).SetCellValue("B胶压力");

                for (int i = 0; i < models.Count; i++)
                {
                    var row = sheet.CreateRow(i + 1);
                    row.CreateCell(0).SetCellValue(models[i].Id);
                    row.CreateCell(0).SetCellValue(models[i].Time);
                    row.CreateCell(1).SetCellValue(models[i].QRCode);
                    row.CreateCell(2).SetCellValue(models[i].Flow);
                    row.CreateCell(3).SetCellValue(models[i].APressure);
                    row.CreateCell(4).SetCellValue(models[i].BPressure);
                }
                workbook.Write(fs);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
