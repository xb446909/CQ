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

        public void Save(string FileName, Model model)
        {
            try
            {

                int idx = FileName.LastIndexOf('.');
                string ext = FileName.Substring(idx + 1);
                NPOI.SS.UserModel.IWorkbook workbook = null;

                if (File.Exists(FileName))
                {
                    using (FileStream fs = new FileStream(FileName, FileMode.Open, FileAccess.Read))
                    {
                        workbook = NPOI.SS.UserModel.WorkbookFactory.Create(fs);
                        int nSheets = workbook.NumberOfSheets;
                        if (nSheets > 0)
                        {
                            var sheet = workbook.GetSheetAt(0);
                            int nRows = sheet.LastRowNum + 1;

                            var row = sheet.CreateRow(nRows);
                            row.CreateCell(0).SetCellValue(model.Id);
                            row.CreateCell(1).SetCellValue(model.Time);
                            row.CreateCell(2).SetCellValue(model.QRCode);
                            row.CreateCell(3).SetCellValue(model.Flow);
                            row.CreateCell(4).SetCellValue(model.APressure);
                            row.CreateCell(5).SetCellValue(model.BPressure);
                        }
                        fs.Close();
                    }
                }
                else
                {
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
                    row0.CreateCell(2).SetCellValue("条码");
                    row0.CreateCell(3).SetCellValue("流量/ml");
                    row0.CreateCell(4).SetCellValue("A胶压力/psi");
                    row0.CreateCell(5).SetCellValue("B胶压力/psi");

                    var row = sheet.CreateRow(1);
                    row.CreateCell(0).SetCellValue(model.Id);
                    row.CreateCell(1).SetCellValue(model.Time);
                    row.CreateCell(2).SetCellValue(model.QRCode);
                    row.CreateCell(3).SetCellValue(model.Flow);
                    row.CreateCell(4).SetCellValue(model.APressure);
                    row.CreateCell(5).SetCellValue(model.BPressure);
                }

                using (FileStream fs_write = new FileStream(FileName, FileMode.Create))
                {
                    workbook.Write(fs_write);
                    fs_write.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Save(string FileName, ObservableCollection<Model> models)
        {
            try
            {

                int idx = FileName.LastIndexOf('.');
                string ext = FileName.Substring(idx + 1);
                NPOI.SS.UserModel.IWorkbook workbook = null;

                if (File.Exists(FileName))
                {
                    using (FileStream fs = new FileStream(FileName, FileMode.Open, FileAccess.Read))
                    {
                        workbook = NPOI.SS.UserModel.WorkbookFactory.Create(fs);
                        int nSheets = workbook.NumberOfSheets;
                        if (nSheets > 0)
                        {
                            var sheet = workbook.GetSheetAt(0);
                            int nRows = sheet.LastRowNum + 1;
                            for (int i = 0; i < models.Count; i++)
                            {
                                var row = sheet.CreateRow(i + nRows);
                                row.CreateCell(0).SetCellValue(models[i].Id);
                                row.CreateCell(1).SetCellValue(models[i].Time);
                                row.CreateCell(2).SetCellValue(models[i].QRCode);
                                row.CreateCell(3).SetCellValue(models[i].Flow);
                                row.CreateCell(4).SetCellValue(models[i].APressure);
                                row.CreateCell(5).SetCellValue(models[i].BPressure);
                            }
                        }
                        fs.Close();
                    }
                }
                else
                {
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
                    row0.CreateCell(2).SetCellValue("条码");
                    row0.CreateCell(3).SetCellValue("流量/ml");
                    row0.CreateCell(4).SetCellValue("A胶压力/psi");
                    row0.CreateCell(5).SetCellValue("B胶压力/psi");

                    for (int i = 0; i < models.Count; i++)
                    {
                        var row = sheet.CreateRow(i + 1);
                        row.CreateCell(0).SetCellValue(models[i].Id);
                        row.CreateCell(1).SetCellValue(models[i].Time);
                        row.CreateCell(2).SetCellValue(models[i].QRCode);
                        row.CreateCell(3).SetCellValue(models[i].Flow);
                        row.CreateCell(4).SetCellValue(models[i].APressure);
                        row.CreateCell(5).SetCellValue(models[i].BPressure);
                    }
                }

                using (FileStream fs_write = new FileStream(FileName, FileMode.Create))
                {
                    workbook.Write(fs_write);
                    fs_write.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
