using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using YS.Export;

namespace YS.Export.Impl.NPoi
{

    public class NpoiExcelWriter : IExcelWriter
    {
        public void AppendData(string file, Dictionary<string, object[][]> excelData)
        {
            using (var fs = new FileStream(file, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new XSSFWorkbook(fs);
                foreach (var kv in excelData)
                {
                    ISheet sheet = workbook.GetSheet(kv.Key);
                    var startRow = sheet.LastRowNum;
                    for (int i = 0; i < kv.Value.Length; i++)
                    {
                        WriteLine(sheet,startRow+i,kv.Value[i]);
                    }
                    
                }
                workbook.Write(fs);
            }
        }

        public void CreateSchemas(string file, Dictionary<string, ColumnInfo[]> schemas)
        {
            using (var fs = new FileStream(file, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                foreach (var kv in schemas)
                {
                    ISheet sheet = workbook.CreateSheet(kv.Key);

                    AppendTitle(sheet, kv.Value);
                }
                workbook.Write(fs);
            }
        }
        private void AppendTitle(ISheet sheet, ColumnInfo[] columns)
        {
            var row = sheet.CreateRow(0);
            for (int i = 0; i < columns.Length; i++)
            {
                var columnInfo = columns[i];
                var cell = row.CreateCell(i);
                cell.SetCellValue(columnInfo.DisplayName);
            }
        }
        private void WriteLine(ISheet sheet, int rowIndex, object[] values)
        {
            var row = sheet.CreateRow(rowIndex);
            for (int i = 0; i < values.Length; i++)
            {
                var cell = row.CreateCell(i);
                SetCellValue(cell, values[i]);
            }

        }
        private void SetCellValue(ICell cell, object value)
        {
            if (value is null) return;
            if (value is string str)
            {
                cell.SetCellValue(str);
            }
            else if (value is bool boolValue)
            {
                cell.SetCellValue(boolValue);
            }
            else if (value is DateTime time)
            {
                cell.SetCellValue(time);
            }
            else if (value is DateTimeOffset timeOffset)
            {
                cell.SetCellValue(timeOffset.DateTime);
            }
            else if (IsNumber(value))
            {
                cell.SetCellValue(Convert.ToDouble(value));
            }
            else
            {
                cell.SetCellValue(value.ToString());
            }
        }
        private bool IsNumber(object value)
        {
            return value is double || value is float || value is int || value is short|| value is byte || value is long || value is decimal|| value is uint || value is ushort || value is ulong;
        }
    }
}
