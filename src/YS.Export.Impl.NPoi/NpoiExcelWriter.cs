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
            throw new NotImplementedException();
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
    }
}
