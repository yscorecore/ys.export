using System;
using System.Collections.Generic;
namespace YS.Export.Core
{
    public interface IExportService
    {
        ExportToken DoExport(ExportInfo exportInfo);

        void Append(DataInfo data);

        Stream Complate(string token);

    }
    public class ExportResult
    {
        string FileName { get; set; }
        Stream OutputStream { get; set; }
    }

    public class ExportInfo
    {
        public Dictionary<string, ColumnInfo> Columns { get; set; }
        public string DataTitle { get; set; }

    }

    public class DataInfo
    {
        public List<Dictionary<string, object>> Data { get; set; }

        public string Token { get; set; }

    }

    public class ExportToken
    {
        public string Token { get; set; }
        public DateTimeOffset Expired { get; set; }
    }
    public class CloumnInfo
    {
        public string DisplayName { get; set; }
        public int ColumnWidth { get; set; }
    }
}
