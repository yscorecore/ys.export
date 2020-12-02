using System;
using System.Collections.Generic;
using System.IO;
using System.ComponentModel;
using System.Globalization;

namespace YS.Export
{
    public interface IExcelExportService
    {
        string CreateToken(ExportInfo exportInfo);

        void AppendData(string token, Dictionary<string, object[][]> excelData);

        ExportResult Complate(string token);


    }

    public interface IExcelWriter
    {
        void CreateSchemas(string file, Dictionary<string, ColumnInfo[]> schemas);
        void AppendData(string file, Dictionary<string, object[][]> excelData);
    }






    public class ExportResult
    {
        string FileType { get; set; }
        string FileName { get; set; }
        Stream FileStream { get; set; }
    }

    public class ExportInfo
    {
        string FileType { get; set; }
        public string DataTitle { get; set; }
        public Dictionary<string, ColumnInfo[]> Sheets { get; set; }
    }

    public class ColumnInfo
    {
        public string DisplayName { get; set; }
        public int ColumnWidth { get; set; }
    }
    public class ColumnInfoTypeConverter : StringConverter
    {
        public override object ConvertFrom(ITypeDescriptorContext context, CultureInfo culture, object value)
        {
            if (value is string columnName)
            {
                return new ColumnInfo
                {
                    DisplayName = columnName
                };
            }
            return base.ConvertFrom(context, culture, value);
        }
    }
}



