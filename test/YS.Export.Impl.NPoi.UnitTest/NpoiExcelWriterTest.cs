using System.IO;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace YS.Export.Impl.NPoi.UnitTest
{
    [TestClass]
    public class NpoiExcelWriterTest
    {
        [TestMethod]
        public void ShouldCreateSchemaSuccess()
        {
            var tempExcelFile = Path.Combine(System.IO.Path.GetTempPath(),
               $"{Path.GetRandomFileName()}");
            var writer = new NpoiExcelWriter();

            var schema = new Dictionary<string, ColumnInfo[]>
            {
                ["Sheet1"] = new[] {
                  new ColumnInfo{ DisplayName="Column1" },
                  new ColumnInfo{ DisplayName="Column2" }
                }
            };

            writer.CreateSchemas(tempExcelFile, schema);
            System.Console.WriteLine(tempExcelFile);
        }
    }
}
