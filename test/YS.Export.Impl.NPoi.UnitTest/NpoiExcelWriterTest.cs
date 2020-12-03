using System.IO;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
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

        [TestMethod]
        public void ShouldWriteDataSuccess()
        {
            var tempExcelFile = Path.Combine(System.IO.Path.GetTempPath(),
               $"{Path.GetRandomFileName()}.xls");
            var writer = new NpoiExcelWriter();

            var schema = new Dictionary<string, ColumnInfo[]>
            {
                ["Sheet1"] = new[] {
                  new ColumnInfo{ DisplayName="Column1" },
                  new ColumnInfo{ DisplayName="Column2" }
                }
            };
            writer.CreateSchemas(tempExcelFile, schema);

            writer.AppendData(tempExcelFile,new Dictionary<string, object[][]>
            {
                ["Sheet1"]=new object[][]
                {
                    new object[]{"a1","a2"},
                    new object[]{"a3","a4"}
                }
            });
            writer.AppendData(tempExcelFile,new Dictionary<string, object[][]>
            {
                ["Sheet1"]=new object[][]
                {
                    new object[]{"b1","b2"},
                    new object[]{"b3","b4"}
                }
            });

            System.Console.WriteLine(tempExcelFile);
        }

         [TestMethod]
        public void ShouldDatetime()
        {
            var code = Type.GetTypeCode(typeof(DateTimeOffset));
            Assert.AreEqual(TypeCode.DateTime,code);
        }
    }
}
