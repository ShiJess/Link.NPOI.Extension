using Link.NPOI.Extension;
using System;
using System.IO;
using Xunit;

namespace Link.NPOI.Extension.Tests
{
 
    public class MappingRelationTest
    {
        [Fact]
        public void SaveMappingTest()
        {
            string filename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test1.xml");
            MappingConfig mapcfg = new MappingConfig();
            mapcfg.Alias = "表名";

            MappingRelation maprel = new MappingRelation();
            maprel.Alias = "字段别名";
            maprel.ColumnName = "FirstProp";
            mapcfg.Relations.Add(maprel);

            MappingRelation maprel2 = new MappingRelation();
            maprel2.Alias = "字段111别名";
            maprel2.ColumnName = "aaa";
            mapcfg.Relations.Add(maprel2);

            MappingConfig.SaveAsXmlFormat(filename, mapcfg);
        }

        [Fact]
        public void ReadMappingTest()
        {
            string filename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xml");

            MappingConfig mapcfg = MappingConfig.ReadFromXmlFormat(filename);
            Console.WriteLine(mapcfg);
        }

    }
}
