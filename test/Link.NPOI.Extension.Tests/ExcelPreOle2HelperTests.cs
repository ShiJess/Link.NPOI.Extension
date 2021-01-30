using Link.NPOI.Extension.DataAnnotations;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Xunit;
using Xunit.Abstractions;

namespace Link.NPOI.Extension.Tests
{
    public class ExcelPreOle2HelperTests
    {
        private ITestOutputHelper outputHelper { get; set; }
        public ExcelPreOle2HelperTests(ITestOutputHelper _outputHelper)
        {
            outputHelper = _outputHelper;
        }

        /// <summary>
        /// 初始化DataTable数据
        /// </summary>
        /// <returns></returns>
        public DataTable InitExportData()
        {
            DataTable dt = new DataTable("我要飞");

            DataColumn dc = new DataColumn();
            dc.Caption = "标题1";
            dc.ColumnName = "列名1";
            dt.Columns.Add(dc);

            DataColumn dcc = new DataColumn();
            dcc.ColumnName = "列名2";
            dt.Columns.Add(dcc);

            //for (int i = 1; i < 10; i++)
            //{
            //    DataRow dr = dt.NewRow();
            //    dr[0] = "列1，行" + i;
            //    dr[1] = "列2，行" + i;
            //    dt.Rows.Add(dr);
            //}

            return dt;
        }

        [Fact]
        public void ExportDataTableToOldExcelTest()
        {
            string fullfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xls");
            bool result = ExcelHelper.ExportDataTableToNewExcel(fullfilename, InitExportData());
            Assert.True(result);
        }

        [Fact]
        public void ExportDataTableToNewExcelTest()
        {
            string fullfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xlsx");
            bool result = ExcelHelper.ExportDataTableToNewExcel(fullfilename, InitExportData());
            Assert.True(result);
        }

        public class Test
        {
            public string FirstProp { get; set; }

            [ColumnHeader("列名2")]
            public string SecondProp { get; set; }
        }
        public class Test1
        {
            public string FirstProp { get; set; }
        }

        [Fact]
        public void ExportOneObjectToNewExcelTest()
        {
            Test test = new Test();
            test.FirstProp = "test";
            List<Test> testlist = new List<Test>();
            testlist.Add(test);

            string filename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xml");
            MappingConfig mapcfg = MappingConfig.ReadFromXmlFormat(filename);

            ExportInfo ex = new ExportInfo();
            ex.data = testlist.Cast<Object>().ToList();
            ex.datatype = typeof(Test);
            ex.Config = mapcfg;

            string fullfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xls");
            bool result = ExcelHelper.ExportObjectToNewExcel(fullfilename, ex);
            Assert.True(result);
        }

        [Fact]
        public void ExportManyObjectToNewExcelTest()
        {
            Test test = new Test();
            test.FirstProp = "test";
            List<Test> testlist = new List<Test>();
            testlist.Add(test);

            string filename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xml");
            MappingConfig mapcfg = MappingConfig.ReadFromXmlFormat(filename);

            ExportInfo ex = new ExportInfo();
            ex.data = testlist.Cast<Object>().ToList();
            ex.datatype = typeof(Test);
            ex.Config = mapcfg;

            Test1 test1 = new Test1();
            test1.FirstProp = "test1";
            List<Test1> testlist1 = new List<Test1>();
            testlist1.Add(test1);

            string filename1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test1.xml");
            MappingConfig mapcfg1 = MappingConfig.ReadFromXmlFormat(filename1);

            ExportInfo ex1 = new ExportInfo();
            ex1.data = testlist1.Cast<Object>().ToList();
            ex1.datatype = typeof(Test1);
            ex1.Config = mapcfg1;

            string fullfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xls");
            bool result = ExcelHelper.ExportObjectToNewExcel(fullfilename, ex, ex1);
            Assert.True(result);
        }



        [Fact(DisplayName = "Excel Biff 2格式读取")]
        public void ImportDataTableFromExcel2()
        {
            string fullfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "testExcel2.xls");

            var dt = ExcelPreOle2Helper.GetDataOnSheet(fullfilename, 1);

            outputHelper.WriteLine(dt.Rows.Count.ToString());
        }


        [Fact(DisplayName = "Excel Biff 2格式读取-测试Codpage")]
        public void ImportDataTableFromExcel2_Codepage()
        {
            string fullfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "testCodepage.xls");

            var dt = ExcelPreOle2Helper.GetDataOnSheet(fullfilename, -1);

            outputHelper.WriteLine(dt.Rows.Count.ToString());
        }

        [Fact]
        public void Import()
        {
            string fullfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "testExcel2.xls");

            List<Test> testlist = ExcelPreOle2Helper.Import<Test>(fullfilename, 0);

            Assert.NotEmpty(testlist);
        }

    }
}
