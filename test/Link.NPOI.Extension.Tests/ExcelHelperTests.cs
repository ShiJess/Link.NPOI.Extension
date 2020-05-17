using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Xunit;

namespace Link.NPOI.Extension.Tests
{
    public class ExcelHelperTests
    {
        /// <summary>
        /// ��ʼ��DataTable����
        /// </summary>
        /// <returns></returns>
        public DataTable InitExportData()
        {
            DataTable dt = new DataTable("��Ҫ��");

            DataColumn dc = new DataColumn();
            dc.Caption = "����1";
            dc.ColumnName = "����1";
            dt.Columns.Add(dc);

            DataColumn dcc = new DataColumn();
            dcc.ColumnName = "����2";
            dt.Columns.Add(dcc);

            for (int i = 1; i < 10; i++)
            {
                DataRow dr = dt.NewRow();
                dr[0] = "��1����" + i;
                dr[1] = "��2����" + i;
                dt.Rows.Add(dr);
            }

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

    }
}