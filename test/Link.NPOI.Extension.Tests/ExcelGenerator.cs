using Jess.DotNet.SmartExcel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Link.NPOI.Extension.Tests
{
    /// <summary>
    /// excel生成器，用于解析测试
    /// </summary>
    public class ExcelGenerator
    {

#if NET452

        /// <summary>
        /// 格式参考：
        /// https://docs.microsoft.com/zh-cn/dotnet/api/microsoft.office.interop.excel.xlfileformat?view=excel-pia
        /// </summary>
        [Fact]
        public void ExportDataTableToOldExcelTest()
        {
            string fullfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test.xls");
            string tempfilename = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "test2.1.xls");
            //bool result = ExcelHelper.ExportDataTableToNewExcel(fullfilename, InitExportData());
            //Assert.True(result);
            if (File.Exists(tempfilename))
            {
                File.Delete(tempfilename);
            }

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.DisplayAlerts = false;
            app.AlertBeforeOverwriting = false;

            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(fullfilename);
            //Microsoft.Office.Interop.Excel.Workbook workbook = new Microsoft.Office.Interop.Excel.Workbook();


            workbook.SaveAs(tempfilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlAddIn8);
            //workbook.SaveAs(tempfilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlAddIn8, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //workbook.SaveAs(tempfilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel5);
            //workbook.SaveAs(tempfilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel4);
            //workbook.Save();

            workbook.Close();

        }

#endif

        [Fact]
        public void ExportDataTableToBiff2ExcelTest()
        {
            SmartExcel excel = new SmartExcel();
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "testExcel2.xls");
            excel.CreateFile(path);
            excel.PrintGridLines = false;

            double height = 1.5;

            excel.SetMargin(MarginTypes.TopMargin, height);
            excel.SetMargin(MarginTypes.BottomMargin, height);
            excel.SetMargin(MarginTypes.LeftMargin, height);
            excel.SetMargin(MarginTypes.RightMargin, height);

            string font = "Arial";
            short fontsize = 12;
            excel.SetFont(font, fontsize, FontFormatting.Italic);

            excel.SetColumnWidth(1, 2, 10);
            byte b1 = 2, b2 = 12;
            short s3 = 18;
            excel.SetColumnWidth(b1, b2, s3);

            string header = "头";
            string footer = "角";
            excel.SetHeader(header);
            excel.SetFooter(footer);

            int row = 1, col = 1, cellformat = 0;
            object title = "列名1";
            excel.WriteValue(ValueTypes.Text, CellFont.Font0, CellAlignment.LeftAlign, CellHiddenLocked.Normal, row, col, title, cellformat);

            col = 2;
            title = "列名2";
            excel.WriteValue(ValueTypes.Text, CellFont.Font0, CellAlignment.LeftAlign, CellHiddenLocked.Normal, row, col, title, cellformat);

            row = 2;
            col = 1;
            title = "FirstValue";
            excel.WriteValue(ValueTypes.Text, CellFont.Font0, CellAlignment.LeftAlign, CellHiddenLocked.Normal, row, col, title, cellformat);

            col = 2;
            title = "SecondValue";
            excel.WriteValue(ValueTypes.Text, CellFont.Font0, CellAlignment.LeftAlign, CellHiddenLocked.Normal, row, col, title, cellformat);

            excel.CloseFile();
        }

    }
}
