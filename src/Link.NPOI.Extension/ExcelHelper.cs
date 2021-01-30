using Link.NPOI.Extension.DataAnnotations;
using Link.NPOI.Extension.Record;
using NPOI.HSSF.Record;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Link.NPOI.Extension
{
    /// <summary>
    /// NPOI-Excel扩展方法集
    /// </summary>
    public sealed class ExcelHelper
    {
        /// <summary>
        /// 将DataTable数据导出至Excel
        /// </summary>
        /// <param name="fullfilename"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        /// <remarks>
        /// Excel中Sheet名默认使用DataTable表名，列名使用DataColumn的Caption标题-Caption为空时，使用ColumnName
        /// </remarks>
        public static bool ExportDataTableToNewExcel(string fullfilename, DataTable data)
        {
            try
            {
                IWorkbook workbook = null;

                if (fullfilename.ToLower().EndsWith(".xlsx"))
                {
                    workbook = new XSSFWorkbook();
                }
                else
                {
                    workbook = new HSSFWorkbook();
                }

                string sheetname = string.IsNullOrWhiteSpace(data.TableName) ? "Sheet1" : data.TableName;
                ISheet sheet = workbook.CreateSheet(sheetname);

                #region 设置表头
                IRow columnheadrow = sheet.CreateRow(0);
                int columnheadnum = 0;
                foreach (DataColumn item in data.Columns)
                {
                    ICell cell = columnheadrow.CreateCell(columnheadnum);
                    cell.SetCellValue(string.IsNullOrWhiteSpace(item.Caption) ? item.ColumnName : item.Caption);
                    columnheadnum++;
                }
                #endregion

                #region 设置数据内容
                int rownum = 1;
                foreach (DataRow item in data.Rows)
                {
                    IRow row = sheet.CreateRow(rownum);
                    int columnnum = 0;
                    foreach (var item1 in item.ItemArray)
                    {
                        ICell cell = row.CreateCell(columnnum);
                        cell.SetCellValue(item1.ToString());
                        columnnum++;
                    }
                    rownum++;
                }
                #endregion

                using (FileStream fs = File.Create(fullfilename))
                {
                    workbook.Write(fs);
                    Console.WriteLine("导出成功！");
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 导出对象集合到excel文件 
        /// </summary>
        /// <param name="fullfilename"></param>
        /// <param name="exportdata">待导出的数据信息集合</param>
        /// <returns></returns>
        public static bool ExportObjectToNewExcel(string fullfilename, params ExportInfo[] exportdata)
        {
            try
            {
                IWorkbook workbook = null;

                if (fullfilename.ToLower().EndsWith(".xlsx"))
                {
                    workbook = new XSSFWorkbook();
                }
                else
                {
                    workbook = new HSSFWorkbook();
                }

                foreach (ExportInfo exportinfo in exportdata)
                {
                    ISheet sheet = string.IsNullOrWhiteSpace(exportinfo.Config.Alias)
                        || workbook.GetSheet(exportinfo.Config.Alias) != null
                        ? workbook.CreateSheet() : workbook.CreateSheet(exportinfo.Config.Alias);

                    #region 设置表头
                    IRow columnheadrow = sheet.CreateRow(0);
                    int columnheadnum = 0;
                    foreach (MappingRelation item in exportinfo.Config.Relations)
                    {
                        if (!item.IsChecked)
                            continue;

                        ICell cell = columnheadrow.CreateCell(columnheadnum);
                        cell.SetCellValue(string.IsNullOrWhiteSpace(item.Alias) ? item.ColumnName : item.Alias);
                        columnheadnum++;
                    }
                    #endregion

                    #region 设置数据内容
                    int rownum = 1;
                    PropertyInfo[] props = exportinfo.datatype.GetProperties();

                    foreach (object item in exportinfo.data)
                    {
                        IRow row = sheet.CreateRow(rownum);
                        int columnnum = 0;
                        foreach (MappingRelation item1 in exportinfo.Config.Relations)
                        {
                            if (!item1.IsChecked)
                                continue;

                            var propsel = from c in props where c.Name.ToLower().Equals(item1.ColumnName.ToLower()) select c;
                            if (propsel == null || propsel.Count() <= 0)
                            {
                                columnnum++;
                                continue;
                            }
                            PropertyInfo prop = propsel.FirstOrDefault();
                            ICell cell = row.CreateCell(columnnum);
                            cell.SetCellValue((prop.GetValue(item, null) ?? "").ToString());
                            columnnum++;
                        }
                        rownum++;
                    }
                    #endregion
                }

                using (FileStream fs = File.Create(fullfilename))
                {
                    workbook.Write(fs);
                    Console.WriteLine("导出成功！");
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 将excel中的内容导入的DataTable中
        /// </summary>
        /// <param name="fullfilename"></param>
        /// <param name="sheetIndex"></param>
        /// <returns></returns>
        public static DataTable ImportExcelToDataTable(string fullfilename, int sheetIndex = 0)
        {
            try
            {
                using (Stream stream = new FileStream(fullfilename, FileMode.Open))
                {


                    IWorkbook workbook = null;

                    if (fullfilename.ToLower().EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(stream);
                    }
                    else
                    {
                        workbook = new HSSFWorkbook(stream);
                    }

                    ISheet sheet = workbook.GetSheetAt(sheetIndex);

                    DataTable table = GetDataOnSheet(workbook, sheet, 0);

                    return table;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        /// <summary>
        /// 获取特定Sheet页的数据
        /// </summary>
        /// <param name="sheet">Sheet页</param>
        /// <param name="headerRowIndex">标题行号</param>
        /// <returns>数据表</returns>
        private static DataTable GetDataOnSheet(IWorkbook workbook, ISheet sheet, int headerRowIndex)
        {
            if (sheet == null
                || headerRowIndex < 0)
                return null;

            DataTable table = new DataTable(sheet.SheetName);
            try
            {

                IRow headerRow = sheet.GetRow(headerRowIndex);
                if (headerRow != null)
                {
                    int cellCount = headerRow.LastCellNum;
                    ICell headerCell = null;
                    string hearName = string.Empty;
                    int index = 1;
                    for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    {
                        headerCell = headerRow.GetCell(i);
                        if (headerCell != null)
                        {
                            if (string.IsNullOrEmpty(headerRow.GetCell(i).StringCellValue))
                            {
                                hearName = "空白列" + index;
                                index++;
                            }
                            else
                            {
                                hearName = headerRow.GetCell(i).StringCellValue.Trim();
                            }

                            DataColumn column = new DataColumn(hearName);
                            table.Columns.Add(column);
                        }
                        else
                        {
                            cellCount = i;
                            break;
                        }
                    }
                    int rowCount = sheet.LastRowNum;
                    IRow row = null;
                    DataRow dataRow = null;
                    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                    {
                        row = sheet.GetRow(i);
                        if (row == null)
                            continue;

                        dataRow = table.NewRow();
                        bool isEffective = false;
                        ICell cell = null;
                        string cellValue = string.Empty;
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            cell = row.GetCell(j);
                            if (cell != null)
                            {
                                //if (string.IsNullOrEmpty(cell.CellFormula) == false)
                                //{
                                //    HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(sheet, workbook);
                                //    cell = evaluator.EvaluateInCell(cell);
                                //}

                                switch (cell.CellType)
                                {
                                    case CellType.Numeric:
                                        if (DateUtil.IsCellDateFormatted(cell))
                                        {
                                            cellValue = cell.DateCellValue.ToString("yyyy-MM-dd");
                                        }
                                        else
                                        {
                                            cellValue = cell.ToString().Trim();
                                        }
                                        break;
                                    //    case HSSFCell.CELL_TYPE_NUMERIC:
                                    //        {
                                    //            if (HSSFDateUtil.IsCellDateFormatted(cell))
                                    //                try
                                    //                {
                                    //                    cellValue = cell.DateCellValue.ToString("yyyy-MM-dd");
                                    //                }
                                    //                catch
                                    //                {
                                    //                    cellValue = cell.ToString().Trim();
                                    //                }
                                    //            else
                                    //                cellValue = cell.NumericCellValue.ToString();
                                    //        }
                                    //        break;
                                    //    case HSSFCell.CELL_TYPE_BOOLEAN:
                                    //        cellValue = cell.BooleanCellValue.ToString();
                                    //        break;
                                    //    case HSSFCell.CELL_TYPE_STRING:
                                    //        cellValue = cell.StringCellValue.Trim();
                                    //        break;
                                    //    case HSSFCell.CELL_TYPE_BLANK:
                                    //    case HSSFCell.CELL_TYPE_ERROR:
                                    //    case HSSFCell.CELL_TYPE_FORMULA:
                                    default:
                                        cellValue = cell.ToString().Trim();
                                        break;
                                }

                                if (isEffective == false && string.IsNullOrEmpty(cellValue) == false)
                                {
                                    isEffective = true;
                                }
                                dataRow[j] = cellValue;
                            }
                        }

                        if (isEffective)
                        {
                            table.Rows.Add(dataRow);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                table = null;
                throw ex;
            }
            return table;
        }


        public static bool Export<T>(string fullfilename, List<T> obj)
        {
            if (obj == null)
                return false;

            try
            {
                IWorkbook workbook = null;

                if (fullfilename.ToLower().EndsWith(".xlsx"))
                {
                    workbook = new XSSFWorkbook();
                }
                else
                {
                    workbook = new HSSFWorkbook();
                }

                PropertyInfo[] props = typeof(T).GetProperties();
                Dictionary<string, string> pairs = new Dictionary<string, string>();
                foreach (var item in props)
                {
                    var ignoreattrs = item.GetCustomAttributes(typeof(ColumnIgnoreAttribute), false).Cast<ColumnIgnoreAttribute>();
                    if (ignoreattrs.Count() > 0)
                        continue;

                    string key = item.Name;
                    string value = key;
                    var attrs = item.GetCustomAttributes(typeof(ColumnHeaderAttribute), false).Cast<ColumnHeaderAttribute>();
                    if (attrs.Count() > 0)
                        value = attrs.First().Name;
                    pairs.Add(key, value);
                }

                ISheet sheet = workbook.CreateSheet();

                #region 设置表头
                IRow columnheadrow = sheet.CreateRow(0);
                int columnheadnum = 0;

                foreach (var item in pairs)
                {
                    ICell cell = columnheadrow.CreateCell(columnheadnum);
                    cell.SetCellValue(item.Value);
                    columnheadnum++;
                }
                #endregion

                #region 设置数据内容
                int rownum = 1;

                foreach (object item in obj)
                {
                    IRow row = sheet.CreateRow(rownum);
                    int columnnum = 0;
                    foreach (var item1 in pairs)
                    {
                        var propsel = from c in props where c.Name.ToLower().Equals(item1.Key.ToLower()) select c;
                        if (propsel == null || propsel.Count() <= 0)
                        {
                            columnnum++;
                            continue;
                        }
                        PropertyInfo prop = propsel.FirstOrDefault();
                        ICell cell = row.CreateCell(columnnum);
                        cell.SetCellValue((prop.GetValue(item, null) ?? "").ToString());
                        columnnum++;
                    }
                    rownum++;
                }
                #endregion

                using (FileStream fs = File.Create(fullfilename))
                {
                    workbook.Write(fs);
                    Console.WriteLine("导出成功！");
                    return true;
                }
            }
            catch
            {
                return false;
            }

        }

        public static List<T> Import<T>(string fullfilename, int sheetIndex = 0)
        {

            try
            {
                using (Stream stream = new FileStream(fullfilename, FileMode.Open))
                {


                    IWorkbook workbook = null;

                    if (fullfilename.ToLower().EndsWith(".xlsx"))
                    {
                        workbook = new XSSFWorkbook(stream);
                    }
                    else
                    {
                        workbook = new HSSFWorkbook(stream);
                    }

                    ISheet sheet = workbook.GetSheetAt(sheetIndex);

                    List<T> table = GetObjDataOnSheet<T>(workbook, sheet, 0);

                    return table;
                }
            }
            catch (Exception ex)
            {
                return null;
            }

        }


        private static List<T> GetObjDataOnSheet<T>(IWorkbook workbook, ISheet sheet, int headerRowIndex)
        {
            if (sheet == null
                || headerRowIndex < 0)
                return null;

            List<T> table = new List<T>();
            try
            {

                PropertyInfo[] props = typeof(T).GetProperties();
                Dictionary<string, string> pairs = new Dictionary<string, string>();
                foreach (var item in props)
                {
                    var ignoreattrs = item.GetCustomAttributes(typeof(ColumnIgnoreAttribute), false).Cast<ColumnIgnoreAttribute>();
                    if (ignoreattrs.Count() > 0)
                        continue;

                    string key = item.Name;
                    string value = key;
                    var attrs = item.GetCustomAttributes(typeof(ColumnHeaderAttribute), false).Cast<ColumnHeaderAttribute>();
                    if (attrs.Count() > 0)
                        value = attrs.First().Name;
                    pairs.Add(key, value);
                }

                Dictionary<int, string> indexes = new Dictionary<int, string>();


                IRow headerRow = sheet.GetRow(headerRowIndex);
                if (headerRow != null)
                {
                    int cellCount = headerRow.LastCellNum;
                    ICell headerCell = null;
                    string hearName = string.Empty;
                    int index = 1;
                    for (int i = headerRow.FirstCellNum; i < cellCount; i++)
                    {
                        headerCell = headerRow.GetCell(i);
                        if (headerCell != null)
                        {
                            hearName = headerRow.GetCell(i).StringCellValue.Trim();

                            var pair = pairs.Where(p => p.Value == hearName);
                            if (pair.Count() > 0)
                                indexes.Add(i, pair.First().Key);
                        }
                        else
                        {
                            cellCount = i;
                            break;
                        }
                    }
                    int rowCount = sheet.LastRowNum;
                    IRow row = null;

                    for (int i = (sheet.FirstRowNum + 1); i <= sheet.LastRowNum; i++)
                    {
                        row = sheet.GetRow(i);
                        if (row == null)
                            continue;

                        //T dataRow = default(T);
                        T dataRow = System.Activator.CreateInstance<T>();
                        bool isEffective = false;
                        ICell cell = null;
                        string cellValue = string.Empty;
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            cell = row.GetCell(j);
                            if (cell != null)
                            {
                                //if (string.IsNullOrEmpty(cell.CellFormula) == false)
                                //{
                                //    HSSFFormulaEvaluator evaluator = new HSSFFormulaEvaluator(sheet, workbook);
                                //    cell = evaluator.EvaluateInCell(cell);
                                //}

                                switch (cell.CellType)
                                {
                                    case CellType.Numeric:
                                        if (DateUtil.IsCellDateFormatted(cell))
                                        {
                                            cellValue = cell.DateCellValue.ToString("yyyy-MM-dd");
                                        }
                                        else
                                        {
                                            cellValue = cell.ToString().Trim();
                                        }
                                        break;
                                    //    case HSSFCell.CELL_TYPE_NUMERIC:
                                    //        {
                                    //            if (HSSFDateUtil.IsCellDateFormatted(cell))
                                    //                try
                                    //                {
                                    //                    cellValue = cell.DateCellValue.ToString("yyyy-MM-dd");
                                    //                }
                                    //                catch
                                    //                {
                                    //                    cellValue = cell.ToString().Trim();
                                    //                }
                                    //            else
                                    //                cellValue = cell.NumericCellValue.ToString();
                                    //        }
                                    //        break;
                                    //    case HSSFCell.CELL_TYPE_BOOLEAN:
                                    //        cellValue = cell.BooleanCellValue.ToString();
                                    //        break;
                                    //    case HSSFCell.CELL_TYPE_STRING:
                                    //        cellValue = cell.StringCellValue.Trim();
                                    //        break;
                                    //    case HSSFCell.CELL_TYPE_BLANK:
                                    //    case HSSFCell.CELL_TYPE_ERROR:
                                    //    case HSSFCell.CELL_TYPE_FORMULA:
                                    default:
                                        cellValue = cell.ToString().Trim();
                                        break;
                                }

                                if (isEffective == false && string.IsNullOrEmpty(cellValue) == false)
                                {
                                    isEffective = true;
                                }
                                //dataRow[j] = cellValue;

                                var propsel = from c in props where c.Name.ToLower().Equals(indexes[j].ToLower()) select c;
                                if (propsel == null || propsel.Count() <= 0)
                                {
                                    continue;
                                }
                                PropertyInfo prop = propsel.FirstOrDefault();
                                prop.SetValue(dataRow, cellValue,null);

                            }
                        }

                        if (isEffective)
                        {
                            //table.Rows.Add(dataRow);
                            table.Add(dataRow);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                table = null;
                throw ex;
            }
            return table;
        }

    }

}