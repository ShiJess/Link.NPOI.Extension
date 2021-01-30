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
    /// Excel 2处理支持 _ Biff 2
    /// </summary>
    public sealed class ExcelPreOle2Helper
    {

        public static bool Export()
        {
            //HelpLink _ 填写Jess.SmartExcel地址
            throw new NotSupportedException("not support export excel 2.1 version xls,suggest use Jess.SmartExcel");
            //return false;
        }

        /// <summary>
        /// 获取旧版excel数据至datatable中
        /// </summary>
        /// <param name="fullfilename"></param>
        /// <param name="headerRowIndex"></param>
        /// <returns></returns>
        public static DataTable GetDataOnSheet(string fullfilename, int headerRowIndex)
        {
            DataTable dt = new DataTable();
            dt.TableName = Path.GetFileNameWithoutExtension(fullfilename);

            using (var stream = new FileStream(fullfilename, FileMode.Open))
            {
                using (var ris = new RecordInputStream(stream))
                {
                    BOFRecordType fileType;
                    int version;

                    while (ris.HasNextRecord)
                    {
                        int sid = ris.GetNextSid();
                        ris.NextRecord();

                        switch (sid)
                        {
                            case BOFRecord.biff2_sid:
                                //文件头
                                {
                                    BOFRecord record = new BOFRecord(ris);
                                    fileType = record.Type;
                                    version = record.Version;
                                }
                                break;
                            case EOFRecord.sid:
                                //文件尾
                                {
                                    EOFRecord record = new EOFRecord(ris);
                                }
                                break;

                            case 30://0x1E
                                {
                                    //FORMAT —— 仅biff2和biff3中有效
                                    //FormatRecord
                                    ris.ReadFully(new byte[ris.Remaining]);
                                }
                                break;
                            case ColumnWidthRecord.biff2_sid:// 36://0x24
                                {
                                    //COLWIDTH —— 仅biff2中有效
                                    ColumnWidthRecord record = new ColumnWidthRecord(ris);

                                    //ris.ReadFully(new byte[ris.Remaining]);

                                    int columnCount = record.LastColumnIndex - record.FirstColumnIndex;
                                    for (int i = 0; i < columnCount; i++)
                                    {
                                        DataColumn dc = new DataColumn();
                                        dt.Columns.Add(dc);
                                    }
                                }
                                break;

                            case FontRecord.sid://0x31
                                {
                                    //FontRecord biff2中与biff5之后版本不一样
                                    ris.ReadFully(new byte[ris.Remaining]);
                                }
                                break;

                            // 0x36 TABLEOP —— 仅biff2中有效
                            //Top left cell of a multiple operations table

                            case PrintGridlinesRecord.sid:
                                //行结束 - 换行 [todo 确认，只有列头结束才有]
                                {
                                    PrintGridlinesRecord record = new PrintGridlinesRecord(ris);
                                }
                                break;

                            case OldNumberRecord.biff2_sid:
                                {
                                    OldNumberRecord record = new OldNumberRecord(ris);

                                    int dtRowIndex = record.Row - headerRowIndex - 1;
                                    if (dt.Rows.Count > dtRowIndex)
                                    {
                                        dt.Rows[dtRowIndex][record.Column] = record.Value;
                                    }
                                    else
                                    {
                                        DataRow dr = dt.NewRow();
                                        dt.Rows.Add(dr);

                                        dr[record.Column] = record.Value;
                                    }
                                }
                                break;

                            case OldLabelRecord.biff2_sid:
                                {
                                    OldLabelRecord record = new OldLabelRecord(ris);
                                    record.SetCodePage(new CodepageRecord() { Codepage = (short)CodePageUtil.CP_GBK });
                                    if (record.Row == headerRowIndex)
                                    {
                                        //列头
                                        //DataColumn dc = new DataColumn(record.Value);
                                        //dt.Columns.Add(dc);
                                        dt.Columns[record.Column].ColumnName = record.Value;
                                    }
                                    else if (record.Row > headerRowIndex)
                                    {
                                        int dtRowIndex = record.Row - headerRowIndex - 1;
                                        if (dt.Rows.Count > dtRowIndex)
                                        {
                                            dt.Rows[dtRowIndex][record.Column] = record.Value;
                                        }
                                        else
                                        {
                                            DataRow dr = dt.NewRow();
                                            dt.Rows.Add(dr);

                                            dr[record.Column] = record.Value;
                                        }
                                    }
                                }
                                break;

                            case OldFormulaRecord.biff2_sid:
                                {
                                    OldFormulaRecord record = new OldFormulaRecord(ris);
                                    var value = record.Value;
                                }
                                break;

                            case OldStringRecord.biff2_sid:
                                {
                                    OldStringRecord record = new OldStringRecord(ris);
                                    record.SetCodePage(new CodepageRecord() { Codepage = (short)CodePageUtil.CP_GBK });

                                    var value = record.GetString();
                                }
                                break;

                            default:
                                ris.ReadFully(new byte[ris.Remaining]);
                                break;
                        }

                    }
                }
            }

            return dt;
        }


        public static List<T> Import<T>(string fullfilename, int headerRowIndex = -1)
        {

            try
            {
                return GetObjDataOnSheet<T>(fullfilename, headerRowIndex);
            }
            catch (Exception ex)
            {
                return null;
            }

        }


        public static List<T> GetObjDataOnSheet<T>(string fullfilename, int headerRowIndex)
        {
            List<T> dt = new List<T>();

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

            using (var stream = new FileStream(fullfilename, FileMode.Open))
            {
                using (var ris = new RecordInputStream(stream))
                {
                    BOFRecordType fileType;
                    int version;

                    while (ris.HasNextRecord)
                    {
                        int sid = ris.GetNextSid();
                        ris.NextRecord();

                        switch (sid)
                        {
                            case BOFRecord.biff2_sid:
                                //文件头
                                {
                                    BOFRecord record = new BOFRecord(ris);
                                    fileType = record.Type;
                                    version = record.Version;
                                }
                                break;
                            case EOFRecord.sid:
                                //文件尾
                                {
                                    EOFRecord record = new EOFRecord(ris);
                                }
                                break;

                            case 30://0x1E
                                {
                                    //FORMAT —— 仅biff2和biff3中有效
                                    //FormatRecord
                                    ris.ReadFully(new byte[ris.Remaining]);
                                }
                                break;
                            case ColumnWidthRecord.biff2_sid:// 36://0x24
                                {
                                    //COLWIDTH —— 仅biff2中有效
                                    ColumnWidthRecord record = new ColumnWidthRecord(ris);

                                    //ris.ReadFully(new byte[ris.Remaining]);

                                    //int columnCount = record.LastColumnIndex - record.FirstColumnIndex;
                                    //for (int i = 0; i < columnCount; i++)
                                    //{
                                    //    DataColumn dc = new DataColumn();
                                    //    dt.Columns.Add(dc);
                                    //}
                                }
                                break;

                            case FontRecord.sid://0x31
                                {
                                    //FontRecord biff2中与biff5之后版本不一样
                                    ris.ReadFully(new byte[ris.Remaining]);
                                }
                                break;

                            // 0x36 TABLEOP —— 仅biff2中有效
                            //Top left cell of a multiple operations table

                            case PrintGridlinesRecord.sid:
                                //行结束 - 换行 [todo 确认，只有列头结束才有]
                                {
                                    PrintGridlinesRecord record = new PrintGridlinesRecord(ris);
                                }
                                break;

                            case OldNumberRecord.biff2_sid:
                                {
                                    OldNumberRecord record = new OldNumberRecord(ris);

                                    if (!indexes.ContainsKey(record.Column))
                                    {
                                        break;
                                    }

                                    int dtRowIndex = record.Row - headerRowIndex - 1;
                                    if (dt.Count() > dtRowIndex)
                                    {
                                        //dt[dtRowIndex][record.Column] = record.Value;
                                        T dataRow = dt[dtRowIndex];

                                        var propsel = from c in props where c.Name.ToLower().Equals(indexes[record.Column].ToLower()) select c;
                                        if (propsel == null || propsel.Count() <= 0)
                                        {
                                            continue;
                                        }
                                        PropertyInfo prop = propsel.FirstOrDefault();
                                        prop.SetValue(dataRow, record.Value, null);
                                    }
                                    else
                                    {
                                        T dataRow = System.Activator.CreateInstance<T>();
                                        dt.Add(dataRow);
                                        //DataRow dr = dt.NewRow();
                                        //dt.Rows.Add(dr);

                                        //dr[record.Column] = record.Value;

                                        var propsel = from c in props where c.Name.ToLower().Equals(indexes[record.Column].ToLower()) select c;
                                        if (propsel == null || propsel.Count() <= 0)
                                        {
                                            continue;
                                        }
                                        PropertyInfo prop = propsel.FirstOrDefault();
                                        prop.SetValue(dataRow, record.Value, null);
                                    }
                                }
                                break;

                            case OldLabelRecord.biff2_sid:
                                {
                                    OldLabelRecord record = new OldLabelRecord(ris);
                                    record.SetCodePage(new CodepageRecord() { Codepage = (short)CodePageUtil.CP_GBK });

                                   
                                    if (record.Row == headerRowIndex)
                                    {
                                        //列头
                                        //DataColumn dc = new DataColumn(record.Value);
                                        //dt.Columns.Add(dc);
                                        //dt.Columns[record.Column].ColumnName = ;

                                        string hearName = record.Value;

                                        var pair = pairs.Where(p => p.Value == hearName);
                                        if (pair.Count() > 0)
                                            indexes.Add(record.Column, pair.First().Key);

                                    }
                                    else if (record.Row > headerRowIndex)
                                    {
                                        if (!indexes.ContainsKey(record.Column))
                                        {
                                            break;
                                        }


                                        int dtRowIndex = record.Row - headerRowIndex - 1;
                                        if (dt.Count > dtRowIndex)
                                        {
                                            //dt.Rows[dtRowIndex][record.Column] = record.Value;

                                            T dataRow = dt[dtRowIndex];

                                            var propsel = from c in props where c.Name.ToLower().Equals(indexes[record.Column].ToLower()) select c;
                                            if (propsel == null || propsel.Count() <= 0)
                                            {
                                                continue;
                                            }
                                            PropertyInfo prop = propsel.FirstOrDefault();
                                            prop.SetValue(dataRow, record.Value, null);

                                        }
                                        else
                                        {
                                            T dataRow = System.Activator.CreateInstance<T>();
                                            dt.Add(dataRow);
                                            //DataRow dr = dt.NewRow();
                                            //dt.Rows.Add(dr);

                                            //dr[record.Column] = record.Value;
                                            var propsel = from c in props where c.Name.ToLower().Equals(indexes[record.Column].ToLower()) select c;
                                            if (propsel == null || propsel.Count() <= 0)
                                            {
                                                continue;
                                            }
                                            PropertyInfo prop = propsel.FirstOrDefault();
                                            prop.SetValue(dataRow, record.Value, null);
                                        }
                                    }
                                }
                                break;

                            case OldFormulaRecord.biff2_sid:
                                {
                                    OldFormulaRecord record = new OldFormulaRecord(ris);
                                    var value = record.Value;
                                }
                                break;

                            case OldStringRecord.biff2_sid:
                                {
                                    OldStringRecord record = new OldStringRecord(ris);
                                    record.SetCodePage(new CodepageRecord() { Codepage = (short)CodePageUtil.CP_GBK });

                                    var value = record.GetString();
                                }
                                break;

                            default:
                                ris.ReadFully(new byte[ris.Remaining]);
                                break;
                        }

                    }
                }
            }

            return dt;
        }

    }
}
