using NPOI.HSSF.Record;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Link.NPOI.Extension.Record
{
    /// <summary>
    /// 列宽记录
    /// </summary>
    public class ColumnWidthRecord
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(ColumnWidthRecord));


        public const short biff2_sid = 0x0024;//36

        /// <summary>
        /// 起始列索引
        /// </summary>
        public int FirstColumnIndex { get; set; }
        /// <summary>
        /// 结束列索引
        /// </summary>
        public int LastColumnIndex { get; set; }

        /// <summary>
        /// 列宽
        /// </summary>
        public int ColumnWidth { get; set; }

        public ColumnWidthRecord(RecordInputStream in1)
        {
            //this.sid = in1.Sid;
            //this.isBiff2 = isBiff2;

            FirstColumnIndex = in1.ReadUByte();
            LastColumnIndex = in1.ReadUByte();
            ColumnWidth = in1.ReadUShort();

            if (in1.Remaining > 0)
            {
                logger.Log(POILogger.INFO,
                        "ColumnWidthRecord data remains: " + in1.Remaining +
                        " : " + HexDump.ToHex(in1.ReadRemainder())
                        );
            }
        }
    }
}
