using NPOI.HSSF.Record;
using NPOI.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Link.NPOI.Extension.Record
{
    /// <summary>
    /// 旧版excel数字读取支持
    /// </summary>
    public class OldNumberRecord : OldCellRecord
    {
        private static POILogger logger = POILogFactory.GetLogger(typeof(OldNumberRecord));


        public const short biff2_sid = 0x0003;


        public OldNumberRecord(RecordInputStream in1) : base(in1, in1.Sid == biff2_sid)
        {
            Value = in1.ReadDouble();

            if (in1.Remaining > 0)
            {
                logger.Log(POILogger.INFO,
                        "OldNumberRecord data remains: " + in1.Remaining +
                        " : " + HexDump.ToHex(in1.ReadRemainder())
                        );
            }
        }


        public double Value { get; set; }


        protected override string RecordName => throw new NotImplementedException();

        protected override void AppendValueText(StringBuilder sb)
        {
            throw new NotImplementedException();
        }

    }
}
