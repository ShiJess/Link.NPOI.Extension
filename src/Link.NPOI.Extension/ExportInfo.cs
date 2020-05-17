using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Link.NPOI.Extension
{
    /// <summary>
    /// 导出信息类
    /// </summary>
    public class ExportInfo
    {
        /// <summary>
        /// 待导出的数据集合
        /// </summary>
        public List<object> data { get; set; }

        /// <summary>
        /// 导出数据的格式
        /// </summary>
        public Type datatype { get; set; }

        /// <summary>
        /// 映射配置文件
        /// </summary>
        public MappingConfig Config { get; set; }
    }

}
