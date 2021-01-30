using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Link.NPOI.Extension.DataAnnotations
{
    /// <summary>
    /// Excel表中对应列头标记
    /// 可以对应多列
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ColumnHeaderAttribute : Attribute
    {

        public ColumnHeaderAttribute() { }

        public ColumnHeaderAttribute(string name)
        {
            Name = name;
        }

        public ColumnHeaderAttribute(string name, int index)
        {
            Name = name;
            Index = index;
        }

        /// <summary>
        /// 列名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 索引对应【Excel文件中列索引】
        /// -1时，表示索引无关
        /// </summary>
        public int Index { get; set; } = -1;

    }
}
