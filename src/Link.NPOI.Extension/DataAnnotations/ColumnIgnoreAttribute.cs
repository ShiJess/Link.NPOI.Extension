using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Link.NPOI.Extension.DataAnnotations
{
    /// <summary>
    /// 列映射忽略标记 
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ColumnIgnoreAttribute : Attribute
    {

    }
}
