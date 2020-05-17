using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace Link.ImportExport.MappingEditor
{

    /// <summary>
    /// 映射配置文件
    /// </summary>
    public class MappingConfig : ObservableObject
    {
        private string alias = string.Empty;
        public string Alias
        {
            get { return alias; }
            set { Set(() => Alias, ref alias, value); }
        }

        /// <summary>
        /// 映射字段关系列表
        /// </summary>
        public BindingList<MappingRelation> Relations { get; set; } = new BindingList<MappingRelation>();

        /// <summary>
        /// 反序列化映射配置xml文件
        /// </summary>
        /// <param name="fullfilename"></param>
        /// <returns></returns>
        public static MappingConfig ReadFromXmlFormat(string fullfilename)
        {
            XmlSerializer xmlformat = new XmlSerializer(typeof(MappingConfig));
            using (Stream fs = new FileStream(fullfilename, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                return xmlformat.Deserialize(fs) as MappingConfig;
            }
        }

        /// <summary>
        /// 序列化映射配置为xml文件
        /// </summary>
        /// <param name="fullfilename">完整文件名路径</param>
        /// <param name="mapcfg">待保存的配置</param>
        public static void SaveAsXmlFormat(string fullfilename, MappingConfig mapcfg)
        {
            XmlSerializer xmlformat = new XmlSerializer(typeof(MappingConfig));
            using (Stream fs = new FileStream(fullfilename, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                xmlformat.Serialize(fs, mapcfg);
            }
        }
    }

    /// <summary>
    /// 映射关系
    /// </summary>
    public class MappingRelation : ObservableObject
    {
        private string columnname = string.Empty;
        public string ColumnName
        {
            get { return columnname; }
            set { Set(() => ColumnName, ref columnname, value); }
        }

        private string alias = string.Empty;
        public string Alias
        {
            get { return alias; }
            set { Set(() => Alias, ref alias, value); }
        }
    }

}
