## Link.NPOI.Extension

[![Link.NPOI.Extension](https://img.shields.io/nuget/dt/Link.NPOI.Extension)](https://www.nuget.org/packages/Link.NPOI.Extension) 

[![Link.NPOI.Crypto](https://img.shields.io/nuget/dt/Link.NPOI.Crypto)](https://www.nuget.org/packages/Link.NPOI.Crypto) 


* tool——部分辅助工具
    * Link.ImportExport.MappingEditor —— 映射关系配置工具

### RoadMap

* Excel-NPOI扩展库
    * datatable导入导出【优先caption，然后columnname】、加载到现有datatable的对象中【即给定部分列，读取部分列】 —— dataset
    * 依据excel模板导入导出、依据xml模板导入导出+模板编辑器
    * 依据特性attribute设置【display，column等特性 or 自定义TilteAttribute】导入导出 —— 对象导入导出
    * 依据fluent api处理
    * 导出前设置全局文档信息【作者，等信息】
    * csv扩展支持 —— 通过继承IWorkbook接口实现
    * 添加workbook创建工厂

* 先实现Helper方式功能
* 后续继承实现
* Biff2使用codepage传入参数


### 参考

* [NPOI](https://github.com/nissl-lab/npoi)
    * [Tony Qu](https://github.com/tonyqus)    
* [Office加密：ooxmlcrypto](https://code.google.com/archive/p/ooxmlcrypto/)

* [Excel文件格式列表](https://docs.microsoft.com/zh-cn/dotnet/api/microsoft.office.interop.excel.xlfileformat)
* Biff2格式Excel解析参考
    * [OldExcelExtractor.cs](https://github.com/nissl-lab/npoi/blob/master/main/HSSF/Extractor/OldExcelExtractor.cs)
    * [Jess.DotNet.SmartExcel](https://github.com/ShiJess/Jess.DotNet.SmartExcel)
        * [excelfileformat.pdf 113页：Worksheet/Workbook Records](https://github.com/ShiJess/Jess.DotNet.SmartExcel/blob/main/docs/excelfileformat.pdf)


### 注意

* `Link.NPOI.Crypto`项目采用`LGPL`开源协议，因为其基于`ooxmlcrypto`修改，而`ooxmlcrypto`采用的是`LGPL`协议
