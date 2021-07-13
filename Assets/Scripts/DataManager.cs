/// <summary>
/// 数据管理器 储存所有数据
/// </summary>
public static class DataManager
{
    //------------------------- Excel Start -------------------------
    public readonly static string ExcelToJsonResourcesPath = "ExcelToJson/";// Json文件夹路径 -> Resources文件夹中Json文件夹所在路径
    public readonly static string ExcelToJsonPath = "/Resources/ExcelToJson/";// Json文件所在文件夹路径
#if UNITY_EDITOR
    public readonly static bool JsonDeBug = false;// 开启JSON DEBUG模式 在编辑器内使用JSON内容
    public readonly static string DefaultExcelFileSuffix = ".xlsx";// 默认Excel后缀
    public readonly static string ExcelTitleCheck = "^";// ExcelTitleCheck + 表名 -> 校验标题头
    public readonly static string ExcelsFolderPath = "/Excels/";// Excel文件所在文件夹路径
    // Excel表格跳过写入规则校验列表 只校验字符串0号位
    public readonly static System.Collections.Generic.List<char> ExcelFilterRule = new System.Collections.Generic.List<char>() {
        '#',
    };
    // Excel表格临时文件规则校验列表 只校验字符串0号位
    public readonly static System.Collections.Generic.List<char> ExcelTempFileFilterRule = new System.Collections.Generic.List<char>() {
        '~',
    };
#endif
    // 声明数据
    public readonly static ExcelData test = ExcelUtil.Read("TEST");
}
