using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using LitJson;
using ExcelDataReader;
using UnityEngine;
using UnityEditor;

/// <summary>
/// ExcelUtil 工具类
/// 默认在编辑器内使用的是Excel文件 打包后使用Json文件
/// 需要在编辑器内使用Json文件请将JsonDeBug设置为true
/// 作者：HEX3F
/// 博客：hex3f.com
/// </summary>
public static class ExcelUtil
{
#if UNITY_EDITOR
    private static bool JsonDeBug = DataManager.JsonDeBug;// 开启JSON DEBUG模式 在编辑器内使用JSON内容
    private static string DefaultExcelFileSuffix = DataManager.DefaultExcelFileSuffix;// 默认Excel后缀
    private static string ExcelTitleCheck = DataManager.ExcelTitleCheck;// ExcelTitleCheck + 表名 -> 校验标题头
    private static string ExcelsFolderPath = Application.dataPath + DataManager.ExcelsFolderPath;// Excel文件所在文件夹路径
    // Excel表格跳过写入规则校验列表 只校验字符串0号位
    private static List<char> ExcelFilterRule = DataManager.ExcelFilterRule;
    // Excel表格临时文件规则校验列表 只校验字符串0号位
    private static List<char> ExcelTempFileFilterRule = DataManager.ExcelTempFileFilterRule;
#endif
    private static string ExcelToJsonPath = Application.dataPath + DataManager.ExcelToJsonPath;// Json文件所在文件夹路径
    private static string ExcelToJsonResourcesPath = DataManager.ExcelToJsonResourcesPath;   // Json文件夹路径 -> Resources文件夹中Json文件夹所在路径

    /// <summary>
    /// 读取Excel数据
    /// </summary>
    /// <param name="_fileName">文件名 - 文件后缀为 DefaultExcelFileSuffix 变量</param>
    /// <param name="_dir">所在文件夹</param>
    /// <returns></returns>
    public static ExcelData Read(string _fileName, string _dir = "")
    {
#if UNITY_EDITOR
        _dir += _dir != "" ? "/" : "";
        string _FilePath = $"{ExcelsFolderPath}{_dir + _fileName}{DefaultExcelFileSuffix}";
        if (JsonDeBug) return ReadJson(ExcelToJsonResourcesPath, _fileName, _dir);
        int titleIndex = -1;
        FileStream stream = null;
        try
        {
            stream = File.Open(_FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelDataReader.AsDataSet();
            stream.Close();
            excelDataReader.Close();
            int columns = result.Tables[0].Columns.Count;
            int rows = result.Tables[0].Rows.Count;
            Dictionary<string, object> value;
            List<string> TitleError = new List<string>();
            Dictionary<string, Dictionary<string, object>> _excelData = new Dictionary<string, Dictionary<string, object>>();
            for (int i = 0; i < rows; i++)
            {
                if (ExcelRule(result.Tables[0].Rows[i][0].ToString(), ExcelFilterRule)) continue;
                if (result.Tables[0].Rows[i][0].ToString() == ExcelTitleCheck + _fileName)
                {
                    titleIndex = i;
                    continue;
                }
                if (_excelData.ContainsKey(result.Tables[0].Rows[i][0].ToString())){ Debug.LogError($"Excel表：{_FilePath}中存在相同的特征：{result.Tables[0].Rows[i][0].ToString()}，已被移除！"); continue; }
                value = new Dictionary<string, object>();
                for (int j = 1; j < columns; j++)
                {
                    if (string.IsNullOrEmpty(result.Tables[0].Rows[titleIndex][j].ToString())) continue;
                    if (!TitleError.Exists(e => e.EndsWith(result.Tables[0].Rows[titleIndex][j].ToString())) && value.ContainsKey(result.Tables[0].Rows[titleIndex][j].ToString()))
                    {
                        TitleError.Add(result.Tables[0].Rows[titleIndex][j].ToString());
                        Debug.LogError($"Excel表：{_FilePath} 中存在相同的标题: {result.Tables[0].Rows[titleIndex][j].ToString()} 已被移除！");
                        continue;
                    }
                    string _value = result.Tables[0].Rows[i][j].ToString();
                    value.Add(result.Tables[0].Rows[titleIndex][j].ToString(), _value);
                }
                _excelData.Add(result.Tables[0].Rows[i][0].ToString(), value);
            }
            ExcelData excelData = new ExcelData(_excelData, _FilePath);
            return excelData;
        }catch(Exception e)
        {
            if (stream != null) stream.Close();
            if (!Directory.Exists(_FilePath))
            {
                Debug.LogError($"不存在Excel表： {_FilePath}");
                return null;
            }
            if (titleIndex == -1)
            {
                Debug.LogError($"无法打开Excel表：{_FilePath} 异常：没有检测到标题头，请检测标题头是否为 {ExcelTitleCheck}{_fileName} ，或者检查是否没有给非表格内容添加{ShowAllFilterRule(ExcelFilterRule)}字符！");
                return null;
            }
            Debug.LogError($"无法打开Excel表：{_FilePath} 异常：{e.Message}");
            return null;
        }
#else
        return ReadJson(ExcelToJsonResourcesPath, _fileName, _dir);
#endif
    }

    /// <summary>
    /// 读取Json数据
    /// </summary>
    /// <param name="_excelsToJsonResourcesPath">Json文件夹路径</param>
    /// <param name="_fileName">文件名</param>
    /// <param name="_dir">所在文件夹</param>
    /// <returns></returns>
    private static ExcelData ReadJson(string _excelsToJsonResourcesPath, string _fileName, string _dir) {
        string FilePathJson = $"{_excelsToJsonResourcesPath}{_dir +_fileName}";
        try
        {
            TextAsset text = Resources.Load<TextAsset>(FilePathJson);
            Dictionary<string, Dictionary<string, object>> _jsonData = JsonMapper.ToObject<Dictionary<string, Dictionary<string, object>>>(text.text);
            ExcelData jsonData = new ExcelData(_jsonData, $"{ExcelToJsonPath}{_fileName}.json");
            return jsonData;
        }
        catch (Exception e)
        {
            if (!Directory.Exists(FilePathJson))
            {
                Debug.LogError($"不存在Json文件： {FilePathJson}");
                return null;
            }
            Debug.LogError($"无法打开Json文件： {FilePathJson} 异常： {e.Message} ");
            return null;
        }
    }

#if UNITY_EDITOR
    /// <summary>
    /// Excel文件转Json文件
    /// </summary>
    public static void ExcelToJson()
    {
        DeleteAllFiles(ExcelToJsonPath);
        List<FileInfo> allExcelFile = new List<FileInfo>();
        GetFolderSuffixFiles(ExcelsFolderPath, DefaultExcelFileSuffix, allExcelFile);
        foreach (var item in allExcelFile)
        {
            int titleIndex = -1;
            FileStream stream = null;
            try
            {
                if (ExcelRule(Path.GetFileNameWithoutExtension(item.FullName), ExcelTempFileFilterRule))
                {
                    Debug.LogWarning($"已跳过临时文件： {item.FullName}");
                    continue;
                }
                //---- 创建文件夹 Start ----
                string excelDir = item.DirectoryName.Replace('\\', '/').Replace(ExcelsFolderPath.Remove(ExcelsFolderPath.Length - 1), "");
                if (excelDir != "")
                {
                    if (Directory.Exists(ExcelToJsonPath + excelDir) == false) Directory.CreateDirectory(ExcelToJsonPath + excelDir);
                    Debug.Log($"创建文件夹成功： {ExcelToJsonPath + excelDir}");
                    excelDir += "/";
                }
                //---- 创建文件夹  End  ----
                stream = File.Open(item.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                IExcelDataReader excelDataReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet result = excelDataReader.AsDataSet();
                stream.Close();
                excelDataReader.Close();
                int columns = result.Tables[0].Columns.Count;
                int rows = result.Tables[0].Rows.Count;
                List<string> TitleError = new List<string>(); ;
                JsonData Json = new JsonData();
                for (int i = 0; i < rows; i++)
                {
                    if (ExcelRule(result.Tables[0].Rows[i][0].ToString(), ExcelFilterRule)) continue;
                    if (result.Tables[0].Rows[i][0].ToString() == ExcelTitleCheck + Path.GetFileNameWithoutExtension(item.FullName))
                    {
                        titleIndex = i;
                        continue;
                    }
                    if (Json.ContainsKey(result.Tables[0].Rows[i][0].ToString())) { Debug.LogError($"Excel表： {item.FullName} 中存在相同的特征: {result.Tables[0].Rows[i][0].ToString()}，已被移除！"); continue; }
                    JsonData tempJson = new JsonData();
                    for (int j = 1; j < columns; j++)
                    {
                        if (string.IsNullOrEmpty(result.Tables[0].Rows[titleIndex][j].ToString())) continue;
                        if (!TitleError.Exists(e => e.EndsWith(result.Tables[0].Rows[titleIndex][j].ToString())) && tempJson.ContainsKey(result.Tables[0].Rows[titleIndex][j].ToString()))
                        {
                            TitleError.Add(result.Tables[0].Rows[titleIndex][j].ToString());
                            Debug.LogError($"Excel表：{item.FullName} 中存在相同的标题: {result.Tables[0].Rows[titleIndex][j].ToString()} 已被移除！");
                            continue;
                        }
                        string _value = result.Tables[0].Rows[i][j].ToString();
                        tempJson[result.Tables[0].Rows[titleIndex][j].ToString()] = _value;
                    }

                    Json[result.Tables[0].Rows[i][0].ToString()] = tempJson;
                }
                string json = Json.ToJson();
                // 转译Json中文
                string CN_JSON = ParseJsonData(json);
                // 创建
                CreatJsonFile(CN_JSON, ExcelToJsonPath + excelDir + $"{Path.GetFileNameWithoutExtension(item.FullName)}.json");
                Debug.Log($"成功生成文件： {ExcelToJsonPath + $"{Path.GetFileNameWithoutExtension(item.FullName)}.json"}");
                AssetDatabase.Refresh();
            }
            catch (Exception e)
            {
                if (stream != null) stream.Close();
                if (titleIndex == -1)
                {
                    Debug.LogError($"无法打开Excel表：{item.FullName} 异常：没有检测到标题头，请检测标题头是否为{ExcelTitleCheck}{Path.GetFileNameWithoutExtension(item.FullName)}，或者检查是否没有给非表格内容添加{ShowAllFilterRule(ExcelFilterRule)}字符！");
                    return;
                }
                Debug.LogError($"转换失败，出现意外：{e.Message}");
            }
        }
    }

    /// <summary>
    /// 返回所有标题头规则为字符串
    /// </summary>
    /// <param name="_rule">规则列表</param>
    /// <returns></returns>
    private static string ShowAllFilterRule(List<char> _rule) {
        string ruleStr = "";
        foreach (var item in _rule)
        {
            ruleStr += ruleStr == "" ? item.ToString() : ", " + item.ToString();
        }
        return ruleStr;
    }

    /// <summary>
    /// Excel 规则校验 校验返回True说明要跳过这个特征
    /// </summary>
    /// <param name="_value">字符串</param>
    /// <param name="_rule">规则列表</param>
    /// <returns></returns>
    private static bool ExcelRule(string _value, List<char> _rule) {
        if (!string.IsNullOrEmpty(_value)) {
            foreach (var item in _rule)
            {
                if (_value[0] == item)
                {
                    return true;
                }
            }
            return false;
        }
        return true;
    }

    /// <summary>
    /// 解析Json文件 Unicode字符串 为 中文字符串
    /// </summary>
    /// <param name="_json">Json字符串</param>
    /// <returns></returns>
    private static string ParseJsonData(string _json)
    {
        Regex reg = new Regex(@"(?i)\\[uU]([0-9a-f]{4})");
        string targetJson = reg.Replace(_json, delegate (Match m) { return ((char)Convert.ToInt32(m.Groups[1].Value, 16)).ToString(); });
        return targetJson;
    }

    /// <summary>
    /// 创建Json文件
    /// </summary>
    /// <param name="_jsonStr">Json String 变量</param>
    /// <param name="_filePath">Json 文件路径</param>
    private static void CreatJsonFile(string _jsonStr, string _filePath)
    {
        StreamWriter streamWriter;
        FileInfo fileInfo = new FileInfo(_filePath);
        if (!fileInfo.Exists)
        {
            streamWriter = fileInfo.CreateText();
        }
        else
        {
            fileInfo.Delete();
            streamWriter = fileInfo.CreateText();
        }
        streamWriter.Write(_jsonStr);
        streamWriter.Close();
        AssetDatabase.Refresh();
    }

    /// <summary>
    /// 获取指定文件夹下的所有指定格式文件
    /// </summary>
    /// <param name="_path">文件夹路径</param>
    /// <param name="_suffix">文件格式</param>
    /// <param name="_allExcelInfo">赋值的列表</param>
    /// <returns></returns>
    public static List<FileInfo> GetFolderSuffixFiles(string _path, string _suffix, List<FileInfo> _allExcelInfo)
    {
        DirectoryInfo directoryInfo = new DirectoryInfo(_path);
        FileInfo[] fileInfo = directoryInfo.GetFiles();
        DirectoryInfo[] inPathDirInfo = directoryInfo.GetDirectories();
        foreach (FileInfo f in fileInfo)
        {
            if (f.Extension.ToString() == _suffix)
            {
                _allExcelInfo.Add(f);
            }
        }
        //获取子文件夹内的文件列表，递归遍历
        foreach (DirectoryInfo d in inPathDirInfo)
        {
            GetFolderSuffixFiles(d.FullName, _suffix, _allExcelInfo);
        }
        return _allExcelInfo;
    }

    /// <summary>
    /// 删除文件夹内所有数据
    /// </summary>
    /// <param name="_path">文件夹路径</param>
    public static void DeleteAllFiles(string _path)
    {
        try
        {
            DirectoryInfo directoryInfo = new DirectoryInfo(_path);
            foreach (FileInfo f in directoryInfo.GetFiles())
            {
                f.Delete();
                Debug.Log($"成功删除文件： {_path + f.Name}");
            }
            foreach (DirectoryInfo d in directoryInfo.GetDirectories())
            {
                d.Delete(true);
                Debug.Log($"成功删除文件夹： {_path + d.Name}");
            }
            AssetDatabase.Refresh();
        }
        catch (Exception e)
        {
            Debug.LogError($"删除时出现异常：{e.Message}");
            return;
        }
    }
#endif
}
/// <summary>
/// Excel 数据类
/// </summary>
public class ExcelData {
    private string FilePath = "";
    private Dictionary<string, Dictionary<string, object>> Data = new Dictionary<string, Dictionary<string, object>>();

    /// <summary>
    /// 构造函数 传Excel/Json数据跟文件路径过来
    /// </summary>
    /// <param name="_data">Excel数据</param>
    /// <param name="_filePath">路径</param>
    public ExcelData(Dictionary<string, Dictionary<string, object>> _data, string _filePath) {
        if (_data == null || string.IsNullOrEmpty(_filePath))
        {
            Debug.LogError($"存在空数据 Data:{_data}, FilePath:{_filePath}");
            return;
        }
        Data = _data;
        FilePath = _filePath;
    }

    /// <summary>
    /// 获得当前数据的文件路径
    /// </summary>
    public string GetFilePath
    {
        get{ return FilePath; }
    }

    /// <summary>
    /// 获得当前数据的文件名 不含后缀
    /// </summary>
    public string GetFileName {
        get { return Path.GetFileNameWithoutExtension(FilePath); }
    }

    /// <summary>
    /// 读取Excel/Json指定内容字符串
    /// </summary>
    /// <param name="_feature">特征头</param>
    /// <param name="_title">标题头</param>
    /// <returns></returns>
    public string Read(string _feature, string _title) {
        try
        {
            string tempData = Data[_feature][_title].ToString();
            tempData = StringRemoveAt(tempData);
            return tempData;
        }
        catch(Exception e)
        {
            Debug.LogError($"数据错误：{e.Message} 文件：{GetFilePath} 来源-> 特性：{_feature}，标题：{_title}");
            return null;
        }
    }

    /// <summary>
    /// 读取Excel/Json指定内容字符串转INT类型 错误返回-1
    /// </summary>
    /// <param name="_feature">特征头</param>
    /// <param name="_title">标题头</param>
    /// <returns></returns>
    public int ReadInt(string _feature, string _title)
    {
        try
        {
            string tempData = Data[_feature][_title].ToString();
            tempData = StringRemoveAt(tempData);
            return Convert.ToInt32(tempData);
        }
        catch (Exception e)
        {
            Debug.LogError($"数据错误：{e.Message} 文件：{GetFilePath} 来源-> 特性：{_feature}，标题：{_title}");
            return -1;
        }
    }

    /// <summary>
    /// 读取Excel/Json指定内容字符串转float类型 错误返回-1
    /// </summary>
    /// <param name="_feature">特征头</param>
    /// <param name="_title">标题头</param>
    /// <returns></returns>
    public float ReadFloat(string _feature, string _title)
    {
        try
        {
            string tempData = Data[_feature][_title].ToString();
            tempData = StringRemoveAt(tempData);
            return Convert.ToSingle(tempData);
        }
        catch (Exception e)
        {
            Debug.LogError($"数据错误：{e.Message} 文件：{GetFilePath} 来源-> 特性：{_feature}，标题：{_title}");
            return -1;
        }
    }

    /// <summary>
    /// 读取Excel/Json指定内容对象
    /// </summary>
    /// <param name="_feature">特征头</param>
    /// <param name="_title">标题头</param>
    /// <returns></returns>
    public object ReadObject(string _feature, string _title)
    {
        try
        {
            return Data[_feature][_title];
        }
        catch (Exception e)
        {
            Debug.LogError($"数据错误：{e.Message} 文件：{GetFilePath} 来源-> 特性：{_feature}，标题：{_title}");
            return null;
        }
    }

    /// <summary>
    /// 格式化字符串,_types不带参数默认为去换行
    /// </summary>
    /// <param name="_strData">字符串</param>
    /// <param name="_types">去除类型列表 new List<string>(){ 这里面指定要去除的字符串 }</param>
    /// <returns></returns>
    public static string StringRemoveAt(string _strData, List<string> _types = null) {
        if (_types == null) _types = new List<string>() {"\n" };
        foreach (var item in _types)
        {
            _strData = _strData.Replace(item, "");
        }
        return _strData;
    }

    /// <summary>
    /// 读取Excel/Json指定内容转字典 Excel格式为： Key:Value,Key1:Value1
    /// </summary>
    /// <param name="_feature">特征头</param>
    /// <param name="_title">标题头</param>
    /// <returns></returns>
    public Dictionary<string,string> ReadToDict(string _feature, string _title)
    {
        try
        {
            if (Data[_feature] == null || Data[_feature][_title] == null) return null;
            return DictFormat(Data[_feature][_title].ToString());
        }
        catch (Exception e)
        {
            Debug.LogError($"数据错误：{e.Message} 文件：{GetFilePath} 来源-> 特性：{_feature}，标题：{_title}");
            return null;
        }
    }

    /// <summary>
    /// 格式化成字典  格式为： Key:Value,Key1:Value1,Key2:Value2
    /// </summary>
    /// <param name="_strData">字符串</param>
    /// <returns></returns>
    public static Dictionary<string, string> DictFormat(string _strData) {
        if (_strData == null) return null;
        Dictionary<string, string> strDict = new Dictionary<string, string>();
        _strData = StringRemoveAt(_strData);
        _strData = _strData.Replace(" ", "");
        foreach (var item in _strData.Split(','))
        {
            strDict.Add(item.Split(':')[0], item.Split(':')[1]);
        }
        return strDict;
    }

    /// <summary>
    /// 读取Excel/Json指定内容转列表 Excel格式为： Value,Value1,Value2
    /// </summary>
    /// <param name="_feature">特征头</param>
    /// <param name="_title">标题头</param>
    /// <returns></returns>
    public List<string> ReadToList(string _feature, string _title)
    {
        try
        {
            if (Data[_feature] == null || Data[_feature][_title] == null) return null;
            return ListFormat(Data[_feature][_title].ToString());
        }
        catch (Exception e)
        {
            Debug.LogError($"数据错误：{e.Message} 文件：{GetFilePath} 来源-> 特性：{_feature}，标题：{_title}");
            return null;
        }
    }

    /// <summary>
    /// 格式化成列表  格式为： Value,Value1,Value2
    /// </summary>
    /// <param name="_strData">字符串</param>
    /// <returns></returns>
    public static List<string> ListFormat(string _strData) {
        if (_strData == null) return null;
        List<string> strList = new List<string>();
        _strData = StringRemoveAt(_strData);
        _strData = _strData.Replace(" ", "");
        foreach (var item in _strData.Split(','))
        {
            strList.Add(item);
        }
        return strList;
    }

    /// <summary>
    /// 读取Excel/Json所有数据
    /// </summary>
    /// <returns></returns>
    public Dictionary<string, Dictionary<string, object>> ReadAll() {
        return Data;
    }

    /// <summary>
    /// 读取Excel/Json的一个特征
    /// </summary>
    /// <param name="_feature">特征头</param>
    /// <returns></returns>
    public Dictionary<string, object> ReadFeature(string _feature) {
        try
        {
            return Data[_feature];
        }
        catch (Exception e)
        {
            Debug.LogError($"数据错误：{e.Message} 文件：{GetFilePath} 来源-> 特性：{_feature}");
            return null;
        }
    }

    /// <summary>
    /// 所属标题内容合集
    /// </summary>
    /// <param name="_title">标题</param>
    /// <returns></returns>
    public List<string> ReadTitle(string _title) {
        try
        {
            List<string> titleList = new List<string>();
            foreach (var feature in ReadAll())
            {
                foreach (var item in feature.Value)
                {
                    if (item.Key == _title) titleList.Add(item.Value.ToString());
                }
            }
            return titleList;
        }
        catch (Exception e)
        {
            Debug.LogError($"数据错误：{e.Message} 文件：{GetFilePath} 来源-> 标题：{_title}");
            return null;
        }
    }
}