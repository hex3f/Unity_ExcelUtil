/// <summary>
/// ���ݹ����� ������������
/// </summary>
public static class DataManager
{
    //------------------------- Excel Start -------------------------
    public readonly static string ExcelToJsonResourcesPath = "ExcelToJson/";// Json�ļ���·�� -> Resources�ļ�����Json�ļ�������·��
    public readonly static string ExcelToJsonPath = "/Resources/ExcelToJson/";// Json�ļ������ļ���·��
#if UNITY_EDITOR
    public readonly static bool JsonDeBug = false;// ����JSON DEBUGģʽ �ڱ༭����ʹ��JSON����
    public readonly static string DefaultExcelFileSuffix = ".xlsx";// Ĭ��Excel��׺
    public readonly static string ExcelTitleCheck = "^";// ExcelTitleCheck + ���� -> У�����ͷ
    public readonly static string ExcelsFolderPath = "/Excels/";// Excel�ļ������ļ���·��
    // Excel�������д�����У���б� ֻУ���ַ���0��λ
    public readonly static System.Collections.Generic.List<char> ExcelFilterRule = new System.Collections.Generic.List<char>() {
        '#',
    };
    // Excel�����ʱ�ļ�����У���б� ֻУ���ַ���0��λ
    public readonly static System.Collections.Generic.List<char> ExcelTempFileFilterRule = new System.Collections.Generic.List<char>() {
        '~',
    };
#endif
    // ��������
    public readonly static ExcelData test = ExcelUtil.Read("TEST");
}
