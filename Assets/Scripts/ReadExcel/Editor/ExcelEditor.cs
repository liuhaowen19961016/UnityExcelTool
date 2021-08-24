using UnityEngine;
using UnityEditor;
using System.IO;
using System.Collections.Generic;
using OfficeOpenXml;
using System.Reflection;
using System;
using System.Text;

/// <summary>
/// Excel编辑器
/// </summary>
public class ExcelEditor : EditorWindow
{
    static EditorWindow window;

    [MenuItem("Excel工具/打开工具")]
    static void OpenWindow()
    {
        if (window == null)
        {
            window = CreateWindow<ExcelEditor>("Excel工具");
        }
        window.minSize = new Vector2(400, 400);
        window.Show();
        window.Focus();
    }

    /// <summary>
    /// 表格映射类模版文件路径
    /// </summary>
    string SheetMappingClassTemplatePath;
    /// <summary>
    /// 代码自动生成的文件夹路径
    /// </summary>
    static string CodeAutoGenDir;

    private void OnEnable()
    {
        SheetMappingClassTemplatePath = Application.dataPath.Replace("Assets", "")
            + "SheetMappingClassTemplate.txt";
        CodeAutoGenDir = Application.dataPath + "/Scripts/CodeAutoGen/";
    }

    private void OnDestroy()
    {
        ClearSheetCache();
    }

    string excelPath;//excel文件路径
    string exportPath;//ScriptableObject导出路径
    string specifiedMappingClassName;//指定表格映射类名称

    bool showSheetList;//是否显示表格列表
    bool useSpecifiedMappingClass;//是否使用指定表格映射类

    private void OnGUI()
    {
        GUILayout.Space(5);

        if (EditorApplication.isCompiling)
        {
            EditorGUILayout.HelpBox("请等待编译完成", MessageType.Warning);
            return;
        }

        GUILayout.Label("当前选择的Excel文件：" + Path.GetFileName(excelPath), "LODBlackBox");
        if (GUILayout.Button("选择Excel文件"))
        {
            excelPath = EditorUtility.OpenFilePanel("选择Excel文件", excelPath, "xlsx");
        }
        if (string.IsNullOrEmpty(excelPath)) return;
        if (GUILayout.Button("读取Excel文件"))
        {
            LoadExcel(excelPath);
        }
        if (sheetCache == null || sheetCache.Count == 0) return;
        #region 表格列表
        showSheetList = EditorGUILayout.Foldout(showSheetList, "表格列表");
        if (showSheetList)
        {
            GUILayout.BeginVertical("box");
            foreach (var si in sheetCache)
            {
                GUILayout.BeginHorizontal("box");
                si.isEnable = GUILayout.Toggle(si.isEnable, "");
                GUILayout.Label(si.sheet.Name);
                GUILayout.EndHorizontal();
            }
            GUILayout.EndVertical();
        }
        #endregion

        GUILayout.Space(20);

        useSpecifiedMappingClass = EditorGUILayout.Toggle("使用指定表格映射类", useSpecifiedMappingClass);
        if (useSpecifiedMappingClass)
        {
            GUILayout.Label("当前选择的指定表格映射类：" + specifiedMappingClassName, "LODBlackBox");
            if (GUILayout.Button("选择指定表格映射类"))
            {
                string specifiedSheetTypePath = EditorUtility.OpenFilePanel("选择指定表格映射类", excelPath, "cs");
                specifiedMappingClassName = Path.GetFileNameWithoutExtension(specifiedSheetTypePath);
            }
        }
        else
        {
            if (GUILayout.Button("生成表格映射类"))
            {
                foreach (var si in sheetCache)
                {
                    if (si.isEnable)
                    {
                        AutoGenMappingClass(si);
                    }
                }
            }
        }

        GUILayout.Space(20);

        GUILayout.Label("当前选择的ScriptableObject导出路径：" + exportPath, "LODBlackBox");
        if (GUILayout.Button("选择ScriptableObject导出路径"))
        {
            exportPath = EditorUtility.OpenFolderPanel("选择ScriptableObject导出路径", Application.dataPath.Replace("Assets", ""), "");
        }
        if (string.IsNullOrEmpty(exportPath)) return;
        if (useSpecifiedMappingClass && string.IsNullOrEmpty(specifiedMappingClassName)) return;
        if (GUILayout.Button("创建ScriptableObject"))
        {
            int successCounter = 0;
            int failCounter = 0;
            foreach (var si in sheetCache)
            {
                if (si.isEnable)
                {
                    if (useSpecifiedMappingClass)
                    {
                        bool result = LoadSheet(si, specifiedMappingClassName);
                        if (result)
                        {
                            successCounter++;
                        }
                        else
                        {
                            failCounter++;
                        }
                    }
                    else
                    {
                        bool result = LoadSheet(si);
                        if (result)
                        {
                            successCounter++;
                        }
                        else
                        {
                            failCounter++;
                        }
                    }
                }
            }
            AssetDatabase.SaveAssets();
            AssetDatabase.Refresh();
            EditorUtility.DisplayDialog("创建ScriptableObject结果"
                , string.Format("创建成功了{0}个，创建失败了{1}个\n\nScriptableObject生成路径为：\n" + exportPath, successCounter, failCounter)
                , "确定");
        }
    }

    List<SheetInfo> sheetCache = new List<SheetInfo>();
    void AddToSheetCache(SheetInfo sheetInfo)
    {
        sheetCache.Add(sheetInfo);
    }
    void ClearSheetCache()
    {
        sheetCache.Clear();
    }

    /// <summary>
    /// 加载Excel文件(得到所有表格)
    /// </summary>
    void LoadExcel(string excelPath)
    {
        ClearSheetCache();

        var file = new FileInfo(excelPath);
        if (!file.Exists)
        {
            Debug.LogError("该路径下找不到Excel文件： " + excelPath);
            return;
        }

        var package = new ExcelPackage(file);
        int sheetCount = package.Workbook.Worksheets.Count;
        for (int i = 1; i <= sheetCount; i++)
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets[i];
            SheetInfo sheetInfo = new SheetInfo(sheet, true, excelPath);
            AddToSheetCache(sheetInfo);
        }
    }

    /// <summary>
    /// 加载表格数据并赋值ScriptableObject实体(使用此表格的映射类赋值数据)
    /// </summary>      
    bool LoadSheet(SheetInfo sheetInfo)
    {
        ClassFieldInfo[] cfi = GetClassFieldInfo(sheetInfo);
        if (cfi == null)
        {
            return false;
        }
        ExcelWorksheet sheet = sheetInfo.sheet;
        if (sheet == null
            || sheet.Dimension == null)
        {
            Debug.LogError("这是一个空表格：" + sheetInfo.ExcelName + "/" + sheetInfo.SheetName);
            return false;
        }

        ScriptableObject obj = ScriptableObject.CreateInstance(sheetInfo.SheetName);
        if (obj == null)
        {
            Debug.LogError("找不到表格映射类：" + sheetInfo.SheetName);
            return false;
        }

        int totalRow = sheet.Dimension.Rows;
        int totalCol = sheet.Dimension.Columns;
        if (cfi.Length != totalCol)
        {
            Debug.LogError("表格中的字段数量与表格映射类中的字段数量不一致：" + sheetInfo.ExcelName + "/" + sheetInfo.SheetName);
            return false;
        }
        BindingFlags flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        FieldInfo dataArray_fi = obj.GetType().GetField("dataArray", flags);
        Type dataArrayType = dataArray_fi.FieldType.GetElementType();
        Array dataArray = Array.CreateInstance(dataArrayType, totalRow - 3);
        for (int row = 4; row <= totalRow; row++)
        {
            var dataItem = Activator.CreateInstance(dataArrayType);
            for (int col = 1; col <= totalCol; col++)
            {
                string fn = cfi[col - 1].fieldName;
                Type ft = cfi[col - 1].fieldType;
                FieldInfo data_fi = dataItem.GetType().GetField(fn);
                if (data_fi == null)
                {
                    Debug.LogError("表格映射类" + sheetInfo.SheetName + "中找不到字段：" + fn + "，请重新生成表格映射类");
                    return false;
                }
                //非数组类型
                if (!ft.Name.Contains("[]"))
                {
                    object value = sheet.GetValue(row, col);
                    try
                    {
                        value = ChangeType(value, ft);
                    }
                    catch
                    {
                        Debug.LogError("表格" + sheetInfo.SheetName + "中的字段值：" + value + "无法转换为类型：" + ft
                            + "，表格中的位置：" + row + "-" + col);
                        return false;
                    }
                    data_fi.SetValue(dataItem, value);
                }
                //数组类型
                else
                {
                    Type fet = data_fi.FieldType.GetElementType();
                    string[] fsa = sheet.GetValue<string>(row, col).Split(' ');
                    Array valueArray = Array.CreateInstance(fet, fsa.Length);
                    for (int i = 0; i < fsa.Length; i++)
                    {
                        object value = fsa[i];
                        try
                        {
                            value = ChangeType(value, fet);
                        }
                        catch
                        {
                            Debug.LogError("表格" + sheetInfo.SheetName + "中的数组字段值：" + value + "无法转换为类型：" + fet
                                + "，表格中的位置：" + row + "-" + col);
                            return false;
                        }
                        valueArray.SetValue(value, i);
                    }
                    data_fi.SetValue(dataItem, valueArray);
                }
            }
            dataArray.SetValue(dataItem, row - 4);
        }
        dataArray_fi.SetValue(obj, dataArray);

        exportPath = exportPath.Replace(Application.dataPath, "Assets");
        AssetDatabase.CreateAsset(obj, exportPath + "/" + sheetInfo.SheetName + ".asset");
        return true;
    }

    /// <summary>
    /// 加载表格数据并赋值ScriptableObject实体(使用指定表格映射类赋值数据)
    /// </summary>      
    bool LoadSheet(SheetInfo sheetInfo, string specifiedMappingClassName)
    {
        ClassFieldInfo[] cfi = GetClassFieldInfo(specifiedMappingClassName);
        if (cfi == null)
        {
            return false;
        }
        ExcelWorksheet sheet = sheetInfo.sheet;
        if (sheet == null
            || sheet.Dimension == null)
        {
            Debug.LogError("这是一个空表格：" + sheetInfo.ExcelName + "/" + sheetInfo.SheetName);
            return false;
        }

        ScriptableObject obj = ScriptableObject.CreateInstance(specifiedMappingClassName);
        if (obj == null)
        {
            Debug.LogError("找不到表格映射类：" + specifiedMappingClassName);
            return false;
        }

        int totalRow = sheet.Dimension.Rows;
        int totalCol = sheet.Dimension.Columns;
        if (cfi.Length != totalCol)
        {
            Debug.LogError("表格中的字段数量与表格映射类中的字段数量不一致：" + sheetInfo.ExcelName + "/" + sheetInfo.SheetName);
            return false;
        }
        BindingFlags flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
        FieldInfo dataArray_fi = obj.GetType().GetField("dataArray", flags);
        Type dataArrayType = dataArray_fi.FieldType.GetElementType();
        Array dataArray = Array.CreateInstance(dataArrayType, totalRow - 3);
        for (int row = 4; row <= totalRow; row++)
        {
            var dataItem = Activator.CreateInstance(dataArrayType);
            for (int col = 1; col <= totalCol; col++)
            {
                string fn = cfi[col - 1].fieldName;
                Type ft = cfi[col - 1].fieldType;
                FieldInfo data_fi = dataItem.GetType().GetField(fn);
                if (data_fi == null)
                {
                    Debug.LogError("表格映射类" + sheetInfo.SheetName + "中找不到字段：" + fn + "，请重新生成表格映射类");
                    return false;
                }
                //非数组类型
                if (!ft.Name.Contains("[]"))
                {
                    object value = sheet.GetValue(row, col);
                    try
                    {
                        value = ChangeType(value, ft);
                    }
                    catch
                    {
                        Debug.LogError("表格" + sheetInfo.SheetName + "中的字段值：" + value + "无法转换为类型：" + ft
                            + "，表格中的位置：" + row + "-" + col);
                        return false;
                    }
                    data_fi.SetValue(dataItem, value);
                }
                //数组类型
                else
                {
                    Type fet = data_fi.FieldType.GetElementType();
                    string[] fsa = sheet.GetValue<string>(row, col).Split(' ');
                    Array valueArray = Array.CreateInstance(fet, fsa.Length);
                    for (int i = 0; i < fsa.Length; i++)
                    {
                        object value = fsa[i];
                        try
                        {
                            value = ChangeType(value, fet);
                        }
                        catch
                        {
                            Debug.LogError("表格" + sheetInfo.SheetName + "中的数组字段值：" + value + "无法转换为类型：" + fet
                                + "，表格中的位置：" + row + "-" + col);
                            return false;
                        }
                        valueArray.SetValue(value, i);
                    }
                    data_fi.SetValue(dataItem, valueArray);
                }
            }
            dataArray.SetValue(dataItem, row - 4);
        }
        dataArray_fi.SetValue(obj, dataArray);

        exportPath = exportPath.Replace(Application.dataPath, "Assets");
        AssetDatabase.CreateAsset(obj, exportPath + "/" + sheetInfo.SheetName + ".asset");
        return true;
    }

    /// <summary>
    /// 自动生成映射类脚本
    /// </summary>
    void AutoGenMappingClass(SheetInfo sheetInfo)
    {
        ClassFieldInfo[] typeFieldInfo = GetClassFieldInfo(sheetInfo);
        if (typeFieldInfo == null)
        {
            return;
        }

        FileInfo fi = new FileInfo(SheetMappingClassTemplatePath);
        if (!fi.Exists)
        {
            Debug.LogError("该路径下找不到表格映射类模版： " + SheetMappingClassTemplatePath);
            return;
        }

        string dir = CodeAutoGenDir + sheetInfo.ExcelName + "/";
        if (!Directory.Exists(dir))
        {
            Directory.CreateDirectory(dir);
        }

        using (StreamWriter sw = File.CreateText(dir + sheetInfo.SheetName + ".cs"))
        {
            string templateStr = File.ReadAllText(SheetMappingClassTemplatePath);
            StringBuilder templateSb = new StringBuilder(templateStr);
            templateSb.Replace("#CLASS_TYPE#", sheetInfo.SheetName);
            templateSb.Replace("#KEY_TYPE#", typeFieldInfo[0].fieldType.Name);
            StringBuilder sheetDataSb = new StringBuilder();
            foreach (var tfi in typeFieldInfo)
            {
                sheetDataSb.AppendFormat("        public {0} {1};", tfi.fieldType, tfi.fieldName);
                sheetDataSb.AppendLine("\n");
            }
            templateSb.Replace("#SHEETDATA#", sheetDataSb.ToString());
            templateSb.Replace("#KEY_NAME#", typeFieldInfo[0].fieldName);

            sw.Write(templateSb.ToString());
            sw.Flush();
            sw.Close();
        }
        AssetDatabase.Refresh();
    }

    /// <summary>
    /// 得到表格中的所有类字段信息(根据当前表格信息)
    /// </summary>
    ClassFieldInfo[] GetClassFieldInfo(SheetInfo sheetInfo)
    {
        ExcelWorksheet sheet = sheetInfo.sheet;
        if (sheet == null
            || sheet.Dimension == null)
        {
            Debug.LogError("这是一个空表格：" + sheetInfo.ExcelName + "/" + sheetInfo.SheetName);
            return null;
        }

        int totalCol = sheet.Dimension.Columns;
        ClassFieldInfo[] cfi = new ClassFieldInfo[totalCol];
        for (int col = 1; col <= totalCol; col++)
        {
            string fn = sheet.GetValue<string>(1, col);
            string fts = sheet.GetValue<string>(2, col);
            Type ft = TypeMapper.GetType(fts);
            if (ft == null)
            {
                return null;
            }
            cfi[col - 1] = new ClassFieldInfo(fn, ft);
        }
        return cfi;
    }

    /// <summary>
    /// 得到表格中的所有类字段信息(根据指定表格映射类)
    /// </summary>
    ClassFieldInfo[] GetClassFieldInfo(string specifiedMappingClassName)
    {
        ScriptableObject obj = ScriptableObject.CreateInstance(specifiedMappingClassName);
        if (obj == null)
        {
            Debug.LogError("找不到此表格映射类：" + specifiedMappingClassName);
            return null;
        }

        FieldInfo[] fis = obj.GetType().GetNestedType("SheetData").GetFields();
        ClassFieldInfo[] cfi = new ClassFieldInfo[fis.Length];
        for (int i = 0; i < cfi.Length; i++)
        {
            string fn = fis[i].Name;
            string fts = fis[i].FieldType.Name;
            Type ft = TypeMapper.GetType(fts);
            if (ft == null)
            {
                return null;
            }
            cfi[i] = new ClassFieldInfo(fn, ft);
        }
        return cfi;
    }

    /// <summary>
    /// 转换类型(支持枚举类型转换)
    /// </summary>
    object ChangeType(object value, Type type)
    {
        if (type.IsEnum)
        {
            return Enum.Parse(type, value.ToString());
        }
        else
        {
            return Convert.ChangeType(value, type);
        }
    }
}

/// <summary>
/// 表格信息
/// </summary>
public class SheetInfo
{
    //对应的Excel文件名称
    public string ExcelName
    {
        get { return Path.GetFileNameWithoutExtension(excelPath); }
    }
    //表格名称
    public string SheetName
    {
        get { return sheet.Name; }
    }

    public ExcelWorksheet sheet;//表格
    public bool isEnable;//是否开启
    public string excelPath;//对应的Excel文件路径

    public SheetInfo(ExcelWorksheet sheet, bool isEnable, string excelPath)
    {
        this.sheet = sheet;
        this.isEnable = isEnable;
        this.excelPath = excelPath;
    }
}

/// <summary>
/// 类字段信息
/// </summary>
public class ClassFieldInfo
{
    public string fieldName;//字段名称
    public Type fieldType;//字段类型

    public ClassFieldInfo(string fieldName, Type fieldType)
    {
        this.fieldName = fieldName;
        this.fieldType = fieldType;
    }
}
