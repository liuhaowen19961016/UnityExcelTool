using UnityEngine;
using System.Collections.Generic;
using System;
using System.Linq;

/// <summary>
/// 通过Excel文件中的表格自动生成的映射类
/// <summary>
//!!!!!自动生成的映射类，禁止手动修改!!!!!
public class Sheet1 : ScriptableObject
{
    [SerializeField]
    SheetData[] dataArray;
    public Dictionary<Int32, SheetData> dataMap;

    [Serializable]
    public class SheetData
    {
        public System.Int32 id;

        public System.String name;

        public System.Int32 age;

        public System.Int32[] score;


    }

    public void Init()
    {
        dataMap = dataArray.ToDictionary(key => key.id, sheetData => sheetData);
    }

    public SheetData GetValue(Int32 key)
    {
        if (dataMap == null || !dataMap.ContainsKey(key))
        {
            Debug.LogError("dataMap中找不到key：" + key);
            return null;
        }
        return dataMap[key];
    }
}
