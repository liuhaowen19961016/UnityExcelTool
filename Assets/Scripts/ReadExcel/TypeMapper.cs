using System.Collections.Generic;
using System;
using UnityEngine;

/// <summary>
/// 类型映射
/// </summary>
public class TypeMapper
{
    /// <summary>
    /// 类型映射
    /// </summary>
    static readonly Dictionary<string, Type> TypeMap = new Dictionary<string, Type>
    {
            //系统内置类型
            { "SByte",typeof(SByte) },
            { "Int16",typeof(Int16) },
            { "UInt16",typeof(UInt16) },
            { "Int32",typeof(Int32) },
            { "UInt32",typeof(UInt32) },
            { "Int64",typeof(Int64) },
            { "UInt64",typeof(UInt64) },
            { "Single",typeof(Single) },
            { "Double",typeof(Double) },
            { "String",typeof(String) },
            { "Boolean",typeof(Boolean) },
            { "SByte[]",typeof(SByte[]) },
            { "Byte[]",typeof(Byte[]) },
            { "Int16[]",typeof(Int16[]) },
            { "UInt16[]",typeof(UInt16[]) },
            { "Int32[]",typeof(Int32[]) },
            { "UInt32[]",typeof(UInt32[]) },
            { "Int64[]",typeof(Int64[]) },
            { "UInt64[]",typeof(UInt64[]) },
            { "Single[]",typeof(Single[]) },
            { "Double[]",typeof(Double[]) },
            { "String[]",typeof(String[]) },
            { "Boolean[]",typeof(Boolean[]) },
            { "sbyte",typeof(SByte) },
            { "byte",typeof(Byte) },
            { "short",typeof(Int16) },
            { "ushort",typeof(UInt16) },
            { "int",typeof(Int32) },
            { "uint",typeof(UInt32) },
            { "long",typeof(Int64) },
            { "ulong",typeof(UInt64) },
            { "float",typeof(Single) },
            { "double",typeof(Double) },
            { "string",typeof(String) },
            { "bool",typeof(Boolean) },
            { "sbyte[]",typeof(SByte[]) },
            { "byte[]",typeof(Byte[]) },
            { "short[]",typeof(Int16[]) },
            { "ushort[]",typeof(UInt16[]) },
            { "int[]",typeof(Int32[]) },
            { "uint[]",typeof(UInt32[]) },
            { "long[]",typeof(Int64[]) },
            { "ulong[]",typeof(UInt64[]) },
            { "float[]",typeof(Single[]) },
            { "double[]",typeof(Double[]) },
            { "string[]",typeof(String[]) },
            { "bool[]",typeof(Boolean[]) },

            //自定义类型
            { "RewardType",typeof(Boolean[]) },
    };

    /// <summary>
    /// 得到类型
    /// </summary
    public static Type GetType(string typeStr)
    {
        if (!TypeMap.ContainsKey(typeStr))
        {
            Debug.LogError("找不到此类型的映射：" + typeStr);
            return null;
        }
        return TypeMap[typeStr];
    }
}