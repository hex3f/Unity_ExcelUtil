using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEditor;
 
public class MenuTool : MonoBehaviour
{
#if UNITY_EDITOR
    [MenuItem("BuildTools/ExcelToJson")]
    static void ExcelToJson()
    {
        ExcelUtil.ExcelToJson();
    }
#endif
}