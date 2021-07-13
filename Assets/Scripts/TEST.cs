using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class TEST : MonoBehaviour
{
    void Start()
    {
        Debug.Log($"读取其中一个内容 A-Title1: {DataManager.test.Read("A", "Title1")}");//读取其中一个内容
        Debug.Log($"读取其中一个内容并转成int类型 A-Title1: {DataManager.test.ReadInt("A", "Title1")}");//读取其中一个内容并转成int类型
        Debug.Log($"读取其中一个内容并转成float类型 A-Title1: {DataManager.test.ReadFloat("A", "Title1")}");//读取其中一个内容并转成float类型

        //读取一个标题的所有内容
        Debug.Log("读取一个标题的所有内容 Titile1");
        foreach (var item in DataManager.test.ReadTitle("Title1"))
        {
            Debug.Log(item);
        }

        //读取一个特征的所有内容
        Debug.Log("读取一个特征的所有内容 A");
        foreach (var item in DataManager.test.ReadFeature("A"))
        {
            Debug.Log($"key: {item.Key} - Value: {item.Value}");
        }

        //读取内容成列表
        Debug.Log("读取内容成列表 E-Title1");
        foreach (var item in DataManager.test.ReadToList("E", "Title1"))
        {
            Debug.Log(item);
        }

        //读取内容成字典
        Debug.Log("读取内容成字典 E-Title2");
        foreach (var item in DataManager.test.ReadToDict("E", "Title2"))
        {
            Debug.Log($"key: {item.Key} - Value: {item.Value}");
        }

        DataManager.test.ReadAll();//读取所有内容

        Debug.Log("更多需求由你扩展......");
        /* 更多需求由你扩展...... */
    }
}
