using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class TEST : MonoBehaviour
{
    void Start()
    {
        Debug.Log($"��ȡ����һ������ A-Title1: {DataManager.test.Read("A", "Title1")}");//��ȡ����һ������
        Debug.Log($"��ȡ����һ�����ݲ�ת��int���� A-Title1: {DataManager.test.ReadInt("A", "Title1")}");//��ȡ����һ�����ݲ�ת��int����
        Debug.Log($"��ȡ����һ�����ݲ�ת��float���� A-Title1: {DataManager.test.ReadFloat("A", "Title1")}");//��ȡ����һ�����ݲ�ת��float����

        //��ȡһ���������������
        Debug.Log("��ȡһ��������������� Titile1");
        foreach (var item in DataManager.test.ReadTitle("Title1"))
        {
            Debug.Log(item);
        }

        //��ȡһ����������������
        Debug.Log("��ȡһ���������������� A");
        foreach (var item in DataManager.test.ReadFeature("A"))
        {
            Debug.Log($"key: {item.Key} - Value: {item.Value}");
        }

        //��ȡ���ݳ��б�
        Debug.Log("��ȡ���ݳ��б� E-Title1");
        foreach (var item in DataManager.test.ReadToList("E", "Title1"))
        {
            Debug.Log(item);
        }

        //��ȡ���ݳ��ֵ�
        Debug.Log("��ȡ���ݳ��ֵ� E-Title2");
        foreach (var item in DataManager.test.ReadToDict("E", "Title2"))
        {
            Debug.Log($"key: {item.Key} - Value: {item.Value}");
        }

        DataManager.test.ReadAll();//��ȡ��������

        Debug.Log("��������������չ......");
        /* ��������������չ...... */
    }
}
