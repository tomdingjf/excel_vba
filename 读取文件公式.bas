Attribute VB_Name = "ģ��3"
Option Explicit

Function ��ȡ�ļ���() As Long
    Open �ļ���λ�� For Input As #1
        �ļ����� = Input$(LOF(1), 1)
        Close #1
        ��ȡ�ļ��� = VBA.Len(�ļ�����)
        
End Function
