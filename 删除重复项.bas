Attribute VB_Name = "ɾ���ظ���"
Option Explicit
Sub del()
    ActiveSheet.Range("a1:b13").RemoveDuplicates Columns:=2, Header:=xlYes
'    Ӱ��ķ�ΧRange("b1:b13")
'    RemoveDuplicatesɾ���ظ����
'   Columns:=1 ��ʾ�Ա���
'   Header:=xlYes��ʾ��ͷ������
End Sub
