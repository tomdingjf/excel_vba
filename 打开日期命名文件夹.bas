Attribute VB_Name = "�����������ļ���"
Option Explicit

Sub �����������ļ���()
    Dim Yue As String, CurrentDate As String
    CurrentDate = Date ' ��ȡ��ǰ����
    Yue = Format(CurrentDate, "mm")
    Workbooks.Open ("\\Server\ʵ����\����������\" & "��������" & Yue)
End Sub


Sub �½����������Ϊ��ַ()
    Dim ss As String
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=("C:\Users\Administrator\Desktop\ss.xlsx")
End Sub


Sub ������ı�����Ᵽ��()
    Sheet3.Protect Password:="123"
    Sheet3.Unprotect Password:="123"
End Sub

Sub �½�һ���ļ��л�ȡ�·�()
    Dim CurrentMonth As Integer, CurrentDate As Date
    CurrentDate = Date
    CurrentMonth = Month(CurrentDate)
    On Error Resume Next
    VBA.MkDir ("C:\Users\Administrator\Desktop\����Ӧ��\" & CurrentMonth)
End Sub

Sub ճ��ֵ()
Attribute ճ��ֵ.VB_Description = "���� ��ת1 ¼�ƣ�ʱ��: 2023/09/01"
Attribute ճ��ֵ.VB_ProcData.VB_Invoke_Func = " 14"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
End Sub
