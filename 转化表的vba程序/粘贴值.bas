Attribute VB_Name = "ճ��ֵ"
Option Explicit
Sub ճ��ֵ()
Dim MsgValue As VbMsgBoxResult
    If Not Application.CutCopyMode = False Then
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Else
        MsgValue = MsgBox("δ���ƣ����ȸ��ơ� ", vbOKOnly + vbInformation, "ճ��ֵ")
    End If
'�˳�����ģʽ
'Application.CutCopyMode = False
End Sub
