Attribute VB_Name = "�ж����"
Option Explicit

Sub �ж����()
Dim PirntExcel As VbMsgBoxResult
PirntExcel = MsgBox("�Ƿ�ȷ�����", vbYesNo + vbInformation, "���")
If PirntExcel = vbYes Then
    MsgBox "1"
Else
    MsgBox "2"
End If
End Sub
