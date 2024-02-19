Attribute VB_Name = "判断输出"
Option Explicit

Sub 判断输出()
Dim PirntExcel As VbMsgBoxResult
PirntExcel = MsgBox("是否确认输出", vbYesNo + vbInformation, "输出")
If PirntExcel = vbYes Then
    MsgBox "1"
Else
    MsgBox "2"
End If
End Sub
