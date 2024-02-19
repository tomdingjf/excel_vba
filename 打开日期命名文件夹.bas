Attribute VB_Name = "打开日期命名文件夹"
Option Explicit

Sub 打开日期命名文件夹()
    Dim Yue As String, CurrentDate As String
    CurrentDate = Date ' 获取当前日期
    Yue = Format(CurrentDate, "mm")
    Workbooks.Open ("\\Server\实验室\订单跟进表\" & "订单跟进" & Yue)
End Sub


Sub 新建工作薄另存为地址()
    Dim ss As String
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:=("C:\Users\Administrator\Desktop\ss.xlsx")
End Sub


Sub 工作表的保护与解保护()
    Sheet3.Protect Password:="123"
    Sheet3.Unprotect Password:="123"
End Sub

Sub 新建一个文件夹获取月份()
    Dim CurrentMonth As Integer, CurrentDate As Date
    CurrentDate = Date
    CurrentMonth = Month(CurrentDate)
    On Error Resume Next
    VBA.MkDir ("C:\Users\Administrator\Desktop\其它应用\" & CurrentMonth)
End Sub

Sub 粘贴值()
Attribute 粘贴值.VB_Description = "宏由 连转1 录制，时间: 2023/09/01"
Attribute 粘贴值.VB_ProcData.VB_Invoke_Func = " 14"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
End Sub
