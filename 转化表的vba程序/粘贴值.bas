Attribute VB_Name = "粘贴值"
Option Explicit
Sub 粘贴值()
Dim MsgValue As VbMsgBoxResult
    If Not Application.CutCopyMode = False Then
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False
    Else
        MsgValue = MsgBox("未复制！请先复制。 ", vbOKOnly + vbInformation, "粘贴值")
    End If
'退出复制模式
'Application.CutCopyMode = False
End Sub
