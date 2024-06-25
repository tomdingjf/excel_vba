Attribute VB_Name = "转化输出程序"
Sub 转化表输出()


Dim PirntExcel As VbMsgBoxResult
    PirntExcel = MsgBox("是否确认进入输出程序第1步，填充大小？", vbYesNo + vbInformation, "填充大小")
    If PirntExcel = vbYes Then
        Call 大小填充.循环向下填充大小
    End If
    
Dim ShuChuExcel As VbMsgBoxResult
    ShuChuExcel = MsgBox("是否确认进入输出程序第2步，输出表格？", vbYesNo + vbInformation, "输出表格")
    If ShuChuExcel = vbYes Then
        Call 输出程序.输出程序
    End If

End Sub
