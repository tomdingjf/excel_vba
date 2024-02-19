Attribute VB_Name = "文件复制与插入"
Option Explicit

Sub cr()
Dim i As Integer, r As Integer

 r = Range("a1").End(xlDown).Row
 With ActiveSheet.PageSetup
    .RightHeader = "&d&""arial,12""" & r
End With

For i = 32 To r Step 31
    Sheets(3).Rows(i).Insert shift:=xlDown
    Sheets(3).Cells(i, 1).Value = "编号"
    Sheets(3).Cells(i, 2).Value = "载体"
    Sheets(3).Cells(i, 3).Value = "位点"
    Sheets(3).Cells(i, 4).Value = "抗性"
    Sheets(3).Cells(i, 5).Value = "其他"
Next i

End Sub

Sub FZ()
Dim yuanwj As String, fuwj As String
yuanwj = "D:\连转转化表\连转表_2023年5月3日.xlsx"
fuwj = "\\Server\实验室\定位表\连转转化表\连转表_2023年5月3日.xlsx"
FileCopy yuanwj, fuwj
End Sub

