Attribute VB_Name = "�ļ����������"
Option Explicit

Sub cr()
Dim i As Integer, r As Integer

 r = Range("a1").End(xlDown).Row
 With ActiveSheet.PageSetup
    .RightHeader = "&d&""arial,12""" & r
End With

For i = 32 To r Step 31
    Sheets(3).Rows(i).Insert shift:=xlDown
    Sheets(3).Cells(i, 1).Value = "���"
    Sheets(3).Cells(i, 2).Value = "����"
    Sheets(3).Cells(i, 3).Value = "λ��"
    Sheets(3).Cells(i, 4).Value = "����"
    Sheets(3).Cells(i, 5).Value = "����"
Next i

End Sub

Sub FZ()
Dim yuanwj As String, fuwj As String
yuanwj = "D:\��תת����\��ת��_2023��5��3��.xlsx"
fuwj = "\\Server\ʵ����\��λ��\��תת����\��ת��_2023��5��3��.xlsx"
FileCopy yuanwj, fuwj
End Sub

