Attribute VB_Name = "�滻�ո�"
Option Explicit

Sub �滻�ո�()

Dim ��� As Range
Dim LastRow As Long

LastRow = Range("g1").End(xlDown).Row

For Each ��� In Range("g1:g" & LastRow)
    If Right(���.Value, 1) = " " Then
        ���.Value = Left(���, Len(���) - 1)
    End If
Next
End Sub
