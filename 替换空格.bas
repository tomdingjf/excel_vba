Attribute VB_Name = "ģ��5"
Option Explicit

Sub �滻�ո�()
Dim ��� As Range
For Each ��� In Range("a1:a10")
    If Right(���.Value, 1) = " " Then
        ���.Value = Left(���, Len(���) - 1)
    End If
Next
End Sub
