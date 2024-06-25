Attribute VB_Name = "Ìæ»»¿Õ¸ñ"
Option Explicit

Sub Ìæ»»¿Õ¸ñ()

Dim ±àºÅ As Range
Dim LastRow As Long

LastRow = Range("g1").End(xlDown).Row

For Each ±àºÅ In Range("g1:g" & LastRow)
    If Right(±àºÅ.Value, 1) = " " Then
        ±àºÅ.Value = Left(±àºÅ, Len(±àºÅ) - 1)
    End If
Next
End Sub
