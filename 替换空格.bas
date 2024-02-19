Attribute VB_Name = "Ä£¿é5"
Option Explicit

Sub Ìæ»»¿Õ¸ñ()
Dim ±àºÅ As Range
For Each ±àºÅ In Range("a1:a10")
    If Right(±àºÅ.Value, 1) = " " Then
        ±àºÅ.Value = Left(±àºÅ, Len(±àºÅ) - 1)
    End If
Next
End Sub
