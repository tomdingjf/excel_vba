Attribute VB_Name = "����ת��"
Option Explicit

Sub ����ת��()
Dim CheckBox1 As Boolean
Dim CheckBox2 As Boolean
Dim CheckBox3 As Boolean
CheckBox1 = ThisWorkbook.Sheets("sheet11").OLEObjects("CheckBox1").Object.Value
CheckBox2 = ThisWorkbook.Sheets("sheet11").OLEObjects("CheckBox2").Object.Value
CheckBox3 = ThisWorkbook.Sheets("sheet11").OLEObjects("CheckBox3").Object.Value
   
    If CheckBox1 = True Then
        [a2:f2].Select
    End If

    If CheckBox1 = True Then
        [a3:f3].Select
    End If

    If CheckBox1 = True Then
        [a4:f4].Select
    End If

End Sub



Sub ����������()
   Union(Range("a1:f1"), Range("a3:f3")).Copy
Range("a20").PasteSpecial


End Sub
