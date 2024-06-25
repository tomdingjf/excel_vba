Attribute VB_Name = "A组定位载体"
Option Explicit

Sub A组定位载体()

    Dim lastrow1 As Long
    lastrow1 = ThisWorkbook.Worksheets(1).Range("b37").End(xlUp).Row
    Range("c2", "d" & lastrow1).Copy
    
    Dim wenjian As String
    wenjian = "\\Server\实验室\定位表\连转_A组载体_定位.xlsx"

    Workbooks.Open (wenjian)
    Dim lastrow2 As Long
    lastrow2 = ActiveSheet.Cells(Rows.Count, 3).End(xlUp).Row + 1
    ActiveSheet.Range("d" & lastrow2).PasteSpecial xlPasteValues
    
    Dim lastrow3 As Long
    lastrow3 = ActiveSheet.Cells(Rows.Count, 5).End(xlUp).Row

    ActiveSheet.Range("e" & lastrow2, "e" & lastrow3).Cut
    ActiveSheet.Range("c" & lastrow2).PasteSpecial xlPasteValues
    ActiveSheet.Range("e" & lastrow2, "e" & lastrow3).Clear
    
End Sub



