Attribute VB_Name = "pucÕ³Ìù"
Option Explicit

Sub pucÕ³Ìù()

Dim SlectRange As Range
Dim ColumnA As Range
Dim ColumnB As Range
Dim ColumnC As Range

Set SlectRange = Selection

Set ColumnA = SlectRange.Columns(1)
Set ColumnB = SlectRange.Columns(2)
Set ColumnC = SlectRange.Columns(3)

ColumnA.Value = "puc57"
ColumnB.Value = "EcoRV"
ColumnC.Value = "AP"

End Sub
