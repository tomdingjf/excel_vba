Sub lick()
    Dim RowC As Integer
        Dim widths As Integer
        Dim heights As Integer
        Dim ss As Range
        Set ss = Selection
        widths = ss.width
    RowC = Selection.Rows.Count
    Selection.Resize(RowC, 3).ClearContents
'    ����ѡ����Ϊ�ο��㣬���¶��������С
'    Range("a1").Resize(2, 3).Select
End Sub
