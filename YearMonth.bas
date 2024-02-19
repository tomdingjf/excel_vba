Attribute VB_Name = "YearMonth"
Option Explicit

Function YearMonth() As String
    Dim Mo As String
    Dim Ye As String
 
    If Range("a1") <> "" Then
        Mo = VBA.Mid(Range("a1").Value, 2, 1)
        If Mo Like "[0-9]" Then
            Mo = "0" & Mo
            GoTo 200
         End If
         
         If Mo = "a" Or Mo = "A" Then
            Mo = "10"
            GoTo 200
         End If
         
         If Mo = "b" Or Mo = "B" Then
            Mo = "11"
            GoTo 200
        End If
            
        If Mo = "c" Or Mo = "C" Then
            Mo = "12"
            GoTo 200
        End If
    End If
200:
    Ye = Year(Date)
    
    YearMonth = Ye & Mo
End Function
