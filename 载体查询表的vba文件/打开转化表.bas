Attribute VB_Name = "打开转化表"
Option Explicit
Sub lastday1()
    Dim rqi As String
    rqi = Format(Date, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday2()
    Dim rqi As String
    rqi = Format(Date - 1, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday3()
    Dim rqi As String
    rqi = Format(Date - 2, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday4()
    Dim rqi As String
    rqi = Format(Date - 3, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday5()
    Dim rqi As String
    rqi = Format(Date - 4, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday6()
    Dim rqi As String
    rqi = Format(Date - 5, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday7()
    Dim rqi As String
    rqi = Format(Date - 6, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday8()
    Dim rqi As String
    rqi = Format(Date - 7, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday9()
    Dim rqi As String
    rqi = Format(Date - 8, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday10()
    Dim rqi As String
    rqi = Format(Date - 9, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday11()
    Dim rqi As String
    rqi = Format(Date - 10, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday12()
    Dim rqi As String
    rqi = Format(Date - 11, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday13()
    Dim rqi As String
    rqi = Format(Date - 12, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday14()
    Dim rqi As String
    rqi = Format(Date - 13, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday15()
    Dim rqi As String
    rqi = Format(Date - 14, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday16()
    Dim rqi As String
    rqi = Format(Date - 15, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday17()
    Dim rqi As String
    rqi = Format(Date - 16, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday18()
    Dim rqi As String
    rqi = Format(Date - 17, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday19()
    Dim rqi As String
    rqi = Format(Date - 18, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday20()
    Dim rqi As String
    rqi = Format(Date - 19, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday21()
    Dim rqi As String
    rqi = Format(Date - 20, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday22()
    Dim rqi As String
    rqi = Format(Date - 21, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub


Sub lastday23()
    Dim rqi As String
    rqi = Format(Date - 22, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub


Sub lastday24()
    Dim rqi As String
    rqi = Format(Date - 23, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub

Sub lastday25()
    Dim rqi As String
    rqi = Format(Date - 24, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub


Sub lastday26()
    Dim rqi As String
    rqi = Format(Date - 25, "yyyy年m月d日")
    Dim wb As Workbook
    On Error Resume Next ' 忽略错误
    Set wb = Workbooks.Open("\\Server\实验室\定位表\连转转化表\连转表_" & rqi & ".xlsx") ' 替换为你的工作簿路径
    If wb Is Nothing Then
        MsgBox "选择的日期工作簿不存在！", vbOKOnly, "温馨提示"
    End If
End Sub


