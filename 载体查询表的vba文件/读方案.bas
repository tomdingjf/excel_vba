Attribute VB_Name = "读方案"
Function 编码(filepath As String)
    Dim Data
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 1
        .LoadFromFile filepath
        Data = .Read
        .Close
    End With
    If IsNull(Data) Then 编码 = "NO BOM UTF-8": Exit Function
    If UBound(Data) < 1 Then 编码 = "NO BOM UTF-8": Exit Function
    If UBound(Data) >= 1 Then
        If UBound(Data) > 1 Then
            If Hex(Data(0)) = "EF" And Hex(Data(1)) = "BB" And Hex(Data(2)) = "BF" Then
                编码 = "UTF-8": Exit Function
            End If
        End If
        Select Case Hex(Data(0)) & Hex(Data(1))
            Case "FEFF"
                编码 = "UTF-16 big endian": Exit Function
            Case "FFFE"
                编码 = "UTF-16 little endian": Exit Function
            Case Else
                CanBeUTF8 = True
                For i = 1 To LenB(Data)
                    FirstByte = AscB(MidB(Data, i, 1))
                    If &H0 <= FirstByte And FirstByte <= &H7F Then
                        FollowingBytesCount = 0
                    ElseIf &HC2 <= FirstByte And FirstByte <= &HDF Then
                        FollowingBytesCount = 1
                    ElseIf &HE0 <= FirstByte And FirstByte <= &HEF Then
                        FollowingBytesCount = 2
                    ElseIf &HF0 <= FirstByte And FirstByte <= &HF4 Then
                        FollowingBytesCount = 3
                    Else
                        CanBeUTF8 = False: Exit For
                    End If
                    For j = 1 To FollowingBytesCount
                        i = i + 1
                        If i > LenB(Data) Then
                            CanBeUTF8 = False: Exit For
                        End If
                        FollowingByte = AscB(MidB(Data, i, 1))
                        If (&H80 <= FollowingByte And FollowingByte <= &HBF) = False Then
                            CanBeUTF8 = False: Exit For: i = LenB(Data) + 1
                        End If
                    Next
                Next
                编码 = IIf(CanBeUTF8, "NO BOM UTF-8", "ANSI"): Exit Function
            End Select
        End If
End Function


Function dufa()
Application.Volatile
Dim 年 As String, 年月文件名 As String, 全编号 As String, 订单分类 As String
Dim 编号月 As String
Dim 订单月 As String

全编号 = [B2] '订单编号
If 全编号 = "" Then Exit Function
全编号 = LCase(全编号) '将订单编号中的大写转换为小写
年 = "20" & VBA.Mid(全编号, 2, 2) '获取订单年份
If VBA.Mid(全编号, 4, 1) = 1 Then '获取订单为那个文件夹
    订单分类 = "金开瑞订单"
Else
    订单分类 = "华美订单"
End If
    编号月 = VBA.Mid(全编号, 5, 1) '获取订单月份文件夹
If 编号月 = "a" Then
    订单月 = "10"
ElseIf 编号月 = "b" Then
    订单月 = "11"
ElseIf 编号月 = "c" Then
    订单月 = "12"
Else
    订单月 = "0" & 编号月
End If
    年月文件名 = 年 & 订单月 '得到订单月份文件夹名称

Dim 订单文件夹 As String, objStream, 文档内容 As String, 简编号 As String, 全地址 As String, 文件编码 As String

简编号 = VBA.Mid(全编号, 4, 6)

订单文件夹 = "\\Server\实验室\订单\"

全地址 = 订单文件夹 & 订单分类 & "\" & 年月文件名 & "\" & 简编号 & "\方案.txt"

If fileexists(全地址) Then Exit Function

文件编码 = 编码(全地址)

If 文件编码 = "ANSI" Then

Open 全地址 For Input As #1
文档内容 = StrConv(InputB(LOF(1), 1), vbUnicode)
Close #1
dufa = 文档内容
 
Else

Set objStream = CreateObject("ADODB.Stream")
objStream.Charset = "utf-8"
objStream.Open
objStream.LoadFromFile (全地址)
文档内容 = objStream.ReadText()
objStream.Close
Set objStream = Nothing
dufa = 文档内容

End If

End Function

Function fileexists(filepath As String) As Boolean
    fileexists = (Dir(filepath) = "")
End Function
