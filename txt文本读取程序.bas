Attribute VB_Name = "txt文本读取程序"
Sub txt文本读取程序()

Dim nian As String, wjm As String, qbh As String
qbh = [a1] '订单编号
qbh = LCase(qbh) '将订单编号中的大写转换为小写
nian = Year(Date) '获取当时的日历
If VBA.Mid(qbh, 4, 1) = 1 Then '获取订单为那个文件夹
dd = "金开瑞订单"
Else
dd = "华美订单"
End If
bh = VBA.Mid(qbh, 5, 1) '获取订单月份文件夹
If bh = "a" Then
Yue = "10"
ElseIf bh = "b" Then
Yue = "11"
ElseIf bh = "c" Then
Yue = "12"
Else
Yue = "0" & bh
End If

wjm = nian & Yue '得到订单月份文件夹名称

Dim t As String, s As String
t = "C:\Users\Administrator\Desktop\其它应用\" & dd & "\" & wjm & "\方案.txt"

Open t For Input As #1
'Charset "UTF-8"
s = StrConv(InputB(LOF(1), 1), vbUnicode)
Close #1
[a10] = s

End Sub

Sub ReadUTFTxt()
 Dim objStream, strData
 Set objStream = CreateObject("ADODB.Stream")
 objStream.Charset = "utf-8"
 objStream.Open
 objStream.LoadFromFile ("C:\Users\Administrator\Desktop\其它应用\金开瑞订单\202305\方案.txt")
 strData = objStream.ReadText()
' [b11] = Split(strData, vbCrLf)
 ' 处理数据
  [b11] = strData
 objStream.Close
 Set objStream = Nothing
End Sub

Sub 读取seq文件()
Dim 文件夹位置 As String
Dim 文件内容 As String
Dim 大小 As String
'路径位置
文件夹位置 = "\\Server\实验室\订单\金开瑞订单\202306\160256\160256.seq"
'读取内容
Open 文件夹位置 For Input As #1
文件内容 = Input$(LOF(1), 1)
Close #1
大小 = VBA.Len(文件内容)
End Sub

Sub 循环向下填充()
Dim cell As Range
Dim 文件夹位置 As String
Dim 文件内容 As String
Dim 大小 As String
Dim 订单1 As String
Dim 订单2 As String
Dim 开始 As Long
Dim 结束 As Long
开始 = InputBox("请输入开始单元格：")
结束 = InputBox("请输入开始单元格：")
For Each cell In Range("g" & 开始, "g" & 结束)
    If cell.Value = "" Then
        订单1 = cell.Offset(0, -6)
        If VBA.Left(订单1, 1) = 1 Then
            订单2 = VBA.Left(订单1, 6)
                    '路径位置
          文件夹位置 = "\\Server\实验室\订单\金开瑞订单\202306\" & 订单2 & "\" & 订单1 & "\" & 订单1 & ".seq"
        Else
            订单2 = VBA.Left(订单1, 5)
                '路径位置
          文件夹位置 = "\\Server\实验室\订单\华美订单\202306\" & 订单2 & "\" & 订单1 & "\" & 订单1 & ".seq"
        End If

        '读取内容
        Open 文件夹位置 For Input As #1
        文件内容 = Input$(LOF(1), 1)
        Close #1
        大小 = VBA.Len(文件内容)
            
                cell.Value = 大小
    Else
'               cell.Value = 2
    End If
Next cell
End Sub


Sub SSSSS()
STTT = 2

If Len(Str("C:\Users\Administrator\Desktop\其它应用\转化表 V_1.8.xlsm")) > 0 Then
    GoTo STTT
End If
STTT = 2
End Sub
