Attribute VB_Name = "订单文件夹"
Option Explicit

Sub 订单文件夹()
Dim 年 As String, 年月文件名 As String, 全编号 As String, 订单分类 As String
Dim 编号月 As String, 订单月 As String
全编号 = [B2] '订单编号
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

全地址 = 订单文件夹 & 订单分类 & "\" & 年月文件名 & "\" & 简编号
Shell "explorer.exe " & 全地址, vbNormalFocus
End Sub

Sub 连转表()
Shell "explorer.exe " & "\\Server\实验室\定位表\连转转化表", vbNormalFocus
End Sub
