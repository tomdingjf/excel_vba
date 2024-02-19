Attribute VB_Name = "模块3"
Option Explicit

Function 读取文件夹() As Long
    Open 文件夹位置 For Input As #1
        文件内容 = Input$(LOF(1), 1)
        Close #1
        读取文件夹 = VBA.Len(文件内容)
        
End Function
