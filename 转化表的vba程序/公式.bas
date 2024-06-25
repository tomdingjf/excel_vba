Attribute VB_Name = "公式"
Function Lastrows() As Long
    Lastrows = Cells(Rows.Count, 7).End(xlUp).Row
End Function
Function Cell() As Range
                '读取内容
                        Open 文件夹位置 For Input As #1
                        文件内容 = Input$(LOF(1), 1)
                        Close #1
                        大小 = VBA.Len(文件内容)
                        文件夹位置 = ""
End Function

