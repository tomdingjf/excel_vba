Attribute VB_Name = "连接PUC57输入"
Sub 连接PUC57输入()

    On Error Resume Next
'    Version 2.5.6 添加此句代码

    Dim StartPucRow As Double
    Dim EndPucRow As Double
    
    StartPucRow = InputBox("请输入连接puc57开始序号：") + 1
    EndPucRow = InputBox("请输入连接puc57结束序号：") + 1
        
    Range("h" & StartPucRow, "h" & EndPucRow).Value = "puc57"
    Range("i" & StartPucRow, "i" & EndPucRow).Value = "EcoRV"
    Range("j" & StartPucRow, "j" & EndPucRow).Value = "Ap"
   
End Sub
