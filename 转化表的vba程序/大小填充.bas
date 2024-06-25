Attribute VB_Name = "大小填充"
Option Explicit
Public Cell As Range
Public 文件夹位置 As String
Public 文件内容 As String
Public 大小 As String
Sub 循环向下填充大小()
'v_2.3.3版本 加入其它12个方案，共计24个方案。
'                     增加时间函数，确定用时。
'                      增加转到输出表，暂不使用
Dim 片段文件夹 As String
Dim 订单文件夹 As String
Dim 年月文件夹 As String
Dim 年文件夹 As String
Dim 月文件夹 As String
Dim 末尾片段号 As String
Dim 开始 As Long
Dim 结束 As Long
Dim LastRow As Long
Dim StartTime As Double
Dim EndTime As Double
Dim UserTime As Double

StartTime = Timer

'v_2.3.9 增加替换空格的程序
Call 替换空格.替换空格

LastRow = 公式.Lastrows
On Error Resume Next
开始 = 2
结束 = LastRow

For Each Cell In Range("k" & 开始, "k" & 结束)
    片段文件夹 = Cell.Offset(0, -4)
    If Cell.Value = "" And VBA.Len(片段文件夹) <= 7 And VBA.Left(片段文件夹, 1) Like "[1-9]" Then
            
            If Cell.Offset(0, -6) = "" Then
                    年文件夹 = "20" & Mid(Cell.Offset(0, -10).Value, 2, 2)
                Else
                    年文件夹 = "20" & Mid(Cell.Offset(0, -6).Value, 2, 2)
                End If
                 
                    If Val(VBA.Mid(Cell.Offset(0, -4).Value, 2, 1)) Like "[1-9]" Then
                        月文件夹 = "0" & VBA.Mid(Cell.Offset(0, -4).Value, 2, 1)
                        GoTo 101
                    End If
                    
                    If VBA.Mid(Cell.Offset(0, -4).Value, 2, 1) Like "[a|A]" Then
                        月文件夹 = "10"
                        GoTo 101
                    End If
                    If VBA.Mid(Cell.Offset(0, -4).Value, 2, 1) Like "[b|B]" Then
                        月文件夹 = "11"
                        GoTo 101
                    End If
                    If VBA.Mid(Cell.Offset(0, -4).Value, 2, 1) Like "[c|C]" Then
                        月文件夹 = "12"
                        GoTo 101
                    End If
101:
            年月文件夹 = 年文件夹 & 月文件夹
            末尾片段号 = VBA.Right(片段文件夹, 1)
            
            If VBA.Left(片段文件夹, 1) = 1 Then
                订单文件夹 = VBA.Left(片段文件夹, 6)
                
                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If

                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & "\" & 片段文件夹 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & "\" & 片段文件夹 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If

                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & "\" & 末尾片段号 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & "\" & 末尾片段号 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & "\" & 片段文件夹 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & "\" & 片段文件夹 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If

                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & "\" & 末尾片段号 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\金开瑞订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & "\" & 末尾片段号 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                

                
            Else
                订单文件夹 = VBA.Left(片段文件夹, 5)
               
                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If

                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If

                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & "\" & 片段文件夹 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & "\" & 片段文件夹 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If

                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & "\" & 末尾片段号 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 片段文件夹 & "\" & 末尾片段号 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & "\" & 片段文件夹 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & "\" & 片段文件夹 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If

                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & "\" & 末尾片段号 & ".txt"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
                
                文件夹位置 = "\\Server\实验室\订单\华美订单\" & 年月文件夹 & "\" & 订单文件夹 & "\" & 末尾片段号 & "\" & 末尾片段号 & ".seq"
                If VBA.Len(Dir(文件夹位置)) > 0 Then
                    大小 = 公式.Cell
                    Cell.Value = 大小
                GoTo 100
                End If
            End If
    End If
100:

Next Cell

UserTime = Timer - StartTime

Dim MsgValue As VbMsgBoxResult

MsgValue = MsgBox("填充完成用时   " & Format(UserTime, "0.0") & " 秒", vbOKOnly + vbInformation, "大小填充")

'------------------------------------------------------------------------------------------------------------------------------------------------------
' 输出转化表
'    If MsgValue = vbNo Then Exit Sub
'
'        转化输出程序.转化表输出
'------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


