Attribute VB_Name = "��С���"
Option Explicit
Public Cell As Range
Public �ļ���λ�� As String
Public �ļ����� As String
Public ��С As String
Sub ѭ����������С()
'v_2.3.3�汾 ��������12������������24��������
'                     ����ʱ�亯����ȷ����ʱ��
'                      ����ת��������ݲ�ʹ��
Dim Ƭ���ļ��� As String
Dim �����ļ��� As String
Dim �����ļ��� As String
Dim ���ļ��� As String
Dim ���ļ��� As String
Dim ĩβƬ�κ� As String
Dim ��ʼ As Long
Dim ���� As Long
Dim LastRow As Long
Dim StartTime As Double
Dim EndTime As Double
Dim UserTime As Double

StartTime = Timer

'v_2.3.9 �����滻�ո�ĳ���
Call �滻�ո�.�滻�ո�

LastRow = ��ʽ.Lastrows
On Error Resume Next
��ʼ = 2
���� = LastRow

For Each Cell In Range("k" & ��ʼ, "k" & ����)
    Ƭ���ļ��� = Cell.Offset(0, -4)
    If Cell.Value = "" And VBA.Len(Ƭ���ļ���) <= 7 And VBA.Left(Ƭ���ļ���, 1) Like "[1-9]" Then
            
            If Cell.Offset(0, -6) = "" Then
                    ���ļ��� = "20" & Mid(Cell.Offset(0, -10).Value, 2, 2)
                Else
                    ���ļ��� = "20" & Mid(Cell.Offset(0, -6).Value, 2, 2)
                End If
                 
                    If Val(VBA.Mid(Cell.Offset(0, -4).Value, 2, 1)) Like "[1-9]" Then
                        ���ļ��� = "0" & VBA.Mid(Cell.Offset(0, -4).Value, 2, 1)
                        GoTo 101
                    End If
                    
                    If VBA.Mid(Cell.Offset(0, -4).Value, 2, 1) Like "[a|A]" Then
                        ���ļ��� = "10"
                        GoTo 101
                    End If
                    If VBA.Mid(Cell.Offset(0, -4).Value, 2, 1) Like "[b|B]" Then
                        ���ļ��� = "11"
                        GoTo 101
                    End If
                    If VBA.Mid(Cell.Offset(0, -4).Value, 2, 1) Like "[c|C]" Then
                        ���ļ��� = "12"
                        GoTo 101
                    End If
101:
            �����ļ��� = ���ļ��� & ���ļ���
            ĩβƬ�κ� = VBA.Right(Ƭ���ļ���, 1)
            
            If VBA.Left(Ƭ���ļ���, 1) = 1 Then
                �����ļ��� = VBA.Left(Ƭ���ļ���, 6)
                
                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If

                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & "\" & Ƭ���ļ��� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & "\" & Ƭ���ļ��� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If

                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & "\" & ĩβƬ�κ� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & "\" & ĩβƬ�κ� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & "\" & Ƭ���ļ��� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & "\" & Ƭ���ļ��� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If

                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & "\" & ĩβƬ�κ� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & "\" & ĩβƬ�κ� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                

                
            Else
                �����ļ��� = VBA.Left(Ƭ���ļ���, 5)
               
                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If

                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If

                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & "\" & Ƭ���ļ��� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & "\" & Ƭ���ļ��� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If

                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & "\" & ĩβƬ�κ� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & Ƭ���ļ��� & "\" & ĩβƬ�κ� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & "\" & Ƭ���ļ��� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & "\" & Ƭ���ļ��� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If

                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & "\" & ĩβƬ�κ� & ".txt"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
                
                �ļ���λ�� = "\\Server\ʵ����\����\��������\" & �����ļ��� & "\" & �����ļ��� & "\" & ĩβƬ�κ� & "\" & ĩβƬ�κ� & ".seq"
                If VBA.Len(Dir(�ļ���λ��)) > 0 Then
                    ��С = ��ʽ.Cell
                    Cell.Value = ��С
                GoTo 100
                End If
            End If
    End If
100:

Next Cell

UserTime = Timer - StartTime

Dim MsgValue As VbMsgBoxResult

MsgValue = MsgBox("��������ʱ   " & Format(UserTime, "0.0") & " ��", vbOKOnly + vbInformation, "��С���")

'------------------------------------------------------------------------------------------------------------------------------------------------------
' ���ת����
'    If MsgValue = vbNo Then Exit Sub
'
'        ת���������.ת�������
'------------------------------------------------------------------------------------------------------------------------------------------------------

End Sub


