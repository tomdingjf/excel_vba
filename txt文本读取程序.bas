Attribute VB_Name = "txt�ı���ȡ����"
Sub txt�ı���ȡ����()

Dim nian As String, wjm As String, qbh As String
qbh = [a1] '�������
qbh = LCase(qbh) '����������еĴ�дת��ΪСд
nian = Year(Date) '��ȡ��ʱ������
If VBA.Mid(qbh, 4, 1) = 1 Then '��ȡ����Ϊ�Ǹ��ļ���
dd = "���𶩵�"
Else
dd = "��������"
End If
bh = VBA.Mid(qbh, 5, 1) '��ȡ�����·��ļ���
If bh = "a" Then
Yue = "10"
ElseIf bh = "b" Then
Yue = "11"
ElseIf bh = "c" Then
Yue = "12"
Else
Yue = "0" & bh
End If

wjm = nian & Yue '�õ������·��ļ�������

Dim t As String, s As String
t = "C:\Users\Administrator\Desktop\����Ӧ��\" & dd & "\" & wjm & "\����.txt"

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
 objStream.LoadFromFile ("C:\Users\Administrator\Desktop\����Ӧ��\���𶩵�\202305\����.txt")
 strData = objStream.ReadText()
' [b11] = Split(strData, vbCrLf)
 ' ��������
  [b11] = strData
 objStream.Close
 Set objStream = Nothing
End Sub

Sub ��ȡseq�ļ�()
Dim �ļ���λ�� As String
Dim �ļ����� As String
Dim ��С As String
'·��λ��
�ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\202306\160256\160256.seq"
'��ȡ����
Open �ļ���λ�� For Input As #1
�ļ����� = Input$(LOF(1), 1)
Close #1
��С = VBA.Len(�ļ�����)
End Sub

Sub ѭ���������()
Dim cell As Range
Dim �ļ���λ�� As String
Dim �ļ����� As String
Dim ��С As String
Dim ����1 As String
Dim ����2 As String
Dim ��ʼ As Long
Dim ���� As Long
��ʼ = InputBox("�����뿪ʼ��Ԫ��")
���� = InputBox("�����뿪ʼ��Ԫ��")
For Each cell In Range("g" & ��ʼ, "g" & ����)
    If cell.Value = "" Then
        ����1 = cell.Offset(0, -6)
        If VBA.Left(����1, 1) = 1 Then
            ����2 = VBA.Left(����1, 6)
                    '·��λ��
          �ļ���λ�� = "\\Server\ʵ����\����\���𶩵�\202306\" & ����2 & "\" & ����1 & "\" & ����1 & ".seq"
        Else
            ����2 = VBA.Left(����1, 5)
                '·��λ��
          �ļ���λ�� = "\\Server\ʵ����\����\��������\202306\" & ����2 & "\" & ����1 & "\" & ����1 & ".seq"
        End If

        '��ȡ����
        Open �ļ���λ�� For Input As #1
        �ļ����� = Input$(LOF(1), 1)
        Close #1
        ��С = VBA.Len(�ļ�����)
            
                cell.Value = ��С
    Else
'               cell.Value = 2
    End If
Next cell
End Sub


Sub SSSSS()
STTT = 2

If Len(Str("C:\Users\Administrator\Desktop\����Ӧ��\ת���� V_1.8.xlsm")) > 0 Then
    GoTo STTT
End If
STTT = 2
End Sub
