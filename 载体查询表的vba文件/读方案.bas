Attribute VB_Name = "������"
Function ����(filepath As String)
    Dim Data
    With CreateObject("ADODB.Stream")
        .Open
        .Type = 1
        .LoadFromFile filepath
        Data = .Read
        .Close
    End With
    If IsNull(Data) Then ���� = "NO BOM UTF-8": Exit Function
    If UBound(Data) < 1 Then ���� = "NO BOM UTF-8": Exit Function
    If UBound(Data) >= 1 Then
        If UBound(Data) > 1 Then
            If Hex(Data(0)) = "EF" And Hex(Data(1)) = "BB" And Hex(Data(2)) = "BF" Then
                ���� = "UTF-8": Exit Function
            End If
        End If
        Select Case Hex(Data(0)) & Hex(Data(1))
            Case "FEFF"
                ���� = "UTF-16 big endian": Exit Function
            Case "FFFE"
                ���� = "UTF-16 little endian": Exit Function
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
                ���� = IIf(CanBeUTF8, "NO BOM UTF-8", "ANSI"): Exit Function
            End Select
        End If
End Function


Function dufa()
Application.Volatile
Dim �� As String, �����ļ��� As String, ȫ��� As String, �������� As String
Dim ����� As String
Dim ������ As String

ȫ��� = [B2] '�������
If ȫ��� = "" Then Exit Function
ȫ��� = LCase(ȫ���) '����������еĴ�дת��ΪСд
�� = "20" & VBA.Mid(ȫ���, 2, 2) '��ȡ�������
If VBA.Mid(ȫ���, 4, 1) = 1 Then '��ȡ����Ϊ�Ǹ��ļ���
    �������� = "���𶩵�"
Else
    �������� = "��������"
End If
    ����� = VBA.Mid(ȫ���, 5, 1) '��ȡ�����·��ļ���
If ����� = "a" Then
    ������ = "10"
ElseIf ����� = "b" Then
    ������ = "11"
ElseIf ����� = "c" Then
    ������ = "12"
Else
    ������ = "0" & �����
End If
    �����ļ��� = �� & ������ '�õ������·��ļ�������

Dim �����ļ��� As String, objStream, �ĵ����� As String, ���� As String, ȫ��ַ As String, �ļ����� As String

���� = VBA.Mid(ȫ���, 4, 6)

�����ļ��� = "\\Server\ʵ����\����\"

ȫ��ַ = �����ļ��� & �������� & "\" & �����ļ��� & "\" & ���� & "\����.txt"

If fileexists(ȫ��ַ) Then Exit Function

�ļ����� = ����(ȫ��ַ)

If �ļ����� = "ANSI" Then

Open ȫ��ַ For Input As #1
�ĵ����� = StrConv(InputB(LOF(1), 1), vbUnicode)
Close #1
dufa = �ĵ�����
 
Else

Set objStream = CreateObject("ADODB.Stream")
objStream.Charset = "utf-8"
objStream.Open
objStream.LoadFromFile (ȫ��ַ)
�ĵ����� = objStream.ReadText()
objStream.Close
Set objStream = Nothing
dufa = �ĵ�����

End If

End Function

Function fileexists(filepath As String) As Boolean
    fileexists = (Dir(filepath) = "")
End Function
