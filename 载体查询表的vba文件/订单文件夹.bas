Attribute VB_Name = "�����ļ���"
Option Explicit

Sub �����ļ���()
Dim �� As String, �����ļ��� As String, ȫ��� As String, �������� As String
Dim ����� As String, ������ As String
ȫ��� = [B2] '�������
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

ȫ��ַ = �����ļ��� & �������� & "\" & �����ļ��� & "\" & ����
Shell "explorer.exe " & ȫ��ַ, vbNormalFocus
End Sub

Sub ��ת��()
Shell "explorer.exe " & "\\Server\ʵ����\��λ��\��תת����", vbNormalFocus
End Sub
