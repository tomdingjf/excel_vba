Attribute VB_Name = "ģ��8"
Option Explicit

Sub cl1()
MsgBox 1
End Sub

Sub cl2()
MsgBox 2
End Sub

Sub cl3()
MsgBox 3
End Sub

Sub PRINT_cll()

Dim Msg1 As VbMsgBoxResult
Dim Msg2 As VbMsgBoxResult
Dim Msg3 As VbMsgBoxResult

Msg1 = MsgBox("��Ҫ��ʼ��", vbYesNo + vbInformation, "Ҫ��ʼ��")
If Msg1 = vbYes Then
    Call cl1
Else
    GoTo 100
End If

Msg2 = MsgBox("��Ҫ��һ����", vbYesNo + vbInformation, "��һ��")
If Msg2 = vbYes Then
    Call cl2
Else
    GoTo 100
End If

Msg3 = MsgBox("��Ҫ��һ����", vbYesNo + vbInformation, "��һ��")
If Msg3 = vbYes Then
    Call cl3
Else
    GoTo 100
End If

100:
End Sub
