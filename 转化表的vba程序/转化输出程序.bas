Attribute VB_Name = "ת���������"
Sub ת�������()


Dim PirntExcel As VbMsgBoxResult
    PirntExcel = MsgBox("�Ƿ�ȷ�Ͻ�����������1��������С��", vbYesNo + vbInformation, "����С")
    If PirntExcel = vbYes Then
        Call ��С���.ѭ����������С
    End If
    
Dim ShuChuExcel As VbMsgBoxResult
    ShuChuExcel = MsgBox("�Ƿ�ȷ�Ͻ�����������2����������", vbYesNo + vbInformation, "������")
    If ShuChuExcel = vbYes Then
        Call �������.�������
    End If

End Sub
