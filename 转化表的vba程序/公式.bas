Attribute VB_Name = "��ʽ"
Function Lastrows() As Long
    Lastrows = Cells(Rows.Count, 7).End(xlUp).Row
End Function
Function Cell() As Range
                '��ȡ����
                        Open �ļ���λ�� For Input As #1
                        �ļ����� = Input$(LOF(1), 1)
                        Close #1
                        ��С = VBA.Len(�ļ�����)
                        �ļ���λ�� = ""
End Function

