Attribute VB_Name = "����PUC57����"
Sub ����PUC57����()

    On Error Resume Next
'    Version 2.5.6 ��Ӵ˾����

    Dim StartPucRow As Double
    Dim EndPucRow As Double
    
    StartPucRow = InputBox("����������puc57��ʼ��ţ�") + 1
    EndPucRow = InputBox("����������puc57������ţ�") + 1
        
    Range("h" & StartPucRow, "h" & EndPucRow).Value = "puc57"
    Range("i" & StartPucRow, "i" & EndPucRow).Value = "EcoRV"
    Range("j" & StartPucRow, "j" & EndPucRow).Value = "Ap"
   
End Sub
