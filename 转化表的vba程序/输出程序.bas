Attribute VB_Name = "�������"
Option Explicit
'Ver 1.2.2
'---------------------------------------------------------------------------
'������Ҫ����
'1 ���������ƶ��������ǰ��
'2 �޸Ķ����������ӿ���������
'3 ����������Ĵ�С����
'4 �������е�������Сһ���ֺţ��������ֱ߿��ӡ����ʾ
'5 �޸Ŀ����й�ʽ�������Լ�
'6 �����ط��޸��Ż�
'---------------------------------------------------------------------------

'Ver 1.2.3
'1.��ʽ������ȡǰ��λ
'2.�����мӴ�
'3.�����й�ʽ�޸�

'Ver 1.2.4
'1.��ʽ�����޸�
'2.�����Ż�

'Ver 1.2.5
'1.���Ӹ��Ƴ���
'2.�Ż����Ƶȳ���

'Ver 1.2.6
'1.���Ӹ����ļ���

'Ver 1.2.7
'1.�Զ�����ļ����ļ���

'Ver 1.2.8
'1.�Ż�����

'Ver 1.2.9
'1.�޸�������ʽ

'Ver 1.3.0
'1.���ӻ����С
'2.�Ż���ʽ

'Ver 1.3.1
'1.�Ż���ʽ����

'Ver 1.3.2
'1.�Ż�����

'Ver 1.3.3
'1.ɾ���������

'Ver 2.4.2
'1.���ʱ����ʾ

'Version 2.4.6-2
'1.�����ȷ�������ѡ��
'2.��Ӵ����ļ��г����Ա��ڲ�ͬ���������
'3.���һ����������ȥ���鷳����

'Version 2.5.3
'1.�ϲ���С���ģ�飬����call�����е���
'2.��ͼ��ת��Ϊvba��ť���������Ż�
'3.����������ģ�黯��ƣ����Ż�

Sub �������()

Dim RiQi As String
Dim CurrentDate As Date
Dim CurrentYear As Integer
Dim CurrentMonth As Integer
Dim CurrentDay As Integer
Dim StartTime As Double
Dim EndTime As Double
Dim UserTime As Double

StartTime = Timer

'--------------------------��ȡ���ڣ��꣬�£���------------------------------

CurrentDate = Date ' ��ȡ��ǰ����
CurrentYear = Year(CurrentDate) ' ��ȡ���
CurrentMonth = Month(CurrentDate) ' ��ȡ�·�
CurrentDay = Day(CurrentDate) ' ��ȡ����

'--------------------------�������
RiQi = CurrentYear & "��" & CurrentMonth & "��" & CurrentDay & "��"
'---------------------------------���-----------------------------------
On Error Resume Next

VBA.MkDir ("D:\��תת����")
VBA.MkDir ("D:\��תת����\" & CurrentMonth & "��")
'VBA.MkDir ("C:\Users\Ding\Desktop\����\�ļ����洦\" & currentMonth & "��")

Dim VerName As String
Dim H%
'--------------------------ѡ������
Sheets(1).Select 'ѡ��

'--------------------------ѡ�����Ʒ�Χ��ֵ
H = Range("g1").End(xlDown).Row

'--------------------------��������İ汾����¼����

VerName = ActiveSheet.Name

'--------------------------���������ݵĵ�Ԫ��
Range(Cells(1, "f"), Cells(H, "k")).Copy

Workbooks.Add '�½�������

Dim Ywj As String

Ywj = "D:\��תת����\" & CurrentMonth & "��" & "\��ת��_" & RiQi & ".xlsx"
'ywj = "C:\Users\Ding\Desktop\����\�ļ����洦\" & currentMonth & "��" & "\��ת��_" & riqi & ".xlsx"

'--------------------------���½��Ĺ�����������

ActiveWorkbook.SaveAs Filename:=Ywj '���½��Ĺ���������ָ���ļ��в�����


Sheets("sheet1").Select 'ѡ���һ��������

Sheets("sheet1").Name = RiQi '����

'--------------------------��֮ǰ���Ƶ�ճ�����½��Ĺ������У���ֻճ��ֵ
Range("a1").PasteSpecial xlPasteValues

'�޸�������ʽ
Range("a1:a" & H).Select '��ʾѡ��a1��h��ֹ�е�����
    Selection.RowHeight = 22
Columns("A:A").Select
    Selection.ColumnWidth = 4.25
Columns("b:b").Select
    Selection.ColumnWidth = 13.88
Columns("c:c").Select
    Selection.ColumnWidth = 41
Columns("d:d").Select
    Selection.ColumnWidth = 14
Columns("e:e").Select
    Selection.ColumnWidth = 6.5
Columns("f:f").Select
    Selection.ColumnWidth = 4.5

'�ر��ļ��У�����
ActiveWorkbook.Close savechanges:=True

Dim Srb As String

VBA.MkDir ("D:\��תת����\" & "��ʱ�洢")
Srb = "D:\��תת����\��ʱ�洢\" & "�����_" & RiQi & ".xlsx"

'srb = "C:\Users\Ding\Desktop\����\�����ļ�\" & "�����_" & riqi & ".xlsx"
FileCopy Ywj, Srb

Workbooks.Open (Ywj)

'-------------------------���У���e�м��е�a��ǰ----------------------------------
Columns("A:A").Select
Columns("e:e").Cut
Selection.Insert Shift:=xlShiftToRight

'--------------------------------�ӱ���----------------------------------
Dim i As Integer, jz As Integer, yu As Integer, shang As Integer, yu2 As Integer

'--------------------------------�����±߽�----------------------------------
shang = (H - 1) \ 30 '----------------ȡ��
yu = (H - 1) Mod 30 '----------------ȡ����
yu2 = (yu > 0) * (-1) '----------------������ת��
jz = (yu2 + shang) * 31 '----------------ͳ�ƽ����к�

'--------------------------------ѭ����ӱ���----------------------------------
For i = 32 To jz Step 31
Sheets(1).Rows(i).Insert Shift:=xlDown
Sheets(1).Cells(i, 1).Value = "����"
Sheets(1).Cells(i, 2).Value = "���"
Sheets(1).Cells(i, 3).Value = "���"
Sheets(1).Cells(i, 4).Value = "����"
Sheets(1).Cells(i, 5).Value = "ø��λ��"
Sheets(1).Cells(i, 6).Value = "��С"
Next i

'--------------------------------��������----------------------------------
Dim bth As Range, xhl As Range, nrq As Range, ztl As Range, xkh As Range, srl As Range, bhl As Range, kxl As Range

Set bth = Range("A1:F1,A32:F32,A63:F63,A94:F94,A125:F125,A156:F156,A187:F187,A218:F218,A249:F249,A280:F280,A311:F311,A342:F342,A373:F373,A404:F404,A435:F435,A466:F466,A497:F497,A528:F528,A559:F559,A590:F590,A621:F621,A652:F652,A683:F683,A714:F714,A745:F745,A776:F776,A807:F807,A838:F838,A869:F869,A900:F900,A931:F931")
Set xhl = Range("B2:B31,B33:B62,B64:B93,B95:B124,B126:B155,B157:B186,B188:B217,B219:B248,B250:B279,B281:B310,B312:B341,B343:B372,B374:B403,B405:B434,B436:B465,B467:B496,B498:B527,B529:B558,B560:B589,B591:B620,B622:B651,B653:B682,B684:B713,B715:B744,B746:B775,B777:B806,B808:B837,B839:B868,B870:B899,B901:B930,B932:B961")
Set nrq = Range("C2:F31,C33:F62,C64:F93,C95:F124,C126:F155,C157:F186,C188:F217,C219:F248,C250:F279,C281:F310,C312:F341,C343:F372,C374:F403,C405:F434,C436:F465,C467:F496,C498:F527,C529:F558,C560:F589,C591:F620,C622:F651,C653:F682,C684:F713,C715:F744,C746:F775,C777:F806,C808:F837,C839:F868,C870:F899,C901:F930,C932:F961")
Set ztl = Range("D2:D31,D33:D62,D64:D93,D95:D124,D126:D155,D157:D186,D188:D217,D219:D248,D250:D279,D281:D310,D312:D341,D343:D372,D374:D403,D405:D434,D436:D465,D467:D496,D498:D527,D529:D558,D560:D589,D591:D620,D622:D651,D653:D682,D684:D713,D715:D744,D746:D775,D777:D806,D808:D837,D839:D868,D870:D899,D901:D930,D932:D961")
Set xkh = Range("C3:F30,C34:F61,C65:F92,C96:F123,C127:F154,C158:F185,C189:F216,C220:F247,C251:F278,C282:F309,C313:F340,C344:F371,C375:F402,C406:F433,C437:F464,C468:F495,C499:F526,C530:F557,C561:F588,C592:F619,C623:F650,C654:F681,C685:F712,C716:F743,C747:F774,C778:F805,C809:F836,C840:F867,C871:F898,C902:F929,C933:F960")
Set srl = Range("G2:G31,G33:G62,G64:G93,G95:G124,G126:G155,G157:G186,G188:G217,G219:G248,G250:G279,G281:G310,G312:G341,G343:G372,G374:G403,G405:G434,G436:G465,G467:G496,G498:G527,G529:G558,G560:G589,G591:G620,G622:G651,G653:G682,G684:G713,G715:G744,G746:G775,G777:G806,G808:G837,G839:G868,G870:G899,G901:G930,G932:G961")
Set bhl = Range("C2:C31,C33:C62,C64:C93,C95:C124,C126:C155,C157:C186,C188:C217,C219:C248,C250:C279,C281:C310,C312:C341,C343:C372,C374:C403,C405:C434,C436:C465,C467:C496,C498:C527,C529:C558,C560:C589,C591:C620,C622:C651,C653:C682,C684:C713,C715:C744,C746:C775,C777:C806,C808:C837,C839:C868,C870:C899,C901:C930,C932:C961")
Set kxl = Range("A2:A31,A33:A62,A64:A93,A95:A124,A126:A155,A157:A186,A188:A217,A219:A248,A250:A279,A281:A310,A312:A341,A343:A372,A374:A403,A405:A434,A436:A465,A467:A496,A498:A527,A529:A558,A560:A589,A591:A620,A622:A651,A653:A682,A684:A713,A715:A744,A746:A775,A777:A806,A808:A837,A839:A868,A870:A899,A901:A930,A932:A961")

'---------------------------------����-----------------------------------
Range("a1:a" & jz).Select '��ʾѡ��a1��a��ֹ�е�����
    Selection.RowHeight = 22
    
'---------------------------------����-----------------------------------

Columns("A:A").Select
    Selection.ColumnWidth = 5.5
Columns("b:b").Select
    Selection.ColumnWidth = 4.25
Columns("c:c").Select
    Selection.ColumnWidth = 13.88
Columns("d:d").Select
    Selection.ColumnWidth = 42
Columns("e:e").Select
    Selection.ColumnWidth = 14
Columns("f:f").Select
    Selection.ColumnWidth = 4.5

'-------------------------------�����������-------------------------------
bth.Select
    Selection.Font.Name = "΢���ź�"
    Selection.Font.Size = 12
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlHAlignCenter '����

'--------------------------------�����������-----------------------------
xhl.Select
    Selection.HorizontalAlignment = xlHAlignCenter
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Bold = True
    Selection.Font.Size = 13

'--------------------------------������������----------------------------
nrq.Select
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 11
    Selection.ShrinkToFit = True '��С�������
    Selection.HorizontalAlignment = xlHAlignCenter '����

'--------------------------------����������----------------------------
bhl.Select
    Selection.Font.Bold = True

'--------------------------------������������----------------------------
ztl.Select
    Selection.HorizontalAlignment = xlHAlignLeft '����
    Selection.Font.Size = 11
    With Selection.Borders(xlEdgeLeft)
        .Weight = xlThin
        
    End With
    
kxl.Select
    Selection.HorizontalAlignment = xlHAlignRight '����
    Selection.Font.Size = 11
    Selection.Font.Bold = True
    Selection.Font.Name = "Times New Roman"

'�����п���
[a:a].Select
    Selection.HorizontalAlignment = xlHAlignRight
    Selection.ShrinkToFit = True '��С�������
    
'--------------------------------�Ӵֱ߿�------------------------------------------------------------------------------------------------
bth.Select
    With Selection.Borders(xlEdgeLeft) '--------------------------��߿�
        .Weight = xlMedium '-------------------------�ֿ���
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeTop) '--------------------------�ϱ߿�
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeBottom) '--------------------------�±߿�
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeRight) '--------------------------�ұ߿�
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With

xhl.Select
    With Selection.Borders(xlEdgeLeft) '--------------------------��߿�
        .Weight = xlThin
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeRight)
        .Weight = xlThin
        .LineStyle = xlContinuous
    End With

 nrq.Select
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeRight)
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With

'-------------------------------��ϸ�߿�------------------------------------------------------------------------------------------------
xkh.Select
    With Selection.Borders(xlEdgeTop) '--------------------------�ϱ߿�
        .Weight = xlThin '-------------------------ϸ����
    End With
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .Weight = xlThin
    End With
    
kxl.Select
    With Selection.Borders(xlEdgeLeft) '--------------------------��߿�
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
  
    
ActiveWorkbook.Save

'---------------------------����ҳüҳ��------------------------------------
Dim ws As Worksheet
For Each ws In Sheets '���й�����
With ws.PageSetup
    '.CenterHorizontally = False '��ֱ���У���ӡҳ��ˮƽ����
    '.LeftHeader = "&""����,�Ӵ���б""&8 &A" '��ǰ���"&"�Ǹ�ʽ������ʡ"&8 &A"���ֺź͸��ľ�������
    '.LeftHeader = "&""����""&8 &A"
    .CenterHeader = "&""����,�Ӵ�""&22 &��ת��"
    .RightHeader = "&""����""&14 &A"
    .LeftFooter = "&""Times New Roman""&8" & VerName
    '.CenterFooter = "&""����""&8 �� &P ҳ/�� &N ҳ"
    .RightFooter = "&""����""&8 �� &P ҳ/�� &N ҳ"
    
End With
Next ws
Set ws = Nothing
ActiveWorkbook.Save

'---------------------------ѡ�������ӡ-----------------------------------
'Dim r%
'r = Range("b1").End(xlDown).Row
'-------�´ֿ��߷��--------------------------------------------------------
'Range(Cells(jz, "a"), Cells(jz, "f")).Select
Range(Cells(jz, "b"), Cells(jz, "f")).Select
    With Selection.Borders(xlEdgeBottom)
            .Weight = xlMedium
            .LineStyle = xlContinuous
    End With
ActiveWorkbook.Save
'-------������--------------------------------------------------------

'-------��ӡ��������-----------------------------------------------------------------------------------------
Range("a1", Cells(jz, "f")).Select
Dim dayin As Range
Set dayin = Selection

With ActiveSheet.PageSetup
    .PrintArea = dayin.Address
    '.Zoom = False
    '.FitToPagesWide = 1 '��ʾÿ����һҳ�Ŀ��
End With

'-------ҳ�߾�����--------------------------------------------------------
With ActiveSheet.PageSetup
    .LeftMargin = Application.InchesToPoints(1)
    .RightMargin = Application.InchesToPoints(0.5)
    '.TopMargin = Application.InchesToPoints(1)
    '.BottomMargin = Application.InchesToPoints(1)
    '-------------------------------------��С������Ӣ������ʾ��1Ӣ�� = 2.54����
End With
'-------�������---------------------------------------------------------------------------------------------

'����Ĺ�����Ϊ����
ActiveSheet.Protect USERINTERFACEONLY:=True

'�ر��ļ��У�����
ActiveWorkbook.Close savechanges:=True

'-------�����ļ���ָ���ļ���--------------------------------------------------------------------------------


Dim Fwj As String

Fwj = "\\Server\ʵ����\��λ��\��תת����\" & "��ת��_" & RiQi & ".xlsx"
'fwj = "C:\Users\Ding\Desktop\����\�ļ����ƴ�\" & "��ת��_" & riqi & ".xlsx"

FileCopy Ywj, Fwj
Workbooks.Open (Ywj)

'-------��ӡ--------------------------------------------------------
'Selection.PrintOut Copies:=1, Collate:=True
 '------------------------------------------------------------------------

'ActiveWorkbook.Close savechanges:=True '�ر��ļ��У�����
'ActiveWorkbook.SaveAs Filename:="G:\cg\��ת��ԭʼ_" & s & ".xlsx"  '���Ϊ
'ActiveWorkbook.Close savechanges:=False '�ر��ļ��У�������

UserTime = Timer - StartTime

Dim MsgValue As VbMsgBoxResult

MsgValue = MsgBox("�����ϣ���ʱ  " & Format(UserTime, "0.0") & " �롣", vbOKOnly + vbInformation, "���")

End Sub
