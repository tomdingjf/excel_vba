Attribute VB_Name = "输出程序"
Option Explicit
'Ver 1.2.2
'---------------------------------------------------------------------------
'更新主要内容
'1 将抗性列移动到编号列前方
'2 修改定义区域，增加抗性列区域
'3 调整各区域的大小分配
'4 将标题行的字体缩小一个字号，否则会出现边框打印不显示
'5 修改抗性列公式，将抗性简化
'6 其他地方修改优化
'---------------------------------------------------------------------------

'Ver 1.2.3
'1.公式抗性列取前两位
'2.抗性列加粗
'3.抗性列公式修改

'Ver 1.2.4
'1.公式抗性修改
'2.其他优化

'Ver 1.2.5
'1.增加复制程序
'2.优化复制等程序

'Ver 1.2.6
'1.增加复制文件夹

'Ver 1.2.7
'1.自动添加文件夹文件夹

'Ver 1.2.8
'1.优化程序

'Ver 1.2.9
'1.修改输入表格式

'Ver 1.3.0
'1.增加基因大小
'2.优化公式

'Ver 1.3.1
'1.优化公式程序

'Ver 1.3.2
'1.优化程序

'Ver 1.3.3
'1.删除清除程序

'Ver 2.4.2
'1.添加时间显示

'Version 2.4.6-2
'1.添加了确认输入的选项
'2.添加创建文件夹程序，以便在不同计算机运行
'3.添加一键导出功能去除麻烦操作

'Version 2.5.3
'1.合并大小填充模块，运用call来进行调用
'2.将图形转化为vba按钮，并进行优化
'3.将输出表进行模块化设计，并优化

Sub 输出程序()

Dim RiQi As String
Dim CurrentDate As Date
Dim CurrentYear As Integer
Dim CurrentMonth As Integer
Dim CurrentDay As Integer
Dim StartTime As Double
Dim EndTime As Double
Dim UserTime As Double

StartTime = Timer

'--------------------------获取日期，年，月，日------------------------------

CurrentDate = Date ' 获取当前日期
CurrentYear = Year(CurrentDate) ' 获取年份
CurrentMonth = Month(CurrentDate) ' 获取月份
CurrentDay = Day(CurrentDate) ' 获取日数

'--------------------------输出日期
RiQi = CurrentYear & "年" & CurrentMonth & "月" & CurrentDay & "日"
'---------------------------------完毕-----------------------------------
On Error Resume Next

VBA.MkDir ("D:\连转转化表")
VBA.MkDir ("D:\连转转化表\" & CurrentMonth & "月")
'VBA.MkDir ("C:\Users\Ding\Desktop\金开瑞\文件保存处\" & currentMonth & "月")

Dim VerName As String
Dim H%
'--------------------------选择工作表
Sheets(1).Select '选择

'--------------------------选定复制范围数值
H = Range("g1").End(xlDown).Row

'--------------------------将工作表的版本名记录下来

VerName = ActiveSheet.Name

'--------------------------复制有数据的单元格
Range(Cells(1, "f"), Cells(H, "k")).Copy

Workbooks.Add '新建工作簿

Dim Ywj As String

Ywj = "D:\连转转化表\" & CurrentMonth & "月" & "\连转表_" & RiQi & ".xlsx"
'ywj = "C:\Users\Ding\Desktop\金开瑞\文件保存处\" & currentMonth & "月" & "\连转表_" & riqi & ".xlsx"

'--------------------------将新建的工作表保存下来

ActiveWorkbook.SaveAs Filename:=Ywj '将新建的工作簿存在指定文件夹并改名


Sheets("sheet1").Select '选择第一个工作表

Sheets("sheet1").Name = RiQi '改名

'--------------------------将之前复制的粘贴到新建的工作表中，且只粘贴值
Range("a1").PasteSpecial xlPasteValues

'修改输入表格式
Range("a1:a" & H).Select '表示选择a1到h截止行的区域
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

'关闭文件夹，保存
ActiveWorkbook.Close savechanges:=True

Dim Srb As String

VBA.MkDir ("D:\连转转化表\" & "暂时存储")
Srb = "D:\连转转化表\暂时存储\" & "输入表_" & RiQi & ".xlsx"

'srb = "C:\Users\Ding\Desktop\金开瑞\输入文件\" & "输入表_" & riqi & ".xlsx"
FileCopy Ywj, Srb

Workbooks.Open (Ywj)

'-------------------------剪切，将e列剪切到a列前----------------------------------
Columns("A:A").Select
Columns("e:e").Cut
Selection.Insert Shift:=xlShiftToRight

'--------------------------------加标题----------------------------------
Dim i As Integer, jz As Integer, yu As Integer, shang As Integer, yu2 As Integer

'--------------------------------计算下边界----------------------------------
shang = (H - 1) \ 30 '----------------取商
yu = (H - 1) Mod 30 '----------------取余数
yu2 = (yu > 0) * (-1) '----------------将余数转化
jz = (yu2 + shang) * 31 '----------------统计截至行号

'--------------------------------循环添加标题----------------------------------
For i = 32 To jz Step 31
Sheets(1).Rows(i).Insert Shift:=xlDown
Sheets(1).Cells(i, 1).Value = "抗性"
Sheets(1).Cells(i, 2).Value = "序号"
Sheets(1).Cells(i, 3).Value = "编号"
Sheets(1).Cells(i, 4).Value = "载体"
Sheets(1).Cells(i, 5).Value = "酶切位点"
Sheets(1).Cells(i, 6).Value = "大小"
Next i

'--------------------------------定义区域----------------------------------
Dim bth As Range, xhl As Range, nrq As Range, ztl As Range, xkh As Range, srl As Range, bhl As Range, kxl As Range

Set bth = Range("A1:F1,A32:F32,A63:F63,A94:F94,A125:F125,A156:F156,A187:F187,A218:F218,A249:F249,A280:F280,A311:F311,A342:F342,A373:F373,A404:F404,A435:F435,A466:F466,A497:F497,A528:F528,A559:F559,A590:F590,A621:F621,A652:F652,A683:F683,A714:F714,A745:F745,A776:F776,A807:F807,A838:F838,A869:F869,A900:F900,A931:F931")
Set xhl = Range("B2:B31,B33:B62,B64:B93,B95:B124,B126:B155,B157:B186,B188:B217,B219:B248,B250:B279,B281:B310,B312:B341,B343:B372,B374:B403,B405:B434,B436:B465,B467:B496,B498:B527,B529:B558,B560:B589,B591:B620,B622:B651,B653:B682,B684:B713,B715:B744,B746:B775,B777:B806,B808:B837,B839:B868,B870:B899,B901:B930,B932:B961")
Set nrq = Range("C2:F31,C33:F62,C64:F93,C95:F124,C126:F155,C157:F186,C188:F217,C219:F248,C250:F279,C281:F310,C312:F341,C343:F372,C374:F403,C405:F434,C436:F465,C467:F496,C498:F527,C529:F558,C560:F589,C591:F620,C622:F651,C653:F682,C684:F713,C715:F744,C746:F775,C777:F806,C808:F837,C839:F868,C870:F899,C901:F930,C932:F961")
Set ztl = Range("D2:D31,D33:D62,D64:D93,D95:D124,D126:D155,D157:D186,D188:D217,D219:D248,D250:D279,D281:D310,D312:D341,D343:D372,D374:D403,D405:D434,D436:D465,D467:D496,D498:D527,D529:D558,D560:D589,D591:D620,D622:D651,D653:D682,D684:D713,D715:D744,D746:D775,D777:D806,D808:D837,D839:D868,D870:D899,D901:D930,D932:D961")
Set xkh = Range("C3:F30,C34:F61,C65:F92,C96:F123,C127:F154,C158:F185,C189:F216,C220:F247,C251:F278,C282:F309,C313:F340,C344:F371,C375:F402,C406:F433,C437:F464,C468:F495,C499:F526,C530:F557,C561:F588,C592:F619,C623:F650,C654:F681,C685:F712,C716:F743,C747:F774,C778:F805,C809:F836,C840:F867,C871:F898,C902:F929,C933:F960")
Set srl = Range("G2:G31,G33:G62,G64:G93,G95:G124,G126:G155,G157:G186,G188:G217,G219:G248,G250:G279,G281:G310,G312:G341,G343:G372,G374:G403,G405:G434,G436:G465,G467:G496,G498:G527,G529:G558,G560:G589,G591:G620,G622:G651,G653:G682,G684:G713,G715:G744,G746:G775,G777:G806,G808:G837,G839:G868,G870:G899,G901:G930,G932:G961")
Set bhl = Range("C2:C31,C33:C62,C64:C93,C95:C124,C126:C155,C157:C186,C188:C217,C219:C248,C250:C279,C281:C310,C312:C341,C343:C372,C374:C403,C405:C434,C436:C465,C467:C496,C498:C527,C529:C558,C560:C589,C591:C620,C622:C651,C653:C682,C684:C713,C715:C744,C746:C775,C777:C806,C808:C837,C839:C868,C870:C899,C901:C930,C932:C961")
Set kxl = Range("A2:A31,A33:A62,A64:A93,A95:A124,A126:A155,A157:A186,A188:A217,A219:A248,A250:A279,A281:A310,A312:A341,A343:A372,A374:A403,A405:A434,A436:A465,A467:A496,A498:A527,A529:A558,A560:A589,A591:A620,A622:A651,A653:A682,A684:A713,A715:A744,A746:A775,A777:A806,A808:A837,A839:A868,A870:A899,A901:A930,A932:A961")

'---------------------------------调行-----------------------------------
Range("a1:a" & jz).Select '表示选择a1到a截止行的区域
    Selection.RowHeight = 22
    
'---------------------------------调列-----------------------------------

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

'-------------------------------调字体标题行-------------------------------
bth.Select
    Selection.Font.Name = "微软雅黑"
    Selection.Font.Size = 12
    Selection.Font.Bold = True
    Selection.HorizontalAlignment = xlHAlignCenter '居中

'--------------------------------调字体序号列-----------------------------
xhl.Select
    Selection.HorizontalAlignment = xlHAlignCenter
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Bold = True
    Selection.Font.Size = 13

'--------------------------------调字体内容区----------------------------
nrq.Select
    Selection.Font.Name = "Times New Roman"
    Selection.Font.Size = 11
    Selection.ShrinkToFit = True '缩小字体填充
    Selection.HorizontalAlignment = xlHAlignCenter '居中

'--------------------------------调字体编号列----------------------------
bhl.Select
    Selection.Font.Bold = True

'--------------------------------调字体载体列----------------------------
ztl.Select
    Selection.HorizontalAlignment = xlHAlignLeft '靠左
    Selection.Font.Size = 11
    With Selection.Borders(xlEdgeLeft)
        .Weight = xlThin
        
    End With
    
kxl.Select
    Selection.HorizontalAlignment = xlHAlignRight '靠右
    Selection.Font.Size = 11
    Selection.Font.Bold = True
    Selection.Font.Name = "Times New Roman"

'抗性列靠右
[a:a].Select
    Selection.HorizontalAlignment = xlHAlignRight
    Selection.ShrinkToFit = True '缩小字体填充
    
'--------------------------------加粗边框------------------------------------------------------------------------------------------------
bth.Select
    With Selection.Borders(xlEdgeLeft) '--------------------------左边框
        .Weight = xlMedium '-------------------------粗框线
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeTop) '--------------------------上边框
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeBottom) '--------------------------下边框
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeRight) '--------------------------右边框
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With

xhl.Select
    With Selection.Borders(xlEdgeLeft) '--------------------------左边框
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

'-------------------------------加细边框------------------------------------------------------------------------------------------------
xkh.Select
    With Selection.Borders(xlEdgeTop) '--------------------------上边框
        .Weight = xlThin '-------------------------细框线
    End With
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .Weight = xlThin
    End With
    
kxl.Select
    With Selection.Borders(xlEdgeLeft) '--------------------------左边框
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
    With Selection.Borders(xlEdgeBottom)
        .Weight = xlMedium
        .LineStyle = xlContinuous
    End With
  
    
ActiveWorkbook.Save

'---------------------------设置页眉页脚------------------------------------
Dim ws As Worksheet
For Each ws In Sheets '所有工作表
With ws.PageSetup
    '.CenterHorizontally = False '垂直居中，打印页面水平居中
    '.LeftHeader = "&""楷体,加粗倾斜""&8 &A" '最前面的"&"是格式，不能省"&8 &A"是字号和跟的具体内容
    '.LeftHeader = "&""楷体""&8 &A"
    .CenterHeader = "&""黑体,加粗""&22 &连转表"
    .RightHeader = "&""楷体""&14 &A"
    .LeftFooter = "&""Times New Roman""&8" & VerName
    '.CenterFooter = "&""楷体""&8 第 &P 页/共 &N 页"
    .RightFooter = "&""楷体""&8 第 &P 页/共 &N 页"
    
End With
Next ws
Set ws = Nothing
ActiveWorkbook.Save

'---------------------------选定区域打印-----------------------------------
'Dim r%
'r = Range("b1").End(xlDown).Row
'-------下粗框线封口--------------------------------------------------------
'Range(Cells(jz, "a"), Cells(jz, "f")).Select
Range(Cells(jz, "b"), Cells(jz, "f")).Select
    With Selection.Borders(xlEdgeBottom)
            .Weight = xlMedium
            .LineStyle = xlContinuous
    End With
ActiveWorkbook.Save
'-------封口完毕--------------------------------------------------------

'-------打印区域设置-----------------------------------------------------------------------------------------
Range("a1", Cells(jz, "f")).Select
Dim dayin As Range
Set dayin = Selection

With ActiveSheet.PageSetup
    .PrintArea = dayin.Address
    '.Zoom = False
    '.FitToPagesWide = 1 '表示每行在一页的宽度
End With

'-------页边距设置--------------------------------------------------------
With ActiveSheet.PageSetup
    .LeftMargin = Application.InchesToPoints(1)
    .RightMargin = Application.InchesToPoints(0.5)
    '.TopMargin = Application.InchesToPoints(1)
    '.BottomMargin = Application.InchesToPoints(1)
    '-------------------------------------大小都是用英寸来表示。1英寸 = 2.54厘米
End With
'-------设置完毕---------------------------------------------------------------------------------------------

'保存的工作表为保护
ActiveSheet.Protect USERINTERFACEONLY:=True

'关闭文件夹，保存
ActiveWorkbook.Close savechanges:=True

'-------复制文件到指定文件夹--------------------------------------------------------------------------------


Dim Fwj As String

Fwj = "\\Server\实验室\定位表\连转转化表\" & "连转表_" & RiQi & ".xlsx"
'fwj = "C:\Users\Ding\Desktop\金开瑞\文件复制处\" & "连转表_" & riqi & ".xlsx"

FileCopy Ywj, Fwj
Workbooks.Open (Ywj)

'-------打印--------------------------------------------------------
'Selection.PrintOut Copies:=1, Collate:=True
 '------------------------------------------------------------------------

'ActiveWorkbook.Close savechanges:=True '关闭文件夹，保存
'ActiveWorkbook.SaveAs Filename:="G:\cg\连转表原始_" & s & ".xlsx"  '另存为
'ActiveWorkbook.Close savechanges:=False '关闭文件夹，不保存

UserTime = Timer - StartTime

Dim MsgValue As VbMsgBoxResult

MsgValue = MsgBox("输出完毕！用时  " & Format(UserTime, "0.0") & " 秒。", vbOKOnly + vbInformation, "输出")

End Sub
