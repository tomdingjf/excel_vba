Attribute VB_Name = "删除重复项"
Option Explicit
Sub del()
    ActiveSheet.Range("a1:b13").RemoveDuplicates Columns:=2, Header:=xlYes
'    影响的范围Range("b1:b13")
'    RemoveDuplicates删除重复项方法
'   Columns:=1 表示对比列
'   Header:=xlYes表示表头的有无
End Sub
