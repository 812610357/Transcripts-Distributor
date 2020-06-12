Attribute VB_Name = "模块2"
Option Explicit
Public ROW, COLUMN, HEAD, DATA, TIMES, i As Integer

Sub 宏()
    ROW = Sheets("控制面板").Range("C3")
    COLUMN = Sheets("控制面板").Range("C4")
    HEAD = Sheets("控制面板").Range("C7")
    DATA = Sheets("控制面板").Range("C8")
    TIMES = Sheets("控制面板").Range("C10")
    Dim k As Integer
    If TIMES = 0 Then
        k = 0
    Else
        k = 1
    End If
    Call 首次
    If k = 1 Then
    Call 插入表头
    Call 叠加
    Else
    COLUMNS("A:B").ClearContents
    Call 插入表头
    Call 输出
    End If
    Sheets("缓存").Cells.Delete Shift:=xlUp
    ActiveSheet.PageSetup.PrintArea = "$C:$AE"
    ActiveWorkbook.Save
    Sheets("控制面板").Select
End Sub

Sub 首次()
'复制到成绩条
    Sheets("导入数据").Select
    Range(Cells(1, 1), Cells(ROW, COLUMN)).Copy
    Sheets("缓存").Select
    Range("C2").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
'来开行距
    For i = HEAD + 4 To ROW + 1 Step 2
        Cells(i, 1) = 1
    Next
    For i = HEAD + 3 To ROW + 1 Step 2
        Cells(i, 2) = 1
    Next
    For i = 1 To HEAD + 1
        COLUMNS("A:B").Select
        Selection.SpecialCells(xlCellTypeConstants, 1).Select
        Selection.EntireRow.Insert
    Next
End Sub

Sub 插入表头()
    Range("B2") = "A"
    For i = HEAD + 3 To (HEAD + 2) * DATA Step HEAD + 2
        Cells(i, 1) = "S"
    Next
    Range(Cells(1, 1), Cells(HEAD + 1, COLUMN + 2)).Copy
    COLUMNS("A:A").Select
    Selection.SpecialCells(xlCellTypeConstants, 2).Select
    ActiveSheet.Paste
    Range("B2").ClearContents
End Sub

Sub 输出()
    Range(Cells(1, 2), Cells((HEAD + 2) * DATA, COLUMN + 2)).Copy
    Sheets("成绩条").Select
    Range("B1").PasteSpecial Paste:=xlPasteColumnWidths, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    ActiveSheet.Paste
    Cells(HEAD + 2, 1) = 1
End Sub

Sub 叠加()
    Sheets("成绩条").Select
    COLUMNS("B:B").Select
    Selection.SpecialCells(xlCellTypeConstants, 2).Select
    Application.CutCopyMode = False
    Selection.EntireRow.Insert
    Sheets("缓存").Select
    Cells(HEAD + 2, 1) = 1
    For i = 1 To TIMES
        COLUMNS("A:B").Select
        Selection.SpecialCells(xlCellTypeConstants, 1).Select
        Selection.EntireRow.Insert
    Next
    Range(Cells(1, 3), Cells((HEAD + TIMES + 2) * DATA, COLUMN + 2)).Copy
    Sheets("成绩条").Select
    Range("C1").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        True, Transpose:=False
    Cells(HEAD + 2, 1) = Cells(HEAD + 2, 1) + 1
End Sub
