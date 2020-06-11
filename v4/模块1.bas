Attribute VB_Name = "模块1"
Sub 检测表格()
'检测数据区大小
    Sheets("导入数据").Select
    Dim ROWS, COLUMNS As Integer
    ROWS = 1048576
    COLUMNS = 0
    Do While Cells(ROWS, 1) = 0
        ROWS = ROWS - 1
    Loop
    Do Until Cells(ROWS, COLUMNS + 1) = 0
        COLUMNS = COLUMNS + 1
    Loop
    Sheets("控制面板").Range("C3") = ROWS
    Sheets("控制面板").Range("C4") = COLUMNS
    Sheets("控制面板").Select
    Sheets("成绩条").Cells.Delete Shift:=xlUp
End Sub

Sub Msg()
    waiting.Show
End Sub

