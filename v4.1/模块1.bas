Attribute VB_Name = "模块1"
Sub 检测表格()
'检测数据区大小
    Sheets("导入数据").Select
    Static ROW, COLUMN As Integer
    ROW = 1048576
    COLUMN = 0
    Do While Cells(ROW, 1) = 0
        ROW = ROW - 1
    Loop
    Do Until Cells(ROW, COLUMN + 1) = 0
        COLUMN = COLUMN + 1
    Loop
    Sheets("控制面板").Range("C3") = ROW
    Sheets("控制面板").Range("C4") = COLUMN
    Sheets("控制面板").Select
    Sheets("成绩条").Cells.Delete Shift:=xlUp
    Sheets("缓存").Cells.Delete Shift:=xlUp
End Sub

Sub Msg()
    waiting.Show
End Sub

