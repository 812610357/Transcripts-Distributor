Attribute VB_Name = "模块1"
Sub 表格检测()
    Do While Sheets("控制面板").Range("C3") = ""
        Sheets("控制面板").Range("C3") = InputBox("请输入表头行数", "提示")
    Loop
    Sheets("导入数据").Select
    Static ROW, COLUMN As Integer
    COLUMN = Sheets("导入数据").UsedRange.COLUMNS.Count
    ROW = Sheets("导入数据").UsedRange.ROWS.Count
    Sheets("控制面板").Range("C7") = ROW
    Sheets("控制面板").Range("C8") = COLUMN
    Sheets("控制面板").Range("C9") = ROW - Sheets("控制面板").Range("C3")
    Sheets("控制面板").Select
End Sub
