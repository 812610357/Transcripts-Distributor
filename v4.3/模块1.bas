Attribute VB_Name = "模块1"
Sub 表格检测()
    Do While Sheets("控制面板").Range("C3") = ""
        Sheets("控制面板").Range("C3") = InputBox("请输入表头行数", "提示")
    Loop
    Sheets("导入数据").Select
    Static ROW, COLUMN As Integer
    ROW = Sheets("控制面板").Range("C3") + 1
    COLUMN = 0
    Do Until Cells(ROW + 1, 1) = 0
        ROW = ROW + 1
    Loop
    Do Until Cells(ROW, COLUMN + 1) = 0
        COLUMN = COLUMN + 1
    Loop
    Sheets("控制面板").Range("C7") = ROW
    Sheets("控制面板").Range("C8") = COLUMN
    Sheets("控制面板").Range("C9") = ROW - Sheets("控制面板").Range("C3")
    Sheets("控制面板").Select
End Sub
