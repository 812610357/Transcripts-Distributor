Attribute VB_Name = "模块3"
Sub 系统重置()
    Sheets("成绩条").Cells.Delete Shift:=xlUp
    Sheets("缓存").Cells.Delete Shift:=xlUp
    Sheets("控制面板").Range("C3") = ""
    Sheets("控制面板").Range("C7") = ""
    Sheets("控制面板").Range("C8") = ""
    Sheets("控制面板").Range("C9") = ""
    Sheets("控制面板").Range("C12") = "=IF(成绩条!C2=0,0,MATCH(""A"",成绩条!A:A)-1)"
    Sheets("控制面板").Range("C13") = "=IF(成绩条!C2=0,0,MATCH(""E"",成绩条!1:1)-3)"
    Sheets("控制面板").Range("C14") = "=COUNTIF(成绩条!A:B,""A"")"
    Sheets("控制面板").Range("C15") = "=INDIRECT(""成绩条!A""&VALUE(控制面板!C12)+2)"
End Sub
