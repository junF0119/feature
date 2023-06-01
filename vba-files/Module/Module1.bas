Attribute VB_Name = "Module1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Sheets("work").Select
    Rows("3:3").Select
    Range("AJ3").Activate
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Rows("3:270").Select
    Range("AJ3").Activate
    ActiveWorkbook.Worksheets("work").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("work").Sort.SortFields.Add2 Key:=Range("AP4:AP270" _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("work").Sort.SortFields.Add2 Key:=Range("BA4:BA270" _
        ), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("work").Sort
        .SetRange Range("A3:JA270")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=93
End Sub
