Attribute VB_Name = "modApplySort"
Option Explicit

Public Sub ApplySort()
Attribute ApplySort.VB_ProcData.VB_Invoke_Func = " \n14"

    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[ColB]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort.SortFields.Add2 _
        Key:=Range("Table1[ColC]"), SortOn:=xlSortOnValues, Order:=xlDescending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
