Attribute VB_Name = "Module1"
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").sort.sortfields.Clear
    With ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").sort.sortfields.Clear
    ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").sort.sortfields.Add( _
        Range("Table1[Gamma]"), xlSortOnFontColor, xlAscending, , xlSortNormal). _
        SortOnValue.Color = RGB(0, 0, 0)
    With ActiveWorkbook.Worksheets("Sheet1").ListObjects("Table1").sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
