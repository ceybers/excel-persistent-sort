Attribute VB_Name = "TestPersistentSortOrder"
'@Folder "SortOrder"
Option Explicit

Public Sub ResetSortOrder()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    lo.Sort.SortFields.Clear
    lo.Sort.Apply
End Sub

Public Sub SortByFirstColumn()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    lo.Sort.SortFields.Clear
    lo.Sort.SortFields.Add Key:=lo.ListColumns(1).Range, Order:=xlAscending
    lo.Sort.Apply
End Sub
