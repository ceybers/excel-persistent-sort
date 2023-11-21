Attribute VB_Name = "ListObjectHelpers"
'@Folder("MVVM.SortOrder.Helpers")
Option Explicit

Public Function GetAllListObjects(ByVal Workbook As Workbook) As Collection
    Set GetAllListObjects = New Collection
    
    If Workbook Is Nothing Then Exit Function
    
    Dim Worksheet As Worksheet
    Dim ListObject As ListObject
    
    For Each Worksheet In Workbook.Worksheets
        For Each ListObject In Worksheet.ListObjects
            GetAllListObjects.Add Item:=ListObject, Key:=ListObject.Name
        Next ListObject
    Next Worksheet
End Function

Public Function HasListColumn(ByVal ListObject As ListObject, ByVal ListColumnName As String) As Boolean
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        If ListColumn.Name = ListColumnName Then
            HasListColumn = True
            Exit Function
        End If
    Next ListColumn
End Function

