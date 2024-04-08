Attribute VB_Name = "AvailableColumnsToListView"
'@Folder "MVVM.ValueConverters"
Option Explicit

Public Sub InitializeListView(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add text:=COLUMN_NAME, Width:=ListView.Width - 16
        .Appearance = cc3D
        .BorderStyle = ccNone
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .HotTracking = False
        .LabelEdit = lvwManual
        'Set .SmallIcons = ImageListHelpers.GetImageList
    End With
End Sub

Public Sub Load(ByVal ViewModel As RemapColumnViewModel, ByVal ListView As ListView)
    ListView.ListItems.Clear
    
    Dim ColumnName As Variant
    For Each ColumnName In ViewModel.ColumnNames
        LoadColumnToListView ListView, ColumnName
    Next ColumnName
End Sub

Private Sub LoadColumnToListView(ByVal ListView As ListView, ByVal ColumnName As String)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(text:=ColumnName)
End Sub
