Attribute VB_Name = "SortOrderToListView"
'@Folder("MVVM.SortOrder.ViewModel")
Option Explicit

Public Sub InitializeListView(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add text:="#", Width:=24
        .ColumnHeaders.Add text:="Column Name", Width:=80
        .ColumnHeaders.Add text:="Direction", Width:=40
        .Appearance = cc3D
        .BorderStyle = ccNone
        .Gridlines = True
        .View = lvwReport
        .FullRowSelect = True
        .HotTracking = False
        .LabelEdit = lvwManual
        Set .SmallIcons = ImageListHelpers.GetImageList
    End With
End Sub
