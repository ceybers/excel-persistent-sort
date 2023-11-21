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

Public Sub Load(ByVal ViewModel As SortOrderViewModel, ByVal ListView As ListView)
    ListView.ListItems.Clear
    If ViewModel.SelectedSortState Is Nothing Then Exit Sub
    
    Dim i As Long
    i = 1
    
    Dim SortFieldState As SortFieldState
    For Each SortFieldState In ViewModel.SelectedSortState.SortFields
        Dim ListItem As ListItem
        Set ListItem = ListView.ListItems.Add(text:=CStr(i), SmallIcon:="AcceptInvitation")
        
        ListItem.ListSubItems.Add text:=SortFieldState.ColumnName
        If SortFieldState.SortOrder = 0 Then
            ListItem.ListSubItems.Add text:="Asc", ReportIcon:="SortUp"
        Else
            ListItem.ListSubItems.Add text:="Desc", ReportIcon:="SortDown"
        End If
        
        If Not ListObjectHelpers.HasListColumn(ByVal ViewModel.ListObject, SortFieldState.ColumnName) Then
            ListItem.SmallIcon = "DeclineInvitation"
        End If
        
        i = i + 1
    Next SortFieldState
    
End Sub
