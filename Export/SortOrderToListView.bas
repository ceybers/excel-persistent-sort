Attribute VB_Name = "SortOrderToListView"
'@Folder "MVVM.ValueConverters"
Option Explicit

Public Sub InitializeListView(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add text:=COLUMN_INDEX, Width:=24
        .ColumnHeaders.Add text:=COLUMN_NAME, Width:=80
        .ColumnHeaders.Add text:=COLUMN_DIRECTION, Width:=40
        .ColumnHeaders.Add text:=COLUMN_SORT_ON, Width:=64
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
    
    Dim SortFieldState As SortFieldState
    For Each SortFieldState In ViewModel.SelectedSortState.SortFields
        LoadSortFieldStateToListView ListView, SortFieldState
    Next SortFieldState
End Sub

Private Sub LoadSortFieldStateToListView(ByVal ListView As ListView, ByVal SortFieldState As SortFieldState)
    Dim ListItem As ListItem
    Set ListItem = ListView.ListItems.Add(text:=CStr(SortFieldState.Priority), SmallIcon:=MSO_COLUMN_EXISTS)
    
    ListItem.ListSubItems.Add text:=SortFieldState.ColumnName
    If SortFieldState.Order = 0 Then
        ListItem.ListSubItems.Add text:=ITEM_DIRECTION_DESC, ReportIcon:=MSO_SORT_DOWN
    Else
        ListItem.ListSubItems.Add text:=ITEM_DIRECTION_ASC, ReportIcon:=MSO_SORT_UP
    End If
    
    Select Case SortFieldState.SortOn
    Case xlSortOnCellColor
        ListItem.ListSubItems.Add text:=ITEM_CELL_COLOR, ReportIcon:=MSO_CELL_COLOR
    Case xlSortOnFontColor
        ListItem.ListSubItems.Add text:=ITEM_FONT_COLOR, ReportIcon:=MSO_FONT_COLOR
    Case xlSortOnIcon
        ListItem.ListSubItems.Add text:=ITEM_ICON, ReportIcon:=MSO_ICON
    Case xlSortOnValues
        If SortFieldState.CustomOrder = Empty Then
            ListItem.ListSubItems.Add text:=ITEM_VALUES, ReportIcon:=MSO_VALUES
        Else
            ListItem.ListSubItems.Add text:=ITEM_VALUES_CUSTOM, ReportIcon:=MSO_VALUES_CUSTOM
        End If
    End Select
    
    If Not SortFieldState.Exists Then
        ListItem.SmallIcon = MSO_COLUMN_NOT_EXISTS
    End If
End Sub
