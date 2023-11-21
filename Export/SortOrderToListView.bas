Attribute VB_Name = "SortOrderToListView"
'@Folder("MVVM.SortOrder.ViewModel")
Option Explicit

Private Const MSO_COLUMN_EXISTS As String = "AcceptInvitation"
Private Const MSO_COLUMN_NOT_EXISTS As String = "DeclineInvitation"
Private Const MSO_SORT_UP As String = "SortUp"
Private Const MSO_SORT_DOWN As String = "SortDown"

Public Sub InitializeListView(ByVal ListView As ListView)
    With ListView
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add text:="#", Width:=24
        .ColumnHeaders.Add text:="Column Name", Width:=80
        .ColumnHeaders.Add text:="Direction", Width:=40
        .ColumnHeaders.Add text:="Sort On", Width:=64
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
        ListItem.ListSubItems.Add text:="Desc", ReportIcon:=MSO_SORT_DOWN
    Else
        ListItem.ListSubItems.Add text:="Asc", ReportIcon:=MSO_SORT_UP
    End If
    
    Select Case SortFieldState.SortOn
    Case xlSortOnCellColor
        ListItem.ListSubItems.Add text:="Cell Color"
    Case xlSortOnFontColor
        ListItem.ListSubItems.Add text:="Font Color"
    Case xlSortOnIcon
        ListItem.ListSubItems.Add text:="Icon"
    Case xlSortOnValues
        If SortFieldState.CustomOrder = Empty Then
            ListItem.ListSubItems.Add text:="Values"
        Else
            ListItem.ListSubItems.Add text:="Values (Custom)"
        End If
    End Select
    
    If Not SortFieldState.Exists Then
        ListItem.SmallIcon = MSO_COLUMN_NOT_EXISTS
    End If
End Sub

