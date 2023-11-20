VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSortOrderView 
   Caption         =   "Sort Order Manager"
   ClientHeight    =   4875
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4740
   OleObjectBlob   =   "frmSortOrderView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSortOrderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.View"
Option Explicit

Private Sub lvSortOrders_DblClick()
    TryApplySortOrder
End Sub

Private Sub UserForm_Initialize()
    InitializeListView
    Dim SortOrders As Collection
    Set SortOrders = GetSavedSortOrders
    
    Dim Item As Variant
    Dim ListItem As ListItem
    For Each Item In SortOrders
        Set ListItem = Me.lvSortOrders.ListItems.Add(text:=Split(Item, ":")(0))
        ListItem.ListSubItems.Add text:=Split(Item, ":")(1)
        ListItem.ListSubItems.Add text:=Split(Item, ":")(2)
    Next Item
End Sub

Private Sub InitializeListView()
    With Me.lvSortOrders
        .View = lvwReport
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add text:="Sheet"
        .ColumnHeaders.Add text:="Table"
        .ColumnHeaders.Add text:="Columns"
        .Gridlines = True
        .HotTracking = False
        .FullRowSelect = True
    End With
End Sub

Private Sub TryApplySortOrder()
    If Me.lvSortOrders.SelectedItem Is Nothing Then Exit Sub
    MsgBox Me.lvSortOrders.SelectedItem.text
    Me.Hide
End Sub
