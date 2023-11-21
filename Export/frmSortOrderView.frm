VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSortOrderView 
   Caption         =   "Sort Order Manager"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7500
   OleObjectBlob   =   "frmSortOrderView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSortOrderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "MVVM.SortOrder.View"
Option Explicit
Implements IView

Private Type TState
    IsCancelled As Boolean
    ViewModel As SortOrderViewModel
End Type
Private This As TState

Private Sub cmbClose_Click()
    OnCancel
End Sub

Private Sub cmbRemove_Click()
    If vbNo = MsgBox("Remove this Sort Order state?", vbExclamation + vbYesNo + vbDefaultButton2) Then
        Exit Sub
    End If
    
    Debug.Assert Not Me.lvSortOrders.SelectedItem Is Nothing
    Dim Index As Long
    Index = Me.lvSortOrders.SelectedItem.Index
    This.ViewModel.RemoveByIndex Index
    InitalizeFromViewModel
    Set Me.lvSortOrders.SelectedItem = Me.lvSortOrders.ListItems.Item(Index - 1)
    Me.lvSortOrders.SetFocus
End Sub

Private Sub cmbRemoveAll_Click()
    If vbNo = MsgBox("Remove ALL Sort Order states?", vbExclamation + vbYesNo + vbDefaultButton2) Then
        Exit Sub
    End If
    
    This.ViewModel.RemoveAll
    InitalizeFromViewModel
End Sub

Private Sub frmSelectedTable_Click()

End Sub

Private Sub lvSortOrders_DblClick()
    TryApplySortOrder
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub

Private Sub OnCancel()
    This.IsCancelled = True
    Me.Hide
End Sub

Private Function IView_ShowDialog(ByVal ViewModel As Object) As Boolean
    Set This.ViewModel = ViewModel
    
    InitalizeLabelPictures
    InitializeListView
    InitalizeFromViewModel
    This.IsCancelled = False
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitalizeFromViewModel()
    ' Update frmSelectedTable
    Me.txtTableName = This.ViewModel.CurrentSortState.ListObjectName
    Me.txtSortOrder = This.ViewModel.CurrentSortState.GetCaption
    Me.cmbSave.Enabled = This.ViewModel.CurrentSortState.HasSortOrder
    If This.ViewModel.CanSave Then
        Me.cmbSave.Enabled = True
        Me.cmbSave.Caption = "Save"
    Else
        Me.cmbSave.Enabled = False
        Me.cmbSave.Caption = "Saved"
    End If
    
    SortOrderToTreeView.InitializeTreeView Me.tvStates
    SortOrderToListView.InitializeListView Me.lvPreview
    This.ViewModel.LoadToTreeView Me.tvStates
    This.ViewModel.LoadToListView Me.lvPreview, "K001"
    Exit Sub
    
    Me.lvSortOrders.ListItems.Clear
    
    Dim SortOrderState As SortOrderState
    For Each SortOrderState In This.ViewModel.SortOrderStates
        LoadSortOrderStateToListView SortOrderState, Me.lvSortOrders
    Next SortOrderState
End Sub

Private Sub InitializeListView()
    Dim ImageList32 As ImageList
    Set ImageList32 = GetImageList

    With Me.lvSortOrders
        .View = lvwReport
        .ColumnHeaders.Clear
        .ColumnHeaders.Add text:="Sheet Name", Width:=80
        .ColumnHeaders.Add text:="Table Name", Width:=80
        .ColumnHeaders.Add text:="Column Sort Order", Width:=240
        .Gridlines = True
        .HotTracking = False
        .FullRowSelect = True
        Set .SmallIcons = ImageList32
    End With
End Sub

Private Sub TryApplySortOrder()
    If Me.lvSortOrders.SelectedItem Is Nothing Then Exit Sub
    This.ViewModel.ApplySortOrderState Me.lvSortOrders.SelectedItem.Index
    Me.Hide
End Sub

Private Sub LoadSortOrderStateToListView(ByVal SortOrderState As SortOrderState, ByVal ListView As ListView)
    Dim ListItem As ListItem
    With ListView
        Set ListItem = .ListItems.Add(text:=SortOrderState.WorksheetName)
        ListItem.SmallIcon = "HeaderFooterSheetNameInsert"
        ListItem.ListSubItems.Add text:=SortOrderState.ListObjectName, ReportIcon:="TableStyleRowHeaders"
        ListItem.ListSubItems.Add text:=SortOrderState.GetCaption, ReportIcon:="SortDialog"
    End With
    
    If Not SortOrderState.CanApply(This.ViewModel.ListObject) Then
        ListItem.ListSubItems.Item(2).ReportIcon = "CancelRequest"
    End If
    
    If SortOrderState.ListObjectName = This.ViewModel.ListObject.Name Then
        ListItem.ListSubItems.Item(1).ReportIcon = "TableStyleColumnHeaders"
    End If
End Sub

Private Sub InitalizeLabelPictures()
    InitalizeLabelPicture Me.lblOptionsPicture, "AdvancedFileProperties"
    InitalizeLabelPicture Me.lblPreviewSortOrderPicture, "SortDialog"
    InitalizeLabelPicture Me.lblSavedSortOrdersPicture, "StarRatedFull"
    InitalizeLabelPicture Me.lblSelectedTablePicture, "TableAutoFormat"
End Sub

Private Sub InitalizeLabelPicture(ByVal Label As MSForms.Label, ByVal ImageMsoName As String)
    Set Label.Picture = Application.CommandBars.GetImageMso(ImageMsoName, 24, 24)
End Sub
