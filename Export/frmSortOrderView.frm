VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSortOrderView 
   Caption         =   "Sort Order Manager"
   ClientHeight    =   7815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8160
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

Private Sub cmbApply_Click()
    This.ViewModel.Apply
    Me.Hide
End Sub

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

Private Sub tvStates_DblClick()
    If This.ViewModel.Apply Then
        Me.Hide
    End If
End Sub

Private Sub tvStates_NodeClick(ByVal Node As MSComctlLib.Node)
    If This.ViewModel.TrySelect(Node.Key) Then
        UpdateListView
    End If
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
    InitalizeFromViewModel
    This.IsCancelled = False
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitalizeFromViewModel()
    UpdateSelectedTable
    
    SortOrderToTreeView.InitializeTreeView Me.tvStates
    UpdateTreeView
    
    SortOrderToListView.InitializeListView Me.lvPreview
    UpdateListView
End Sub

Private Sub UpdateSelectedTable()

    Me.txtTableName = This.ViewModel.CurrentSortState.ListObjectName
    Me.txtSortOrder = This.ViewModel.CurrentSortState.GetCaption
    Me.cmbSave.Enabled = This.ViewModel.CurrentSortState.HasSortOrder
    
    Me.cmbSave.Enabled = This.ViewModel.CanSave
    Me.cmbSave.Caption = IIf(This.ViewModel.CanSave, "Save", "Saved")
End Sub

Private Sub UpdateTreeView()
    SortOrderToTreeView.Load This.ViewModel, Me.tvStates
End Sub

Private Sub UpdateListView()
    SortOrderToListView.Load This.ViewModel, Me.lvPreview
    If This.ViewModel.SelectedSortState Is Nothing Then Exit Sub
    
    Me.cmbApply.Enabled = This.ViewModel.SelectedSortState.CanApply(This.ViewModel.ListObject)
    Me.cmbApply.Caption = "Apply"
    
    If Not This.ViewModel.CurrentSortState Is Nothing Then
        If This.ViewModel.SelectedSortState.Equals(This.ViewModel.CurrentSortState) Then
            Me.cmbApply.Enabled = False
            Me.cmbApply.Caption = "Applied"
        End If
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
