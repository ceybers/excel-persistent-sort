VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRemapColumn 
   Caption         =   "Remap Column Name"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmRemapColumn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRemapColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("MVVM.SortOrder.Views")
Option Explicit
Implements IView

Private Type TState
    IsCancelled As Boolean
    ViewModel As RemapColumnViewModel
End Type

Private This As TState

Private Sub cmbRemap_Click()
    Me.Hide
End Sub

Private Sub cmbCancel_Click()
    OnCancel
End Sub

Private Sub lvRemapTo_ItemClick(ByVal Item As MSComctlLib.ListItem)
    This.ViewModel.TrySelect Item.text
    UpdateControls
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
    UpdateControls
    This.IsCancelled = False
    
    Me.Show
    
    IView_ShowDialog = Not This.IsCancelled
End Function

Private Sub InitalizeFromViewModel()
    AvailableColumnsToListView.InitializeListView Me.lvRemapTo
    AvailableColumnsToListView.Load This.ViewModel, Me.lvRemapTo
End Sub

Private Sub InitalizeLabelPictures()
    InitalizeLabelPicture Me.lblCurrentColumnPicture, "GroupFieldsAndColumns"
    InitalizeLabelPicture Me.lblRemapToPicture, "DatasheetColumnRename"
End Sub

Private Sub InitalizeLabelPicture(ByVal Label As MSForms.Label, ByVal ImageMsoName As String)
    Set Label.Picture = Application.CommandBars.GetImageMso(ImageMsoName, 24, 24)
End Sub

Private Sub UpdateControls()
    Me.txtColumnName = This.ViewModel.CurrentColumnName
    Me.cmbRemap.Enabled = (This.ViewModel.SelectedColumnName <> Empty)
End Sub
