Attribute VB_Name = "Main"
'@Folder "PersistentSortOrder"
Option Explicit

'@EntryPoint "Open UI for PersistentSortOrderTool"
Public Sub RunPersistentSortOrderTool()
    On Error GoTo ErrorHandler
    
    If Selection.ListObject Is Nothing Then
        MsgBox MSG_SELECT_TABLE_FIRST, vbExclamation + vbOKOnly, APP_TITLE
        Exit Sub
    End If
    
    Dim ViewModel As SortOrderViewModel
    Set ViewModel = New SortOrderViewModel
    ViewModel.Load Selection.ListObject
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = New frmSortOrderView
    
    ViewAsInterface.ShowDialog ViewModel
    Exit Sub
    
ErrorHandler:
    MsgBox ERR_UNKNOWN_ERROR, vbExclamation + vbOKOnly, APP_TITLE
End Sub

