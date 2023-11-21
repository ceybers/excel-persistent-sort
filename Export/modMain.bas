Attribute VB_Name = "modMain"
'@Folder "SortOrderState"
Option Explicit

'@EntryPoint "Open UI for PersistentSortOrderTool"
Public Sub PersistentSortOrderTool()
    If Selection.ListObject Is Nothing Then
        MsgBox "Select a table before running Persistent Sort Order Tool.", vbExclamation, "Persistent Sort Order Tool"
        Exit Sub
    End If
    
    Dim ViewModel As SortOrderViewModel
    Set ViewModel = New SortOrderViewModel
    ViewModel.Load Selection.ListObject
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = New frmSortOrderView
    
    ViewAsInterface.ShowDialog ViewModel
End Sub

