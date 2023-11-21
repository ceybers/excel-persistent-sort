Attribute VB_Name = "modMain"
'@Folder "SortOrderState"
Option Explicit
Public Sub aaa()
    Dim sort As sort
    Set sort = Selection.ListObject.sort
    
    Dim sortfield As sortfield
    Set sortfield = sort.sortfields.Item(1)
    
    Stop
End Sub
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

