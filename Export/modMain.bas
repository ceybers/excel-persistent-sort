Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

Public Sub TestMVVM()
    ThisWorkbook.Worksheets(1).Activate
    ThisWorkbook.Worksheets(1).Range("A2").Activate
    If Selection.ListObject Is Nothing Then
        Exit Sub
    End If
    
    Dim ViewModel As SortOrderViewModel
    Set ViewModel = New SortOrderViewModel
    ViewModel.Load Selection.ListObject
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = New frmSortOrderView
    
    If ViewAsInterface.ShowDialog(ViewModel) Then
        'Debug.Print "ShowDialog = TRUE"
    Else
        'Debug.Print "ShowDialog = FALSE"
    End If
End Sub
