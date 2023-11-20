Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

Public Sub TestMVVM()
    Dim ViewModel As SortOrderViewModel
    Set ViewModel = New SortOrderViewModel
    ViewModel.Load ActiveWorkbook
    
    Dim View As frmSortOrderView
    Set View = New frmSortOrderView
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = View
    
    If ViewAsInterface.ShowDialog(ViewModel) Then
        Debug.Print "ShowDialog = TRUE"
    Else
        Debug.Print "ShowDialog = FALSE"
    End If
    
End Sub

Public Sub TestA()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim ASortOrderState As SortOrderState
    Set ASortOrderState = New SortOrderState
    ASortOrderState.LoadFromListObject lo
    
    Dim Base64String As String
    Base64String = ASortOrderState.ToBase64
    
    Stop
End Sub

Public Sub TestB()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim ASortOrderState As SortOrderState
    Set ASortOrderState = New SortOrderState
    ASortOrderState.LoadFromString "Sheet1:Table1:Q29sQg==,1;Q29sQw==,2"
   
    Stop
End Sub
