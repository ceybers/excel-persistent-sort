Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

Public Sub TestMVVM()
    ThisWorkbook.Worksheets(1).Activate
    ThisWorkbook.Worksheets(1).Range("A2").Activate
    
    Dim ViewModel As SortOrderViewModel
    Set ViewModel = New SortOrderViewModel
    
    Dim SelectedListObject As ListObject
    If Not Selection.ListObject Is Nothing Then
        ViewModel.Load Selection.ListObject
    Else
        Exit Sub
    End If
    
    Dim View As frmSortOrderView
    Set View = New frmSortOrderView
    
    Dim ViewAsInterface As IView
    Set ViewAsInterface = View
    
    If ViewAsInterface.ShowDialog(ViewModel) Then
        'Debug.Print "ShowDialog = TRUE"
    Else
        'Debug.Print "ShowDialog = FALSE"
    End If
    
End Sub

Public Sub TestA()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(2).ListObjects(1)
    
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
