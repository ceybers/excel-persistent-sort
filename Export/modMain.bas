Attribute VB_Name = "modMain"
'@Folder("VBAProject")
Option Explicit

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
