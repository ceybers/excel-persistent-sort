Attribute VB_Name = "CollectionHelpers"
'@Folder "Helpers"
Option Explicit

Public Function ConcatCollection(ByVal Collection As Collection, ByVal Delimiter As String) As String
    Dim Result As String
    
    Dim Item As Variant
    For Each Item In Collection
        Result = Result & Item & Delimiter
    Next Item
    
    If Len(Result) > 1 Then
        Result = Left$(Result, Len(Result) - 1)
    End If
    
    ConcatCollection = Result
End Function
