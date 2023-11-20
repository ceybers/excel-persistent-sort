VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortOrderViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.SortOrder.ViewModel"
Option Explicit

Event PropertyChanged(ByVal PropertyName As String)

Private Type TState
    ListObject As ListObject
    SortOrderStates As Collection
End Type
Private This As TState

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Set ListObject(ByVal vNewValue As ListObject)
    Set This.ListObject = vNewValue
End Property

Public Property Get SortOrderStates() As Collection
    Set SortOrderStates = This.SortOrderStates
End Property

Public Property Get IsValid() As Boolean
    'If This.AvailableTables.Selected Is Nothing Then Exit Property
    'If This.AvailableColumns.Selected Is Nothing Then Exit Property
    'If CanSelectNone = False Then Exit Property
    
    IsValid = True
End Property

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    Reload
End Sub

Private Sub Reload()
    Dim SerializedSortOrders As Collection
    Set SerializedSortOrders = GetSavedSortOrders
    
    ' TODO Should clear it instead
    Set This.SortOrderStates = New Collection
    
    Dim SerializedSortOrder As Variant
    For Each SerializedSortOrder In SerializedSortOrders
        Dim SortOrderState As SortOrderState
        Set SortOrderState = New SortOrderState
        SortOrderState.LoadFromString SerializedSortOrder
        This.SortOrderStates.Add SortOrderState
    Next SerializedSortOrder
End Sub

Private Sub Class_Initialize()
    Set This.SortOrderStates = New Collection
End Sub

Public Sub ApplySortOrderState(ByVal Index As Long)
    Dim SortOrderState As SortOrderState
    Set SortOrderState = This.SortOrderStates.Item(Index)
    SortOrderState.Apply This.ListObject
End Sub

Public Sub RemoveByIndex(ByVal Index As Long)
    'MsgBox "NYI RemoveByIndex"
    RemoveSavedSortOrders Index
    Reload
End Sub

Public Sub RemoveAll()
    RemoveAllSavedSortOrders
    Reload
End Sub
