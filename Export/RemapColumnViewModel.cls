VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RemapColumnViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.SortOrder.ViewModel")
Option Explicit

Private Type TState
    CurrentColumnName As String
    SelectedColumnName As String
    ColumnNames As Collection
End Type

Private This As TState

Public Property Get IsValid() As Boolean
    If This.ColumnNames Is Nothing Then Exit Property
    If This.ColumnNames.Count = 0 Then Exit Property
    IsValid = True
End Property

Public Property Get CurrentColumnName() As String
    CurrentColumnName = This.CurrentColumnName
End Property

Public Property Let CurrentColumnName(ByVal vNewValue As String)
    This.CurrentColumnName = vNewValue
End Property

Public Property Get SelectedColumnName() As String
    SelectedColumnName = This.SelectedColumnName
End Property

Public Property Let SelectedColumnName(ByVal vNewValue As String)
    This.SelectedColumnName = vNewValue
End Property

Public Property Get ColumnNames() As Collection
    Set ColumnNames = This.ColumnNames
End Property

Public Sub Load(ByVal ColumnNames As Collection)
    Set This.ColumnNames = ColumnNames
End Sub

Private Sub Class_Initialize()
    Set This.ColumnNames = New Collection
End Sub

Public Function TrySelect(ByVal Key As String) As Boolean
    This.SelectedColumnName = Key
    TrySelect = True
End Function

