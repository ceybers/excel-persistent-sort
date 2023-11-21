VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortFieldState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SortOrderState")
Option Explicit

Private Type TState
    ColumnName As String
    SortOrder As Long
    Index As Long
    Exists As Boolean
End Type

Private This As TState

Public Property Get ColumnName() As String
    ColumnName = This.ColumnName
End Property

Public Property Let ColumnName(ByVal vNewValue As String)
    This.ColumnName = vNewValue
End Property

Public Property Get SortOrder() As Long
    SortOrder = This.SortOrder
End Property

Public Property Let SortOrder(ByVal vNewValue As Long)
    This.SortOrder = vNewValue
End Property

Public Property Get Index() As Long
    Index = This.Index
End Property

Public Property Let Index(ByVal vNewValue As Long)
    This.Index = vNewValue
End Property

Public Property Get Exists() As Boolean
    Exists = This.Exists
End Property

Public Property Let Exists(ByVal vNewValue As Boolean)
    This.Exists = vNewValue
End Property

