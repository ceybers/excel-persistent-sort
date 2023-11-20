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
