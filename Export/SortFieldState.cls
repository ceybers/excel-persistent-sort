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
    SortOn As Long
    Order As Long
    CustomOrder As String
    Priority As Long
    Exists As Boolean
End Type

Private This As TState

Public Property Get ColumnName() As String
    ColumnName = This.ColumnName
End Property

Public Property Get SortOn() As Long
    SortOn = This.SortOn
End Property

Public Property Get Order() As Long
    Order = This.Order
End Property

Public Property Get CustomOrder() As String
    CustomOrder = This.CustomOrder
End Property

Public Property Get Priority() As Long
    Priority = This.Priority
End Property

Public Property Get Exists() As Boolean
    Exists = This.Exists
End Property

Public Property Let Exists(ByVal vNewValue As Boolean)
    This.Exists = vNewValue
End Property

Public Function ToString() As String
    ToString = StringtoBase64(This.ColumnName) & "," _
                                             & CStr(This.SortOn) & "," _
                                             & CStr(This.Order) & "," & _
                                               StringtoBase64(This.CustomOrder)
End Function

Public Sub FromString(ByVal SortFieldStateString As String, ByVal Priority As Long)
    Dim SplitString() As String
    SplitString = Split(SortFieldStateString, ",")
    If UBound(SplitString) <> 3 Then Exit Sub
    
    This.ColumnName = Base64toString(SplitString(0))
    This.SortOn = CLng(SplitString(1))
    This.Order = CLng(SplitString(2))
    This.CustomOrder = Base64toString(SplitString(3))
    This.Priority = Priority
End Sub

Public Sub Load(ByVal SortField As SortField)
    If SortField.Key.ListObject Is Nothing Then
        Exit Sub
    End If
    
    Dim HeaderName As String
    If Not TryRangeToListHeader(SortField.Key.ListObject, SortField.Key, HeaderName) Then
        Exit Sub
    End If
    
    With This
        .ColumnName = HeaderName
        .SortOn = SortField.SortOn
        .Order = SortField.Order
        .Priority = SortField.Priority
        .Exists = True
    End With
    
    Select Case This.SortOn
    Case xlSortOnCellColor
        This.CustomOrder = CStr(SortField.SortOnValue.Color)
    Case xlSortOnFontColor
        This.CustomOrder = CStr(SortField.SortOnValue.Color)
    Case xlSortOnIcon
        This.CustomOrder = SortField.SortOnValue.Parent.ID & "." & SortField.SortOnValue.Index
    Case xlSortOnValues
        If SortField.CustomOrder <> CLng(0) Then
            This.CustomOrder = SortField.CustomOrder
        End If
    End Select
End Sub

Public Sub Add(ByVal ListObject As ListObject)
    Dim SortFields As SortFields
    Set SortFields = ListObject.sort.SortFields
    
    If Not ListObjectHelpers.HasListColumn(ListObject, This.ColumnName) Then
        Exit Sub
    End If
    
    Dim KeyRange As Range
    Set KeyRange = ListObject.ListColumns.Item(This.ColumnName).DataBodyRange
    
    Select Case This.SortOn
    Case xlSortOnCellColor
        SortFields.Add(KeyRange, xlSortOnCellColor, This.Order, , xlSortNormal).SortOnValue.Color = CLng(This.CustomOrder)
    Case xlSortOnFontColor
        SortFields.Add(KeyRange, xlSortOnFontColor, This.Order, , xlSortNormal).SortOnValue.Color = CLng(This.CustomOrder)
    Case xlSortOnIcon
        Dim IconSetsIndex As Long
        IconSetsIndex = CLng(Split(This.CustomOrder, ".")(0))
        Dim ItemID As Long
        ItemID = CLng(Split(This.CustomOrder, ".")(1))
        SortFields.Add(KeyRange, xlSortOnIcon, This.Order, , xlSortNormal).SetIcon Icon:=ActiveWorkbook.IconSets.Item(IconSetsIndex).Item(ItemID)
    Case xlSortOnValues
        If This.CustomOrder = vbNullString Then
            SortFields.Add KeyRange, This.SortOn, This.Order
        Else
            SortFields.Add KeyRange, This.SortOn, This.Order, This.CustomOrder
        End If
    End Select
End Sub


