VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortOrderState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "SortOrderState"
Option Explicit

Private Const TRIM_FIELD_NAME_LENGTH As Long = 16
Private Const NO_SORT_ORDER As String = "(no sort order)"

Private Type TState
    WorksheetName As String
    ListObjectName As String
    sortfields As Collection
End Type

Private This As TState

Public Property Get HasSortOrder() As Boolean
    HasSortOrder = (This.sortfields.Count > 0)
End Property

Public Property Get WorksheetName() As String
    WorksheetName = This.WorksheetName
End Property

Public Property Let WorksheetName(ByVal vNewValue As String)
    This.WorksheetName = vNewValue
End Property

Public Property Get ListObjectName() As String
    ListObjectName = This.ListObjectName
End Property

Public Property Let ListObjectName(ByVal vNewValue As String)
    This.ListObjectName = vNewValue
End Property

Public Property Get sortfields() As Collection
    Set sortfields = This.sortfields
End Property

Private Sub Class_Initialize()
    Set This.sortfields = New Collection
End Sub

Public Sub LoadFromListObject(ByVal ListObject As ListObject)
    Debug.Assert Not ListObject Is Nothing
    This.WorksheetName = ListObject.Parent.Name
    This.ListObjectName = ListObject.Name
    
    Dim i As Long
    For i = 1 To ListObject.sort.sortfields.Count
        Dim sortfield As sortfield
        Set sortfield = ListObject.sort.sortfields.Item(i)
        AddSortField ListObject, sortfield, i
    Next i
End Sub

Public Sub LoadFromString(ByVal SerializedString As String, ByVal ListObject As ListObject)
    Dim SplitText() As String
    SplitText = Split(SerializedString, ":")
    
    This.WorksheetName = SplitText(0)
    This.ListObjectName = SplitText(1)
    
    Dim SplitFields() As String
    SplitFields = Split(SplitText(2), ";")
    
    Dim i As Long
    For i = LBound(SplitFields) To UBound(SplitFields)
        AddSortFieldSerialized SplitFields(i), i, ListObject
    Next i
End Sub

Public Function CanApply(ByVal ListObject As ListObject) As Boolean
    Dim SortFieldState As SortFieldState
    
    For Each SortFieldState In This.sortfields
        If ListObjectHelpers.HasListColumn(ListObject, SortFieldState.ColumnName) Then
            CanApply = True
            Exit Function
        End If
    Next SortFieldState
End Function

Public Function IsPartialMatch(ByVal ListObject As ListObject) As Boolean
    Dim SortFieldState As SortFieldState
    
    Dim n As Long
    
    For Each SortFieldState In This.sortfields
        If ListObjectHelpers.HasListColumn(ListObject, SortFieldState.ColumnName) Then
            n = n + 1
        End If
    Next SortFieldState
    
    IsPartialMatch = (n > 0) And (n < This.sortfields.Count)
End Function

Public Function Apply(ByVal ListObject As ListObject) As Boolean
    With ListObject.sort
        .sortfields.Clear
        
        Dim SortFieldState As SortFieldState
        For Each SortFieldState In This.sortfields
            ApplySortFieldState ListObject, SortFieldState
        Next SortFieldState
        
        .Header = xlYes
        .MatchCase = False
        .Apply
    End With
    
    Apply = True
End Function

Private Sub ApplySortFieldState(ByVal ListObject As ListObject, ByVal SortFieldState As SortFieldState)
    If Not ListObjectHelpers.HasListColumn(ListObject, SortFieldState.ColumnName) Then
        Exit Sub
    End If
    
    Dim KeyRange As Range
    Set KeyRange = ListObject.ListColumns.Item(SortFieldState.ColumnName).DataBodyRange
    
    ListObject.sort.sortfields.Add Key:=KeyRange, Order:=SortFieldState.SortOrder
End Sub

Private Sub AddSortField(ByVal ListObject As ListObject, ByVal sortfield As sortfield, ByVal Index As Long)
    If sortfield.SortOn <> xlSortOnValues Then Exit Sub
    Dim SortFieldState As SortFieldState
    Set SortFieldState = New SortFieldState
    
    Dim HeaderName As String
    TryRangeToListHeader ListObject, sortfield.Key, HeaderName
    
    With SortFieldState
        .ColumnName = HeaderName
        .SortOrder = sortfield.Order
        .Index = Index
    End With
    
    This.sortfields.Add Item:=SortFieldState
End Sub

Private Sub AddSortFieldSerialized(ByVal SerializedSortField As String, ByVal Index As Long, ByVal ListObject As ListObject)
    Dim SplitText() As String
    SplitText = Split(SerializedSortField, ",")
    
    Dim SortFieldState As SortFieldState
    Set SortFieldState = New SortFieldState

    With SortFieldState
        .ColumnName = Base64toString(SplitText(0))
        .SortOrder = CLng(SplitText(1))
        .Index = (Index + 1)
        .Exists = ListObjectHelpers.HasListColumn(ListObject, .ColumnName)
    End With
    
    This.sortfields.Add Item:=SortFieldState
End Sub

Public Function ToBase64() As String
    Dim Result As String
    Result = Result & This.WorksheetName & ":"
    Result = Result & This.ListObjectName & ":"
    
    Dim SortFieldState As SortFieldState
    For Each SortFieldState In This.sortfields
        Result = Result & StringtoBase64(SortFieldState.ColumnName) & "," & CStr(SortFieldState.SortOrder) & ";"
    Next SortFieldState
    
    Result = Left$(Result, Len(Result) - 1)
    
    ToBase64 = Result
End Function

Public Function GetCaption() As String
    Dim Result As String
    
    Dim SortFieldState As SortFieldState
    For Each SortFieldState In This.sortfields
        Dim FieldName As String
        FieldName = SortFieldState.ColumnName
        If Len(FieldName) > TRIM_FIELD_NAME_LENGTH Then
            FieldName = Left$(FieldName, TRIM_FIELD_NAME_LENGTH - 3) & "..."
        End If
        Result = Result & FieldName & " > "
    Next SortFieldState
    
    If Len(Result) > 0 Then
        Result = Left$(Result, Len(Result) - 3)
    Else
        Result = NO_SORT_ORDER
    End If
    GetCaption = Result
End Function

Public Function Equals(ByVal RHS As SortOrderState) As Boolean
    Equals = (Me.ToBase64 = RHS.ToBase64)
End Function

