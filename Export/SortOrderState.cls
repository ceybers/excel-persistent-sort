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
Private Const NO_SORT_ORDER = "(no sort order)"

Private Type TState
    WorksheetName As String
    ListObjectName As String
    SortFields As Collection
    'CreatedTime As Date
    'ModifiedTime as Date
    'LastUsedTime as Date
    'Comments as String ' base64 and multiline
End Type
Private This As TState

Public Property Get HasSortOrder() As Boolean
    HasSortOrder = (This.SortFields.Count > 0)
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

Public Property Get SortFields() As Collection
    Set SortFields = This.SortFields
End Property

Private Sub Class_Initialize()
    Set This.SortFields = New Collection
End Sub

Public Sub LoadFromListObject(ByVal ListObject As ListObject)
    Debug.Assert Not ListObject Is Nothing
    This.WorksheetName = ListObject.Parent.Name
    This.ListObjectName = ListObject.Name
    
    Dim SortField As SortField
    For Each SortField In ListObject.Sort.SortFields
        AddSortField ListObject, SortField
    Next SortField
End Sub

Public Sub LoadFromString(ByVal SerializedString As String)
    Dim SplitText() As String
    SplitText = Split(SerializedString, ":")
    
    This.WorksheetName = SplitText(0)
    This.ListObjectName = SplitText(1)
    
    Dim SplitFields() As String
    SplitFields = Split(SplitText(2), ";")
    
    Dim i As Long
    For i = LBound(SplitFields) To UBound(SplitFields)
        AddSortFieldSerialized SplitFields(i)
    Next i
End Sub

Public Function CanApply(ByVal ListObject As ListObject) As Boolean
    Dim SortFieldState As SortFieldState
    
    For Each SortFieldState In This.SortFields
        If ListColumnExists(ListObject, SortFieldState.ColumnName) Then
            CanApply = True
            Exit Function
        End If
    Next SortFieldState
End Function

Public Function Apply(ByVal ListObject As ListObject) As Boolean
    With ListObject.Sort
        .SortFields.Clear
        
        Dim SortFieldState As SortFieldState
        For Each SortFieldState In This.SortFields
            ApplySortFieldState ListObject, SortFieldState
        Next SortFieldState
        
        .Header = xlYes
        .MatchCase = False
        .Apply
    End With
End Function

Private Sub ApplySortFieldState(ByVal ListObject As ListObject, ByVal SortFieldState As SortFieldState)
    If Not ListColumnExists(ListObject, SortFieldState.ColumnName) Then
        Exit Sub
    End If
    
    Dim KeyRange As Range
    Set KeyRange = ListObject.ListColumns(SortFieldState.ColumnName).DataBodyRange
    
    ListObject.Sort.SortFields.Add Key:=KeyRange, Order:=SortFieldState.SortOrder
End Sub

Private Sub AddSortField(ByVal ListObject As ListObject, ByVal SortField As SortField)
    If SortField.SortOn <> xlSortOnValues Then Exit Sub
    Dim SortFieldState As SortFieldState
    Set SortFieldState = New SortFieldState
    
    Dim HeaderName As String
    TryRangeToListHeader ListObject, SortField.Key, HeaderName
    
    With SortFieldState
        .ColumnName = HeaderName
        .SortOrder = SortField.Order
    End With
    
    This.SortFields.Add Item:=SortFieldState
End Sub

Private Sub AddSortFieldSerialized(ByVal SerializedSortField As String)
    Dim SplitText() As String
    SplitText = Split(SerializedSortField, ",")
    
    Dim SortFieldState As SortFieldState
    Set SortFieldState = New SortFieldState

    With SortFieldState
        .ColumnName = Base64toString(SplitText(0))
        .SortOrder = CLng(SplitText(1))
    End With
    
    This.SortFields.Add Item:=SortFieldState
End Sub

Public Function ToBase64() As String
    Dim Result As String
    Result = Result & This.WorksheetName & ":"
    Result = Result & This.ListObjectName & ":"
    
    Dim SortFieldState As SortFieldState
    For Each SortFieldState In This.SortFields
        Result = Result & StringtoBase64(SortFieldState.ColumnName) & "," & CStr(SortFieldState.SortOrder) & ";"
    Next SortFieldState
    
    Result = Left$(Result, Len(Result) - 1)
    
    ToBase64 = Result
End Function

Public Function GetCaption()
    Dim Result As String
    
    Dim SortFieldState As SortFieldState
    For Each SortFieldState In This.SortFields
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
