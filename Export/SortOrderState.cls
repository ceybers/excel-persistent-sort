VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortOrderState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "MVVM.Model"
Option Explicit

Private Const TRIM_FIELD_NAME_LENGTH As Long = 16

Private Type TState
    WorksheetName As String
    ListObjectName As String
    SortFields As Collection
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
    
    Dim i As Long
    For i = 1 To ListObject.sort.SortFields.Count
        Dim SortField As SortField
        Set SortField = ListObject.sort.SortFields.Item(i)
        AddSortField ListObject, SortField
    Next i
End Sub

Public Function TryLoadFromString(ByVal SerializedString As String, ByVal ListObject As ListObject) As Boolean
    Dim SplitText() As String
    SplitText = Split(SerializedString, ":")
    
    If UBound(SplitText) <> 2 Then Exit Function
    
    This.WorksheetName = SplitText(0)
    This.ListObjectName = SplitText(1)
    
    Dim SplitFields() As String
    SplitFields = Split(SplitText(2), ";")
    
    Dim i As Long
    For i = LBound(SplitFields) To UBound(SplitFields)
        AddSortFieldSerialized SplitFields(i), (i + 1), ListObject
    Next i
    
    TryLoadFromString = True
End Function

Public Function CanApply(ByVal ListObject As ListObject) As Boolean
    Dim SortFieldState As SortFieldState
    
    For Each SortFieldState In This.SortFields
        If ListObjectHelpers.HasListColumn(ListObject, SortFieldState.ColumnName) Then
            CanApply = True
            Exit Function
        End If
    Next SortFieldState
End Function

Public Function IsPartialMatch(ByVal ListObject As ListObject) As Boolean
    Dim SortFieldState As SortFieldState
    
    Dim n As Long
    
    For Each SortFieldState In This.SortFields
        If ListObjectHelpers.HasListColumn(ListObject, SortFieldState.ColumnName) Then
            n = n + 1
        End If
    Next SortFieldState
    
    IsPartialMatch = (n > 0) And (n < This.SortFields.Count)
End Function

Public Function Apply(ByVal ListObject As ListObject) As Boolean
    With ListObject.sort
        .SortFields.Clear
        
        Dim SortFieldState As SortFieldState
        For Each SortFieldState In This.SortFields
            SortFieldState.Add ListObject
        Next SortFieldState
        
        .Header = xlYes
        .MatchCase = False
        .Apply
    End With
    
    Apply = True
End Function

Private Sub AddSortField(ByVal ListObject As ListObject, ByVal SortField As SortField)
    Dim SortFieldState As SortFieldState
    Set SortFieldState = New SortFieldState
    
    Dim HeaderName As String
    TryRangeToListHeader ListObject, SortField.Key, HeaderName
    
    SortFieldState.Load SortField
    
    This.SortFields.Add Item:=SortFieldState
End Sub

Private Sub AddSortFieldSerialized(ByVal SerializedSortField As String, ByVal Priority As Long, ByVal ListObject As ListObject)
    Dim SortFieldState As SortFieldState
    Set SortFieldState = New SortFieldState

    SortFieldState.FromString SerializedSortField, Priority
    
    SortFieldState.Exists = ListObjectHelpers.HasListColumn(ListObject, SortFieldState.ColumnName)
    
    This.SortFields.Add Item:=SortFieldState
End Sub

Public Function ToBase64() As String
    Dim Result As String
    Result = Result & This.WorksheetName & ":"
    Result = Result & This.ListObjectName & ":"
    
    Dim SortFieldState As SortFieldState
    For Each SortFieldState In This.SortFields
        Result = Result & SortFieldState.ToString() & ";"
    Next SortFieldState
    
    Result = Left$(Result, Len(Result) - 1)
    
    ToBase64 = Result
End Function

Public Function GetCaption() As String
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
        Result = CAPTION_NO_SORT_ORDER
    End If
    
    GetCaption = Result
End Function

Public Function Equals(ByVal RHS As SortOrderState) As Boolean
    Equals = (Me.ToBase64 = RHS.ToBase64)
End Function

