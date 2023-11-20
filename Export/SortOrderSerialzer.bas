Attribute VB_Name = "SortOrderSerialzer"
'@Folder("VBAProject")
Option Explicit

' Returns a string representing the sort order of a ListObject.
' Each column name is base64 encoded and paired with the sort order (0 asc, 1 desc)
' separated by a comma. The columns are separated by semicolons.
Public Function SerializeSortOrder(ByVal ListObject As ListObject) As String
    Dim Results As Collection
    Set Results = New Collection
    
    Dim SortField As SortField
    For Each SortField In ListObject.Sort.SortFields
        Results.Add Item:=SerializeSortField(SortField)
    Next SortField
    
    SerializeSortOrder = ConcatCollection(Results, ";")
End Function

Private Function SerializeSortField(ByVal SortField As SortField) As String
    If SortField.SortOn <> xlSortOnValues Then Exit Function
    
    Dim ListObject As ListObject
    Set ListObject = SortField.Parent.Parent.Parent
    
    Dim HeaderName As String
    If Not TryRangeToListHeader(ListObject, SortField.Key, HeaderName) Then
        Exit Function
    End If
            
    SerializeSortField = StringtoBase64(HeaderName) & "," & CStr(SortField.Order)
End Function

' Attempts to deserialize a serialized sort order string onto a list object.
Public Sub DeserializeSortOrder(ByVal ListObject As ListObject, ByVal SortOrders As String)
    Dim SortFieldStrings() As String
    SortFieldStrings = Split(SortOrders, ";")
    
    ListObject.Sort.SortFields.Clear
    
    Dim i As Long
    For i = 0 To UBound(SortFieldStrings)
        DeserializeSortField ListObject, SortFieldStrings(i)
    Next i
    
    ListObject.Sort.Apply
End Sub

Private Function DeserializeSortField(ByVal ListObject As ListObject, ByVal SortFieldString As String) As Boolean
    Dim SplitText() As String
    SplitText = Split(SortFieldString, ",")
    
    Dim HeaderName As String
    HeaderName = Base64toString(SplitText(0))
    
    If Not ListColumnExists(ListObject, HeaderName) Then
        Exit Function
    End If
    
    Dim KeyRange As Range
    Set KeyRange = ListObject.ListColumns(HeaderName).DataBodyRange
    
    Dim SortOrder As Long
    SortOrder = CInt(SplitText(1))
    
    ListObject.Sort.SortFields.Add Key:=KeyRange, Order:=SortOrder
    
    DeserializeSortField = True
End Function


