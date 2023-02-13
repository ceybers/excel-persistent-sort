Attribute VB_Name = "PersistentSortOrder"
'@Folder "SortOrder"
' Persistent Sort Order tool
' Craig Eybers
' Mon 13 February 2023

Option Explicit

Private Const SAVENAME As String = "caeSortOrder"
Private Const NO_TABLE_SELECTED As String = "No table selected. Cannot save or restore Sort Order!"
Private Const SAVED_SORT_ORDER As String = "Sort Order saved successfully."
Private Const RESTORED_SORT_ORDER As String = "Sort Order restored successfully."
Private Const CANNOT_RESTORE As String = "Cannot restore - No Sort Order was saved!"
Private Const NO_CUSTOM_SORT_ORDERS As String = "Custom Sort Orders are not supported!"
Private Const TABLE_MISMATCH As String = "Cannot restore Sort Order - table does not match!"
Private Const RESTORE_CAPTION As String = "Restore Table Sort Order"
Private Const SAVE_CAPTION As String = "Save Table Sort Order"

Public Sub SaveSortOrderSelected()
    Dim lo As ListObject
    Set lo = Selection.ListObject

    If lo Is Nothing Then
        MsgBox NO_TABLE_SELECTED, vbCritical + vbOKOnly, SAVE_CAPTION
        Exit Sub
    End If

    SaveSortOrder lo
End Sub

Public Sub RestoreSortOrderSelected()
    Dim lo As ListObject
    Set lo = Selection.ListObject

    If lo Is Nothing Then
        MsgBox NO_TABLE_SELECTED, vbCritical + vbOKOnly, RESTORE_CAPTION
        Exit Sub
    End If

    RestoreSortOrder lo
End Sub

Private Sub SaveSortOrder(ByVal lo As ListObject)
    Dim sortString As String
    sortString = SortToString(lo)

    Dim ws As Worksheet
    Set ws = lo.Parent

    Dim n As Name
    For Each n In ws.Parent.Names
        If n.Name Like ("*" & SAVENAME) Then GoTo FoundSaveName
    Next n

    Set n = ws.Names.Add(Name:=SAVENAME, RefersTo:=lo.Range)

FoundSaveName:
    n.Comment = sortString
    n.Visible = False

    MsgBox SAVED_SORT_ORDER, vbInformation + vbOKOnly, SAVE_CAPTION
End Sub

Private Sub RestoreSortOrder(ByVal lo As ListObject)
    Dim ws As Worksheet
    Set ws = lo.Parent

    Dim n As Name
    For Each n In ws.Parent.Names
        If n.Name Like ("*" & SAVENAME) Then GoTo FoundSaveName
    Next n

    MsgBox CANNOT_RESTORE, vbCritical + vbOKOnly, RESTORE_CAPTION
    Exit Sub

FoundSaveName:
    Dim sortString As String
    sortString = n.Comment
    StringToSort sortString, lo
End Sub

Private Function SortToString(ByVal lo As ListObject) As String
    Debug.Assert Not lo Is Nothing
    
    Dim s As Sort
    Set s = lo.Sort

    Dim result As String
    result = "ORDER " & s.Rng.Address
    If s.SortFields.Count = 0 Then GoTo NoSortFields
    result = result & " BY "

    Dim sf As SortField
    Dim i As Double
    For i = 1 To s.SortFields.Count
        Set sf = s.SortFields.Item(i)
        result = result & sf.Key.Address & " "
        If sf.Order = xlAscending Then
            result = result & "ASC"
        ElseIf sf.Order = xlDescending Then
            result = result & "DESC"
        Else
            MsgBox NO_CUSTOM_SORT_ORDERS, vbExclamation + vbOKOnly, SAVE_CAPTION
        End If
        result = result & ", "
    Next i

    result = Left$(result, Len(result) - 2)

NoSortFields:
    result = result & ";"
    SortToString = result
End Function

Private Sub StringToSort(ByVal sortString As String, ByVal lo As ListObject)
    Debug.Assert Left$(sortString, 6) = "ORDER "
    Debug.Assert Right$(sortString, 1) = ";"
    
    Dim ws As Worksheet
    Set ws = lo.Parent
    
    Dim tokens As Variant
    tokens = Split(sortString, " ")

    If lo.Range.EntireColumn.Address <> ws.Range(tokens(1)).EntireColumn.Address Then
        MsgBox TABLE_MISMATCH, vbCritical + vbOKOnly, RESTORE_CAPTION
        Exit Sub
    End If

    lo.Sort.SortFields.Clear

    Dim i As Double
    For i = 3 To UBound(tokens) Step 2
        Dim sortOrder As Double
        If tokens(i + 1) Like "ASC?" Then
            sortOrder = xlAscending
        ElseIf tokens(i + 1) Like "DESC?" Then
            sortOrder = xlDescending
        Else
            MsgBox NO_CUSTOM_SORT_ORDERS, vbExclamation + vbOKOnly, SAVE_CAPTION
            GoTo NextSortField
        End If

        Dim chkRange As Range
        Set chkRange = Application.Intersect(lo.Range, ws.Range(tokens(i)).EntireColumn)
        If Not chkRange Is Nothing Then
            lo.Sort.SortFields.Add Key:=chkRange, Order:=sortOrder
        Else
            Debug.Print "Sort column missing in table! Range = '" & ws.Range(tokens(i)).Address & "'"
        End If
NextSortField:
    Next i

    lo.Sort.Apply
    MsgBox RESTORED_SORT_ORDER, vbInformation + vbOKOnly, RESTORE_CAPTION
End Sub

