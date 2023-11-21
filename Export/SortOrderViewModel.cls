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

Private Const XML_SETTINGS_NAME As String = "PersistentSortOrder"
Private Const SETTING_COLLECTION_NAME As String = "SortOrderStates"
Private Const SETTING_LAST_UPDATED As String = "LastUpdated"
Private Const DO_ASSOCIATE_ON_APPLY As String = "DO_ASSOCIATE_ON_APPLY"
Private Const DO_PARTIAL_MATCH As String = "DO_PARTIAL_MATCH"
Private Const DO_PARTIAL_APPLY As String = "DO_PARTIAL_APPLY"
Private Const DO_CLOSE_ON_APPLY As String = "DO_CLOSE_ON_APPLY"

Private Type TState
    Workbook As Workbook
    ListObject As ListObject
    CurrentSortState As SortOrderState
    SelectedSortState As SortOrderState
    SortOrderStates As Collection
    Settings As ISettingsModel
End Type

Private This As TState

Public Property Get DoAssociateOnApply() As Boolean
    DoAssociateOnApply = This.Settings.Workbook.GetFlag(DO_ASSOCIATE_ON_APPLY)
End Property

Public Property Let DoAssociateOnApply(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_ASSOCIATE_ON_APPLY, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get DoPartialMatch() As Boolean
    DoPartialMatch = This.Settings.Workbook.GetFlag(DO_PARTIAL_MATCH)
End Property

Public Property Let DoPartialMatch(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_PARTIAL_MATCH, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get DoPartialApply() As Boolean
    DoPartialApply = This.Settings.Workbook.GetFlag(DO_PARTIAL_APPLY)
End Property

Public Property Let DoPartialApply(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_PARTIAL_APPLY, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get DoCloseOnApply() As Boolean
    DoCloseOnApply = This.Settings.Workbook.GetFlag(DO_CLOSE_ON_APPLY)
End Property

Public Property Let DoCloseOnApply(ByVal vNewValue As Boolean)
    This.Settings.Workbook.SetFlag DO_CLOSE_ON_APPLY, vNewValue
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Property

Public Property Get CurrentSortState() As SortOrderState
    Dim Result As SortOrderState
    Set Result = New SortOrderState
    Result.LoadFromListObject This.ListObject
    Set CurrentSortState = Result
End Property

Public Property Get SelectedSortState() As SortOrderState
    Set SelectedSortState = This.SelectedSortState
End Property

Public Property Get Workbook() As Workbook
    Set Workbook = This.Workbook
End Property

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Get SortOrderStates() As Collection
    Set SortOrderStates = This.SortOrderStates
End Property

Public Property Get IsValid() As Boolean
    IsValid = (Not This.ListObject Is Nothing)
End Property

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    Set This.Workbook = This.ListObject.Parent.Parent
    
    LoadSettings
    
    Reload
End Sub

Private Sub LoadSettings()
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
                           Workbook:=This.Workbook, _
                           RootNode:=XML_SETTINGS_NAME)
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
        .AddWorkbookSettings(WorkbookSettings)
    
    Set This.Settings = ASettingsModel
    
    If This.Settings.Workbook.GetSetting(SETTING_LAST_UPDATED) = Empty Then
        ApplyDefaultSettings
    End If
End Sub

Private Sub ApplyDefaultSettings()
    With This.Settings.Workbook
        .SetFlag DO_ASSOCIATE_ON_APPLY, True
        .SetFlag DO_PARTIAL_MATCH, True
        .SetFlag DO_PARTIAL_APPLY, True
        .SetFlag DO_CLOSE_ON_APPLY, True
        .SetSetting SETTING_LAST_UPDATED, CStr(Now())
    End With
End Sub

Private Sub Reload()
    Dim SerializedSortOrders As Collection
    Set SerializedSortOrders = This.Settings.Workbook.GetCollection(SETTING_COLLECTION_NAME)

    CollectionHelpers.CollectionClear This.SortOrderStates
    
    Dim SerializedSortOrder As Variant
    For Each SerializedSortOrder In SerializedSortOrders
        Dim SortOrderState As SortOrderState
        Set SortOrderState = New SortOrderState
        SortOrderState.LoadFromString SerializedSortOrder, This.ListObject
        This.SortOrderStates.Add SortOrderState
    Next SerializedSortOrder
End Sub

Private Sub Class_Initialize()
    Set This.SortOrderStates = New Collection
End Sub

Public Function Apply() As Boolean
    If This.SelectedSortState Is Nothing Then Exit Function
    If Not This.SelectedSortState.CanApply(This.ListObject) Then Exit Function
    
    This.SelectedSortState.Apply This.ListObject
    
    If Me.DoAssociateOnApply And IsSortOrderStateOrphaned(This.SelectedSortState) Then
        AssociateSortOrderState This.SelectedSortState
    End If
    
    Apply = True
End Function

Public Sub RemoveSelected()
    If This.SelectedSortState Is Nothing Then Exit Sub
    
    Dim i As Long
    For i = This.SortOrderStates.Count To 1 Step -1
        Dim SortOrderState As SortOrderState
        Set SortOrderState = This.SortOrderStates.Item(i)
        If SortOrderState Is This.SelectedSortState Then
            This.SortOrderStates.Remove i
        End If
    Next i
    
    Dim SavedCollection As Collection
    Set SavedCollection = This.Settings.Workbook.GetCollection(SETTING_COLLECTION_NAME)
    For i = SavedCollection.Count To 1 Step -1
        If SavedCollection.Item(i) = This.SelectedSortState.ToBase64 Then
            SavedCollection.Remove i
        End If
    Next i
    
    This.Settings.Workbook.SetCollection SETTING_COLLECTION_NAME, SavedCollection
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())

    Set This.SelectedSortState = Nothing
End Sub

Public Sub RemoveAll()
    Set This.SelectedSortState = Nothing
    
    CollectionHelpers.CollectionClear This.SortOrderStates
    
    Dim EmptyCollection As Collection
    Set EmptyCollection = New Collection
    This.Settings.Workbook.SetCollection SETTING_COLLECTION_NAME, EmptyCollection
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Sub

Public Function CanSave() As Boolean
    Dim RHSSortOrderState As SortOrderState
    For Each RHSSortOrderState In This.SortOrderStates
        If CurrentSortState.Equals(RHSSortOrderState) Then
            Exit Function
        End If
    Next RHSSortOrderState
    CanSave = True
End Function

Public Function TrySelect(ByVal Key As String) As Boolean
    If Key = "UNSAVED" Then
        Set This.SelectedSortState = CurrentSortState
        TrySelect = True
        Exit Function
    End If
    
    Dim SortOrderState As SortOrderState
    For Each SortOrderState In This.SortOrderStates
        If SortOrderState.ToBase64 = Key Then
            Set This.SelectedSortState = SortOrderState
            TrySelect = True
            Exit Function
        End If
    Next SortOrderState
    
    Set This.SelectedSortState = Nothing
End Function

Public Sub Save()
    Debug.Assert Not CurrentSortState Is Nothing
    
    Dim SavedCollection As Collection
    Set SavedCollection = This.Settings.Workbook.GetCollection(SETTING_COLLECTION_NAME)
    
    SavedCollection.Add Item:=CurrentSortState.ToBase64
    
    This.Settings.Workbook.SetCollection SETTING_COLLECTION_NAME, SavedCollection
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
    
    Reload
End Sub

Private Function IsSortOrderStateOrphaned(ByVal SortOrderState As SortOrderState)
    Dim AllListObjects As Collection
    Set AllListObjects = ListObjectHelpers.GetAllListObjects(This.Workbook)
    
    Dim ListObject As ListObject
    For Each ListObject In AllListObjects
        If ListObject.Name = SortOrderState.ListObjectName Then
            Exit Function
        End If
    Next ListObject
    
    IsSortOrderStateOrphaned = True
End Function

Private Sub AssociateSortOrderState(ByVal SortOrderState As SortOrderState)
    Dim OldKey As String
    OldKey = SortOrderState.ToBase64
    
    SortOrderState.WorksheetName = This.ListObject.Parent.Name
    SortOrderState.ListObjectName = This.ListObject.Name

    Dim SavedCollection As Collection
    Set SavedCollection = This.Settings.Workbook.GetCollection(SETTING_COLLECTION_NAME)
    
    Dim NewCollection As Collection
    Set NewCollection = New Collection
    
    Dim SortOrderString As Variant
    For Each SortOrderString In SavedCollection
        If SortOrderString = OldKey Then
            NewCollection.Add SortOrderState.ToBase64
        Else
            NewCollection.Add SortOrderString
        End If
    Next SortOrderString
    
    This.Settings.Workbook.SetCollection SETTING_COLLECTION_NAME, NewCollection
    This.Settings.Workbook.SetSetting SETTING_LAST_UPDATED, CStr(Now())
End Sub
