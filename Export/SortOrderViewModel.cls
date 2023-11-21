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
Event PropertyChanged(ByVal PropertyName As String)

Private Type TState
    Workbook As Workbook
    ListObject As ListObject
    CurrentSortState As SortOrderState
    SelectedSortState As SortOrderState
    SortOrderStates As Collection
    
    Settings As ISettingsModel
End Type
Private This As TState

Public Property Get CurrentSortState() As SortOrderState
    Set CurrentSortState = This.CurrentSortState
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

Public Property Set ListObject(ByVal vNewValue As ListObject)
    Set This.ListObject = vNewValue
End Property

Public Property Get SortOrderStates() As Collection
    Set SortOrderStates = This.SortOrderStates
End Property

Public Property Get IsValid() As Boolean
    'If This.AvailableTables.Selected Is Nothing Then Exit Property
    'If This.AvailableColumns.Selected Is Nothing Then Exit Property
    'If CanSelectNone = False Then Exit Property
    
    IsValid = True
End Property

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    
    Set This.CurrentSortState = New SortOrderState
    This.CurrentSortState.LoadFromListObject This.ListObject
    
    LoadSettings
    Reload
End Sub

Private Sub LoadSettings()
    Set This.Workbook = This.ListObject.Parent.Parent
    
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
      Workbook:=This.Workbook, _
      RootNode:=XML_SETTINGS_NAME)
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
      .AddWorkbookSettings(WorkbookSettings)
    
    Set This.Settings = ASettingsModel
    
    'WorkbookSettings.Load
End Sub

Private Sub Reload()
    Dim SerializedSortOrders As Collection
    Set SerializedSortOrders = This.Settings.Workbook.GetCollection(SETTING_COLLECTION_NAME)
    
    ' TODO Should clear it instead
    Set This.SortOrderStates = New Collection
    
    Dim SerializedSortOrder As Variant
    For Each SerializedSortOrder In SerializedSortOrders
        Dim SortOrderState As SortOrderState
        Set SortOrderState = New SortOrderState
        SortOrderState.LoadFromString SerializedSortOrder
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
    Apply = True
End Function

Public Sub ZZZ_ApplySortOrderState(ByVal Index As Long)
    Dim SortOrderState As SortOrderState
    Set SortOrderState = This.SortOrderStates.Item(Index)
    SortOrderState.Apply This.ListObject
End Sub

Public Sub RemoveByIndex(ByVal Index As Long)
    'MsgBox "NYI RemoveByIndex"
    RemoveSavedSortOrders Index
    Reload
End Sub

Public Sub RemoveAll()
    RemoveAllSavedSortOrders
    Reload
End Sub

Public Function LoadToTreeView(ByVal TreeView As TreeView)
    TreeView.Nodes.Clear
    
    Dim Node As Node
    Set Node = TreeView.Nodes.Add(text:=This.Workbook.Name, Image:="FileSaveAsExcelXlsx")
    Node.Expanded = True
    
    Dim Node2 As Node
    Set Node2 = TreeView.Nodes.Add(Relative:=Node, relationship:=tvwChild, text:="Table1", Image:="CreateTable")
    Node2.Expanded = True
    
    Dim Node3 As Node
    Set Node3 = TreeView.Nodes.Add(Relative:=Node2, relationship:=tvwChild, text:="Col A > Col B", Image:="SortDialog")
    Set Node3 = TreeView.Nodes.Add(Relative:=Node2, relationship:=tvwChild, text:="Col B > Col A", Image:="SortDialog")
    
    Set Node2 = TreeView.Nodes.Add(Relative:=Node, relationship:=tvwChild, text:="Table2", Image:="CreateTable")
    Node2.Expanded = True
    Set Node3 = TreeView.Nodes.Add(Relative:=Node2, relationship:=tvwChild, text:="foo > bar", Image:="SortDialog")
    
    Set Node2 = TreeView.Nodes.Add(Relative:=Node, relationship:=tvwChild, text:="(Orphaned)", Image:="CreateTable")
    Node2.Expanded = True
    Set Node3 = TreeView.Nodes.Add(Relative:=Node2, relationship:=tvwChild, text:="apple > banana", Image:="SortDialog")
    Set Node3 = TreeView.Nodes.Add(Relative:=Node2, relationship:=tvwChild, text:="zebra > xray", Image:="SortDialog")
    
    'Node2.Expanded = True
    
End Function

Public Function CanSave() As Boolean
    Dim RHSSortOrderState As SortOrderState
    For Each RHSSortOrderState In This.SortOrderStates
        If This.CurrentSortState.Equals(RHSSortOrderState) Then
            Exit Function
        End If
    Next RHSSortOrderState
    CanSave = True
End Function

Public Function TrySelect(ByVal Key As String) As Boolean
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
