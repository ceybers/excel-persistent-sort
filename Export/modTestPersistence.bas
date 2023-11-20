Attribute VB_Name = "modTestPersistence"
'@Folder("VBAProject")
Option Explicit

Private Const XML_SETTINGS_NAME As String = "PersistentSortOrder"

Public Function GetSavedSortOrders() As Collection
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
      Workbook:=ThisWorkbook, _
      RootNode:=XML_SETTINGS_NAME)
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
      .AddWorkbookSettings(WorkbookSettings)
      
    WorkbookSettings.Load

    Set GetSavedSortOrders = ASettingsModel.Workbook.GetCollection("SortOrderStates")
    'MsgBox "Found " & GetSavedSortOrders.Count & " sort order state(s).", vbInformation + vbOKOnly
    
    Set ASettingsModel = Nothing
    Set WorkbookSettings = Nothing
End Function

Public Sub SetSavedSortOrders()
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
      Workbook:=ThisWorkbook, _
      RootNode:=XML_SETTINGS_NAME)
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
      .AddWorkbookSettings(WorkbookSettings)
      
    WorkbookSettings.Load
    ASettingsModel.Workbook.Reset

    Dim SortOrderStates As Collection
    Set SortOrderStates = New Collection
    SortOrderStates.Add Item:="Sheet2:Table2:R2FtbWE=,2"
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQg==,1;Q29sQw==,2"
    SortOrderStates.Add Item:="Sheet2:Table2:VmVyeUxvbmdDb2x1bW5OYW1l,1;TG9uZ0NvbHVtbk5hbWU=,2"
    
    ASettingsModel.Workbook.SetCollection "SortOrderStates", SortOrderStates
    
    Set ASettingsModel = Nothing
    Set WorkbookSettings = Nothing
    
    MsgBox "Reset sort order states to hard-coded test values.", vbInformation + vbOKOnly
End Sub

Public Sub RemoveSavedSortOrders(ByVal Index As Long)
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
      Workbook:=ThisWorkbook, _
      RootNode:=XML_SETTINGS_NAME)
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
      .AddWorkbookSettings(WorkbookSettings)
    
    Dim SortOrderStates As Collection
    Set SortOrderStates = ASettingsModel.Workbook.GetCollection("SortOrderStates")
    
    Dim NewCollection As Collection
    Set NewCollection = New Collection
    
    Dim i As Long
    For i = 1 To SortOrderStates.Count
        If i <> Index Then
            NewCollection.Add Item:=SortOrderStates.Item(i)
        End If
    Next i
    
    ASettingsModel.Workbook.SetCollection "SortOrderStates", NewCollection
    
    Set ASettingsModel = Nothing
    Set WorkbookSettings = Nothing
End Sub

Public Sub RemoveAllSavedSortOrders()
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
      Workbook:=ThisWorkbook, _
      RootNode:=XML_SETTINGS_NAME)
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
      .AddWorkbookSettings(WorkbookSettings)
      
    WorkbookSettings.Load
    ASettingsModel.Workbook.Reset
    
    'Dim EmptyCollection As Collection
    'Set EmptyCollection = New Collection
    
    'ASettingsModel.Workbook.SetCollection "SortOrderStates", EmptyCollection
    
    Set ASettingsModel = Nothing
    Set WorkbookSettings = Nothing
End Sub
