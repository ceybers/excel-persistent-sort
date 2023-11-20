Attribute VB_Name = "modTestPersistence"
'@Folder("VBAProject")
Option Explicit

Public Function GetSavedSortOrders() As Collection
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
      Workbook:=ThisWorkbook, _
      RootNode:="PersistentSortOrder")
    
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
      RootNode:="PersistentSortOrder")
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
      .AddWorkbookSettings(WorkbookSettings)
      
    WorkbookSettings.Load
    ASettingsModel.Workbook.Reset

    Dim SortOrderStates As Collection
    Set SortOrderStates = New Collection
    SortOrderStates.Add Item:="Sheet2:Table2:R2FtbWE=,2"
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQg==,1;Q29sQw==,2"
 
    ASettingsModel.Workbook.SetCollection "SortOrderStates", SortOrderStates
    
    Set ASettingsModel = Nothing
    Set WorkbookSettings = Nothing
    
    MsgBox "Reset sort order states to hard-coded test values.", vbInformation + vbOKOnly
End Sub

