Attribute VB_Name = "modTESTDate"
'@Folder("SortOrderState")
Option Explicit

Private Const XML_SETTINGS_NAME As String = "PersistentSortOrder"

'@EntryPoint "Debug to seed test data"
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
    
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQg==,0,2,eHJheSx5YWNodCx6ZWJyYQ==;Q29sQQ==,0,1,"
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQQ==,0,1,;Q29sQg==,0,2,eHJheSx5YWNodCx6ZWJyYQ=="
    SortOrderStates.Add Item:="Sheet1:Table1:R2FtbWE=,0,1,;Q29sQg==,0,1,;Q29sQQ==,0,1,"
    SortOrderStates.Add Item:="Sheet1:Orphan:R2FtbWE=,0,1,;Q29sQg==,0,1,;Q29sQQ==,0,1,"
    SortOrderStates.Add Item:="Sheet1:Orphan:R2FtbWE=,0,1,"
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQQ==,0,1,"
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQw==,1,1,MTM1NTE2MTU=;Q29sQQ==,0,1,"
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQw==,3,1,MTYuMg==;Q29sQQ==,0,1,"
    
    ASettingsModel.Workbook.SetCollection "SortOrderStates", SortOrderStates
    
    Set ASettingsModel = Nothing
    Set WorkbookSettings = Nothing
    
    MsgBox "Reset sort order states to hard-coded test values.", vbInformation + vbOKOnly
End Sub


