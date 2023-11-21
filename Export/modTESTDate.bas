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
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQg==,1;Q29sQw==,2"
    SortOrderStates.Add Item:="Sheet2:Table2:VmVyeUxvbmdDb2x1bW5OYW1l,1;TG9uZ0NvbHVtbk5hbWU=,2"
    SortOrderStates.Add Item:="Sheet2:OrphanTable:R2FtbWE=,2"
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQw==,2;Q29sQg==,1"
    SortOrderStates.Add Item:="Sheet1:Table1:Q29sQw==,2;Q29sQg==,1;TG9uZ0NvbHVtbk5hbWU=,2"
    
    ASettingsModel.Workbook.SetCollection "SortOrderStates", SortOrderStates
    
    Set ASettingsModel = Nothing
    Set WorkbookSettings = Nothing
    
    MsgBox "Reset sort order states to hard-coded test values.", vbInformation + vbOKOnly
End Sub


