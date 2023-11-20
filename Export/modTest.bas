Attribute VB_Name = "modTest"
'@Folder("VBAProject")
Option Explicit

Public Sub BBB()
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
      Workbook:=ThisWorkbook, _
      RootNode:="PersistentSortOrder")
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
      .AddWorkbookSettings(WorkbookSettings)
      
    WorkbookSettings.Load
    
    ASettingsModel.Table("Table1").SetSetting "TestSetting", "TestSetting works OK"
    Debug.Print ASettingsModel.Table("Table1").GetSetting("TestSetting")
    
    Dim SortOrderStates As Collection
    Set SortOrderStates = ASettingsModel.Table("Table1").GetCollection("SortOrderStates")
    'SortOrderStates.Add Key:="Key1", Item:="hello"
    ASettingsModel.Table("Table1").SetCollection "SortOrderStates", SortOrderStates
    'Debug.Print SortOrderStates.Item("Key1")
End Sub

Public Sub AAA()
    Dim lo As ListObject
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    
    Dim Payload As String
    'Payload = SerializeSortOrder(lo)
    Payload = "Q29sQg==,1;Q29sQw==,2"
    'Debug.Print Payload
    
    DeserializeSortOrder lo, Payload
End Sub
