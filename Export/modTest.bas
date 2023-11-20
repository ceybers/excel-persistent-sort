Attribute VB_Name = "modTest"
'@Folder("VBAProject")
Option Explicit

Public Sub AAA()
    Dim WorkbookSettings As XMLSettings
    Set WorkbookSettings = XMLSettingsFactory.CreateWorkbookSettings( _
      Workbook:=ThisWorkbook, _
      RootNode:="TestPersistentStorage")
    
    Dim ASettingsModel As ISettingsModel
    Set ASettingsModel = SettingsModel.Create() _
      .AddWorkbookSettings(WorkbookSettings)
      
    WorkbookSettings.Load
    'ASettingsModel.Table("Table1").SetSetting "Hello", "World"
    Debug.Print ASettingsModel.Table("Table1").GetSetting("Hello")
End Sub
