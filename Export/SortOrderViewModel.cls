VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortOrderViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.TableSplit.ViewModel"
Option Explicit

Event PropertyChanged(ByVal PropertyName As String)

Private Type TState
    'AvailableTables As AvailableTables
    'AvailableColumns As AvailableColumns
    'TargetSheets As TargetSheets
    
    ShowHiddenColumns As Boolean
    ShowUnsuitableColumns As Boolean
    DeleteExistingSheets As Boolean
    RemoveOtherSheets As Boolean
    
    'Settings As ISettings
    'MRU As IMostRecentlyUsed
End Type
Private This As TState

Public Property Get IsValid() As Boolean
    'If This.AvailableTables.Selected Is Nothing Then Exit Property
    'If This.AvailableColumns.Selected Is Nothing Then Exit Property
    'If CanSelectNone = False Then Exit Property
    
    IsValid = True
End Property

Public Sub Load(ByVal Workbook As Workbook)
    'noop
End Sub
