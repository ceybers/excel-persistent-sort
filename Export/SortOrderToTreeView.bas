Attribute VB_Name = "SortOrderToTreeView"
'@Folder("MVVM.SortOrder.ViewModel")
Option Explicit

Public Sub InitializeTreeView(ByVal TreeView As TreeView)
    With TreeView
        .Nodes.Clear
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = tvwManual
        .LineStyle = tvwTreeLines
        .Style = tvwTreelinesPictureText
        Set .ImageList = ImageListHelpers.GetImageList
        .Indentation = 16
    End With
End Sub

