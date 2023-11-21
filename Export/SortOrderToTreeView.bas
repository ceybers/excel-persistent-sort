Attribute VB_Name = "SortOrderToTreeView"
'@Folder("MVVM.SortOrder.ViewModel")
Option Explicit

Private Const ORPHAN_LISTOBJECT_NAME As String = "(Orphaned)"
Private Const GREY_TEXT_COLOR As Long = 12632256 'RGB(192,192,192)
Private Const SUFFIX_CURRENTLY_ACTIVE  As String = " (active)"
Private Const SUFFIX_SELECTED_LISTOBJECT As String = " (selected)"
Private Const UNSAVED_SORTORDER As String = "(current sort order)"
Private Const NO_STATES_FOUND As String = "No saved Sort Order States found."
Private Const MSO_WORKBOOK As String = "FileSaveAsExcelXlsx"
Private Const MSO_LISTOBJECT As String = "CreateTable"
Private Const MSO_LISTOBJECT_SELECTED As String = "GroupTableDrawBorders"
Private Const MSO_SORTORDER As String = "SortDialog"

Public Sub InitializeTreeView(ByVal TreeView As TreeView)
    With TreeView
        .Nodes.Clear
        .FullRowSelect = False
        .HideSelection = False
        .LabelEdit = tvwManual
        .LineStyle = tvwTreeLines
        .Style = tvwTreelinesPictureText
        Set .ImageList = ImageListHelpers.GetImageList
        .Indentation = 16
    End With
End Sub

Public Sub Load(ByVal ViewModel As SortOrderViewModel, ByVal TreeView As TreeView)
    TreeView.Nodes.Clear
    LoadWorkbookNode ViewModel, TreeView
    LoadListObjectNodes ViewModel, TreeView
    AddUnsavedSortStateNode ViewModel, TreeView
    LoadSortOrderStateNodes ViewModel, TreeView
    UpdateListObjectIcons TreeView
    CheckNoSortOrderStatesFound TreeView
    RemoveEmptyListObjectNodes TreeView
    TrySelectSelectedNode ViewModel, TreeView
End Sub

Private Sub LoadWorkbookNode(ByVal ViewModel As SortOrderViewModel, ByVal TreeView As TreeView)
    Dim Node As Node
    Set Node = TreeView.Nodes.Add(Key:="ROOT", text:=ViewModel.Workbook.Name, Image:=MSO_WORKBOOK)
    Node.Expanded = True
End Sub

Private Sub LoadListObjectNodes(ByVal ViewModel As SortOrderViewModel, ByVal TreeView As TreeView)
    Dim ListObjectNames As Collection
    Set ListObjectNames = New Collection
    ListObjectNames.Add ViewModel.ListObject.Name
    
    Dim HasOrphans As Boolean
    
    Dim AllListObjects As Collection
    Set AllListObjects = GetAllListObjects(ViewModel.Workbook)
    
    Dim ListObjectName As String
    Dim SortOrderState  As SortOrderState
    For Each SortOrderState In ViewModel.SortOrderStates
        ListObjectName = SortOrderState.ListObjectName
        If Not ExistsInCollection(ListObjectNames, ListObjectName) Then
            If ExistsInCollection(AllListObjects, ListObjectName) Then
                ListObjectNames.Add ListObjectName
            Else
                HasOrphans = True
            End If
        End If
    Next SortOrderState
    
    If HasOrphans Then
        ListObjectNames.Add ORPHAN_LISTOBJECT_NAME
    End If
    
    Dim ParentNode As Node
    Set ParentNode = TreeView.Nodes.Item("ROOT")
    
    Dim Node As Node
    Dim ListObjectNameVariant As Variant
    For Each ListObjectNameVariant In ListObjectNames
        Set Node = TreeView.Nodes.Add(Relative:=ParentNode, _
                                      Relationship:=tvwChild, _
                                      Key:=ListObjectNameVariant, _
                                      text:=ListObjectNameVariant, _
                                      Image:=MSO_LISTOBJECT)
        Node.Expanded = True
    Next ListObjectNameVariant
    
    If Node.text = ORPHAN_LISTOBJECT_NAME Then
        Node.ForeColor = GREY_TEXT_COLOR
    End If
End Sub

Private Sub AddUnsavedSortStateNode(ByVal ViewModel As SortOrderViewModel, ByVal TreeView As TreeView)
    If ViewModel.CurrentSortState.SortFields.Count > 0 Then
        Dim Node As Node
        Set Node = TreeView.Nodes.Add(Relative:=TreeView.Nodes.Item(2), _
                                      Relationship:=tvwChild, _
                                      Key:="UNSAVED", _
                                      text:=UNSAVED_SORTORDER, _
                                      Image:=MSO_SORTORDER)
        Node.Bold = True
        Node.Selected = True
        ViewModel.TrySelect "UNSAVED"
    End If
End Sub

Private Sub LoadSortOrderStateNodes(ByVal ViewModel As SortOrderViewModel, ByVal TreeView As TreeView)
    Dim AllListObjects As Collection
    Set AllListObjects = GetAllListObjects(ViewModel.Workbook)
    
    Dim SortOrderState  As SortOrderState
    For Each SortOrderState In ViewModel.SortOrderStates
        Dim ParentNode As Node
        If ExistsInCollection(AllListObjects, SortOrderState.ListObjectName) Then
            Set ParentNode = TreeView.Nodes.Item(SortOrderState.ListObjectName)
        Else
            Set ParentNode = TreeView.Nodes.Item(ORPHAN_LISTOBJECT_NAME)
        End If
       
        Dim Node As Node
        Set Node = TreeView.Nodes.Add(Relative:=ParentNode, _
                                      Relationship:=tvwChild, _
                                      Key:=SortOrderState.ToBase64, _
                                      text:=SortOrderState.GetCaption, _
                                      Image:=MSO_SORTORDER)
       If TreeView.SelectedItem Is Nothing Then
        Node.Selected = True
       End If
       
        If Not SortOrderState.CanApply(ViewModel.ListObject) Then
            Node.ForeColor = GREY_TEXT_COLOR
        Else
            If Not ViewModel.DoPartialApply Then
                If SortOrderState.IsPartialMatch(ViewModel.ListObject) Then
                    Node.Selected = False
                End If
            End If
        End If
       
        If Not ViewModel.CurrentSortState Is Nothing Then
            If SortOrderState.Equals(ViewModel.CurrentSortState) Then
                Node.text = Node.text & SUFFIX_CURRENTLY_ACTIVE
                Node.Bold = True
                Node.Selected = True
                ' Make sure that selecting a sort order to preview will never update the treeview
                ' list of all sort orders, or it will start a recursive loop.
                ViewModel.TrySelect SortOrderState.ToBase64
                
                TreeView.Nodes.Remove "UNSAVED"
            End If
        End If
        
        If Not ViewModel.DoPartialMatch Then
            If SortOrderState.IsPartialMatch(ViewModel.ListObject) Then
                TreeView.Nodes.Remove Node.Key
            End If
            If Not SortOrderState.CanApply(ViewModel.ListObject) Then
                TreeView.Nodes.Remove Node.Key
            End If
        End If
    Next SortOrderState
End Sub

Private Sub UpdateListObjectIcons(ByVal TreeView As TreeView)
    With TreeView.Nodes.Item(2)
        .text = .text & SUFFIX_SELECTED_LISTOBJECT
        .Image = MSO_LISTOBJECT_SELECTED
    End With
End Sub

Private Sub CheckNoSortOrderStatesFound(ByVal TreeView As TreeView)
    If TreeView.Nodes.Count > 2 Then Exit Sub
    
    ' Remove manually added node for target ListObject
    If TreeView.Nodes.Count = 2 Then
        TreeView.Nodes.Remove (2)
    End If
    
    Dim Node As Node
    Set Node = TreeView.Nodes.Add(Relative:=TreeView.Nodes.Item("ROOT"), _
                                  Relationship:=tvwChild, _
                                  text:=NO_STATES_FOUND)
    Node.ForeColor = GREY_TEXT_COLOR
End Sub


Private Sub RemoveEmptyListObjectNodes(ByVal TreeView As TreeView)
    Dim NodesToDelete As Collection
    Set NodesToDelete = New Collection
    
    Dim Node As Node
    For Each Node In TreeView.Nodes
        If Node.Image = MSO_LISTOBJECT Then
            If Node.Children = 0 Then
                NodesToDelete.Add Node.Key
            End If
        End If
    Next Node
    
    Dim NodeKey As Variant
    For Each NodeKey In NodesToDelete
        TreeView.Nodes.Remove NodeKey
    Next NodeKey
End Sub

Private Sub TrySelectSelectedNode(ByVal ViewModel As SortOrderViewModel, ByVal TreeView As TreeView)
    If Not TreeView.SelectedItem Is Nothing Then
        ViewModel.TrySelect TreeView.SelectedItem.Key
    End If
End Sub


