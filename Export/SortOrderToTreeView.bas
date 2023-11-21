Attribute VB_Name = "SortOrderToTreeView"
'@Folder("MVVM.SortOrder.ViewModel")
Option Explicit

Private Const ORPHAN_LISTOBJECT_NAME As String = "(Orphaned)"
Private Const GREY_TEXT_COLOR As Long = 12632256 'RGB(192,192,192)
Private Const SUFFIX_CURRENTLY_ACTIVE  As String = " (active)"

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

Public Sub Load(ByVal ViewModel As SortOrderViewModel, ByVal TreeView As TreeView)
    TreeView.Nodes.Clear
    LoadWorkbookNode ViewModel, TreeView
    LoadListObjectNodes ViewModel, TreeView
    LoadSortOrderStateNodes ViewModel, TreeView
End Sub

Private Sub LoadWorkbookNode(ByVal ViewModel As SortOrderViewModel, ByVal TreeView As TreeView)
    Dim Node As Node
    Set Node = TreeView.Nodes.Add(Key:="ROOT", text:=ViewModel.Workbook.Name, Image:="FileSaveAsExcelXlsx")
    Node.Expanded = True
End Sub

Private Sub LoadListObjectNodes(ByVal ViewModel As SortOrderViewModel, ByVal TreeView As TreeView)
    Dim ListObjectNames As Collection
    Set ListObjectNames = New Collection
    
    Dim AllListObjects As Collection
    Set AllListObjects = GetAllListObjects(ViewModel.Workbook)
    
    Dim ListObjectName As String
    Dim SortOrderState  As SortOrderState
    For Each SortOrderState In ViewModel.SortOrderStates
        ListObjectName = SortOrderState.ListObjectName
        Debug.Print ListObjectName
        If Not ExistsInCollection(ListObjectNames, ListObjectName) Then
            If ExistsInCollection(AllListObjects, ListObjectName) Then
                ListObjectNames.Add ListObjectName
            Else
                If Not ExistsInCollection(ListObjectNames, ORPHAN_LISTOBJECT_NAME) Then
                    ListObjectNames.Add ORPHAN_LISTOBJECT_NAME
                End If
            End If
        End If
    Next SortOrderState
    
    Dim ParentNode As Node
    Set ParentNode = TreeView.Nodes.Item("ROOT")
    
    Dim Node As Node
    Dim ListObjectNameVariant As Variant
    For Each ListObjectNameVariant In ListObjectNames
        Debug.Print "Creating tree node " & ListObjectNameVariant
        Set Node = TreeView.Nodes.Add(Relative:=ParentNode, relationship:=tvwChild, Key:=ListObjectNameVariant, text:=ListObjectNameVariant, Image:="FileSaveAsExcelXlsx")
        Node.Expanded = True
    Next ListObjectNameVariant
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
       Set Node = TreeView.Nodes.Add(Relative:=ParentNode, relationship:=tvwChild, Key:=SortOrderState.ToBase64, text:=SortOrderState.GetCaption, Image:="SortDialog")
       
       If Not SortOrderState.CanApply(ViewModel.ListObject) Then
        Node.ForeColor = GREY_TEXT_COLOR
       End If
       
       If Not ViewModel.CurrentSortState Is Nothing Then
        If SortOrderState.Equals(ViewModel.CurrentSortState) Then
            Node.text = Node.text & SUFFIX_CURRENTLY_ACTIVE
            Node.Bold = True
            Node.Selected = True
            ' Make sure that selecting a sort order to preview will never update the treeview
            ' list of all sort orders, or it will start a recursive loop.
            ViewModel.TrySelect SortOrderState.ToBase64
        End If
       End If
    Next SortOrderState
End Sub
