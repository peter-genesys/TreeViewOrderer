
'Option Strict On

Imports System.Windows.Forms

Imports System.Drawing

Imports System.ComponentModel

Public Class TreeViewDraggableNodes2Levels
    Inherits TreeView
#Region "Events"
    Public Event NodeMovedByDrag As EventHandler(Of NodeMovedByDragEventArgs)
    Protected Overridable Sub OnNodeMovedByDrag(ByVal e As NodeMovedByDragEventArgs)
        RaiseEvent NodeMovedByDrag(Me, e)
    End Sub

    Public Event NodeMovingByDrag As EventHandler(Of NodeMovingByDragEventArgs)
    Protected Overridable Sub OnNodeMovingByDrag(ByVal e As NodeMovingByDragEventArgs)
        RaiseEvent NodeMovingByDrag(Me, e)
    End Sub

    Public Event NodeDraggingOver As EventHandler(Of NodeDraggingOverEventArgs)
    Protected Overridable Sub OnNodeDraggingOver(ByVal e As NodeDraggingOverEventArgs)
        RaiseEvent NodeDraggingOver(Me, e)
    End Sub

#End Region

    Sub New()
        MyBase.AllowDrop = True
        MyBase.BackColor = Color.AliceBlue
        AddHandler MyBase.AfterCheck, AddressOf Me.AfterCheck
        'MyBase.DrawMode = TreeViewDrawMode.OwnerDrawAll
        'InitializeComponent()
        'MyBase.InitializeComponent()
    End Sub

    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public Shadows ReadOnly Property DrawMode() As System.Windows.Forms.TreeViewDrawMode
        Get
            Return MyBase.DrawMode
        End Get
    End Property

    '<DefaultValue(True)> _
    <Browsable(False), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)> _
    Public Shadows ReadOnly Property AllowDrop() As Boolean
        Get
            Return MyBase.AllowDrop
        End Get
    End Property

    Protected Overrides Sub OnItemDrag(ByVal e As System.Windows.Forms.ItemDragEventArgs)
        MyBase.DoDragDrop(e.Item, DragDropEffects.Move)

        MyBase.OnItemDrag(e)
    End Sub

    Protected Overrides Sub OnDragOver(ByVal drgevent As System.Windows.Forms.DragEventArgs)

        'Get the node we are currently over
        'while dragging another node
        Dim targetNode As TreeNode = MyBase.GetNodeAt(MyBase.PointToClient(New Point(drgevent.X, drgevent.Y)))

        'Get the node being dragged
        Dim dragNode As TreeNode = FindNodeInDataObject(drgevent.Data)

        Dim eaDraggingOver As New NodeDraggingOverEventArgs(dragNode, targetNode)
        OnNodeDraggingOver(eaDraggingOver)

        If eaDraggingOver.DropIsLegal = False Then
            drgevent.Effect = DragDropEffects.None
            Return
        End If


        'If we are not currently dragging over
        'a node...
        If targetNode Is Nothing Then

            'Let no node be selected
            MyBase.SelectedNode = Nothing

            'Allow the move because its valid
            'to drag a node over the TreeView itself
            'the drop will place the node being dragged
            'in the root
            drgevent.Effect = DragDropEffects.Move

            'Get out
            Return
        End If

        'This would only be nothing if something is being
        'dragged over the TreeView that isn't a node
        If dragNode IsNot Nothing Then

            'Illegal to drop nodes 
            'onto themselves
            'inside their descendants
            'level1 onto level2 (Special for this 2Level control)
            If targetNode Is dragNode _
                OrElse IsNodeDescendant(targetNode, dragNode) _
                OrElse (IsNodeLevel1(dragNode) And Not IsNodeLevel1(targetNode)) Then      '

                'Prevents a drop
                drgevent.Effect = DragDropEffects.None
            Else

                'Allows a drop
                drgevent.Effect = DragDropEffects.Move
                MyBase.SelectedNode = targetNode
            End If

        End If

        MyBase.OnDragOver(drgevent)
    End Sub

    Protected Overrides Sub OnDragDrop(ByVal drgevent As System.Windows.Forms.DragEventArgs)
        Dim dragNode As TreeNode = FindNodeInDataObject(drgevent.Data)

        If dragNode IsNot Nothing Then
            'Dim dragNode As TreeNode = DirectCast(drgevent.Data.GetData(GetType(TreeNode)), TreeNode)

            'Get the parent of the node before moving
            Dim prevParent As TreeNode = dragNode.Parent
            Dim parentToBe As TreeNode = If(MyBase.SelectedNode Is Nothing, Nothing, MyBase.SelectedNode)

            Dim eaNodeMoving As New NodeMovingByDragEventArgs(dragNode, prevParent, parentToBe)

            OnNodeMovingByDrag(eaNodeMoving)

            If eaNodeMoving.CancelMove = False Then

                If IsNodeLevel1(dragNode) Then
                    dragNode.Remove()
                    If MyBase.SelectedNode IsNot Nothing Then
                        MyBase.Nodes.Insert(MyBase.SelectedNode.Index, dragNode) 'Move Dragged Level1 node, to be sibling above Selected Level1 node
                    Else
                        MyBase.Nodes.Add(dragNode) 'Move Dragged Level1 node to last Level1 posn

                    End If
                    OnNodeMovedByDrag(New NodeMovedByDragEventArgs(dragNode, prevParent))
                    MyBase.SelectedNode = dragNode

                Else
                    If MyBase.SelectedNode Is Nothing Then
                        'Do not move the node must have a selected parent or sibling

                    Else
                        'A node has been selected
                        dragNode.Remove()
                        If IsNodeLevel1(MyBase.SelectedNode) Then
                            'Dragging level2 to a new parent level1 node
                            MyBase.SelectedNode.Nodes.Add(dragNode)
                        Else
                            'Dragging level2 to be sibling above Selected Level2 node  
                            MyBase.SelectedNode.Parent.Nodes.Insert(MyBase.SelectedNode.Index, dragNode)
                        End If
                        OnNodeMovedByDrag(New NodeMovedByDragEventArgs(dragNode, prevParent))
                        MyBase.SelectedNode = dragNode
                    End If
  

                End If


                End If

            End If


        MyBase.OnDragDrop(drgevent)
    End Sub

 

    Private Function IsNodeDescendant(ByVal node As TreeNode, ByVal potentialElder As TreeNode) As Boolean
        Dim n As TreeNode

        If node Is Nothing OrElse potentialElder Is Nothing Then Return False

        Do
            n = node.Parent

            If n IsNot Nothing Then
                If n Is potentialElder Then
                    Return True
                Else
                    node = n
                End If
            End If
        Loop Until n Is Nothing

        Return False
    End Function

    Private Function IsNodeLevel1(ByVal node As TreeNode) As Boolean
        Return node.Parent Is Nothing
    End Function




    Private Function FindNodeInDataObject(ByVal dataObject As IDataObject) As TreeNode

        For Each format As String In dataObject.GetFormats
            Dim data As Object = dataObject.GetData(format)

            If GetType(TreeNode).IsAssignableFrom(data.GetType) Then
                Return DirectCast(data, TreeNode)
            End If
        Next

        Return Nothing
    End Function


    Private Sub CheckChildNodes(treeNode As TreeNode, nodeChecked As Boolean)
        Dim node As TreeNode
        For Each node In treeNode.Nodes
            node.Checked = nodeChecked
            If nodeChecked Then
                node.ExpandAll()
            End If
        Next node
    End Sub

    Private Shadows Sub AfterCheck(sender As Object, e As TreeViewEventArgs)

        CheckChildNodes(e.Node, e.Node.Checked)

    End Sub

    Public Sub RemoveChildlessLevel1Nodes()
        Dim node As TreeNode
        For i As Integer = MyBase.Nodes.Count - 1 To 0 Step -1
            node = MyBase.Nodes(i)

            If node.Nodes.Count = 0 Then
                MyBase.Nodes.Remove(node)
            End If
        Next

    End Sub

    Public Sub ReadTags(ByRef givenNodes As TreeNodeCollection,
                               ByRef iChosenPatches As Collection,
                               ByVal listAllRoots As Boolean,
                               ByVal listNoRoots As Boolean,
                               ByVal listTags As Boolean,
                               ByVal checkedItemsOnly As Boolean,
                      Optional ByVal listFullPath As Boolean = False,
                      Optional ByVal rootNode As String = "")
        'This routine assumes usage of a 2level control like TreeViewDraggableNodes2Levels
        Dim node As TreeNode
        For Each node In givenNodes
            If node.Parent Is Nothing Then
                'Level1
                If (node.Nodes.Count > 0 Or listAllRoots) And Not listNoRoots Then
                    If listTags Then
                        iChosenPatches.Add(node.Tag, node.Tag.ToString)
                    Else
                        iChosenPatches.Add(node.Text, node.Text)
                    End If
                End If
                If listFullPath Then
                    ReadTags(node.Nodes, iChosenPatches, listAllRoots, listNoRoots, listTags, checkedItemsOnly, listFullPath, If(listTags, node.Tag.ToString, node.Text) & MyBase.PathSeparator)
                Else
                    ReadTags(node.Nodes, iChosenPatches, listAllRoots, listNoRoots, listTags, checkedItemsOnly)
                End If


            Else
                'Level2
                If node.Checked Or Not checkedItemsOnly Then
                    If listTags Then
                        iChosenPatches.Add(rootNode & node.Tag, rootNode & node.Tag.ToString)
                    Else
                        iChosenPatches.Add(rootNode & node.Text, rootNode & node.Text)
                    End If
                End If
            End If


        Next node


    End Sub


    Public Sub ReadTags(ByRef iChosenPatches As Collection,
                        ByVal listAllRoots As Boolean,
                        ByVal listNoRoots As Boolean,
                        ByVal listTags As Boolean,
                        ByVal checkedItemsOnly As Boolean,
               Optional ByVal listFullPath As Boolean = False,
               Optional ByVal rootNode As String = "")

        ReadTags(MyBase.Nodes, iChosenPatches, listAllRoots, listNoRoots, listTags, checkedItemsOnly, listFullPath, rootNode)

    End Sub
    Public Function CategoryExists(ByVal category As String) As Boolean
        Dim node As TreeNode
        For Each node In MyBase.Nodes
            If node.Text = category Then
                Return True
            End If
        Next
        Return False

    End Function




    'A Category is a root node, with formatting
    Public Function AddCategory(ByVal category As String) As Boolean
        If CategoryExists(category) Then
            Return True
        End If
 
        Dim newNode As TreeNode = New TreeNode(category)
        newNode.Tag = category
        newNode.BackColor = Color.Aqua

        MyBase.Nodes.Add(newNode)

        Return True

    End Function

    'A Category is a root node, with formatting
    Public Function PrependCategory(ByVal category As String) As Boolean

        If CategoryExists(category) Then
            Return True
        End If


        Dim newNode As TreeNode = New TreeNode(category)
        newNode.Tag = category
        newNode.BackColor = Color.Aqua
        MyBase.Nodes.Insert(0, newNode) 'Add node at start of nodes list.

        Return True

    End Function


    Public Function AddFileToCategory(ByVal category As String, ByVal label As String, ByVal path As String, Optional nodeChecked As Boolean = False) As Boolean

        Dim newNode As TreeNode = New TreeNode(label)
        newNode.Tag = path
        newNode.Checked = nodeChecked

        'add a new category
        AddCategory(category)

        Dim node As TreeNode
        For Each node In MyBase.Nodes
            If node.Text = category Then
                node.Nodes.Add(newNode)
                Return True
            End If
        Next

        Return False


    End Function



    Public Sub populateTreeFromCollection(ByRef categorisedItemList As Collection, Optional ByVal checked As Boolean = False)

        MyBase.PathSeparator = "#"
        MyBase.Nodes.Clear()

        'copy each item from listbox
        Dim found As Boolean = False
        Dim patch As String = Nothing
        For Each categorisedItem As String In categorisedItemList

            Dim category As String = categorisedItem.Split(MyBase.PathSeparator)(0)
            Dim item As String = categorisedItem.Split(MyBase.PathSeparator)(1)

            'find or create each node for item
            found = AddFileToCategory(category, item, item, checked)

        Next

    End Sub


End Class

Public Class NodeDraggingOverEventArgs
    Inherits EventArgs

    Private _DropLegal As Boolean
    Private _MovingNode As TreeNode
    Private _TargetNode As TreeNode

    Public Sub New(ByVal movingNode As TreeNode, ByVal targetNode As TreeNode)
        _DropLegal = True
        _MovingNode = movingNode
        _TargetNode = targetNode
    End Sub

    Public ReadOnly Property TargetNode() As TreeNode
        Get
            Return _TargetNode
        End Get
    End Property
    Public ReadOnly Property MovingNode() As TreeNode
        Get
            Return _MovingNode
        End Get
    End Property

    'Use this to disallow a drop
    Public Property DropIsLegal() As Boolean
        Get
            Return _DropLegal
        End Get
        Set(ByVal value As Boolean)
            _DropLegal = value
        End Set
    End Property


End Class

Public Class NodeMovingByDragEventArgs
    Inherits EventArgs

    Private _MovingNode As TreeNode
    Private _CurParent As TreeNode
    Private _ParentToBe As TreeNode

    Private _CancelMove As Boolean

    Public Sub New(ByVal nodeMoving As TreeNode, ByVal prevParent As TreeNode, ByVal parentToBe As TreeNode)
        _MovingNode = nodeMoving
        _CurParent = prevParent
        _ParentToBe = parentToBe
    End Sub
    Public Property CancelMove() As Boolean
        Get
            Return _cancelMove
        End Get
        Set(ByVal value As Boolean)
            _CancelMove = value
        End Set
    End Property
    Public ReadOnly Property MovingNode() As TreeNode
        Get
            Return _MovingNode
        End Get
    End Property
    Public ReadOnly Property CurrentParent() As TreeNode
        Get
            Return _CurParent
        End Get
    End Property
    Public ReadOnly Property ParentToBe() As TreeNode
        Get
            Return _ParentToBe
        End Get
    End Property
End Class

Public Class NodeMovedByDragEventArgs
    Inherits EventArgs

    Private _MovedNode As TreeNode
    Private _PreviousParent As TreeNode

    Public Sub New(ByVal nodeMoved As TreeNode, ByVal prevParent As TreeNode)
        _MovedNode = nodeMoved
        _PreviousParent = prevParent
    End Sub
    Public ReadOnly Property MovedNode() As TreeNode
        Get
            Return _MovedNode
        End Get
    End Property
    Public ReadOnly Property PreviousParent() As TreeNode
        Get
            Return _PreviousParent
        End Get
    End Property

End Class

