Attribute VB_Name = "ModTVNodeShifting"
'******************************************************************
'***************Copyright PSST 2004********************************
'***************Written by MrBobo**********************************
'This code was submitted to Planet Source Code (www.planetsourcecode.com)
'If you downloaded it elsewhere, they stole it and I'll eat them alive

Option Explicit
Public Sub ShiftNode(TargetTV As TreeView, TargetNode As Node, MoveMode As TreeRelationshipConstants)
    Dim NewNode As Node
    'Firstly we add a new node in the position specified in the MoveMode parameter
    Select Case MoveMode
        Case tvwFirst
            If TargetNode.Previous Is Nothing Then Exit Sub
            Set NewNode = TargetTV.Nodes.Add(TargetNode.FirstSibling, tvwFirst)
        Case tvwPrevious
            If TargetNode.Previous Is Nothing Then Exit Sub
            Set NewNode = TargetTV.Nodes.Add(TargetNode.Previous, tvwPrevious)
        Case tvwNext
            If TargetNode.Next Is Nothing Then Exit Sub
            Set NewNode = TargetTV.Nodes.Add(TargetNode.Next, tvwNext)
        Case tvwLast
            If TargetNode.Next Is Nothing Then Exit Sub
            Set NewNode = TargetTV.Nodes.Add(TargetNode.LastSibling, tvwLast)
        Case Else
            Exit Sub
    End Select
    'Copy all of it's attributes to the new node
    CopyNodeAttrributes TargetNode, NewNode
    'If there are any children of the node being moved we need to copy those as well
    If TargetNode.Children <> 0 Then IterateChildren TargetTV, TargetNode, NewNode
    'Finally we remove the original node
    TargetTV.Nodes.Remove TargetNode.Index
    Set TargetNode = Nothing
End Sub
Public Sub IterateChildren(TV As TreeView, TargetNode As Node, ParentNode As Node)
    Dim z As Long
    Dim NewNode As Node
    Dim TargetChild As Node
    'Cycle through each of the child nodes
    Set TargetChild = TargetNode.Child
    For z = 1 To TargetNode.Children
        'Add a child to our new node
        Set NewNode = TV.Nodes.Add(ParentNode, tvwChild)
        'Copy all of it's attributes to the new child
        CopyNodeAttrributes TargetChild, NewNode
        'If the child has children itself then get those too
        If TargetChild.Children <> 0 Then IterateChildren TV, TargetChild, NewNode
        'Change our target to the next child
        Set TargetChild = TargetChild.Next
    Next
End Sub
Public Sub CopyNodeAttrributes(TargetNode As Node, NewNode As Node)
On Error Resume Next
    'Copies the attributes of one node to another node
    With TargetNode
        'We need to change it's key before we use it again
        .Key = "TemporaryNodeKey" & .Key
        NewNode.Key = Right(.Key, Len(.Key) - 16)
        NewNode.BackColor = .BackColor
        NewNode.Bold = .Bold
        NewNode.Checked = .Checked
        NewNode.Expanded = .Expanded
        NewNode.ExpandedImage = .ExpandedImage
        NewNode.ForeColor = .ForeColor
        NewNode.Image = .Image
        NewNode.Selected = .Selected
        NewNode.SelectedImage = .SelectedImage
        NewNode.Tag = .Tag
        NewNode.Text = .Text
    End With
End Sub

