VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeqT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule
'@PredeclaredId
'@Exposed
'@Folder("Seq")
Option Explicit
' Code for InsertNode, FindNode, DeleteNode, rotateright and rotateleft taken from https://www.geeksforgeeks.org/implementation-of-search-insert-and-delete-in-treap/
Private Type State
    Root                            As SeqTNode
    Count                           As Long
    Size                            As Long
    
End Type

Private s                           As State


'@DefaultMember
Public Function Deb() As SeqT
Attribute Deb.VB_UserMemId = 0
    With New SeqT
        Set Deb = .ConstructInstance()
    End With
End Function

Public Function ConstructInstance() As SeqT
    VBA.Randomize
    s.Size = 0
    Set s.Root = Nothing
    Set ConstructInstance = Me
End Function


Public Function Add(ByRef ipItem As Variant) As Long
    s.Count = 0
    InsertNode s.Root, ipItem
    s.Size = s.Size + 1
    Add = s.Count
End Function

Public Function Count() As Long
    Count = s.Size
End Function

Public Function Insert(ByVal ipIndex As Variant) As SeqT

    
End Function

Private Function FindNode(ByRef ipNode As SeqTNode, ByRef ipItem As Variant) As SeqTNode
    s.Count = s.Count + 1
    ' Base Cases: root is nothing or Item is present in Node
    If ipNode Is Nothing Or Comparers.EQ(ipNode.Item, ipItem) Then
        Set FindNode = ipNode
        Exit Function
    End If
     
    ' Key is greater than root's key
    If Comparers.MT(ipItem, ipNode.Item) Then
        Set FindNode = FindNode(ipNode.Right, ipItem)
        Exit Function
    End If
    ' Key is smaller than root's key
    Set FindNode = FindNode(ipNode.Left, ipItem)
    
     
End Function

Private Function InsertNode(ByRef ipNode As SeqTNode, ByRef ipItem As Variant) As SeqTNode
   
    s.Count = s.Count + 1
    ' ipNode can only be nothing when we pass an uninitialised s.root
    If ipNode Is Nothing Then
        Set ipNode = SeqTNode(ipItem, GetPriority)
        Set InsertNode = ipNode
        Exit Function
    End If
    
    ' If Item is smaller than ipNode
    If Comparers.LTEQ(ipItem, ipNode.Item) Then
        ' Insert in left subtree
        Set ipNode.Left = InsertNode(ipNode.Left, ipItem)
 
        ' Fix Heap property if it is violated
        If ipNode.Left.Priority > ipNode.Priority Then
            Set ipNode = rightRotate(ipNode)
        End If
    Else '  If key is greater
        ' Insert in right subtree
            Set ipNode.Right = InsertNode(ipNode.Right, ipItem)
 
        ' Fix Heap property if it is violated
        If ipNode.Right.Priority > ipNode.Priority Then
            Set ipNode = leftRotate(ipNode)
        End If
        
    End If
     
    Set InsertNode = ipNode
    
End Function

Private Function GetPriority() As Long
        GetPriority = VBA.CLng(VBA.Rnd(1) * (Globals.maxLong - 1))
End Function

Private Function deleteNode(ByRef ipNode As SeqTNode, ByRef ipItem As Variant) As Variant

    Dim myTmp As SeqTNode
    ' Base case
    If ipNode Is Nothing Then
        Set deleteNode = ipNode
        Exit Function
    End If
 
    ' IF ipItemS IS NOT AT ipNode
    If Comparers.LT(ipItem, ipNode.Item) Then
        Set ipNode.Left = deleteNode(ipNode.Left, ipItem)
    ElseIf Comparers.MT(ipItem, ipNode.Item) Then
        Set ipNode.Right = deleteNode(ipNode.Right, ipItem)
 
    ' IF ipItem IS AT ipNode
    ' If left is nothing
    Else
        If ipNode.Left Is Nothing Then
           Set myTmp = ipNode.Right
           Set ipNode = Nothing
           Set deleteNode = myTmp   ' Make right child as ipNode
            Exit Function
       ' If Right is nothing
       ElseIf ipNode.Right Is Nothing Then
           Set myTmp = ipNode.Left
           Set ipNode = Nothing
           Set deleteNode = myTmp  ' Make left child as ipNode
            Exit Function
       ' If ipItem is at ipNode and both left and right are not NULL
       ElseIf ipNode.Left.Priority < ipNode.Right.Priority Then
           Set ipNode = leftRotate(ipNode)
           Set ipNode.Left = deleteNode(ipNode.Left, ipItem)
       Else
           Set ipNode = rightRotate(ipNode)
           Set ipNode.Right = deleteNode(ipNode.Right, ipItem)
       End If
    End If
 
    s.Size = s.Size - 1
    Set deleteNode = ipNode
    
End Function


Private Function rightRotate(ByRef y As SeqTNode) As SeqTNode
    Dim x As SeqTNode: Set x = y.Left
    Dim myTmp As SeqTNode: Set myTmp = x.Right
     
    ' Perform rotation
    Set x.Right = y
    Set y.Left = myTmp
     
    ' Return new root
    Set rightRotate = x
End Function

Private Function leftRotate(ByRef x As SeqTNode) As SeqTNode
    Dim y As SeqTNode: Set y = x.Right
    Dim myTmp As SeqTNode: Set myTmp = y.Left
     
    ' Perform rotation
    Set y.Left = x
    Set x.Right = myTmp
     
    ' Return new root
    Set leftRotate = y
    
End Function


Public Sub PrintByPriority()
    inorder s.Root
End Sub
Private Sub inorder(ByRef ipNode As SeqTNode)

    If ipNode Is Nothing Then
        Exit Sub
    End If
    
    inorder ipNode.Left
    Debug.Print "key:", ipNode.Item, "| priority:", ipNode.Priority, " "
    If Not ipNode.Left Is Nothing Then
        Debug.Print " | left child:", ipNode.Left.Item
    End If
    If Not ipNode.Right Is Nothing Then
        Debug.Print " | right child:", ipNode.Right.Item, " "
    End If
    Debug.Print
    inorder ipNode.Right
End Sub
