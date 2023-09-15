VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeqTNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Seq")
'https://www.codeproject.com/KB/vb/TreapVb/Treapvb_src.zip
Option Explicit

'' SeqTNode is used to manage the treap structure in SeqT
'' SeqT allows items to be deleted by Index, where index is
'' the order of addition to the SeqT.
'' When a node is deleted by index the deletion will fail
'' if the pointer to the node in the parent node cannot be
'' updated appropritely

Private Type State
    
    Item                                As Variant
    Priority                            As Long
    Left                                As SeqTNode
    Right                               As SeqTNode
    Order                               As Long
  '  Parent                              As SeqTNode
    
End Type

Private s                               As State

'@DefaultMember
Public Function Deb _
( _
    ByRef ipItem As Variant, _
    ByRef ipPriority As Variant, _
    Optional ByRef ipOrder As Long, _
    Optional ByRef ipLeft As SeqTNode = Nothing, _
    Optional ByRef ipRight As SeqTNode = Nothing _
) As SeqTNode
Attribute Deb.VB_UserMemId = 0
    With New SeqTNode
        Set Deb = .ConstructInstance(ipItem, ipPriority, ipOrder, ipLeft, ipRight)
    End With
End Function


Public Function ConstructInstance _
( _
    ByRef ipItem As Variant, _
    ByRef ipPriority As Variant, _
    ByRef ipOrder As Long, _
    ByRef ipLeft As SeqTNode, _
    ByRef ipRight As SeqTNode _
) As SeqTNode

    If VBA.IsObject(ipItem) Then
        Set s.Item = ipItem
    Else
        s.Item = ipItem
    End If
    
    s.Priority = ipPriority
    s.Order = ipOrder
    Set s.Left = ipLeft
    Set s.Right = ipRight
    
    Set ConstructInstance = Me
    
End Function

Public Property Get Item() As Variant
    If VBA.IsObject(s.Item) Then
        Set Item = s.Item
    Else
        Item = s.Item
    End If
End Property

Public Property Set Item(ByVal ipItem As Variant)
    Set s.Item = ipItem
End Property

Public Property Let Item(ByVal ipItem As Variant)
    s.Item = ipItem
End Property


Public Property Get Priority() As Long
    Priority = s.Priority
End Property

Public Property Let Priority(ByVal ipPriority As Long)
    s.Priority = ipPriority
End Property


Public Property Get Left() As SeqTNode
    Set Left = s.Left
End Property

Public Property Set Left(ByVal ipLeft As SeqTNode)
    Set s.Left = ipLeft
End Property


Public Property Get Right() As SeqTNode
    Set Right = s.Right
End Property

Public Property Set Right(ByVal ipRight As SeqTNode)
    Set s.Right = ipRight
End Property

Public Property Get Parent() As SeqTNode
    Set Parent = s.Parent
End Property

Public Property Set Parent(ByVal ipParent As SeqTNode)
    Set s.Parent = ipParent
End Property

Public Property Get Order() As Long
    Order = s.Order
End Property

Public Property Let Order(ByVal ipOrder As Long)
    If ipOrder < 1 Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.SeqTNode.Let Order", _
            Fmt.Text("Expecting an integer greater than zero.  Got '{0}'.", ipOrder)
    End If
    
    s.Order = ipOrder
    
End Property

Public Function RotateRight() As SeqTNode

    Dim myTmp As SeqTNode: Set myTmp = Me.Left
    Set Me.Left = Me.Left.Right
    Set myTmp.Right = Me

    Set RotateRight = myTmp

End Function

Public Function RotateLeft() As SeqTNode

    Dim myTmp As SeqTNode: Set myTmp = Me.Right
    Set Me.Right = Me.Right.Left
    Set myTmp.Left = Me

    Set RotateLeft = myTmp

End Function
