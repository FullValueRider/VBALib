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
Option Explicit

Private Type State
    
    Item                                As Variant
    Priority                            As Long
    Left                                As SeqTNode
    Right                               As SeqTNode
    
End Type

Private s                               As State

'@DefaultMember
Public Function Deb _
( _
    ByRef ipItem As Variant, _
    ByVal ipPriority As Variant, _
    Optional ByRef ipLeft As SeqTNode = Nothing, _
    Optional ByRef ipRight As SeqTNode = Nothing _
) As SeqTNode
Attribute Deb.VB_UserMemId = 0
    With New SeqTNode
        Set Deb = .ConstructInstance(ipItem, ipPriority, ipLeft, ipRight)
    End With
End Function


Public Function ConstructInstance _
( _
    ByRef ipItem As Variant, _
    ByVal ipPriority As Variant, _
    ByRef ipLeft As SeqTNode, _
    ByRef ipRight As SeqTNode _
) As SeqTNode

    If VBA.IsObject(ipItem) Then
        Set s.Item = ipItem
    Else
        s.Item = ipItem
    End If
    
    s.Priority = ipPriority
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
