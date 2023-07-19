VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeqLNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'Folder("set")
Option Explicit


Private Type State
    Pred                    As SeqLNode
    Succ                    As SeqLNode
    Item                    As Variant
End Type

Private s                   As State

'@DefaultMember
Public Function Deb(Optional ByRef ipPred As SeqLNode = Nothing, Optional ByRef ipSucc As SeqLNode = Nothing, Optional ByRef ipItem As Variant = Null) As SeqLNode
Attribute Deb.VB_UserMemId = 0
    With New SeqLNode
        Set Deb = .ConstructInstance(ipPred, ipSucc, ipItem)
    End With
End Function

Friend Function ConstructInstance(ByRef ipPred As SeqLNode, ByRef ipSucc As SeqLNode, ByRef ipItem As Variant) As SeqLNode

    Set s.Pred = ipPred
    Set s.Succ = ipSucc
   
    If VBA.IsObject(ipItem) Then
        Set s.Item = ipItem
    Else
        s.Item = ipItem
    End If
    
    
    Set ConstructInstance = Me
    
End Function


Public Property Get Succ() As SeqLNode
    Set Succ = s.Succ
End Property

Public Property Set Succ(ByVal ipSeqLNode As SeqLNode)
    Set s.Succ = ipSeqLNode
End Property


Public Property Get Pred() As SeqLNode
    Set Pred = s.Pred
End Property

Public Property Set Pred(ByVal ipSeqLNode As SeqLNode)
    Set s.Pred = ipSeqLNode
End Property


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
