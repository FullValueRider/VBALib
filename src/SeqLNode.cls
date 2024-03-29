VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeqLNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Seq")
Option Explicit

Private Const MY_TYPENAME               As String = "SeqLNode"

 
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

Private Sub Terminate()
    Set s.Pred = Nothing
    Set s.Succ = Nothing
End Sub


Public Property Get Succ() As SeqLNode
    Set Succ = s.Succ
End Property


Public Property Set Succ(ByVal ipListNode As SeqLNode)
    Set s.Succ = ipListNode
End Property


Public Property Get Pred() As SeqLNode
    Set Pred = s.Pred
End Property


Public Property Set Pred(ByVal ipListNode As SeqLNode)
    Set s.Pred = ipListNode
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


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
