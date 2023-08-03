VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("vbWatchDog")
Option Explicit

Private Const MY_TYPENAME               As String = "ListNode"

 
Private Type State
    Pred                    As ListNode
    Succ                    As ListNode
    Item                    As Variant
End Type

Private s                   As State

'@DefaultMember
Public Function Deb(Optional ByRef ipPred As ListNode = Nothing, Optional ByRef ipSucc As ListNode = Nothing, Optional ByRef ipItem As Variant = Null) As ListNode
Attribute Deb.VB_UserMemId = 0
    With New ListNode
        Set Deb = .ConstructInstance(ipPred, ipSucc, ipItem)
    End With
End Function

Friend Function ConstructInstance(ByRef ipPred As ListNode, ByRef ipSucc As ListNode, ByRef ipItem As Variant) As ListNode

    Set s.Pred = ipPred
    Set s.Succ = ipSucc
   
    If VBA.IsObject(ipItem) Then
        Set s.Item = ipItem
    Else
        s.Item = ipItem
    End If
    
    
    Set ConstructInstance = Me
    
End Function


Public Property Get Succ() As ListNode
    Set Succ = s.Succ
End Property

Public Property Set Succ(ByVal ipListNode As ListNode)
    Set s.Succ = ipListNode
End Property


Public Property Get Pred() As ListNode
    Set Pred = s.Pred
End Property

Public Property Set Pred(ByVal ipListNode As ListNode)
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


Public Property Get Typename() As String
    Typename = MY_TYPENAME
End Property
