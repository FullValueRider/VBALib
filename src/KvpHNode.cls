VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpHNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Kvp")
Option Explicit

Private Const MY_TYPENAME               As String = "KvpHNode"

 
Private Type State
    Pred                    As KvpHNode
    Succ                    As KvpHNode
    Item                    As Variant
    Key                     As Variant
    Hash                    As Long
    Order                   As Long
End Type


Private s                   As State


'@DefaultMember
Public Function Deb _
( _
Optional ByRef ipPred As KvpHNode = Nothing, _
Optional ByRef ipSucc As KvpHNode = Nothing, _
Optional ByRef ipHash As Long = -1, _
Optional ByRef ipKey As Variant = Null, _
Optional ByRef ipItem As Variant = Null _
) As KvpHNode
Attribute Deb.VB_UserMemId = 0
    With New KvpHNode
        Set Deb = .ConstructInstance(ipPred, ipSucc, ipHash, ipKey, ipItem)
    End With
End Function


Friend Function ConstructInstance(ByRef ipPred As KvpHNode, ByRef ipSucc As KvpHNode, ByRef ipHash As Long, ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpHNode

    Set s.Pred = ipPred
    Set s.Succ = ipSucc
    
    If VBA.IsObject(ipItem) Then
        Set s.Item = ipItem
    Else
        s.Item = ipItem
    End If
    
    If VBA.IsObject(ipKey) Then
        Set s.Key = ipKey
    Else
        s.Key = ipKey
    End If
    
    s.Hash = ipHash
    
'    If s.Pred Is Nothing Then
'        s.Order = 1
'    Else
'        s.Order = s.Pred.Order + 1
'
'    End If
    
    Me.AdjustNodeOrder
    
    Set ConstructInstance = Me
    
End Function


Public Property Get Succ() As KvpHNode
    Set Succ = s.Succ
End Property


Public Property Set Succ(ByVal ipHDNode As KvpHNode)
    Set s.Succ = ipHDNode
    AdjustNodeOrder
End Property


Public Property Get Pred() As KvpHNode
    Set Pred = s.Pred
End Property


Public Property Set Pred(ByVal ipHDNode As KvpHNode)
    Set s.Pred = ipHDNode
    
    AdjustNodeOrder
    
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


Public Property Get Key() As Variant
    If VBA.IsObject(s.Key) Then
        Set Key = s.Key
    Else
        Key = s.Key
    End If
End Property


Public Property Set Key(ByVal ipKey As Variant)
    Set s.Key = ipKey
End Property


Public Property Let Key(ByVal ipKey As Variant)
    s.Key = ipKey
End Property


Public Property Get Typename() As String
    Typename = MY_TYPENAME
End Property


Public Property Get Hash() As Variant
    Hash = s.Hash
End Property


Public Property Let Hash(ByVal ipHash As Variant)
    s.Hash = ipHash
End Property


Public Property Get Order() As Variant
    Order = s.Order
End Property


Public Property Let Order(ByVal ipOrder As Variant)
    s.Order = ipOrder
End Property


Public Sub AdjustNodeOrder()

    If s.Pred Is Nothing Then
        s.Order = 1
    Else
        s.Order = s.Pred.Order + 1
    End If
    
    If s.Succ Is Nothing Then
        Exit Sub
    End If
    
    s.Succ.AdjustNodeOrder
    
End Sub


Private Sub Class_Terminate()
    Set s.Succ = Nothing
    Set s.Pred = Nothing
End Sub
