VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpHNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Kvp")
Option Explicit

Private Const MY_TYPENAME               As String = "KvpHNode"

' kvpHNode uses two sets of pointers to manage the connection of nodes via the hash/item conbination
' and the connection via the order field

Private Type Properties
    HashPred                    As KvpHNode
    HashSucc                    As KvpHNode
    OrderSucc                   As KvpHNode
    OrderPred                   As KvpHNode
    Item                        As Variant
    Key                         As Variant
    Hash                        As Long
    Order                       As Long
End Type

Private p                       As Properties


'@DefaultMember
Public Function Deb _
( _
Optional ByRef ipHash As Long = -1, _
Optional ByRef ipKey As Variant = Null, _
Optional ByRef ipItem As Variant = Null, _
Optional ByRef ipHashPred As KvpHNode = Nothing, _
Optional ByRef ipHashSucc As KvpHNode = Nothing, _
Optional ByRef ipOrder As Long = -1 _
) As KvpHNode
Attribute Deb.VB_UserMemId = 0
    With New KvpHNode
        Set Deb = .ConstructInstance(ipHash, ipKey, ipItem, ipHashPred, ipHashSucc, ipOrder)
    End With
End Function


Friend Function ConstructInstance _
( _
    ByRef ipHash As Long, _
    ByRef ipKey As Variant, _
    ByRef ipItem As Variant, _
    ByRef ipHashPred As KvpHNode, _
    ByRef ipHashSucc As KvpHNode, _
    ByRef ipOrder As Long _
) As KvpHNode

    Set p.HashPred = ipHashPred
    Set p.HashSucc = ipHashSucc
    
    Set p.OrderPred = Nothing
    Set p.OrderSucc = Nothing
    
    If VBA.IsObject(ipItem) Then
        Set p.Item = ipItem
    Else
        p.Item = ipItem
    End If
    
    If VBA.IsObject(ipKey) Then
        Set p.Key = ipKey
    Else
        p.Key = ipKey
    End If
    
    p.Hash = ipHash
    p.Order = ipOrder
    
    Set ConstructInstance = Me
    
End Function

Public Sub UpdateOrder()

    If p.OrderPred Is Nothing Then
        p.Order = 1
    Else
        p.Order = p.OrderPred.Order + 1
    End If
    
    If p.OrderSucc Is Nothing Then
        Exit Sub
    Else
        p.OrderSucc.UpdateOrder
    End If
    
End Sub

Public Property Get HashSucc() As KvpHNode
    Set HashSucc = p.HashSucc
End Property


Public Property Set HashSucc(ByVal ipKvpHNode As KvpHNode)
    Set p.HashSucc = ipKvpHNode
End Property




Public Property Get HashPred() As KvpHNode
    Set HashPred = p.HashPred
End Property


Public Property Set HashPred(ByVal ipKvpHNode As KvpHNode)
    Set p.HashPred = ipKvpHNode
End Property


Public Property Get OrderSucc() As KvpHNode
    Set OrderSucc = p.OrderSucc
End Property


Public Property Set OrderSucc(ByVal ipKvpHNode As KvpHNode)
    Set p.OrderSucc = ipKvpHNode
End Property


Public Property Get OrderPred() As KvpHNode
    Set OrderPred = p.OrderPred
End Property


Public Property Set OrderPred(ByVal ipKvpHNode As KvpHNode)
    Set p.OrderPred = ipKvpHNode
End Property


Public Property Get Item() As Variant
    If VBA.IsObject(p.Item) Then
        Set Item = p.Item
    Else
        Item = p.Item
    End If
End Property


Public Property Set Item(ByVal ipItem As Variant)
    Set p.Item = ipItem
End Property


Public Property Let Item(ByVal ipItem As Variant)
    p.Item = ipItem
End Property


Public Property Get Key() As Variant
    If VBA.IsObject(p.Key) Then
        Set Key = p.Key
    Else
        Key = p.Key
    End If
End Property


Public Property Set Key(ByVal ipKey As Variant)
    Set p.Key = ipKey
End Property


Public Property Let Key(ByVal ipKey As Variant)
    p.Key = ipKey
End Property


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property


Public Property Get Hash() As Variant
    Hash = p.Hash
End Property


Public Property Let Hash(ByVal ipHash As Variant)
    p.Hash = ipHash
End Property


Public Property Get Order() As Variant
    Order = p.Order
End Property


Public Property Let Order(ByVal ipOrder As Variant)
    p.Order = ipOrder
End Property

Private Sub Class_Terminate()
    Set p.HashSucc = Nothing
    Set p.HashPred = Nothing
    Set p.OrderPred = Nothing
    Set p.OrderSucc = Nothing
End Sub
