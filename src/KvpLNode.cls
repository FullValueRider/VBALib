VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpLNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Kvp")
Option Explicit


Private Const MY_TYPENAME               As String = "KvpLNode"

 
Private Type Properties
    Pred                    As KvpLNode
    Succ                    As KvpLNode
    Item                    As Variant
    Key                     As Variant
    User                    As Variant  ' a valiable that allows user data to be stored in a node
                                        ' typically used to pass back index information
                                        ' in the getnodebyitem and getnodebykey methods
End Type


Private p                   As Properties


'@DefaultMember
Public Function Deb(ByRef ipKey As Variant, ByRef ipItem As Variant, Optional ByRef ipPred As KvpLNode = Nothing, Optional ByRef ipSucc As KvpLNode = Nothing) As KvpLNode
Attribute Deb.VB_UserMemId = 0
    With New KvpLNode
        Set Deb = .ConstructInstance(ipKey, ipItem, ipPred, ipSucc)
    End With
End Function


Friend Function ConstructInstance(ByRef ipKey As Variant, ByRef ipItem As Variant, ByRef ipPred As KvpLNode, ByRef ipSucc As KvpLNode) As KvpLNode


    Guard.KeyIsAdmin ipKey, MY_TYPENAME, "Deb"
    
    VariantCopy p.Key, ipKey
    VariantCopy p.Item, ipItem
    
'    If VBA.IsObject(ipKey) Then
'        Set p.Key = ipKey
'    Else
'        p.Key = ipKey
'    End If
'
'    If VBA.IsObject(ipItem) Then
'        Set p.Item = ipItem
'    Else
'        p.Item = ipItem
'    End If
    
    Set p.Pred = ipPred
    Set p.Succ = ipSucc

    Set ConstructInstance = Me
    
End Function

Private Sub Terminate()
    Set p.Pred = Nothing
    Set p.Succ = Nothing
End Sub


Public Property Get Succ() As KvpLNode
    Set Succ = p.Succ
End Property


Public Property Set Succ(ByVal ipListNode As KvpLNode)
    Set p.Succ = ipListNode
End Property


Public Property Get Pred() As KvpLNode
    Set Pred = p.Pred
End Property


Public Property Set Pred(ByVal ipListNode As KvpLNode)
    Set p.Pred = ipListNode
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



Public Property Get User() As Variant
    If VBA.IsObject(p.User) Then
        Set User = p.User
    Else
        User = p.User
    End If
End Property


Public Property Set User(ByVal ipUser As Variant)
    Set p.User = ipUser
End Property


Public Property Let User(ByVal ipUser As Variant)
    p.User = ipUser
End Property

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
