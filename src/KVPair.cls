VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KVPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A simple single item dictionary"
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========COption Explicit
'@Moduledescription("A simple single item dictionary")
'@PredeclaredId
'@Exposed
'@Folder("Kvp")
Option Explicit


Private Type Properties

    Keys                                 As Collection
    Items                                As Collection
        
End Type

Private p               As Properties

'@DefaultMember
Public Function Deb(Optional ByRef ipKVPairOrKey As Variant = Empty, Optional ByRef ipItem As Variant = Empty) As KVPair
Attribute Deb.VB_UserMemId = 0
    With New KVPair
        Set Deb = .ConstructInstance(ipKVPairOrKey, ipItem)
    End With
End Function
 
Friend Function ConstructInstance(ByRef ipKVPairOrKey As Variant, ByRef ipItem As Variant) As KVPair

    Set p.Keys = New Collection
    Set p.Items = New Collection
    Set ConstructInstance = Me
    
    If VBA.TypeName(ipKVPairOrKey) = "KVPair" Then
        Set ConstructInstance = ConstructInstanceUsingKVPair(ipKVPairOrKey)
        
    Else
        Set ConstructInstance = ConstructInstanceUsingKeyAndItem(ipKVPairOrKey, ipItem)
    End If
            
End Function

Private Function ConstructInstanceUsingKVPair(ByRef ipKVpair As Variant) As KVPair

    p.Keys.Add ipKVpair.Key
    p.Items.Add ipKVpair.Item
    Set ConstructInstanceUsingKVPair = Me
    
End Function

Private Function ConstructInstanceUsingKeyAndItem(ByRef ipKey As Variant, ByRef ipItem As Variant) As KVPair

    p.Keys.Add ipKey
    p.Items.Add ipItem
    Set ConstructInstanceUsingKeyAndItem = Me
    
End Function

Public Property Get Key() As Variant
    Key = p.Keys.Item(1)
End Property

Public Property Let Key(ByVal ipKey As Variant)
    p.Keys.Add ipKey
    p.Keys.Remove 1
End Property

Public Property Set Key(ByVal ipKey As Variant)
    p.Keys.Add ipKey
    p.Keys.Remove 1
End Property


Public Property Get Item() As Variant

    If VBA.IsObject(p.Items.Item(1)) Then
        Set Item = p.Items.Item(1)
    Else
        Item = p.Items.Item(1)
    End If
    
End Property

Public Property Let Item(ByVal ipItem As Variant)
    p.Items.Add ipItem
    p.Items.Remove 1
End Property

Public Property Set Item(ByVal ipItem As Variant)
    p.Items.Add ipItem
    p.Items.Remove 1
End Property

'Public Function ToString(Optional ByVal ipSeparator As Variant) As String
'    If VBA.IsMissing(ipSeparator) Then ipSeparator = chars.twComma
'        Debug.Print ipSeparator, p.Key.Item(0), p.Item.Item(0)
'    ToString = Fmt.Txt("{1}{0}{2}", ipSeparator, p.Key.Item(0), p.Item.Item(0))
'End Function

Public Function Name() As String
    Name = TypeName(Me)
End Function

'This may seem stupid but a KVPair can be treated as a single item dictionary
' So to conform with other dictionary type we implement Keys and Items functions
'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = p.Keys.[_NewEnum]
End Function

Public Function Keys() As Variant
    Set Keys = Array(p.Keys.Item(1))
End Function

Public Function Items() As Variant
    Set Items = Array(p.Items.Item(1))
End Function
