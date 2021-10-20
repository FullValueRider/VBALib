VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KVPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========COption Explicit
'@Moduledescription("A simple single item dictionary")
' Originally this class was implemented to emulate a tuple.  However it was an 
' oddity when trying to triage a paramarray and consequently it was transformed 
' into a single item dictionary so that it could be treated a member of the 
' IterableKeysByEnum.

'@PredeclaredId
'@Exposed
Option Explicit


Private Type Properties

    Key                                 As ArrayList
    Item                                As ArrayList
    
End Type

Private p               As Properties

Public Function Deb(Optional ByVal ipKeyOrKVPair As Variant = Empty, Optional ByVal ipValue As Variant = Empty) As KVPair

    With New KVPair
    
        Set Deb = .ReadyToUseInstance(ipKeyOrKVPair, ipValue)
        
    End With
    
End Function


Public Function ReadyToUseInstance(ByVal ipKeyOrKVPair As Variant, ByVal ipValue As Variant) As KVPair
    
    Set p.Key = New ArrayList
    Set p.Item = New ArrayList
    '@Ignore ImplicitDefaultMemberAccess
    If Types.IsKVPair(ipKeyOrKVPair) Then
    
        p.Key.Add ipKeyOrKVPair.Key
        p.Item.Add ipKeyOrKVPair.Item
        
    Else
    
        p.Key.Add ipKeyOrKVPair
        p.Item.Add ipValue
        
    End If
        
    Set ReadyToUseInstance = Me
    
End Function


Public Property Get Key() As Variant
    Key = p.Key.Item(0)
End Property

Public Property Let Key(ByVal ipKey As Variant)
    p.Key.Item(0) = ipKey
End Property


Public Property Get Item(Optional ByVal ipKey As Long = 0) As Variant
    Item = p.Item.Item(0)
End Property

Public Property Let Item(ByVal ipValue As Variant, Optional ByVal ipKey As Long = 0)
    p.Item.Item(0) = ipValue
End Property


Public Function ToString(Optional ByVal ipSeparator As String = Char.Comma) As String
    ToString = Fmt.Txt("{1}{0}{2}", ipSeparator, p.Key.Item(0), p.Item.Item(0))
End Function

'@DefaultMember
Public Function Name() As String
Attribute Name.VB_UserMemId = 0
    Name = TypeName(Me)
End Function

'This may seem stupid but a KVPair can be treated as a single item dictionary
' So to conform with other dictionary type we implement Keys and Items functions

Public Function Keys() As Variant
    Set Keys = p.Key.toarray
End Function

Public Function Items() As Variant
    Set Items = p.Item.toarray
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Dim enumerator As IUnknown
    Set enumerator = p.Key.GetEnumerator(0, p.Key.Count)
    Set NewEnum = enumerator

End Function