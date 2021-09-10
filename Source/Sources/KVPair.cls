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
'@PredeclaredId
'@Exposed
Option Explicit


Private Type Properties

    Key                 As Variant
    Value               As Variant
    
End Type

Private p               As Properties

Public Function Deb(ByVal ipKeyOrKVPair As Variant, Optional ByVal ipValue As Variant = Empty) As KVPair

    With New KVPair
    
        Set Deb = .ReadyToUseInstance(ipKeyOrKVPair, ipValue)
        
    End With
    
End Function


Public Function ReadyToUseInstance(ByVal ipKeyOrKVPair As Variant, ByVal ipValue As Variant) As KVPair
    
    '@Ignore ImplicitDefaultMemberAccess
    If Types.IsTypeByName(Types.Name.OfKVPair, ipKeyOrKVPair) Then
    
        p.Key = ipKeyOrKVPair.Key
        p.Value = ipKeyOrKVPair.Value
        
    Else
    
        p.Key = ipKeyOrKVPair
        p.Value = ipValue
        
    End If
        
    Set ReadyToUseInstance = Me
    
End Function


Public Property Get Key() As Variant
    Key = p.Key
End Property

Public Property Let Key(ByVal ipKey As Variant)
    p.Key = ipKey
End Property


Public Property Get Value() As Variant
    Value = p.Value
End Property

Public Property Let Value(ByVal ipValue As Variant)
    p.Value = ipValue
End Property


Public Function ToString(Optional ByVal ipSeparator As String = Char.Comma) As String
    ToString = Fmt.Txt("{1}{0}{2}", ipSeparator, p.Key, p.Value)
End Function

'@DefaultMember
Public Function Name() As String
Attribute Name.VB_UserMemId = 0
    Name = TypeName(Me)
End Function
