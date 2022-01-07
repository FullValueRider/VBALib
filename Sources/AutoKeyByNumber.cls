VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoKeyByNumber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@PredeclaredId
'@Exposed
'@Folder("AutoKeys")
Option Explicit

Implements IAutoKey

Private Type ReadOnly
  
  DefaultKey                                  As Long
  
End Type

Private Type Properties

    FirstUseKey                               As Variant
    CurrentKey                                As Variant
    FirstUseKeyType                           As String
    
    
End Type

Private p                                     As Properties

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class admin
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


Public Function Deb(Optional ByVal ipFirstUseKey As Variant = 0&) As AutoKeyByNumber
    ' Variant is used to allow any number type to be used as a Key
    With New AutoKeyByNumber
            
        Set Deb = .ReadyToUseInstance(ipFirstUseKey)

    End With
        
End Function

Public Function ReadyToUseInstance(ByVal ipFirstUseKey As Variant) As AutoKeyByNumber
  
  'Set Globals.Res = Result.Deb.SetOkay
      
  If Types.Group.IsNotNumber(ipFirstUseKey) Then
      
     Globals.Res.SetNotOkay Msg.InvalidType
     Exit Function
              
  End If
      
  p.FirstUseKey = ipFirstUseKey
  p.CurrentKey = Empty
  'p.FirstUse = p.FirstKey = 0
  p.FirstUseKeyType = TypeName(p.FirstUseKey)
  Set ReadyToUseInstance = Me
        
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Interface
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Property Get IAutoKey_FirstUseKey() As Variant
    IAutoKey_FirstUseKey = FirstUseKey
End Property

Private Property Let IAutoKey_FirstUseKey(ByVal ipValue As Variant)
    FirstUseKey = ipValue
End Property

Private Property Get IAutoKey_CurrentKey() As Variant
    IAutoKey_CurrentKey = CurrentKey
End Property

Private Property Let IAutoKey_CurrentKey(ByVal ipValue As Variant)
    CurrentKey = ipValue
End Property

Public Property Get IAutoKey_FirstUseKeyType() As String
    IAutoKey_FirstUseKeyType = p.FirstUseKeyType
End Property

Private Function IAutoKey_Clone() As IAutoKey
    Set IAutoKey_Clone = Clone
End Function

Private Function IAutoKey_GetNextKey() As Variant
    IAutoKey_GetNextKey = GetNextKey
End Function

Private Function IAutoKey_Name() As String
    IAutoKey_Name = Name
End Function

Private Function IAutoKey_GetKeySequence(ByVal ipCount As Long) As Lyst
    Set IAutoKey_GetKeySequence = GetKeySequence(ipCount)
End Function

Private Function IAutoKey_Res() As Result
    Set IAutoKey_Res = Res
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class body
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Public Property Get FirstUseKey() As Variant
    FirstUseKey = p.FirstUseKey
End Property

Public Property Let FirstUseKey(ByVal ipValue As Variant)
    p.FirstUseKey = ipValue
End Property

Public Property Get CurrentKey() As Variant
    CurrentKey = p.CurrentKey
End Property

Public Property Let CurrentKey(ByVal ipValue As Variant)
    p.CurrentKey = ipValue
End Property

'Public Property Get FirstUse() As Boolean
'    FirstUse = p.FirstUse
'End Property
'
'Public Property Let FirstUse(ByVal ipValue As Boolean)
'    p.FirstUse = ipValue
'End Property


Public Function Clone() As IAutoKey
    ' we don't use current key as the first key in the clone becuase
    ' this will already exist in the Host ordered Dictionary

    Dim myAutoKey As AutoKeyByNumber
    Set myAutoKey = Me.Deb(p.FirstUseKey)
        
    myAutoKey.CurrentKey = p.CurrentKey
    'myAutoKey.FirstUse = p.FirstUse
        
    Set Clone = myAutoKey
        
End Function

'Public Function GetDefaultFirstKey() As Variant
'    GetDefaultFirstKey = DefaultFirstKey
'End Function

Public Property Get GetNextKey() As Variant

    If VBA.IsEmpty(p.CurrentKey) Then
        
        p.CurrentKey = p.FirstUseKey
            
    Else
        
        p.CurrentKey = Inc(p.CurrentKey)
            
    End If
        
    GetNextKey = p.CurrentKey
        
End Property

Public Function Name() As String
    Name = TypeName(Me)
End Function

Public Property Get FirstUseKeyType() As String
    FirstUseKeyType = p.FirstUseKeyType
End Property

Public Function GetKeySequence(ByVal ipCount As Long) As Lyst
        
    Dim myList As Lyst
    Set myList = Lyst.Deb
    Set GetKeySequence = myList
        
    If ipCount <= 0 Then Exit Function
            
    Dim myCounter As Long
    For myCounter = 1 To ipCount
            
        myList.Add GetNextKey
            
    Next
        
    Set GetKeySequence = myList
        
End Function

Public Function Res() As Result
  Set Res = Globals.Res
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class private members
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Function Inc(ByVal ipKey As Variant) As Variant
    Inc = ipKey + 1
End Function


