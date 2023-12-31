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

Private Type Properties

    FirstKey                As Variant
    CurrentKey              As Variant
'    FirstUse                As Boolean
    FirstKeyType            As String
End Type

Private p                   As Properties

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class admin
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Function Deb(Optional ByVal ipFirstKey As Variant) As AutoKeyByNumber

    With New AutoKeyByNumber
        
        Set Deb = .ReadyToUseInstance(IIf(VBA.IsMissing(ipFirstKey), 0, ipFirstKey))

    End With
    
End Function


Public Function ReadyToUseInstance(ByVal ipFirstKey As Variant) As AutoKeyByNumber

    
    If Types.IsObject(ipFirstKey) Then
    
        Err.Raise _
            17 + vbObjectError, _
            TypeName(Me) & ": ReadyToUseInstance", _
            Fmt.Txt("The input variable Type must be one of {nl2}{0}'", VBA.Join(Types.Group.Numbers.ToArray, Char.comma)) 'TypesAid.ShortNumberTypes)
            
    End If
    
    If Types.Group.IsNotNumber(ipFirstKey) Then
    
        ' Autokey wasn't a VBANumericType so that's a whooopsy
        Err.Raise _
            17 + vbObjectError, _
            TypeName(Me) & ": ReadyToUseInstance", _
            "Unexpected type (not a type defined by Types.SimpleNumberTypes) "
            
    End If
    
    p.FirstKey = ipFirstKey
    p.CurrentKey = Empty
    'p.FirstUse = p.FirstKey = 0
    p.FirstKeyType = TypeName(p.FirstKey)
    Set ReadyToUseInstance = Me
    
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Interface
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Property Get IAutoKey_FirstKey() As Variant
    IAutoKey_FirstKey = FirstKey
End Property

Private Property Let IAutoKey_FirstKey(ByVal ipValue As Variant)
    FirstKey = ipValue
End Property


Private Property Get IAutoKey_CurrentKey() As Variant
    IAutoKey_CurrentKey = CurrentKey
End Property

Private Property Let IAutoKey_CurrentKey(ByVal ipValue As Variant)
    CurrentKey = ipValue
End Property

Public Property Get IAutoKey_FirstKeyType() As String
    IAutoKey_FirstKeyType = p.FirstKeyType
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
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class body
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Public Property Get FirstKey() As Variant
    FirstKey = p.FirstKey
End Property
    
Public Property Let FirstKey(ByVal ipValue As Variant)
    p.FirstKey = ipValue
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
    Set myAutoKey = Me.Deb(p.FirstKey)
    
    myAutoKey.CurrentKey = p.CurrentKey
    'myAutoKey.FirstUse = p.FirstUse
    
    Set Clone = myAutoKey
    
End Function

'Public Function GetDefaultFirstKey() As Variant
'    GetDefaultFirstKey = DefaultFirstKey
'End Function

Public Property Get GetNextKey() As Variant

    If VBA.IsEmpty(p.CurrentKey) Then
    
        p.CurrentKey = p.FirstKey
        
    Else
    
        p.CurrentKey = Inc(p.CurrentKey)
        
    End If
    
    GetNextKey = p.CurrentKey
    
End Property


Public Function Name() As String
    Name = TypeName(Me)
End Function

Public Property Get FirstKeyType() As String
    FirstKeyType = p.FirstKeyType
End Property

Public Function GetKeySequence(ByVal ipCount As Long) As Lyst
    
    Dim myList As Lyst
    Set myList = Lyst.Deb
    Set GetKeySequence = myList
    
    If ipCount <= 0 Then Exit Function
        
    Dim mycounter As Long
    For mycounter = 1 To ipCount
        
        myList.Add GetNextKey
        
    Next
    
    Set GetKeySequence = myList
    
 End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class private members
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Function Inc(ByVal ipKey As Variant) As Variant
    Inc = ipKey + 1
End Function


