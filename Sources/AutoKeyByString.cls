VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoKeyByString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Preamble
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' DefaultKeyCharSequence is used for incrementing string keys.
' e.g 0009->000A->000z->0010.
' Incrementing is done by the Index position and not by character value
' hence for a char sequence of "a0b2c3d4etc" we get 0002->000c->0003->000d
' Characters not in the sequence are ignored for the purposes of
' incrementing BUT characters are inserted after the ignored character
' if required by a rollover e.g. #z ->#10 not 1#0
'

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@PredeclaredId
'@Exposed
'@Folder("AutoKeys")

Option Explicit

Implements IAutoKey
Private Type ReadOnly
  
  DefaultFirstUseKey                        As String
  DefaultAllowedCharListForKey              As ROC
  
End Type

Private r                                   As ReadOnly
Private Type Properties

    FirstUseKey                             As Variant
    CurrentKey                              As Variant
    FirstUse                                As Boolean
    KeyAllowedChars                         As ROC
    KeyAsChars                              As Lyst
    FirstUseKeyType                         As String
    Result                                  As Result
End Type

Private p                                   As Properties

Private Sub Class_Initialise()

   r.DefaultFirstUseKey = "0000"
   Set r.DefaultAllowedCharListForKey = Strs.AlphaNumList

End Sub

Public Function Deb _
( _
    Optional ByVal ipFirstUseKey As String, _
    Optional ByVal ipKeyCharList As String _
) As AutoKeyByString

    With New AutoKeyByString
        
        Set Deb = .ReadyToUseInstance(ipFirstUseKey, ipKeyCharList) ' Why is ipKeyCharsequence a null string
            
    End With
        
End Function

Public Function ReadyToUseInstance _
( _
    ByVal ipFirstUseKey As String, _
    ByVal ipKeyCharList As String _
) As AutoKeyByString
    r.DefaultFirstUseKey = "0000"
    p.FirstUseKey = IIf(VBA.Len(ipFirstUseKey) = 0, r.DefaultFirstUseKey, ipFirstUseKey)
    If VBA.Len(ipKeyCharList) = 0 Then
   
      Set p.KeyAllowedChars = r.DefaultAllowedCharListForKey
      
      
    Else
      
      Set p.KeyAllowedChars = ROC.Deb(ipKeyCharList)
      
      
    End If
  
    Set p.Result = Result.Deb
    p.CurrentKey = Empty
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
    IAutoKey_Name = TypeName(Me)
End Function

Private Function IAutoKey_GetKeySequence(ByVal ipCount As Long) As Lyst
    Set IAutoKey_GetKeySequence = GetKeySequence(ipCount)
End Function

Private Function IAutoKey_Res() As Result
  Set IAutoKey_Res = Res
End Function
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class public members
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

Public Property Get KeyCharList() As ROC
    Set KeyCharList = p.KeyAllowedChars
End Property

Public Property Let KeyCharList(ByVal ipValue As ROC)
    Set p.KeyAllowedChars = ipValue
End Property

Public Property Get FirstUseKeyType() As String
    FirstUseKeyType = p.FirstUseKeyType
End Property

'Public Property Get FirstUse() As Boolean
'    FirstUse = p.FirstUse
'End Property
'
'Public Property Let FirstUse(ByVal ipValue As Boolean)
'    p.FirstUse = ipValue
'End Property


Public Function Clone() As AutoKeyByString
    ' we don't use current key as the first key in the clone becuase
    ' this will already exist in the Host ordered Dictionary

    Dim myAutoKey As AutoKeyByString
    Set myAutoKey = Me.Deb(p.FirstUseKey, p.KeyAllowedChars.ToString)
    'myAutoKey.CurrentKey = p.CurrentKey
    'myAutoKey.FirstUse = p.FirstUse
    myAutoKey.FirstUseKey = p.FirstUseKey
        
    Set Clone = myAutoKey
        
End Function

Public Property Get GetNextKey() As Variant

    If VBA.IsEmpty(p.CurrentKey) Then
        
        p.CurrentKey = p.FirstUseKey
            
    Else
        
        p.CurrentKey = Strs.Inc(p.CurrentKey, p.KeyAllowedChars)
            
    End If
        
    GetNextKey = p.CurrentKey
        
End Property

Public Function Name() As String
    Name = TypeName(Me)
End Function

Public Function GetKeySequence(ByVal ipCount As Long) As Lyst
        
    Dim myList As Lyst
    Set myList = Lyst.Deb
    Set GetKeySequence = myList
        
    If ipCount <= 0 Then Exit Function
            
    Dim myCounter As Long
    For myCounter = 1 To ipCount
            
        myList.Add GetNextKey
            
    Next
        
End Function

Public Function Res() As Result
  Set Res = p.Result
End Function
