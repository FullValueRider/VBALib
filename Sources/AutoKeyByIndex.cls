VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoKeyByIndex"
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
  
  DefaultFirstUseIndex                      As Long
  
End Type

Private r                                   As ReadOnly

Private Type Properties

    FirstUseIndex                           As Long
    CurrentIndex                            As Variant
    KeysList                                As Lyst
    FirstUseKeyType                         As String
    Result                                  As Result
        
End Type

Private p                                   As Properties

Private Sub Class_Initialise()
  'If Me Is AutoKeyByIndex Then Exit Sub
  r.DefaultFirstUseIndex = 1
End Sub

Public Function Deb(Optional ByRef ipKeys As Lyst, Optional ByVal ipFirstuseIndex As Long = 0) As AutoKeyByIndex
        
    With New AutoKeyByIndex
        
        Set Deb = .ReadyToUseInstance(ipKeys, ipFirstuseIndex)
            
    End With
        
End Function

Public Function ReadyToUseInstance(ByVal ipKeys As Lyst, ByVal ipFirstuseIndex As Long) As AutoKeyByIndex
  
    Set p.Result = Globals.Res.SetNotOkay
    Set ReadyToUseInstance = Nothing
    'If BailOut.When(Types.Group.IsNotIterable(ipKeys), Msg.InvalidType) Then Exit Function
    If BailOut.When(ipKeys Is Nothing, Msg.HasNoItems) Then Exit Function
    'If BailOut.When(ipKeys.HasNoItems, Msg.HasNoItems) Then Exit Function
        
    Set p.KeysList = ipKeys
    
    If BailOut.When(ipFirstuseIndex > p.KeysList.LastIndex) Then Exit Function
    If ipFirstuseIndex = 0 Then
        
        p.FirstUseIndex = p.KeysList.FirstIndex
        
    Else
        p.FirstUseIndex = ipFirstuseIndex
        
    End If
        
    p.CurrentIndex = Empty
    p.FirstUseKeyType = TypeName(p.KeysList.Item(p.FirstUseIndex))
    'Debug.Print "Assigned first firstuseindex", p.FirstUseIndex, p.CurrentIndex
    Globals.Res.SetOkay
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

Private Function IAutoKey_Clone() As IAutoKey
    Set IAutoKey_Clone = Clone
End Function

Private Function IAutoKey_GetNextKey() As Variant
    IAutoKey_GetNextKey = GetNextKey
End Function

Private Function IAutoKey_Name() As String
    IAutoKey_Name = Name
End Function

Private Property Get IAutoKey_FirstUseKeyType() As String
    IAutoKey_FirstUseKeyType = p.FirstUseKeyType
End Property

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
'FirstUseKey is an Interface method but in reality
'should be an 'Long' index, not an actual key
'@Description("For AutokeyByIndex Returns the Index of the first used key")
Public Property Get FirstUseKey() As Variant
Attribute FirstUseKey.VB_Description = "For AutokeyByIndex Returns the Index of the first used key"
    FirstUseKey = p.FirstUseIndex
End Property
'@Description("For AutoKeyByIndex ipvalue is the first key to use, not the index of the first key")
Public Property Let FirstUseKey(ByVal ipValue As Variant)
Attribute FirstUseKey.VB_Description = "For AutoKeyByIndex ipvalue is the first key to use, not the index of the first key"
  
  'Debug.Print TypeName(ipValue), TypeName(p.KeysList.Item(p.FirstUseIndex))
  'Debug.Print "Resetting firstuseindex to ", ipValue, p.CurrentIndex
  
  
  If Types.IsSameType(p.KeysList.Item(p.FirstUseIndex), ipValue) Then
    
    p.FirstUseIndex = p.KeysList.IndexOf(ipValue)
    
    
  Else
    
    p.FirstUseIndex = p.KeysList.FirstIndex
    
    
  End If

  p.FirstUseKeyType = VBA.TypeName(p.KeysList.Item(p.FirstUseIndex))
  'Debug.Print "Reset firstuseindex to ", ipValue, p.CurrentIndex
End Property


Public Property Get CurrentKey() As Variant
    If VBA.IsEmpty(p.CurrentIndex) Then
        p.CurrentIndex = p.FirstUseIndex
    End If
    CurrentKey = p.KeysList.Item(p.CurrentIndex)
End Property

Public Property Let CurrentKey(ByVal ipValue As Variant)
    p.CurrentIndex = p.KeysList.IndexOf(ipValue)
End Property

Public Function Clone() As IAutoKey

    Dim myAutoKey As AutoKeyByIndex
    Set myAutoKey = Me.Deb(p.KeysList.ToArray, p.FirstUseIndex)
        
    myAutoKey.CurrentKey = p.KeysList.Item(p.CurrentIndex)
    'myAutoKey.FirstUse = p.FirstUse
        
    Set Clone = myAutoKey
        
End Function

Public Property Get GetNextKey() As Variant

    If VBA.IsEmpty(p.CurrentIndex) Then
        
        p.CurrentIndex = p.FirstUseIndex
            
    Else
        
        p.CurrentIndex = Inc(p.CurrentIndex)
            
    End If
        
    GetNextKey = p.KeysList.Item(p.CurrentIndex)
        
End Property

Public Function Name() As String
    Name = VBA.TypeName(Me)
End Function

Public Property Get FirstUseKeyType() As String
    FirstUseKeyType = p.FirstUseKeyType
End Property

' the list must be set by deb so the code below makes no sense
' Public Function SetKeysList(ByVal ipKeysList As Lyst) As AutoKeyByIndex
        
'     Set p.KeysList = ipKeysList
'     p.FirstUseKeyType = TypeName(p.KeysList.Item(1))
'     Set SetKeysList = Me
        
' End Function

' Public Function SetFirstUseIndex(Optional ByVal ipFirstuseIndex As Long = 1) As AutoKeyByIndex
        
'     If ipFirstuseIndex < p.KeysList.FirstIndex Then
      
'       p.FirstUseIndex = p.KeysList.FirstIndex
      
'     ElseIf p.FirstUseIndex > p.KeysList.LastIndex Then
      
'       p.FirstUseIndex = p.KeysList.LastIndex
      
'     Else
      
'       p.FirstUseIndex = ipFirstuseIndex
      
'     End If
    
'     p.FirstUseKeyType = VBA.TypeName(p.KeysList.Item(p.FirstUseIndex))
    
'     Set SetFirstUseIndex = Me
        
' End Function

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
  Set Res = p.Result
End Function

' Public Function SetKeysList(ByVal ipKeys As Lyst) As AutoKeyByIndex
    
'     Set p.KeysList = ipKeys
'     Set p.firstusekey = ipkeys.index.p.
'     Set SetKeysList = Me
' End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class private members
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Function Inc(ByVal ipKeyIndex As Long) As Long

    If ipKeyIndex = p.KeysList.LastIndex Then
        
        p.CurrentIndex = p.KeysList.FirstIndex
        
        
    Else
        
        p.CurrentIndex = p.CurrentIndex + 1
        
            
    End If
    
    Inc = p.CurrentIndex

End Function
