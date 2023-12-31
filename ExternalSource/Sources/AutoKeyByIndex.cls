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

Private Type Properties

    FirstIndex                              As Long
    CurrentIndex                            As Variant
    KeysList                                As VBALib.Lyst
    FirstKeyType                            As String
    
End Type

Private p                                   As Properties

Public Function Deb(Optional ByVal ipKeys As Variant, Optional ByVal ipFirstIndex As Long = 0) As AutoKeyByIndex

    With New AutoKeyByIndex
    
        Set Deb = .ReadyToUseInstance(ipKeys, ipFirstIndex)
        
    End With
    
End Function

Public Function ReadyToUseInstance(ByVal ipKeys As Variant, ByVal ipFirstIndex As Long) As AutoKeyByIndex

    Set ReadyToUseInstance = Me
    
    If ipFirstIndex < 0 Then
            
        Err.Raise _
            17 + vbObjectError, _
            TypeName(Me) & ": ReadyToUseInstance", _
            Fmt.Txt("The Keys Index cannot be less than 0: Found is '{0}'", ipFirstIndex)
            
        Exit Function
    
    End If
    
    p.FirstIndex = ipFirstIndex
    p.CurrentIndex = Empty
    
    If Types.Iterable.HasNoItems(ipKeys) Then Exit Function
    
    Set p.KeysList = Types.Iterable.ToLyst(ipKeys)
    p.FirstKeyType = TypeName(p.KeysList.Item(ipFirstIndex))
    
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

Private Function IAutoKey_Clone() As IAutoKey
    Set IAutoKey_Clone = Clone
End Function

Private Function IAutoKey_GetNextKey() As Variant
    IAutoKey_GetNextKey = GetNextKey
End Function

Private Function IAutoKey_Name() As String
    IAutoKey_Name = Name
End Function

Private Property Get IAutoKey_FirstKeyType() As String
    IAutoKey_FirstKeyType = p.FirstKeyType
End Property

Private Function IAutoKey_GetKeySequence(ByVal ipCount As Long) As Lyst
    Set IAutoKey_GetKeySequence = GetKeySequence(ipCount)
End Function


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class public members
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Property Get FirstKey() As Variant
    FirstKey = p.KeysList.Item(p.FirstIndex)
End Property

Public Property Let FirstKey(ByVal ipValue As Variant)
    p.FirstIndex = p.KeysList.IndexOf(ipValue)
End Property

Public Property Get CurrentKey() As Variant
    CurrentKey = p.KeysList.Item(p.CurrentIndex)
End Property

Public Property Let CurrentKey(ByVal ipValue As Variant)
    p.CurrentIndex = p.KeysList.IndexOf(ipValue)
End Property


Public Function Clone() As IAutoKey

    Dim myAutoKey As AutoKeyByIndex
    Set myAutoKey = Me.Deb(p.KeysList.ToArray, p.FirstIndex)
    
    myAutoKey.CurrentKey = p.KeysList.Item(p.CurrentIndex)
    'myAutoKey.FirstUse = p.FirstUse
    
    Set Clone = myAutoKey
    
End Function

Public Property Get GetNextKey() As Variant

    If VBA.IsEmpty(p.CurrentIndex) Then
    
        p.CurrentIndex = p.FirstIndex
        
    Else
    
        p.CurrentIndex = Inc(p.CurrentIndex)
        
    End If
    
    GetNextKey = p.KeysList.Item(p.CurrentIndex)
    
End Property

Public Function Name() As String
    Name = TypeName(Me)
End Function

Public Property Get FirstKeyType() As String
    FirstKeyType = p.FirstKeyType
End Property

Public Function SetKeysList(ByVal ipKeysList As Lyst) As AutoKeyByIndex
    
    Set p.KeysList = ipKeysList
    p.FirstKeyType = TypeName(p.KeysList.Item(0))
    Set SetKeysList = Me
    
End Function

Public Function SetFirstIndex(Optional ByVal ipFirstIndex As Long = 0) As AutoKeyByIndex
    
    If ipFirstIndex < 0 Then
        
        Err.Raise _
        17 + vbObjectError, _
        TypeName(Me) & ": ReadyToUseInstance", _
        Fmt.Txt("The Keys Index cannot be less than 0: Found is '{0}'", ipFirstIndex)
        
    Exit Function
        
    End If
    
    p.FirstIndex = ipFirstIndex
    Set SetFirstIndex = Me
    
End Function

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
Private Function Inc(ByVal ipKeyIndex As Long) As Long

    If ipKeyIndex >= p.KeysList.Count - 1 Then
    
       Inc = 0
       
    Else
    
        Inc = p.CurrentIndex + 1
        
    End If

End Function
        

