VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoKeyByString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
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
'Exposed

Option Explicit

Implements IAutoKey

Private Type Properties

    FirstKey                                As Variant
    CurrentKey                              As Variant
'    FirstUse                                As Boolean
    AllowedKeyChars                         As lyst
    KeyAsChars                              As lyst
    FirstKeyType                            As String
End Type

Private p                                   As Properties

Private Sub Class_Initialise()
    
    ' If Me Is AutoKeyByString Then
        
    '     Debug.Print "Me is PredeclaredId"
        
    ' Else
        
    '     Debug.Print "Other instance"
        
    ' End If
    
End Sub

Public Function Deb _
( _
    Optional ByVal ipFirstKey As String = AutoKeyAid.DefaultKeyForStringKey, _
    Optional ByVal ipKeyCharSequence As String = AutoKeyAid.DefaultKeyCharSequence _
) As AutoKeyByString

    With New AutoKeyByString
    
        Set Deb = .ReadyToUseInstance(ipFirstKey, ipKeyCharSequence) ' Why is ipKeyCharsequence a null string
        
    End With
    
End Function

Public Function ReadyToUseInstance _
( _
    ByVal ipFirstKey As String, _
    ByVal ipKeyCharSequence As String _
) As AutoKeyByString

    p.FirstKey = IIf(ipFirstKey = vbNullString, AutoKeyAid.DefaultKeyForStringKey, ipFirstKey)
    Set p.AllowedKeyChars = VBAlib.Strings.ToCharLyst(IIf(ipKeyCharSequence = vbNullString, DefaultKeyCharSequence, ipKeyCharSequence))
    
    p.CurrentKey = Empty
    ' if the initial key was not the default key then we have had first use
    'p.FirstUse = p.FirstKey = AutoKeyAid.DefaultKeyForString
    
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
' class public members
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


Public Property Get KeyCharSequence() As Variant
    KeyCharSequence = p.AllowedKeyChars.ToString(vbNullString)
End Property

Public Property Let KeyCharSequence(ByVal ipValue As Variant)
    KeyCharSequence = ipValue
End Property

Public Property Get FirstKeyType() As String
    FirstKeyType = p.FirstKeyType
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

    Dim myAutoKey As AutoKeyByString
    Set myAutoKey = AutoKeyByString.Deb(p.FirstKey, Join(p.AllowedKeyChars.ToArray, vbNullString))
    'myAutoKey.CurrentKey = p.CurrentKey
    'myAutoKey.FirstUse = p.FirstUse
    myAutoKey.FirstKey = p.FirstKey
    
    Set Clone = myAutoKey
    
End Function


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


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class private members
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' the harder option. strings are treated as numbers. The sequence
' of numbers allowed for each character is defined by
' AutoKeyCharSequence, which should be one character per string. If
' a character  is encountered that is not in myAutoKeyStrings then key characters
' will be inserted to the rightt of the unknown character to allow the known
' part of the string to increment. in numeric terms /9 would go to
' /10 where the 1 has been inserted to allow incrementing.
'
' The string containing the characters to be used in incrementing
' the key is read left to right as it appears on screen. The key
' string is read in the opposite direction ( right to left ) so
' that it increments in the same pattern as a number when the
' number is displayed on screen.
'
' Example of incrementing where '/' is not in the list of
' characters that can be used in a Key '0000' -> '0001' '000x' ->
' '000y' '000/' -> '000/1' i.e. an implied 0 after the / 
' '000/0' ->'000/1' '000z' -> '0010' '/zzzz' ->'/10000'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Private Function Inc(ByVal ipKey As String) As String
        
        Dim myKeyAsChars As Lyst
        Set myKeyAsChars = VBAlib.Strings.ToCharLyst(ipKey)
        
        ' Dim myKeyAsCharsLastIndex As Long
        ' myKeyAsCharsLastIndex = p.AllowedKeyChars.LastIndex
        
        ' The key is read last index to first index so start we need the index of 
        ' the last character of myKeyAsChars        
        Dim myKeyAsCharsCurrentIndex As Long
        myKeyAsCharsCurrentIndex = myKeyAsChars.LastIndex
        
        ' A flag to indicate that the current character in myKeyChars as been
        ' rolled over from the last char in the allowed char sequence to the 
        ' first char in the allowed char sequence
        Dim myCarryFlag As Boolean
        myCarryFlag = False
        
        ' normalise for the case of the last character of the key
        ' not being included in the sequence of chars allowed for 
        ' a key i.e. for the default key char sequence the key
        ' 'aaa/' becomes 'aaa/0'. This simplifies code later as we 
        ' now don't need to check for the last char of the key 
        ' not being in the allowed key chars.
        If p.AllowedKeyChars.LacksItem(myKeyAsChars.Last) Then myKeyAsChars.Add(p.AllowedKeyChars.First)

        Do
            myCarryFlag = False

            If myKeyAsCharsCurrentIndex < 0 Then
            
                myKeyAsChars.InsertAt(0, p.AllowedKeyChars.Item(1))
                Exit Do
                
            Else
            
                If p.AllowedKeyChars.LacksItem(myKeyAsChars.Item(myKeyAsCharsCurrentIndex)) Then
            
                    myKeyAsChars.InsertAt(myKeyAsCharsCurrentIndex + 1, p.AllowedKeyChars.Item(1))
                    Exit Do
                    
                
                ElseIf myKeyAsChars.Item(myKeyAsCharsCurrentIndex) = p.AllowedKeyChars.Last Then
                
                    myKeyAsChars.Item(myKeyAsCharsCurrentIndex) = p.AllowedKeyChars.Item(0)
                    myCarryFlag = True
                    myKeyAsCharsCurrentIndex = myKeyAsCharsCurrentIndex - 1
                    
                    
                Else
                
                    Dim myNextAllowedKeyCharIndex As String
                    myNextAllowedKeyCharIndex = p.AllowedKeyChars.IndexOf(myKeyAsChars.Item(myKeyAsCharsCurrentIndex)) + 1
                    
                    myKeyAsChars.Item(myKeyAsCharsCurrentIndex) = p.AllowedKeyChars.Item(myNextAllowedKeyCharIndex)
                    Exit Do
                    
                    
                End If
                
            End If
            
        Loop While myCarryFlag

        Inc = myKeyAsChars.ToString(vbNullString)
        
    End Function


