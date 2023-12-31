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

Private Type Properties

    FirstKey                                As Variant
    CurrentKey                              As Variant
'    FirstUse                                As Boolean
    KeyAllowedChars                         As Lyst
    KeyAsChars                              As Lyst
    FirstKeyType                            As String
End Type

Private p                                   As Properties

'Private Sub Class_Initialise()
'
'    ' If Me Is AutoKeyByString Then
'
'    '     Debug.Print "Me is PredeclaredId"
'
'    ' Else
'
'    '     Debug.Print "Other instance"
'
'    ' End If
'
'End Sub

Public Function Deb _
( _
    Optional ByVal ipFirstKey As String = AutoKeyAid.DefaultInitialStringKey, _
    Optional ByVal ipKeyCharList As String = AutoKeyAid.DefaultAllowedCharListForStringKey _
) As AutoKeyByString

    With New AutoKeyByString
    
        Set Deb = .ReadyToUseInstance(ipFirstKey, ipKeyCharList) ' Why is ipKeyCharsequence a null string
        
    End With
    
End Function

Public Function ReadyToUseInstance _
( _
    ByVal ipFirstKey As String, _
    ByVal ipKeyCharList As String _
) As AutoKeyByString

    p.FirstKey = IIf(VBA.Len(ipFirstKey) = 0, AutoKeyAid.DefaultInitialStringKey, ipFirstKey)
    Set p.KeyAllowedChars = VBALib.Strings.ToCharLyst(IIf(VBA.Len(ipKeyCharList) = 0, AutoKeyAid.DefaultAllowedCharListForStringKey, ipKeyCharList))
    
    p.CurrentKey = Empty
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


Public Property Get KeyCharList() As Variant
    KeyCharList = p.KeyAllowedChars.ToString(vbNullString)
End Property

Public Property Let KeyCharList(ByVal ipValue As Variant)
    KeyCharList = ipValue
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


Public Function Clone() As AutoKeyByString
    ' we don't use current key as the first key in the clone becuase
    ' this will already exist in the Host ordered Dictionary

    Dim myAutoKey As AutoKeyByString
    Set myAutoKey = Me.Deb(p.FirstKey, Join(p.KeyAllowedChars.ToArray, vbNullString))
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
        
    Dim mycounter As Long
    For mycounter = 1 To ipCount
        
        myList.Add GetNextKey
        
    Next
    
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
        
        Dim myKeyAsCharList As Lyst
        Set myKeyAsCharList = VBALib.Strings.ToCharLyst(ipKey)
        
        ' Dim myKeyAsCharsLastIndex As Long
        ' myKeyAsCharsLastIndex = p.AllowedKeyChars.LastIndex
        
        ' The key is read last index to first index so start we need the index of
        ' the last character of myKeyAsChars
        Dim myKeyAsCharsCurrentIndex As Long
        myKeyAsCharsCurrentIndex = myKeyAsCharList.LastIndex
        
        ' A flag to indicate that the current character in myKeyChars as been
        ' rolled over from the last char in the allowed char sequence to the
        ' first char in the allowed char sequence
        Dim myCarryFlag As Boolean
        'myCarryFlag = False
        
        ' normalise for the case of the last character of the key
        ' not being included in the sequence of chars allowed for
        ' a key i.e. for the default key char sequence the key
        ' 'aaa/' becomes 'aaa/0'. This simplifies code later as we
        ' now don't need to check for the last char of the key
        ' not being in the allowed key chars.
        If p.KeyAllowedChars.LacksItem(myKeyAsCharList.Last) Then myKeyAsCharList.Add p.KeyAllowedChars.First

        Do
            myCarryFlag = False

            If myKeyAsCharsCurrentIndex < 0 Then
            
                myKeyAsCharList.InsertAt 0, p.KeyAllowedChars.Item(1)
                Exit Do
                
            End If
            
            Dim myCurChar As String
            myCurChar = myKeyAsCharList.Item(myKeyAsCharsCurrentIndex)
            If p.KeyAllowedChars.LacksItem(myCurChar) Then
        
                myKeyAsCharList.InsertAt myKeyAsCharsCurrentIndex + 1, p.KeyAllowedChars.Item(1)
                Exit Do
                
            
            ElseIf myKeyAsCharList.Item(myKeyAsCharsCurrentIndex) = p.KeyAllowedChars.Last Then
            
                myKeyAsCharList.Item(myKeyAsCharsCurrentIndex) = p.KeyAllowedChars.Item(0)
                myCarryFlag = True
                myKeyAsCharsCurrentIndex = myKeyAsCharsCurrentIndex - 1
                
                
            Else
            
                Dim myNextAllowedKeyCharIndex As String
                myNextAllowedKeyCharIndex = p.KeyAllowedChars.IndexOf(myKeyAsCharList.Item(myKeyAsCharsCurrentIndex)) + 1
                
                myKeyAsCharList.Item(myKeyAsCharsCurrentIndex) = p.KeyAllowedChars.Item(myNextAllowedKeyCharIndex)
                Exit Do
                
                
            End If
                
            
        Loop While myCarryFlag

        Inc = VBA.Join(myKeyAsCharList.ToArray, vbNullString)
        
    End Function


