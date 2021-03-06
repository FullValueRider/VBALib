Attribute VB_Name = "Strings"
Option Explicit
'@Folder("VBALib")
    
    
Public Const WhiteSpace                 As String = Char.Space & Char.Period & Char.SemiColon & Char.Colon & vbTab & vbCrLf
'@Ignore ConstantNotUsed
Public Const NumberChars                As String = "0123456789"
'@Ignore ConstantNotUsed
Public Const LCaseChars                 As String = "abcdefghijklmnopqrstuvwxyz"
'@Ignore ConstantNotUsed
Public Const UCaseChars                 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


Public Function Dedup(ByVal ipSource As String, ByVal ipDedup As String) As String

    Dim mySource As String
    mySource = ipSource
    Dim MyDedupDedup As String
    MyDedupDedup = ipDedup & ipDedup
    
    Do
    
        Dim myLen As Long
        myLen = Len(mySource)
        mySource = Replace(mySource, MyDedupDedup, ipDedup)
        
    Loop Until myLen = Len(mySource)
    
    Dedup = mySource
    
End Function

Public Function Trimmer(ByVal ipString As String, Optional ByVal ipTrimChars As String = " ,;" & vbCrLf & vbTab) As String

    Dim myString As String
    myString = ipString
    
    'Dim myIndex As Long
   
    
    If VBA.Len(myString) = 0 Then
        
        Trimmer = myString
        Exit Function
        
    End If

    Do While VBA.InStr(ipTrimChars, VBA.Left$(myString, 1)) > 0
            
        DoEvents ' Always put a do event statement in a do loop
        
        myString = VBA.Mid$(myString, 2)
        
    Loop
    
    Do While VBA.InStrRev(ipTrimChars, VBA.Right$(myString, 1)) > 0
            
        DoEvents ' Always put a do event statement in a do loop
        
        myString = VBA.Left$(myString, VBA.Len(myString) - 1)
        
    Loop
        
    Trimmer = myString
    
End Function

Public Function PadLeft(ByVal ipString As String, ByVal ipWidth As Long, Optional ByVal ipChar As String = Char.Space) As String

    If Len(ipString) >= ipWidth Then
    
        PadLeft = ipString
        Exit Function
        
    End If
    
    Dim myReturn As String
    myReturn = VBA.String$(ipWidth, ipChar)
    LSet myReturn = ipString
    PadLeft = myReturn
    
End Function


Public Function PadRight(ByVal ipString As String, ByVal ipWidth As Long, Optional ByVal ipChar As String = " ") As String
    
    If Len(ipString) >= ipWidth Then
    
        PadRight = ipString
        Exit Function
        
    End If
    
    Dim myReturn As String
    myReturn = String$(ipWidth, ipChar)
    RSet myReturn = ipString
    PadRight = myReturn
    
End Function


Public Function Count(ByVal ipString As String, ByVal ipChar As String) As Long
    Count = Len(ipString) - Len(Replace(ipString, ipChar, vbNullString))
End Function



'Public Function NullStrArr(ByVal ipCount As Long) As Variant
'
'    'Dim myKvp As Kvp
'    ' vbNullString gives an EMpty variant, "" gives ""
'    '@Ignore EmptyStringLiteral
'    NullStrArr = Kvp.Deb.Add("", ipCount).GetValues
'
'End Function


'@Description("Takes string in the form of X,Y and returns array containing X Long, Y Long")
Public Function CoordsToXY(ByVal ipCoord As String) As Variant
Attribute CoordsToXY.VB_Description = "Takes string in the form of X,Y and returns array containing X Long, Y Long"
    CoordsToXY = Array(CLng(Split(ipCoord, ",")(0)), CLng(Split(ipCoord, ",")(1)))
End Function

'@Description("Converts a string to an Lyst of trimmed substrings")
Public Function ToSubStrLyst(ByVal ipString As String, Optional ByVal ipSeparator As String = Char.comma, Optional ByVal ipDeleteChars As String = WhiteSpace) As Lyst
Attribute ToSubStrLyst.VB_Description = "Converts a string to an Lyst of trimmed substrings"

    Dim myArray As Variant
    Dim myString As String
    
    myString = Strings.Replacer(ipString, ipDeleteChars)
    If InStr(myString, ipSeparator) = 0 Then
        
        myArray = Array(myString)
        
    Else
        
        myArray = VBA.Split(myString, ipSeparator)
        
    End If
    
    Dim myItem As Variant
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    For Each myItem In myArray
    
        myLyst.Add myItem
        
    Next

    Set ToSubStrLyst = myLyst
    
End Function

Public Function Replacer(ByVal ipString As String, Optional ByVal ipReplaceChars As String = WhiteSpace) As String
    
    Dim myString As String
    myString = ipString
    
    If VBA.Len(ipReplaceChars) = 0 Then Exit Function
    
    'Dim myResult As String
    Dim myIndex As Long
    For myIndex = 1 To Len(ipReplaceChars)
        
       
        myString = VBA.Replace(myString, VBA.Mid$(ipReplaceChars, myIndex, 1), vbNullString)
        
    Next

    Replacer = myString
    
End Function

Public Function ToAnsiBytes(ByVal ipString As String) As Lyst

    If VBA.Len(ipString) = 0 Then
    
        Set ToAnsiBytes = Lyst.Deb
        
    Else
    
        Set ToAnsiBytes = Lyst.Deb(Split(StrConv(ipString, vbFromUnicode)))
        
    End If
    
End Function

Public Function ToUnicodeBytes(ByVal ipString As String) As Lyst

    If VBA.Len(ipString) = 0 Then
    
        Set ToUnicodeBytes = Lyst.Deb
        
    Else
    
        Dim myBytes() As Byte
        myBytes = ipString
        Set ToUnicodeBytes = Lyst.Deb.AddRange(myBytes)
        
    End If
    
End Function

 
Public Function ToCharLyst(ByVal ipString As String) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set ToCharLyst = myLyst
    
    Dim myLen As Long
    myLen = VBA.Len(ipString)
    If myLen = 0 Then Exit Function
    
    Dim myIndex As Long
    For myIndex = 1 To myLen
    
        myLyst.Add VBA.Mid$(ipString, myIndex, 1)
        
    Next
    
    Set ToCharLyst = myLyst

End Function
