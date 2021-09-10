Attribute VB_Name = "Strings"
Option Explicit
'@Folder("VBALib")
    
    
Public Const WhiteSpace                 As String = Char.Space & Char.comma & Char.Period & Char.SemiColon & Char.Colon
Public Const NumberChars                As String = "0123456789"
Public Const LCaseChars                 As String = "abcdefghijklmnopqrstuvwxyz"
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
    
    Dim myIndex As Long
   
    
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

Public Function PadLeft(ByVal ipString As String, ByVal ipWidth As Long, Optional ByVal ipChar As String = char.space) As String

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
Public Function ToSubStrLyst(ByVal ipString As String, Optional ByVal ipSeparator As String = Char.comma, Optional ipTrimChars As String = Whitespace) As Lyst

    Dim myarray As Variant
    myarray = VBA.Split(VBAlib.Strings.Trimmer(ipString), ipSeparator)
    
    Dim myItem As Variant
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    For Each myItem In myarray
    
        myLyst.Add VBAlib.Strings.Trimmer(myItem)
        
    Next

    Set ToSubStrLyst = myLyst
    
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
        Set ToUnicodeBytes = Lyst.Deb(myBytes)
        
    End If
    
End Function

 
Public Function ToCharLyst(ByVal ipString As String) As Lyst

    Dim myLen As Long
    myLen = VBA.Len(ipString)
    
    If myLen = 0 Then
    
        Set ToCharLyst = Lyst.Deb
        Exit Function
        
    End If
    
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Dim myIndex As Long
    For myIndex = 1 To myLen
    
        myLyst.Add VBA.Mid$(ipString, myIndex, 1)
        
    Next
    
    '@Ignore UnassignedVariableUsage
    Set ToCharLyst = myLyst

End Function
