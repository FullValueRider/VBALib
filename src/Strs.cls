VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Strs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule FunctionReturnValueDiscarded, IndexedDefaultMemberAccess
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Strings")
' The strs class is for methods that process **strings**
' To convert to a string, see class 'Stringifier'


'@Description("Converts a string repreentation of a binary number to Variant(Byte/Integer/Long/Longlong))
Public Function BinToNum(ByVal ipBinary As String, Optional ByRef ipSeperator As String = Chars.twLowBar) As Variant

    Dim myBinary As String
    myBinary = VBA.Replace(ipBinary, ipSeperator, twNullStr)
    
    Dim myType As Long
    
    
    Dim myLen As Long
    myLen = VBA.Len(myBinary)

    Dim mySign As Long
    mySign = 0
    
    ' Bytes are unsigned, Integers,Longs and LongLongs are Unsigned
    'Dim myStartIndex As Long
    If myLen < 9 Then
        
        mySign = 1
        myType = 1
        
    ElseIf myLen < 16 Then
    
        mySign = 1
        myType = 2
        
    ElseIf myLen = 16 Then
        mySign = IIf(VBA.Left$(myBinary, 1) = "1", -1, 1)
        myType = 2
        myBinary = VBA.Mid$(myBinary, 2)
        
    ElseIf myLen < 32 Then
        
        mySign = 1
        myType = 3
        
    ElseIf myLen = 32 Then
        
        mySign = IIf(VBA.Left$(myBinary, 1) = "1", -1, 1)
        myType = 3
        myBinary = VBA.Mid$(myBinary, 2)
        
    ElseIf myLen < 64 Then
    
        mySign = 1
        myType = 4
        
    ElseIf myLen = 64 Then
    
        mySign = IIf(VBA.Left$(myBinary, 1) = "1", -1, 1)
        myType = 4
        myBinary = VBA.Mid$(myBinary, 2)
        
    Else
    
        Err.Raise 17 + vbObjectError, _
        "Strs.BinToNum", _
        "Number too big"


    End If
    
    Dim myHost As Variant
    myHost = 0
    Dim myIndex As Long
    For myIndex = 1 To VBA.Len(myBinary)
        
        Dim myChar As String
        myChar = VBA.Mid$(myBinary, myIndex, 1)
        
        myHost = myHost * 2
        If myChar = "1" Then
            
            myHost = myHost + 1
            
        End If
            
    Next

    Select Case myType
        Case 1:     BinToNum = VBA.CByte(myHost)
        Case 2:     BinToNum = VBA.CInt(myHost * mySign)
        Case 3:     BinToNum = VBA.CLng(myHost * mySign)
        Case 4:     BinToNum = VBA.CLngLng(myHost * mySign)
    End Select
        
End Function


Public Function ToChars(ByRef ipString As String) As Variant

    Dim myLen As Long
    myLen = VBA.Len(ipString)
    
    If myLen = 0 Then
        ToChars = Array()
        Exit Function
    End If
    
    Dim myResult As Variant
    ReDim myResult(1 To myLen)
    
    Dim myIndex As Long
    For myIndex = 1 To myLen
        myResult(myIndex) = VBA.Mid$(ipString, myIndex, 1)
    Next
    
    ToChars = myResult
    
End Function


'@Description("Removes duplicates of substrings in ipDedup.")
Public Function Dedup(ByRef ipString As String, Optional ByRef ipDedupItems As Variant = Empty) As String
Attribute Dedup.VB_Description = "Removes duplicates of substrings in ipDedup."
    
    If VBA.Len(ipString) = 0 Then
        Dedup = ipString
        Exit Function
    End If
    
    Dim myDedupItems As Variant
    If VBA.IsEmpty(ipDedupItems) Then
        myDedupItems = Array(Chars.twSpace)
    Else
        myDedupItems = ipDedupItems
    End If
    
    '    If ArrayOp.LacksItems(myDedupItems) Then
    '        myDedupItems = Array(Chars.twSpace)
    '    End If
    
    Dim myString As String
    myString = ipString
    
    Dim mySubStr As Variant
    For Each mySubStr In myDedupItems
        Dim myDedupStr As String
        myDedupStr = mySubStr & mySubStr
        
        Do
            Dim myStringLen As Long
            myStringLen = VBA.Len(myString)
            myString = Replace(myString, myDedupStr, mySubStr)
        Loop Until myStringLen = VBA.Len(myString)
    
    Next
    
    Dedup = myString
          
End Function


'@Description("Removes ipTrimchars from the ends of a string")
Public Function Trimmer(ByRef ipString As String, Optional ByRef ipTrimChars As String = Chars.WhiteSpace, Optional ByVal ipLHS As Boolean = True, Optional ByVal ipRhs As Boolean = True) As String
Attribute Trimmer.VB_Description = "Removes ipTrimchars from the ends of a string"
      
    If VBA.Len(ipString) = 0 Then
        Trimmer = ipString
        Exit Function
    End If
    
    Dim myTrimchars As String
    myTrimchars = IIf(VBA.Len(ipTrimChars) = 0, Chars.WhiteSpace, ipTrimChars)
    
    Dim myString As String
    myString = ipString
    
    '    ' Check for empty SeqC
    '    If myTrimchars.LacksItems Then
    '        Set myTrimchars = ToChars(Chars.WhiteSpace)
    '    End If
     
    'Trim the RHS of string
    If ipRhs Then
        Do While VBA.Len(myString) > 1
            'Debug.Print Asc(VBA.Right$(myString, 1))
            If InStr(myTrimchars, VBA.Right$(myString, 1)) = 0 Then
                Exit Do
            End If
            myString = VBA.Left$(myString, VBA.Len(myString) - 1)
        Loop
    End If
    
    
    'Trim the LHS of string
    If ipLHS Then
        Do While VBA.Len(myString) > 1
            If InStr(myTrimchars, VBA.Left$(myString, 1)) = 0 Then
                Exit Do
            End If
            myString = VBA.Mid$(myString, 2)
        Loop
    End If
    
    Trimmer = myString
          
End Function


'@Description("Adds padding to the RHS of a string to achieve the desired width of string"
Public Function PadRight(ByVal ipString As String, ByVal ipWidth As Long, Optional ByVal ipPadChars As String) As String

    Dim myReturn As String
    myReturn = ipString
    
    If VBA.Len(ipString) >= ipWidth Then
        PadRight = myReturn
        Exit Function
    End If
        
    Dim myPadchars As String
    myPadchars = IIf(VBA.Len(ipPadChars) = 0, Chars.twSpace, ipPadChars)
    
    Do While VBA.Len(myReturn) < ipWidth
        myReturn = myReturn & myPadchars
    Loop

    PadRight = VBA.Left$(myReturn, ipWidth)
          
End Function


'@Description("Adds padding to the LHS of a string to achieve the desired width of string"
Public Function PadLeft(ByVal ipString As String, ByVal ipWidth As Long, Optional ByVal ipPadChars As String) As String
    
    Dim myReturn As String
    myReturn = ipString
    
    If VBA.Len(ipString) >= ipWidth Then
        PadLeft = myReturn
        Exit Function
    End If
        
    Dim myPadchars As String
    myPadchars = IIf(VBA.Len(ipPadChars) = 0, Chars.twSpace, ipPadChars)
    
    Do While VBA.Len(myReturn) < ipWidth
        myReturn = myPadchars & myReturn
    Loop

    PadLeft = VBA.Right$(myReturn, ipWidth)
        
End Function


'@Description("Counts the ocurrence of a non overlapping substring")
Public Function Countof(ByVal ipCountIn As String, ByVal ipCountMe As String) As Long
Attribute Countof.VB_Description = "Counts the ocurrence of a non overlapping substring"

    If VBA.Len(ipCountMe) = 0 Then
        Countof = -1
        Exit Function
    End If
    
    If VBA.Len(ipCountIn) = 0 Then
        Countof = -1
        Exit Function
    End If
  
    Countof = (VBA.Len(ipCountIn) - VBA.Len(VBA.Replace(ipCountIn, ipCountMe, vbNullString))) / VBA.Len(ipCountMe)
    
End Function


'@Description("Splits a string into a SeqC of trimmed substrings")
Public Function ToSubStr _
( _
ByRef ipString As String, _
Optional ByRef ipSeparator As String = Chars.twDefaultSeparator, _
Optional ByRef ipTrimChars As String = twNullStr _
) As Variant
Attribute ToSubStr.VB_Description = "Splits a string into a SeqC of trimmed substrings"

    If VBA.Len(ipString) = 0 Then
        Set ToSubStr = Array()
        Exit Function
    End If
    
    Dim myTrimchars As String
    If VBA.Len(ipTrimChars) = 0 Then
        myTrimchars = WhiteSpace
    Else
        myTrimchars = ipTrimChars
    End If
    
    Dim myArray As Variant
    myArray = VBA.Split(ipString, ipSeparator)
    
    Dim myA As Variant
    ReDim myA(1 To UBound(myArray) - LBound(myArray) + 1)
        
    Dim myIndex As Long
    For myIndex = LBound(myArray) To UBound(myArray)
        myA(myIndex + 1) = Trimmer(VBA.CStr(myArray(myIndex)), myTrimchars)
    Next
    
    ToSubStr = myA
      
End Function


'@Description("Creates a string composed of ipCount repetitions of ipString")
Public Function Repeat(ByRef ipString As String, ByRef ipCount As Long) As String
Attribute Repeat.VB_Description = "Creates a string composed of ipCount repetitions of ipString"

    If VBA.Len(ipString) = 0 Then
        Repeat = vbNullString
        Exit Function
    End If
    
    If ipCount < 0 Then
        Repeat = vbNullString
        Exit Function
    End If
    
    Dim myRepeat As Variant
    ReDim myRepeat(1 To ipCount)
    Dim myCount As Long
    For myCount = 1 To ipCount
        myRepeat(myCount) = ipString
    Next
    
    Repeat = Join(myRepeat, Chars.twNullStr)
  
End Function


'@Description("Repeats the replace action until the string no longer changes in lengtth")
Public Function Replacer(ByRef ipString As String, Optional ByRef ipFind As String = Chars.twSpace, Optional ByVal ipReplace As String = Chars.twNullStr) As String
Attribute Replacer.VB_Description = "Repeats the replace action until the string no longer changes in lengtth"
  
    If VBA.Len(ipString) = 0 Then
        Replacer = ipString
        Exit Function
    End If
    
    If VBA.Len(ipFind) = 0 Then
        ipFind = Chars.twSpace
    End If
    
    Dim myReturn As String
    myReturn = ipString
    Do
    
        DoEvents
        
        Dim myLen As Long
        myLen = VBA.Len(myReturn)
        myReturn = VBA.Replace(myReturn, ipFind, ipReplace)
        
    Loop Until myLen = VBA.Len(myReturn)

    Replacer = myReturn
      
End Function


'@Description("Applies multiple, repeat find and replace actions using  string pairs either of the form Array(Array(Find,Repace),Array(Find,Replace)...) or Dictionary of Find vs Replace")
Public Function MultiReplacer(ByRef ipString As String, ParamArray ipFindReplacePairs() As Variant) As String
Attribute MultiReplacer.VB_Description = "Applies multiple, repeat find and replace actions using  string pairs either of the form Array(Array(Find,Repace),Array(Find,Replace)...) or Dictionary of Find vs Replace"

    If VBA.Len(ipString) = 0 Then
        MultiReplacer = ipString
        Exit Function
    End If
    
    Dim myFindReplacePairs As Variant: myFindReplacePairs = ArrayOp.Splat(VBA.CVar(ipFindReplacePairs))
    
    If ArrayOp.HoldsOneItem(myFindReplacePairs) Then
        If GroupInfo.IsItemByKey(myFindReplacePairs(0)) Then
            Set myFindReplacePairs = myFindReplacePairs(0)
        End If
    End If
    
    Dim myFRPairs() As Variant
    
    Select Case GroupInfo.Id(VBA.CVar(myFindReplacePairs))
    
        Case m_Admin
            myFRPairs = Array(Array(Chars.twSpace, vbNullString))
    
        Case m_array
            If ArrayOp.LacksItems(myFindReplacePairs) Then
                myFRPairs = Array(Array(Chars.twSpace, Chars.twNullStr))
            Else
'                Fmt.Dbg "{0}", Array(myFindReplacePairs)
'                Debug.Print ArrayOp.IsNotArray(myFindReplacePairs(0))
                If ArrayOp.Count(myFindReplacePairs) = 2 And ArrayOp.IsNotArray(myFindReplacePairs(0)) Then
                    myFRPairs = Array(myFindReplacePairs)
                Else
                    myFRPairs = myFindReplacePairs
                End If
            End If
            
        Case m_ItemByKey
            If myFindReplacePairs.LacksItems Then
                myFRPairs = Array(Array(Chars.twSpace, Chars.twNullStr))
            Else
                myFRPairs = myFindReplacePairs.Pairs
            End If
            
        Case Else
        
            Err.Raise 17 + vbObjectError, _
                "VBALib.Strs.MultiReplace", _
                Fmt.Text("Expecting Array of arrays or ItemByKey.  Got {0}", VBA.TypeName(ipFindReplacePairs))

    End Select
    
    Dim myReturn As String
    myReturn = ipString
    
    Dim myPair As Variant
    For Each myPair In myFRPairs
    
        If Not VBA.IsArray(myPair) Then
            Err.Raise 17 + vbObjectError, _
            "strs.Multireplacer", _
            "Expecting array.  Got " & VBA.TypeName(myPair)
        End If
        
        If LBound(myPair) <> 0 Or UBound(myPair) <> 1 Then
            Err.Raise 17 + vbObjectError, _
            "strs.Multireplacer", _
            "Expecting bounds of 0 to 1. Got bounds" & LBound(myPair) & " to " & UBound(myPair)
        End If
        
        myReturn = Replacer(myReturn, VBA.CStr(myPair(0)), VBA.CStr(myPair(1)))
        
    Next
    
    MultiReplacer = myReturn
  
End Function


Public Function ToAscB(ByVal ipString As String) As Variant
  
    If VBA.Len(ipString) = 0 Then
      
        Set ToAscB = Array()
          
    Else
    
        Dim myLen As Long: myLen = VBA.Len(ipString)
        Dim myA As Variant
        ReDim myA(1 To myLen)
        Dim myIndex As Long
        For myIndex = 1 To myLen
            myA(myIndex) = VBA.AscB(VBA.Mid$(ipString, myIndex, 1))
        Next
    
    End If
  
    ToAscB = myA
  
End Function


Public Function ToUnicodeBytes(ByVal ipString As String) As Variant

    If VBA.Len(ipString) = 0 Then
      
        ToUnicodeBytes = Array()
          
    Else
      
        Dim myvariant() As Byte
        myvariant = ipString
        ToUnicodeBytes = myvariant
          
    End If
      
End Function


Public Function ToUnicodeIntegers(ByVal ipString As String) As Variant

    If VBA.Len(ipString) = 0 Then
        
        Set ToUnicodeIntegers = SeqC.Deb
            
    Else
        
        Dim myLen As Long: myLen = VBA.Len(ipString)
        Dim myA As Variant
        ReDim myA(1 To myLen)
        Dim myIndex As Long
        For myIndex = 1 To myLen
            myA(myIndex) = VBA.AscW(VBA.Mid$(ipString, myIndex, 1))
        Next
        
    End If
  
    ToUnicodeIntegers = myA
  
End Function


Public Function Sort(ByRef ipString As String) As String

    If VBA.Len(ipString) = 0 Then
        Sort = ipString
        Exit Function
    End If
    
    Dim myChars As Variant
    myChars = Strs.ToChars(ipString)
    Sorters.ShakerSortArrayByIndex myChars
    Sort = Join(myChars, vbNullString)
  
End Function


'@Description("increments characters using the SeqCuence in the ipAllowedChar SeqC")
Public Function Inc(ByRef ipString As String, Optional ByVal ipAllowedChars As String) As String
Attribute Inc.VB_Description = "increments characters using the SeqCuence in the ipAllowedChar SeqC"
  
    If VBA.Len(ipString) = 0 Then
        Exit Function
    End If
    
    ' Define the set of characters which can be incremented
    Dim myAllowedChars As String
    If VBA.Len(ipAllowedChars) = 0 Then
        myAllowedChars = Chars.AlphaNumeric
    Else
        myAllowedChars = ipAllowedChars
    End If

    
    ' check if the last character of iopString can be incremented
    ' if not append the second character in myAllowedchars to ipSring
    If InStr(myAllowedChars, Right$(ipString, 1)) = 0 Then
        Inc = ipString & Mid$(myAllowedChars, 2, 1)
        Exit Function
    End If
    
    ' we now need an index into ipString
    Dim myStrIndex As Long
    myStrIndex = VBA.Len(ipString)
    
    Do
    
        ' get the index of the next allowed character
        Dim myAllowedIndex As Long
        myAllowedIndex = InStr(myAllowedChars, VBA.Mid$(ipString, myStrIndex, 1))
        myAllowedIndex = myAllowedIndex + 1
        
        ' if allowed do the increment
        If myAllowedIndex <= VBA.Len(myAllowedChars) Then
            Mid$(ipString, myStrIndex, 1) = VBA.Mid$(myAllowedChars, myAllowedIndex, 1)
            Inc = ipString
            Exit Function
        End If
        
        ' Allowedindex is too big
        ' so at this point we have a roll over situation
        ' the first step is to replace the character at strIndex
        ' with the first allowed character
        
        Mid$(ipString, myStrIndex, 1) = VBA.Mid$(myAllowedChars, 1, 1)
        myStrIndex = myStrIndex - 1
        
        ' we now have three paths
        ' strIndex = 0 - update by prepending second character in allowed chars and exit loop
        ' the character as strindex is incrementable- continue loop
        ' the character at strindex is not incrementable insert second allowed char and exit loop
        
        ' check for strindex =0
        If myStrIndex = 0 Then
            Inc = VBA.Mid$(myAllowedChars, 2, 1) & ipString
            Exit Function
        End If
        
        ' check for non incrementable
        If InStr(myAllowedChars, VBA.Mid$(ipString, myStrIndex, 1)) = 0 Then
            Inc = VBA.Mid$(ipString, 1, myStrIndex) & VBA.Mid$(myAllowedChars, 2, 1) & VBA.Mid$(ipString, myStrIndex + 1)
            Exit Function
        End If
        
    Loop
  
End Function

'@description("Removes the specified number of characters from the LHS and RHS of ipString")
Public Function Inner(ByRef ipString As String, Optional ByRef ipLHS As Long = 1, Optional ByRef ipRhs As Long = 1) As String
Attribute Inner.VB_Description = "Removes the specified number of characters from the LHS and RHS of ipString"

    Inner = vbNullString
    
    Dim myCount As Long: myCount = VBA.Len(ipString)
    If myCount = 0 Then
        Exit Function
    End If
    
    If ipLHS > 0 Then
        If ipLHS + 1 > myCount Then
            Exit Function
        End If
    End If
    
    Dim myStr As String: myStr = VBA.Mid$(ipString, ipLHS + 1)
    '@Ignore AssignmentNotUsed
    myCount = VBA.Len(myStr)
    
    If ipRhs > 0 Then
        If ipRhs + 1 > myCount Then
            Inner = myStr
            Exit Function
        End If
    End If
    
    Inner = VBA.Mid$(myStr, 1, myCount - ipRhs)
            
End Function
