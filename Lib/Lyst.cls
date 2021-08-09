VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lyst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' List Class
'
' This class is a wrapper for the ArrayList class
' the principle differences being
'
' 1.    it omits parts of the arraylistt class that cannot be acheved in VBA
' 2.    it return its Me instance so that methods can be chained
' 3.    it allows intellisense for the class
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit
'@PredeclaredId
'@exposed
'@Folder("VBALib")

Private Type State

    Host                                As ArrayList
    
End Type

Private s                               As State


Public Function Deb(Optional ByVal ipIterable As Variant) As Lyst
    
    With New Lyst
    
        Set Deb = .ReadyToUseInstance(ipIterable)
        
    End With
    
End Function


Public Function ReadyToUseInstance(ByVal ipIterable As Variant) As Lyst

    Set ReadyToUseInstance = Me
    Set s.Host = New ArrayList
    Set ReadyToUseInstance = Me
    
    '@Ignore IsMissingOnInappropriateArgument
    If VBA.IsMissing(ipIterable) Then Exit Function
    
    If Types.Group.IsIterableItemsByEnum(ipIterable) Then
    
        Dim myItem As Variant
        For Each myItem In ipIterable
        
            
            s.Host.Add myItem
            
        Next
        
    End If
    
End Function

'@Enumerator
Public Function NewEnum() As IUnknown


    Dim enumerator As IUnknown
    Set enumerator = s.Host.GetEnumerator(0, s.Host.Count)
    Set NewEnum = enumerator

End Function
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Properties
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Property Get Count() As Long
    Count = s.Host.Count
End Property

'@DefaultMember
Public Property Get Item(ByVal ipIndex As Long) As Variant


    If VBA.IsObject(s.Host.Item(ipIndex)) Then
    
        Set Item = s.Host.Item(ipIndex)
        
    Else
    
        Item = s.Host.Item(ipIndex)
        
    End If
    
End Property

Public Property Let Item(ByVal ipIndex As Long, ByVal ipValue As Variant)
    
    If VBA.IsObject(ipValue) Then
    
        Set s.Host.Item(ipIndex) = ipValue
        
    Else
    
        s.Host.Item(ipIndex) = ipValue
        
    End If
    
End Property


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Methods
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Function Add(ByVal ipValue As Variant) As Lyst

    s.Host.Add ipValue
    Set Add = Me
    
End Function



'@Description("Adds the values of the iterable to the Lyst.  The iterable should be of type IterablItemsByEnum")
Public Function AddRange(ByVal ipIterable As Variant) As Lyst
'Attribute AddRange.VB_Description = "Adds the values of the iterable to the Lyst.  The iterable should be of type IterablItemsByEnum"

    Set AddRange = Me
    
    If Types.Group.IsNotIterableItemsByEnum(ipIterable) Then Exit Function
    ' The add range method for ArrayList requires the input range to support
    ' ICollection which doesn't cover variants containing an array
    ' so we just add item by item
    Dim myItem As Variant
    For Each myItem In ipIterable
        
        s.Host.Add myItem
    
    Next

End Function


'**Error**
'Public Function BinarySearch(ByVal ipValue As Variant) As Long
'' Produces Invalid Procedure call or argument
'' Probably needs a comparer which cannot be provided by VBA
'
'    s.Host.Sort
'    BinarySearch = s.Host.BinarySearch(ipValue)
'End Function

Public Function Clear() As Lyst
    
    Set Clear = Me
    s.Host.Clear
    
End Function


Public Function Clone() As Lyst
    Set Clone = Lyst.Deb(s.Host.ToArray)
End Function


Public Function HoldsValue(ByVal ipValue As Variant) As Boolean
    HoldsValue = s.Host.Contains(ipValue)
End Function


Public Function LacksValue(ByVal ipValue As Variant) As Boolean
    LacksValue = Not s.Host.Contains(ipValue)
End Function


Public Sub CopyTo(ByRef iparray As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1)

    If Types.Arrays.Ranks(iparray).value <> 1 Then Exit Sub
    If ipStartIndex < 0 Then Exit Sub
    If ipStartIndex >= s.Host.Count Then Exit Sub
    If ipCount = 0 Then Exit Sub
    If ipStartIndex + ipCount - 1 >= s.Host.Count Then Exit Sub
    
    Dim mycount As Long
    mycount = IIf(ipCount = -1, s.Host.Count - ipStartIndex, ipCount)
    
    Dim myArrayIndex As Long
    myArrayIndex = LBound(iparray)
    
    
    
    Dim myOffset As Long
    For myOffset = 0 To mycount - 1
    
        iparray(myArrayIndex + myOffset) = s.Host.Item(ipStartIndex + myOffset)
        
    Next
    
End Sub


Public Function GetRange(ByVal ipStartIndex As Long, ByVal ipCount As Long) As Lyst

    Dim myList As Lyst
    Set myList = Lyst.Deb
    
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex > s.Host.Count Then Exit Function
    If ipCount <= 0 Then Exit Function
    If ipCount - 1 + ipStartIndex > s.Host.Count Then Exit Function
    
    Dim myIndex As Long
    For myIndex = ipStartIndex To ipCount
    
        myList.Add s.Host.Item(myIndex)
        
    Next
    
    Set GetRange = myList
    
End Function
    
    
Public Function IndexOf(ByVal ipValue As Variant, Optional ByVal ipStartIndex As Long = 0) As Long
    ' The default method exposed to VBA requires a start position for the search
    
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex > s.Host.Count Then Exit Function
    'If ipCount = 0 Then Exit Function
    'If ipCount + 1 + ipStartIndex >= s.Host.Count Then Exit Function
    
    'Dim myCount As Long
    'If ipCount = -1 Then myCount = s.Host.Count - ipStartIndex
    
    IndexOf = s.Host.IndexOf(ipValue, ipStartIndex)
    
End Function


Public Function InsertAt(ByVal ipIndex As Long, ByVal ipValue As Variant) As Lyst

    Set InsertAt = Me
    If ipIndex < 0 Then Exit Function
    If ipIndex > s.Host.Count Then Exit Function
    
    If ipIndex = s.Host.Count Then
    
        s.Host.Add ipValue
        
    Else
    
        s.Host.Insert ipIndex, ipValue
        
    End If
    
End Function


Public Function InsertRange(ByVal ipIndex As Long, ByVal ipIterable As Variant) As Lyst

    Set InsertRange = Me
    
    If Types.Iterable.HasNoItems(ipIterable) Then Exit Function
    If Types.Group.IsNotIterableItemsByEnum(ipIterable) Then Exit Function
    If ipIndex < 0 Then Exit Function
    If ipIndex > s.Host.Count Then Exit Function
    
    Dim myItem As Variant
    
    If ipIndex = s.Host.Count Then
    
        For Each myItem In ipIterable
        
            s.Host.Add myItem
            
        Next
        
    Else
   
        Dim myIndex As Long
        myIndex = ipIndex
        For Each myItem In ipIterable
        
            s.Host.Insert myIndex, myItem
            myIndex = myIndex + 1
            
        Next
        
    End If
    
End Function

 
Public Function LastIndexOf(Optional ByVal ipValue As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Long
    
    LastIndexOf = IIf(Count = 0, 0, Count - 1)
    If VBA.IsMissing(ipValue) Then Exit Function
        
    LastIndexOf = -1
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex >= s.Host.Count Then Exit Function
    If ipStartIndex + ipCount - 1 >= s.Host.Count Then Exit Function
         
    LastIndexOf = 0
    If ipCount = 0 Then Exit Function
    
    
    Dim mycount As Long
    mycount = IIf(ipCount = -1, s.Host.Count, ipCount + 1) - ipStartIndex
    ' ArrayList.Lastindexof will only accept a value from VBA
    ' so we use GetRange to first extract the range for lastindexof
    ' for the
    ' object,int32 and
    ' object, int32, int32 overloads
    '
    LastIndexOf = s.Host.GetRange(ipStartIndex, mycount).LastIndexOf(ipValue) + ipStartIndex
    
End Function

Public Function RemoveValue(ByVal ipValue As Variant) As Lyst

    Set RemoveValue = Me
    s.Host.Remove ipValue
    
End Function

Public Function RemoveAt(ByVal ipIndex As Variant) As Lyst

    Set RemoveAt = Me
    If Types.Group.IsNotNumber(ipIndex) Then Exit Function
    
    If ipIndex < 0 Then Exit Function
    If ipIndex >= s.Host.Count Then Exit Function
    s.Host.RemoveAt ipIndex
    
End Function

Public Function RemoveRange(ByVal ipStartIndex As Long, ByVal ipCount As Long) As Lyst
    
    Set RemoveRange = Me
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex >= s.Host.Count Then Exit Function
    If ipCount < 1 Then Exit Function
    If ipStartIndex + ipCount - 1 >= s.Host.Count Then Exit Function
    
    s.Host.RemoveRange ipStartIndex, ipCount
    
End Function

' See Sequence.GetSeries, Sequence.GetConstSeries
''@Description("Adds a repeated value to a Lyst (ArrayList equivalent is Lyst.Deb.Repeat)"
'Public Function Repeat(ByVal ipValue As Variant, ByVal ipCount As Long) As Lyst
'
'    Set Repeat = Me
'    If ipCount < 0 Then Exit Function
'
'    Dim myCount As Long
'    For myCount = 1 To ipCount
'
'        s.Host.Add ipValue
'
'    Next
'
'End Function

Public Function Reverse(Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = 0) As Lyst
    Set Reverse = Me
    
    If ipStartIndex < 0 Then Exit Function
    If ipCount < 0 Then Exit Function
    
    If ipStartIndex = 0 Then
    
        s.Host.Reverse
        Exit Function
        
    End If
    
    Dim myRange As Lyst
    Set myRange = Lyst.Deb(s.Host.GetRange(ipStartIndex, ipCount))
    myRange.Reverse
    s.Host.SetRange ipStartIndex, myRange.ToArrayList
    
    
End Function

Public Function SetRange(ByVal ipStartIndex As Long, ByVal ipIterable As Variant) As Lyst
    
    Dim myStartIndex As Long
    myStartIndex = ipStartIndex
    Set SetRange = Me
    If myStartIndex < 0 Then Exit Function
    If Types.Group.IsNotIterableByIndex(ipIterable) Then Exit Function
    If myStartIndex > s.Host.Count - Types.Iterable.GetSize(ipIterable) + 1 Then Exit Function
    
    If Types.IsTypeByName(Types.Name.OfArrayList, ipIterable) Then
    
        s.Host.SetRange myStartIndex, ipIterable
        
    Else
    
        s.Host.SetRange myStartIndex, Types.Iterable.ToArrayList(ipIterable)
        
    End If

End Function

'@Description("Sorts the whole list. Does not support sorting part of the list")
'ToDo: implement the missing ArrayList sort overloads
' 1 sort from index to end
' 2 sort from index through count
Public Function Sort(Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Lyst

    
    Set Sort = Me
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex >= s.Host.Count Then Exit Function
    If ipCount = 0 Then Exit Function
    
'    Dim myCount As Long
'    myCount = IIf(ipCount = -1, s.Host.Count - ipStartIndex, ipCount)
    
    
   ' If ipStartIndex + myCount - 1 >= s.Host.Count Then Exit Function
    
'    If ipStartIndex = 0 And ipCount = 0 Then
'
'        s.Host.Sort
'
'    ElseIf ipCount = 0 Then
    
        s.Host.Sort 'ipStartIndex, myCount
        
'    Else
'
'        s.Host.Sort ipStartIndex, ipCount
'
'    End If
    
End Function

Public Function ToArray() As Variant
    ToArray = s.Host.ToArray
End Function

Public Function ToArrayList() As ArrayList
    Set ToArrayList = s.Host.Clone
End Function

Public Function ToString(Optional ByVal ipSeparator As String = Char.comma) As String

'    Dim myString As String
'
'    Dim myItem As Variant
'    For Each myItem In s.Host
'
'        'ToDo: Update cstr to use stringifier
'        myString = myString & IIf(VBA.Len(myItem) = 0, CStr(myItem), Char.comma & CStr(myItem))
'
'    Next
    Stringifier.SetItemSeparator ipSeparator
    ToString = Stringifier.ToString(s.Host)  ' myString
    Stringifier.SetItemSeparator
    
End Function
