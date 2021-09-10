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
' 2.    it returns its Me instance so that methods can be chained
' 3.    it allows intellisense for the class
'
'ToDo: Convert range methods to use the range functions of arraylist
'      currently there is a net bug which revents this in twinbasic but
'      not in VBA.
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit
'@PredeclaredId
'@exposed
'@Folder("VBALib")


Private Const mcLocation                  As String = "VBALib.Lyst."


Private Type State

    Host                                As ArrayList
    
End Type

Private s                               As State


Private Sub Class_Initialize()
    ' Debug.Print TypeName(Me), TypeName(Lyst)
    ' Debug.Print (Me Is Lyst)
    'Guard InvalidNew, Not (Me Is Lyst), mcLocation & "Class_Initialize", TypeName(Me)
End Sub


Public Function Deb(Optional ByVal ipIterable As Variant) As Lyst
    
    'Guard InvalidIterable, Types.Group.IsNotIterable(ipIterable), Location & "Deb", ipIterable
    Guard UsePredeclaredIdInstance, Not (Me Is Lyst), mcLocation & "Deb"
        
    With New Lyst
    
        Set Deb = .ReadyToUseInstance(ipIterable)
        
    End With
    
End Function


Public Function ReadyToUseInstance(ByVal ipIterable As Variant) As Lyst

    Set ReadyToUseInstance = Me
    Set s.Host = New ArrayList
    'Set ReadyToUseInstance = Me
    
    '@Ignore IsMissingOnInappropriateArgument
    If Types.IsMissing(ipIterable) Then Exit Function
    
    If Types.Group.IsIterableItemsByEnum(ipIterable) Then
    
        Dim myItem As Variant
        For Each myItem In ipIterable
            
            s.Host.Add myItem
            
        Next
        
    End If
    
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

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
Attribute Item.VB_UserMemId = 0

    Guard GuardClause.IndexLessThanFirstIndex, ipIndex < ArrayListFirstIndex, mcLocation & "Set Item"
    Guard GuardClause.IndexMoreThanLastIndex, ipIndex > Me.LastIndex, mcLocation & "Get Item"
    Dim myList As ArrayList
    Set myList = s.Host

        If VBA.IsObject(myList.Item(ipIndex)) Then
        
            Set Item = s.Host.Item(ipIndex)
            
        Else
        
            Item = s.Host.Item(ipIndex)
            
        End If
    
End Property

Public Property Let Item(ByVal ipIndex As Long, ByVal ipValue As Variant)
    
    Guard GuardClause.IndexLessThanFirstIndex, ipIndex < ArrayListFirstIndex, mcLocation & "Set Item"
    Guard GuardClause.IndexMoreThanLastIndex, ipIndex > Me.LastIndex, mcLocation & "Set Item"
    
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


Public Function LacksItem(ByVal ipValue As Variant) As Boolean
    LacksItem = Not s.Host.Contains(ipValue)
End Function

Public Sub CopyTo(ByRef iparray As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1)

    If ipCount = 0 Then Exit Sub
        
    Guard NotSingleDimension, Arrays.Ranks(iparray) <> 1, mcLocation & "CopyTo", TypeName(iparray)
    Guard IndexOutOfBounds, ipStartIndex < 0, mcLocation & "CopyTo", "too low"
    Guard IndexOutOfBounds, ipStartIndex >= s.Host.Count, mcLocation & "CopyTo", "too high"
   

    '-1 is a flag to signify copy from start index to the end of the lyst
    ' mycount is calculated to allow the guard to be used without being confused by the -1 flag
    ' thus we are appropriately guarding againt ipcount being inappropriately large
    Dim myCount As Long
    myCount = IIf(ipCount = -1, s.Host.Count - ipStartIndex, ipCount)
    Guard IndexWillOverrun, ipStartIndex + myCount > s.Host.Count, mcLocation & "CopyTo"
    
    Dim myArraySize As Long
    myArraySize = Arrays.Count(iparray)
    Guard DestinationTooSmall, myArraySize < myCount, mcLocation & "CopyTo", Array(myCount, myArraySize)
    
    Dim myLbound As Long
    myLbound = LBound(iparray)
    
    'todo revisit the code below when twinbasic has getrange options working
    Dim myOffset As Long
    For myOffset = 0 To myCount - 1
    
        iparray(myLbound + myOffset) = s.Host.Item(ipStartIndex + myOffset)
        
    Next
    
End Sub


Public Function GetRange(ByVal ipStartIndex As Long, Optional ByVal ipCount As Long = -1) As Lyst

    Dim myList As Lyst
    Set myList = Lyst.Deb
    Set GetRange = myList
    
    If ipStartIndex < 0 Then Exit Function 'Guard.NegativeIndexing
    If ipStartIndex >= s.Host.Count Then Exit Function ' Guard.OutOfRangeIndexing
    If ipCount - 1 + ipStartIndex > s.Host.Count Then Exit Function 'Guard.IndexingOverrun
    
        'if count is -1 we count from ipstartindex to the end of the list
    If ipCount = -1 Then ipCount = s.Host.count - ipStartIndex
        
    Dim myIndex As Long
    For myIndex = ipStartIndex To ipStartIndex + ipCount - 1
    
        myList.Add s.Host.item(myIndex)
        
    Next
    ' Dim myArrayList As ArrayList
    ' Set myArrayList = s.Host
    ' Set myArrayList = myArrayList.getrange(ipStartIndex, ipCount)' (ipStartIndex, IIf(ipCount = 0, Count, ipCount))
    
    
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
    
    If Types.Iterable.LacksItems(ipIterable) Then Exit Function
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

Public Function LastIndex() As Long
    
    LastIndex = IIf(s.Host.count = 0, 0, s.Host.count - 1)
    
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
    
    Dim myCounter As Long
    For myCounter = ipStartIndex To ipStartIndex + ipCount - 1
        
        s.Host.removeat ipStartIndex
        
    Next
    
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
    
    If ipStartIndex = 0 And ipCount = 0 Then
    
        s.Host.Reverse
        Exit Function
        
    End If
    
    Dim myArrayList As ArrayList
    Set myArrayList = s.Host.GetRange(ipStartIndex, ipCount)
    
    myArrayList.Reverse
    s.Host.SetRange ipStartIndex, myArrayList
    
End Function

Public Function SetRange(ByVal ipStartIndex As Long, ByVal ipIterable As Variant) As Lyst
    
    Dim myStartIndex As Long
    myStartIndex = ipStartIndex
    Set SetRange = Me
    
    Dim myArrayList As ArrayList
    Set myArrayList = New ArrayList
    If myStartIndex < 0 Then Exit Function
    Dim mySize As Long
    mySize = Types.Iterable.Count(ipIterable)
    If mySize = Failed Then Exit Function
    If myStartIndex > s.Host.Count - mySize + 1 Then Exit Function
           
    If Types.Group.IsNotIterable(ipIterable) Then
        
        myArrayList.add ipIterable
        
    ElseIf Types.IsTypeByName(Types.Name.OfArrayList, ipIterable) Then
    
        Set myArrayList = ipIterable
        
    Else
        
        Set myArrayList = Types.Iterable.ToArrayList(ipIterable)
       
    End If

    Dim myItem As Variant
    For Each myItem In myArrayList
        's.Host.SetRange myStartIndex, myArrayList
        s.Host.item(ipStartIndex) = myItem
        ipStartIndex = ipStartIndex + 1
    Next
    
End Function

'@Description("Sorts the whole list. Does not support sorting part of the list")
'ToDo: implement the missing ArrayList sort overloads
' 1 sort from index to end
' 2 sort from index through count
Public Function Sort(Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Lyst
Attribute Sort.VB_Description = "Sorts the whole list. Does not support sorting part of the list"
    
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

Public Function First() As Variant
    
   Guard NullReference, Types.IsSameInstance(s.Host, Nothing), mcLocation & "Last"
    
    If VBA.IsObject(s.Host.item(0)) Then
        
        Set First = s.Host.item(0)
        
    Else
        
        First = s.Host.item(0)
        
    End If
        
End Function

Public Function Last() As Variant
    
    Guard NullReference, Types.IsSameInstance(s.Host, Nothing), mcLocation & "Last"
        
    If VBA.IsObject(s.Host.item(Me.LastIndex)) Then
        
        Set Last = s.Host.item(Me.LastIndex)
        
    Else
        
        Last = s.Host.item(Me.LastIndex)
        
    End If
        
End Function
