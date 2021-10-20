VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lyst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A Wrapper for the ArrayList object which provides intellisense and extends functionality"
'@IgnoreModule FunctionReturnValueAlwaysDiscarded, SuspiciousPredeclaredInstanceAccess
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
'@ModuleDescription("A Wrapper for the ArrayList object which provides intellisense and extends functionality")
'@PredeclaredId
'@exposed
'@Folder("VBALib")


Private Type InternalConstants
    
    Location                    As String
    
End Type

Private i                       As InternalConstants


Public Enum RangeStatus
    
    
    NoError = 0
    
    StartIndexAfterLastIndex = -3
    RangeOfZero = -4
    RangeExceededLastIndex = -5
    CountExceededLastIndex = -6
    NegativeCount = -7
    
End Enum

Private Type State

    Host                                    As ArrayList
    RangeStatus                             As RangeStatus
    
End Type

Private s                               As State


Private Sub Class_Initialize()
    
    If VBA.Len(i.Location) = 0 Then
        
        i.Location = "VBALib.Lyst."
        
    End If
    
    'Guard InvalidNew, Not (Me Is Lyst), i.Location & "Class_Initialize", TypeName(Me)
    
End Sub

'@Description("Returns a new instance of the Lyst Class. Optionally populates the instance with the result of For Each applied to ipIterable.")
Public Function Deb(Optional ByVal ipIterable As Variant = Empty) As Lyst
Attribute Deb.VB_Description = "Returns a new instance of the Lyst Class. Optionally populates the instance with the result of For Each applied to ipIterable."
    
    Guard UsePredeclaredIdInstance, Not (Me Is Lyst), i.Location & "Deb"
    
    If Not VBA.IsEmpty(ipIterable) Then
        
        Guard InvalidIterable, Types.Group.Iterables.LacksItem(ipIterable), i.Location & "Deb"
        
        
    End If
    
    With New Lyst
    
        Set Deb = .ReadyToUseInstance(ipIterable)
        
        
    End With
    
End Function

'@Description("For internal use only.  Instantiates a new instance of the lyst class")
Friend Function ReadyToUseInstance(ByVal ipIterable As Variant) As Lyst
Attribute ReadyToUseInstance.VB_Description = "For internal use only.  Instantiates a new instance of the lyst class"
    
    Set ReadyToUseInstance = Me
    Set s.Host = New ArrayList
    
    If VBA.IsEmpty(ipIterable) Then Exit Function
    
    Dim myItem As Variant
    For Each myItem In ipIterable
        
        s.Host.Add myItem
        
        
    Next
        
End Function

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4

    Dim enumerator As IUnknown
    Set enumerator = s.Host.GetEnumerator  '  (0, s.Host.Count)
    Set NewEnum = enumerator

End Function
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Properties
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@Description("Returns the number of Items in the Lyst instance)")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of Items in the Lyst instance)"
    Count = s.Host.Count
End Property

'@Description("The status")
Public Function Status() As RangeStatus
Attribute Status.VB_Description = "The status"
    Status = s.RangeStatus
End Function


'@DefaultMember
'@Description("Allows read and write of an item value")
Public Property Get Item(ByVal ipIndex As Long) As Variant
Attribute Item.VB_Description = "Allows read and write of an item value"
Attribute Item.VB_UserMemId = 0

    Guard ResultCode.IndexLessThanFirstIndex, ipIndex < ArrayListFirstIndex, i.Location & "Set Item"
    Guard ResultCode.IndexMoreThanLastIndex, ipIndex > Me.LastIndex, i.Location & "Get Item"
    Dim myList As ArrayList
    Set myList = s.Host

        If VBA.IsObject(myList.Item(ipIndex)) Then
        
            Set Item = s.Host.Item(ipIndex)
            
        Else
        
            Item = s.Host.Item(ipIndex)
            
        End If
    
End Property

Public Property Let Item(ByVal ipIndex As Long, ByVal ipValue As Variant)
    
    Guard ResultCode.IndexLessThanFirstIndex, ipIndex < ArrayListFirstIndex, i.Location & "Set Item"
    Guard ResultCode.IndexMoreThanLastIndex, ipIndex > Me.LastIndex, i.Location & "Set Item"
    
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
' Lyst cannot use Is<Group> methos of the TypesGroup class as this leads to using
' Lysts of typenames before they have been populated
' thus in the Lyst class we use the pcLacksItem and pvHoldsItem Methods
' Public Function pvHoldsItem(ByVal ipGroupString As String, ByVal ipItem As Variant) As Boolean
'     pvHoldsItem = VBA.InStr(ipGroupString, TypeName(ipItem)) > 0
' End Function

' Public Function pvLacksItem(ByVal ipGroupString As String, ByVal ipItem As Variant) As Boolean
'     pvLacksItem = Not pvHoldsItem(ipGroupString, ipItem)
' End Function

'@Description("Adds a comma separated list of items to the Lyst.  )")
Public Function Add(ParamArray ipArgs() As Variant) As Lyst
Attribute Add.VB_Description = "Adds a comma separated list of items to the Lyst.  )"

    Set Add = Me
    If Arrays.IsEmptyParamArray(ipArgs) Then Exit Function
        
    Dim myItem As Variant
    For Each myItem In ipArgs
            
        s.Host.Add myItem
    
    Next
    
End Function


Friend Function AddValidatedItem(ByVal ipItem As Variant) As Lyst
    
    s.Host.Add ipItem
    Set AddValidatedItem = Me
    
End Function


'@Description("Implements add for a single item produced by ParseParamArrayToLyst<test> thus does not retest via ParseParaArray... thereby avoiding an infinite recursion")
Friend Function AddValidatedIterable(ByVal ipIterable As Variant) As Lyst
Attribute AddValidatedIterable.VB_Description = "Implements add for a single item produced by ParseParamArrayToLyst<test> thus does not retest via ParseParaArray... thereby avoiding an infinite recursion"

    Dim myItem As Variant
    For Each myItem In ipIterable
        
        s.Host.Add myItem
        
    Next
    
    Set AddValidatedIterable = Me
    
End Function


'@Description("Adds the Items of a single iterable to the Lyst.")
Public Function AddRange(ByVal ipIterable As Variant) As Lyst
Attribute AddRange.VB_Description = "Adds the Items of a single iterable to the Lyst."
' The arraylist equivalent requires that ipiterable supports ICollection
    
    Guard InvalidIterable, Types.Group.IsNotIterable(ipIterable), i.Location & "AddRange"
    
    Set AddRange = Me
    
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

'@Description("Removes all items from the Lyst")
Public Function Clear() As Lyst
Attribute Clear.VB_Description = "Removes all items from the Lyst"
    
    Set Clear = Me
    s.Host.Clear
    
End Function

'@Description("Returns a shallow copy of the Lyst")
Public Function Clone() As Lyst
Attribute Clone.VB_Description = "Returns a shallow copy of the Lyst"
    Set Clone = Lyst.Deb.AddValidatedIterable(s.Host)
End Function

'@Description("Returns True if the Item is found in the Lyst. For compatibility with ArrayList syntax")
Public Function Contains(ByVal ipItem As Variant) As Boolean
Attribute Contains.VB_Description = "Returns True if the Item is found in the Lyst. For compatibility with ArrayList syntax"
    Contains = s.Host.Contains(ipItem)
End Function

'@Description("Returns True if the Item is found in the Lyst. Alternative syntax used within VBALib")
Public Function HoldsItem(ByVal ipValue As Variant) As Boolean
Attribute HoldsItem.VB_Description = "Returns True if the Item is found in the Lyst. Alternative syntax used within VBALib"
    HoldsItem = s.Host.Contains(ipValue) ' Does this need to be typename
End Function

'@Description("Returns True if the Item is not found in the Lyst.  Alternative syntax for 'Not Contains'")
Public Function LacksItem(ByVal ipValue As Variant) As Boolean
Attribute LacksItem.VB_Description = "Returns True if the Item is not found in the Lyst.  Alternative syntax for 'Not Contains'"
    LacksItem = Not HoldsItem(ipValue)
End Function

'@Description("Copies the specified range to the byref array provided as the first parameter. Default is to copy the whole Lyst")
Public Sub CopyTo(ByRef iparray As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipRun As Long = -1)
Attribute CopyTo.VB_Description = "Copies the specified range to the byref array provided as the first parameter. Default is to copy the whole Lyst"

    If ipRun = 0 Then Exit Sub
        
    Guard NotListArray, Arrays.Ranks(iparray) <> 1, i.Location & "CopyTo", TypeName(iparray)
    Guard IndexOutOfBounds, ipStartIndex < 0, i.Location & "CopyTo", "too low"
    Guard IndexOutOfBounds, ipStartIndex >= s.Host.Count, i.Location & "CopyTo", "too high"
   

    '-1 is a flag to signify copy from start index to the end of the lyst
    ' mycount is calculated to allow the guard to be used without being confused by the -1 flag
    ' thus we are appropriately guarding againt ipcount being inappropriately large
    Dim myCount As Long
    '@Ignore IIfSideEffect
    myCount = IIf(ipRun = -1, s.Host.Count - ipStartIndex, ipRun)
    Guard IndexWillOverrun, ipStartIndex + myCount > s.Host.Count, i.Location & "CopyTo"
    
    Dim myArraySize As Long
    myArraySize = Arrays.Count(iparray)
    Guard DestinationTooSmall, myArraySize < myCount, i.Location & "CopyTo", Array(myCount, myArraySize)
    
    Dim myLbound As Long
    myLbound = LBound(iparray)
    
    'todo revisit the code below when twinbasic has getrange options working
    Dim myOffset As Long
    For myOffset = 0 To myCount - 1
    
        iparray(myLbound + myOffset) = s.Host.Item(ipStartIndex + myOffset)
        
    Next
    
End Sub

'@Description("Returns a new Lyst instance containing a shallow copy of the specified range")
Public Function GetRange(Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Lyst
Attribute GetRange.VB_Description = "Returns a new Lyst instance containing a shallow copy of the specified range"

    Dim myLyst As Lyst
    Set myLyst = Deb
    Set GetRange = myLyst
    
    If ipCount = 0 Then Exit Function
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex >= s.Host.Count Then Exit Function
    
    
    Dim myCount As Long
    
    If ipCount < 0 Then
        
        myCount = s.Host.Count
        
 
    ElseIf ipStartIndex + ipCount - 1 >= s.Host.Count Then
        
        myCount = s.Host.Count - ipStartIndex + 1
        
        
    Else
        
        myCount = ipCount
        
    
    End If
    
    Dim myAL As ArrayList
    Set myAL = s.Host.GetRange(ipStartIndex, myCount).Clone
    
    myLyst.InjectHostArrayList myAL
    
End Function
    
'@Description("Returns the Index of the first Item found in the Lyst after the specified starting point.  The default action for count (-1) is to start at index 0 and use the whole list")
Public Function IndexOf(ByVal ipItem As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Long
Attribute IndexOf.VB_Description = "Returns the Index of the first Item found in the Lyst after the specified starting point.  The default action for count (-1) is to start at index 0 and use the whole list"
    ' Updated to allow index of objects to be returned using the ObjPtr method.
    
    IndexOf = -1
   
    
    
    If VBA.IsObject(ipItem) Then
        
        IndexOf = IndexOfObject(ipItem, ipStartIndex, ipCount)
        
        
    Else
        
        IndexOf = IndexOfValue(ipItem, ipStartIndex, ipCount)
        
        
    End If
 
End Function



Friend Function IndexOfValue(ByVal ipValue As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Long
    ' The following caveats apply for using ArrayList from twinbasic
    ' TB/VBA gives an error if the method call does not contain a startindex value
    ' TB/VBA gives an error if the count parameter is used
    ' The Index returned is counted from start of list not start of range
    ' if ipValue is not found  then -1 is returned
    ' ipCount is used to limit the range over which the search is conducted
    
    ' As the count parameter is not supported a subrange is extracted so we can just use Index of(value)
    Dim myRange As ArrayList
    Set myRange = s.Host.GetRange(ipStartIndex, ipCount)
    
    Dim myIndex As Long
    ' Remember that a start index is required in TB/VBA
    myIndex = myRange.IndexOf(ipValue, 0)
    If myIndex = -1 Then Exit Function
    IndexOfValue = myIndex + ipStartIndex
    
End Function


Friend Function IndexOfObject(ByVal ipItem As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Long
    
    IndexOfObject = -1
    Dim myRange As ArrayList
    Set myRange = s.Host.GetRange(ipStartIndex, ipCount)
    
    Dim myIndex As Long
    ' Remember that a start index is required in TB/VBA
    
    For myIndex = 0 To myRange.Count - 1
      
        If VBA.ObjPtr(ipItem) = VBA.ObjPtr(s.Host.Item(myIndex)) Then
            
            IndexOfObject = myIndex + ipStartIndex
            Exit Function
            
            
        End If
        
 
    Next
    
End Function


Public Function InsertAt(ByVal ipIndex As Long, ParamArray ipArgs() As Variant) As Lyst

    Set InsertAt = Me
    If ipIndex < 0 Then Exit Function
    If ipIndex > s.Host.Count Then Exit Function
    If Arrays.HasNoItems(ipArgs) Then Exit Function
    
    Dim myArrayList As ArrayList
    If UBound(ipArgs) = 0 Then
        
        If Types.Group.IsNotIterable(ipArgs(0)) Then
        
            Set myArrayList = Types.Iterable.ToArrayList(ipArgs(0))
            
            
        Else
            
           Set myArrayList = Types.Iterable.ToArrayList(ipArgs(0))
               
        
        End If
        
    Else
        
        Set myArrayList = Types.Iterable.ToArrayList(ipArgs)
        
    End If
 
    If ipIndex = s.Host.Count Then
    
        s.Host.AddRange myArrayList
        
    Else
    
        s.Host.InsertRange ipIndex, myArrayList
        
    End If
    
End Function


Public Function InsertRange(ByVal ipIndex As Long, ByVal ipIterable As Variant) As Lyst

    Set InsertRange = Me
    If ipIndex < 0 Then Exit Function
    If ipIndex > s.Host.Count Then Exit Function
    If Types.Group.IsNotIterableItemsByEnum(ipIterable) Then Exit Function
    
    If ipIndex = s.Host.Count Then
        
            s.Host.AddRange Types.Iterable.ToArrayList(ipIterable)

    Else
   
        s.Host.InsertRange ipIndex, Types.Iterable.ToArrayList(ipIterable)
        
    End If
    
End Function

'@Description("Returns the index of the last item in the Lyst or -1 if the Lyst count is 0")
Public Function LastIndex() As Long
Attribute LastIndex.VB_Description = "Returns the index of the last item in the Lyst or -1 if the Lyst count is 0"
    
    Dim myCount As Long
    myCount = s.Host.Count
    LastIndex = IIf(myCount = 0, -1, myCount - 1)
    
End Function

Public Function LastIndexOf(Optional ByVal ipValue As Variant, Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Long
    
    '@Ignore IIfSideEffect
    LastIndexOf = IIf(Count = 0, 0, Count - 1)
    If VBA.IsMissing(ipValue) Then Exit Function
        
    LastIndexOf = -1
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex >= s.Host.Count Then Exit Function
    If ipStartIndex + ipCount - 1 >= s.Host.Count Then Exit Function
         
    LastIndexOf = 0
    If ipCount = 0 Then Exit Function
    
    
    Dim myCount As Long
    '@Ignore IIfSideEffect
    myCount = IIf(ipCount = -1, s.Host.Count, ipCount + 1) - ipStartIndex
    ' ArrayList.Lastindexof will only accept a value from VBA
    ' so we use GetRange to first extract the range for lastindexof
    ' for the
    ' object,int32 and
    ' object, int32, int32 overloads
    '
    LastIndexOf = s.Host.GetRange(ipStartIndex, myCount).LastIndexOf(ipValue) + ipStartIndex
    
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

Public Function RemoveRange(Optional ByVal ipStartIndex As Long = 0, Optional ByVal ipCount As Long = -1) As Lyst
    
    Set RemoveRange = Me
    If ipCount < 0 Then
        
        s.Host.Clear
        Exit Function
        
        
    End If
    
    If ipCount = 0 Then Exit Function
    If ipStartIndex < 0 Then Exit Function
    If ipStartIndex >= s.Host.Count Then Exit Function
    
    Dim myCount As Long
    
    If ipStartIndex + ipCount - 1 >= s.Host.Count Then
        
        myCount = s.Host.Count - ipStartIndex + 1
        
        
    Else
        
        myCount = ipCount
        
    
    End If
    
    s.Host.RemoveRange ipStartIndex, myCount
    
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
    
    '@Ignore MemberNotOnInterface
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
    If mySize = Sys.Failed Then Exit Function
    If myStartIndex > s.Host.Count - mySize + 1 Then Exit Function
           
    If Types.Group.IsNotIterable(ipIterable) Then
        
        myArrayList.Add ipIterable
        
        
    ElseIf Types.IsTypeByName(Types.Name.OfArrayList, ipIterable) Then
    
        Set myArrayList = ipIterable
        
        
    Else
        
        Set myArrayList = Types.Iterable.ToArrayList(ipIterable)
       
        
    End If

    Dim myItem As Variant
    For Each myItem In myArrayList
        
        's.Host.SetRange myStartIndex, myArrayList
        s.Host.Item(myStartIndex) = myItem
        myStartIndex = myStartIndex + 1
        
        
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

    s.Host.Sort 'ipStartIndex, myCount
        
    
End Function

Public Function ToArray() As Variant
    ToArray = s.Host.ToArray
End Function


Public Function ToArrayList() As ArrayList
    Set ToArrayList = s.Host.Clone
End Function


Friend Function InjectHostArrayList(ByVal ipArrayList As ArrayList) As Lyst
    Set s.Host = ipArrayList.Clone
    Set InjectHostArrayList = Me
End Function


Public Function ToString(Optional ByVal ipSeparator As String = vbNullString) As String

    Stringifier.ResetItemSeparator ipSeparator
    Stringifier.ResetArrayMarkup vbNullString, vbNullString, vbNullString
    Stringifier.ResetObjectMarkup vbNullString, vbNullString, vbNullString
    ToString = Stringifier.ToString(s.Host)  ' myString
    
    
End Function


Public Function First() As Variant
    
   Guard NullReference, s.Host Is Nothing, i.Location & "Last"
    
    If VBA.IsObject(s.Host.Item(0)) Then
        
        Set First = s.Host.Item(0)
        
        
    Else
        
        First = s.Host.Item(0)
        
        
    End If
        
End Function


Public Function Last() As Variant
    
    Guard NullReference, Types.IsSameInstance(s.Host, Nothing), i.Location & "Last"
        
    If VBA.IsObject(s.Host.Item(Me.LastIndex)) Then
        
        Set Last = s.Host.Item(Me.LastIndex)
        
        
    Else
        
        Last = s.Host.Item(Me.LastIndex)
        
        
    End If
        
End Function
