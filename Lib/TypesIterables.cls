VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypesIterables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Lib")
Option Explicit
'@PredeclaredId
'@Exposed

Public Function GetFirstIndex(ByVal ipIterable As Variant) As Long
    GetFirstIndex = Types.Iterable.TryGetFirstIndex(ipIterable).value
End Function


Public Function TryGetFirstIndex(ByVal ipIterable As Variant, Optional ByRef iopResult As ResultLong) As ResultLong

    If iopResult Is Nothing Then Set iopResult = ResultLong.Deb.SetFailure
    Set TryGetFirstIndex = iopResult
    
    If Types.Group.IsNotIterableByIndex(ipIterable) Then
    
        Exit Function
    
    ElseIf Types.Iterable.HasNoItems(ipIterable) Then
    
        Exit Function
        
    ElseIf Types.IsArray(ipIterable) Then
    
        iopResult.SetSuccess LBound(ipIterable)

    Else
     
        iopResult.SetSuccess IIf(Types.IsTypeByName(Types.Name.OfCollection, ipIterable), 1, 0)
        
    End If
    
End Function


Public Function GetLastIndex(ByVal ipIterable As Variant) As Long
    GetLastIndex = Types.Iterable.TryGetLastIndex(ipIterable).value
End Function

Public Function TryGetLastIndex(ByVal ipIterable As Variant, Optional ByRef iopResult As ResultLong) As ResultLong

    If iopResult Is Nothing Then Set iopResult = ResultLong.Deb.SetFailure
    Set TryGetLastIndex = iopResult
    
    'Dim myIndex As Long
    
    If Types.Group.IsNotIterableByIndex(ipIterable) Then
    
        Exit Function
        
    ElseIf Types.Iterable.HasNoItems(ipIterable) Then
    
        Exit Function
    
    ElseIf Types.IsArray(ipIterable) Then

        iopResult.SetSuccess UBound(ipIterable)
        
    Else
        
        iopResult.SetSuccess ipIterable.Count - IIf(Types.IsTypeByName(Types.Name.OfCollection, ipIterable), 0, 1)
        
    End If

End Function


Public Function GetSize(ByVal ipIterable As Variant, Optional ByVal ipRank As Long = 1) As Long
    GetSize = Types.Iterable.TryGetSize(ipIterable, ipRank:=ipRank).value
End Function

'@Description("Returns the count of items in an iterable")
Public Function TryGetSize(ByVal ipIterable As Variant, Optional ByVal ipRank As Long = 1, Optional ByRef iopResult As ResultLong) As ResultLong
Attribute TryGetSize.VB_Description = "Returns the count of items in an iterable"

    If iopResult Is Nothing Then Set iopResult = ResultLong.Deb.SetFailure
    Set TryGetSize = iopResult
    
    If Types.Group.IsNotIterable(ipIterable) Then Exit Function
    
    If Types.IsArray(ipIterable) Then
    
        Types.Arrays.TryGetSize ipIterable, ipRank, iopResult
 
    Else
    
        iopResult.SetSuccess ipIterable.Count
        
    End If
     
End Function


Public Function Sum(ByVal ipIterable As Variant) As Variant
    
    If Types.Group.IsNotIterable(ipIterable) Then
    
        If Types.Group.IsNotNumber(ipIterable) Then Exit Function
        
        Sum = -1
        Exit Function
        
    End If
    
    Dim myItem As Variant
    Dim mySum As Double
    For Each myItem In ipIterable
    
        mySum = mySum + myItem
        
    Next
    
    Sum = mySum
    
End Function


Public Function MinMax(ByVal ipIterable As Variant) As Variant
        
    If Types.Group.IsNotIterable(ipIterable) Then
    
        If Types.Group.IsNotNumber(ipIterable) Then Exit Function
        
        MinMax = Array(ipIterable, ipIterable)
        Exit Function
        
    End If
   
    Dim myItem As Variant
    Dim myMax As Variant
    Dim myMin As Variant
    For Each myItem In ipIterable
        
        If VBA.IsEmpty(myMin) Then
        
            myMax = myItem
            myMin = myItem
            
        End If
        
        If myItem < myMin Then myMin = myItem
        If myItem > myMax Then myMax = myItem
        
    Next

    MinMax = Array(myMin, myMax)
    
End Function



Public Function IsSameSize(ParamArray ipArgs() As Variant) As Boolean
    IsSameSize = pvIsSameSizeByCVar(CVar(ipArgs))
End Function

Public Function IsNotSameSize(ParamArray ipArgs() As Variant) As Boolean
    IsNotSameSize = Not pvIsSameSizeByCVar(CVar(ipArgs))
End Function

Public Function pvIsSameSizeByCVar(ByVal ipArgs As Variant) As Boolean

    pvIsSameSizeByCVar = False
    If UBound(ipArgs) = -1 Then Exit Function
    
    Dim mySize As Long
    mySize = Types.Iterable.GetSize(ipArgs(0))
    
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If mySize <> GetSize(myArg) Then Exit Function
    
    Next
    
    pvIsSameSizeByCVar = True
    
End Function


'@Description("Returns true if each item in the iterable is the same type")
Public Function IsUniform(ByVal ipIterable As Variant) As Boolean
Attribute IsUniform.VB_Description = "Returns true if each item in the iterable is the same type"

    IsUniform = False
    
    If Not Types.Group.IsIterable(ipIterable) Then
    
        Err.Raise _
            17, _
            "VBALib.Types.IsIterableUniform", _
            "Expected an Iterable by value type"
            
        Exit Function
        
    End If
    
    ' Typename is used to avoid issues of vartype getting the type of the default
    ' member if ipIterable(x) is an object with a default member
    Dim myType  As String
    myType = TypeName(ipIterable(0))
    
    Dim myItem As Variant
    For Each myItem In ipIterable
    
        If TypeName(myItem) <> myType Then Exit Function
    
    Next
    
    IsUniform = True

End Function


Public Function IsNotUniform(ByVal iparray As Variant) As Boolean
    IsNotUniform = Not IsUniform(iparray)
End Function



'Public Function LacksItems(ByVal ipArg As Variant, Optional ByVal ipRank As Long = 1, Optional iopResult As ResultLong) As ResultLong
'
'    If iopResult Is Nothing Then Set iopResult = ResultLong.Deb.SetSuccess
'    Set LacksItems = iopResult
'
'    If Types.Group.IsNotIterable(ipArg) Then
'
'        iopResult.SetFailure rcNotIterable
'
'    ElseIf VBA.IsArray(ipArg) Then
'
'        If Types.Arrays.TryGetLbound(ipArg, ipRank, iopResult).Succeeded Then iopResult.SetFailure
'
'    Else
'
'        If ipRank <> 1 Then
'
'            iopResult.SetFailure rcInvalidRank, Fmt.Txt("A rank greater than 1 is not valid for Type '{0}'", TypeName(ipArg))
'
'        Else
'
'            If ipArg.Count <> 0 Then iopResult.SetFailure
'
'        End If
'
'    End If
'
'End Function

'@Description("Returns true if arg is an array and has ubound of 0 or more, or is an object with a count greater than 0")

Public Function HasNoItems(ByVal ipArg As Variant, Optional ByVal ipRank As Long = 1) As Boolean
Attribute HasNoItems.VB_Description = "Returns true if arg is an array and has ubound of 0 or more, or is an object with a count greater than 0"
    HasNoItems = Not HasItems(ipArg, ipRank)
End Function
Public Function HasItems(ByVal ipArg As Variant, Optional ByVal ipRank As Long = 1) As Boolean
 
    ' Uninitialised variables will be objects or arrays
    ' not initialised means
    ' an object that is nothing
    ' an array that gives a ubound error
    ' and array that gives a ubound of -1
    
    HasItems = False
    
    If Types.Group.IsNotIterable(ipArg) Then
    
        Exit Function
        
    ElseIf VBA.IsArray(ipArg) Then
        
        HasItems = Types.Arrays.TryGetLbound(ipArg, ipRank).Status
        
    Else
    
        If ipRank <> 1 Then
        
            Err.Raise _
                17 + vbObjectError, _
                "VBALib.TypesIterables.HasItems", _
                Fmt.Txt("A rank greater than 1 is not valid for Type '{0}'", TypeName(ipArg))
            
        Else
        
            If ipArg.Count > 0 Then HasItems = True
        
        End If

    End If
    
End Function

Public Function ToArray(ParamArray ipArgs() As Variant) As Variant

    Dim myThingList As ArrayList
    Set myThingList = New ArrayList
    ToArray = myThingList.ToArray
    
    If UBound(ipArgs) = -1 Then Exit Function
    
    Dim myThingArray As Variant
    myThingArray = pvGetToThingArray(CVar(ipArgs))
    
    Dim myItem As Variant
    For Each myItem In myThingArray(0)
    
        myThingList.Add myItem
        
    Next
           
    ToArray = myThingList.ToArray
    
End Function


Public Function ToQueue(ParamArray ipArgs() As Variant) As Queue
    
    Dim myThingQueue As Queue
    Set myThingQueue = New Queue
    Set ToQueue = myThingQueue

    If UBound(ipArgs) = -1 Then Exit Function
    
    Dim myThingArray As Variant
    myThingArray = pvGetToThingArray(CVar(ipArgs))
    
    Dim myItem As Variant
    For Each myItem In myThingArray(0)
    
        myThingQueue.enqueue myItem
        
    Next
           
    Set ToQueue = myThingQueue
            
End Function


Public Function ToStack(ParamArray ipArgs() As Variant) As Stack
    
    Dim myThingStack As Stack
    Set myThingStack = New Stack
    Set ToStack = myThingStack

    If UBound(ipArgs) = -1 Then Exit Function
    
    Dim myThingArray As Variant
    myThingArray = pvGetToThingArray(CVar(ipArgs))
    
    Dim myItem As Variant
    For Each myItem In myThingArray(0)
    
        myThingStack.Push myItem
        
    Next
           
    Set ToStack = myThingStack
            
End Function


''@Description("For the rare case where we wish to have a single iterable as a single iterable)
'Public Function SingleItemList(ByVal ipValue As Variant) As VBALib.Lyst
'
'    Dim myList As VBALib.Lyst
'    Set myList = New VBALib.Lyst
'    myList.Add ipValue
'    Set SingleItemList = myList
'End Function

Public Function ToArrayList(ParamArray ipArgs() As Variant) As ArrayList

    Dim myThingArrayList As ArrayList
    Set myThingArrayList = New ArrayList
    Set ToArrayList = myThingArrayList

    If UBound(ipArgs) = -1 Then Exit Function
    
    Dim myThingArray As Variant
    myThingArray = pvGetToThingArray(CVar(ipArgs))
    
    Dim myItem As Variant
    For Each myItem In myThingArray(0)
    
        myThingArrayList.Add myItem
        
    Next
           
    Set ToArrayList = myThingArrayList
            
End Function


Public Function ToLyst(ParamArray ipArgs() As Variant) As Lyst

    Dim myThingLyst As Lyst
    Set myThingLyst = Lyst.Deb
    Set ToLyst = myThingLyst

    If UBound(ipArgs) = -1 Then Exit Function
    
    Dim myThingArray As Variant
    myThingArray = pvGetToThingArray(CVar(ipArgs))
    
    Dim myItem As Variant
    For Each myItem In myThingArray(0)
    
        myThingLyst.Add myItem
        
    Next
           
    Set ToLyst = myThingLyst
            
End Function


Public Function ToCollection(ParamArray ipArgs() As Variant) As Collection
    
    Dim myThingCollection As Collection
    Set myThingCollection = New Collection
    Set ToCollection = myThingCollection

    If UBound(ipArgs) = -1 Then Exit Function
    
    Dim myThingArray As Variant
    myThingArray = pvGetToThingArray(CVar(ipArgs))
    
    Dim myItem As Variant
    For Each myItem In myThingArray(0)
    
        myThingCollection.Add myItem
        
    Next
           
    Set ToCollection = myThingCollection
    
End Function


Public Function pvGetToThingArray(ByVal ipArgs As Variant) As Variant
    
    ' ipArgs has already been triaged for empty
    ' ipArgs is a parameter list from a 'ToThing' method where thing is an iterable array/object
    ' the intent of this function is to process ipArgs to an array
    ' in item 1 of the list.
    ' In the case of a single item being provided to ipargs, and the single item is an
    ' iterable type, the contents of the iterable type are added item by item
    '
    ' In the odd case where a single item is an iterable, which we want to add as
    ' as single item, then that item should be encapsulated into an array
    ' when provided as a parameter
    
    Dim myList As ArrayList
    Set myList = New ArrayList
    
    ' if ipArgs is a single item
    If UBound(ipArgs) = 0 Then
    
        If Types.Group.IsPrimitive(ipArgs(0)) Then
        
            myList.Add Array(ipArgs(0))
            
        ElseIf Types.Group.IsIterableByMethod(ipArgs(0)) Then
            ' captures object like stack and queue
            Dim myarray As Variant
            myarray = ipArgs(0).ToArray
            myList.Add myarray
            
        ElseIf Types.Group.IsIterableKeysByEnum(ipArgs(0)) Then
            ' captures objects like Scripting.Dictionary and Kvp
            myList.Add ipArgs(0).Items
            
        ElseIf Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
            ' Captures arrays and other object that can be enumerated
            myList.Add ipArgs(0)
            
        End If
        
    Else
    
        myList.Add ipArgs
        
    End If
    
    pvGetToThingArray = myList.ToArray

End Function
