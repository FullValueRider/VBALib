VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypesIterables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("VBALib.Types")
'@PredeclaredId
'@Exposed
Option Explicit

Public Enum IterablesConst

    TryFirstIndexResult = 0
    TryLastIndexResult = 0
    
End Enum

Const mcLocation                    As String = "VBALib.TypesIterables."
Const mcNotCountable                As Long = -1


Public Enum IterableRangeInputType
    
    IsAdd = -1          ' Input at last index. No change in order of input. Special Case of Insert
    IsInsert = -2       ' Input at any index/key. No change in order of input
    IsPush = -3         ' Input at any Index/Key. Order of input is reversed
    
    
End Enum

'Private Type State
'
'    Results                As Result
'
'End Type
'
'Private s                 As State

Private Sub Class_Initialize()
    Guard InvalidNew, Types.IsNotSameInstance(Me, TypesIterables), mcLocation & "New"
End Sub


' Public Function GetFirstIndex(ByVal ipIterable As Variant, Optional opResult As ResultLong, Optional ipRank As Long = 1) As Long
    
'     Guard InvalidIterable, Types.Group.IsNotIterable(ipIterable), mcLocation & "GetSize", TypeName(ipIterable)
'     Guard InvalidIterable, Types.Group.IsNotIterableByIndex, mcLocation & "GetSize", TypeName(ipIterable)
'     GetFirstIndex = Types.Iterable.GetFirstIndex(ipIterable, ipRank)
    
' End Function
 
Public Function GetConst(ByVal ipConst As IterablesConst) As IterablesConst
    GetConst = ipConst
End Function

Public Function TryExtent(ByVal ipIterable As Variant, Optional ByRef opResult As Result, Optional ByVal ipRank As Long = ArrayFirstRank) As Result
    ' Returns opResult in opResult and by TryExtent
    If opResult Is Nothing Then Set opResult = opResult.Deb
    Set TryExtent = opResult
    
    If Types.Group.IsNotIterable(ipIterable) Then Exit Function
    
   
    Select Case True
        
        Case VBA.IsArray(ipIterable)
        
            If Arrays.TryExtent(ipIterable, opResult, ipRank).Failed Then Exit Function
                
                
        Case Types.Group.IterableByMethod(ipIterable)
            ' Queues ans stack type iterables don't allow indexes
            opResult.Status = True
            'opResult.Reason = NotIndexed
            opResult.Item(Extent.IsFirstIndex) = Empty
            opResult.Item(Extent.IsLastIndex) = Empty
            opResult.Item(Extent.IsCount) = ipIterable.Count
            
            
        Case Types.Group.IsIterable(ipIterable)
            
            opResult.Item(Extent.IsCount) = ipIterable.Count
            If opResult.Item(Extent.IsCount) = 0 Then Exit Function
            opResult.Status = True
            opResult.Item(Extent.IsFirstIndex) = IIf(Types.IsTypeByName(Types.Name.OfCollection, ipIterable), 1, 0) 'item 0 is firstindex
            opResult.Item(Extent.IsLastIndex) = opResult.Item(Extent.IsCount) - IIf(opResult.Item(Extent.IsFirstIndex) = 0, 1, 0)
            
        
    End Select
    
End Function

'Public Function TryFirstIndex(ByVal ipIterable As Variant, Optional ByRef opResult As Result, Optional ByVal ipRank As Long = ArrayFirstRank) As Result
'
'    Set opResult = opResult.Deb
'    Set TryFirstIndex = opResult
'
'    If Types.Group.IsNotIterable(ipIterable) Then Exit Function
'
'    Dim myResult As Result
'    Select Case True
'
'        Case VBA.IsArray(ipIterable)
'
'            If Arrays.TryLBound(ipIterable, myResult, ipRank).Failed Then Exit Function
'
'
'        Case Types.Group.IsIterableByIndex(ipIterable)
'
'            If ipIterable Is Nothing Then Exit Function
'            If ipIterable.Count = 0 Then Exit Function
'           ' Debug.Print TypeName(ipIterable), Types.Name.OfCollection
'            myResult.Item(0) = IIf(Types.IsTypeByName(Types.Name.OfCollection, ipIterable), 1, 0) 'item 0 is firstindex
'
'        Case Else
'
'            Exit Function
'
'
'    End Select
'
'    TryFirstIndex = True
'
'End Function
'
''@Description("Returns the LBound of valid arrays and the first index of collection objects. Returns -1 for invalid arrays and unpopulaated or uninitialised collection objects")
'Public Function TryLastIndex(ByVal ipIterable As Variant, Optional ByRef opResult As Result, Optional ByVal ipRank As Long = ArrayFirstRank) As Result
'
'    Set opResult = Result.Deb
'    Set TryLastIndex = opResult
'
'
'    'If Types.Group.IsNotIterable(ipIterable) Then Exit Function
'
'    Select Case True
'
'        Case VBA.IsArray(ipIterable)
'
'            Arrays.TryUBound ipIterable, opResult, ipRank
'
'
'        Case Types.Group.IsIterableByIndex(ipIterable)
'
'            Dim myLastIndex As Long
'            myLastIndex = ipIterable.Count
'            If myLastIndex = 0 Then Exit Function
'            ' If mylastindex>0 then we have valid first and last indexes
'            ' so we now adjust myLastIndex to account for the value of first index (0 or 1)
'            Dim myFirstIndexResult As Result
'            TryFirstIndex ipIterable, myFirstIndexResult
'            opResult.Status = True
'            opResult.Item(0) = myLastIndex - IIf(myFirstIndexResult.Item(0) = 0, 1, 0)
'
'
'    End Select
'
'End Function


' Public Function GetSize(ByVal ipIterable As Variant, Optional ByVal ipRank As Long = 1) As Long
    
'     Guard GuardClause.IsNotIterable, Types.Group.IsNotIterable(ipIterable), mcLocation & "GetSize", TypeName(ipIterable)
'     GetSize = Types.Iterable.TryGetSize(ipIterable, ipRank:=ipRank).value
' End Function
'@Description("Returns the number of items in a valid array or colllection object. An unitialised array returns -1. A collection object returns -1 if nothing and 0 if unpoulated0")")
Public Function Count(ByVal ipIterable As Variant, Optional ByVal ipRank As Long = ArrayFirstRank) As Long

    Count = mcNotCountable
    
    'If VBA.IsEmpty(ipIterable) Then Exit Function
    
    Select Case True
            
        Case VBA.IsArray(ipIterable)
        
            If Arrays.HasNoItems(ipIterable) Then Exit Function
        
            Count = Arrays.Count(ipIterable, ipRank)
                
        
        Case Types.Group.IsIterable(ipIterable)
        
            If ipIterable Is Nothing Then Exit Function
            Count = ipIterable.Count
        
    End Select
     
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



Public Function IsSameSize _
( _
    ByRef ipTestIterable As Variant, _
    ByRef ipRefIterable As Variant, _
    Optional ByVal ipTestRank As Long = ArrayFirstRank, _
    Optional ByVal ipRefRank As Long = ArrayFirstRank _
) As Boolean
   
    Dim myLocation As String
    myLocation = mcLocation & "IsSameSize"
    Guard InvalidIterable, Types.Group.IsNotIterable(ipRefIterable, ipTestIterable), myLocation
    
    IsSameSize = False
    
    Dim myRefSize As Long
    Dim myTestSize As Long
   
    myRefSize = Count(ipRefIterable, ipRefRank)
    If myRefSize = Sys.Failed Then Exit Function
        
    myTestSize = Count(ipTestIterable, ipTestRank)
    If myTestSize = Sys.Failed Then Exit Function
    
    IsSameSize = myTestSize = myRefSize
    
End Function
        
        
Public Function IsNotSameSize _
( _
    ipTestIterable As Variant, _
    ipRefIterable As Variant, _
    Optional ByVal ipTestRank As Long = ArrayFirstRank, _
    Optional ByVal ipRefRank As Long = ArrayFirstRank _
) As Boolean
    IsNotSameSize = Not IsSameSize(ipTestIterable, ipRefIterable, ipTestRank, ipRefRank)
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
'        If Arrays.TryGetLbound(ipArg, ipRank, iopResult).Succeeded Then iopResult.SetFailure
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


Public Function HasNoItems(ByVal ipArg As Variant, Optional ByVal ipRank As Long = 1) As Boolean
    HasNoItems = Not HasItems(ipArg, ipRank)
End Function

'@Description("Returns true if arg is an array and has ubound of 0 or more, or is an object with a count greater than 0")
Public Function HasItems(ByVal ipArg As Variant, Optional ByVal ipRank As Long = 1) As Boolean
Attribute HasItems.VB_Description = "Returns true if arg is an array and has ubound of 0 or more, or is an object with a count greater than 0"
 
    ' Uninitialised variables will be objects or arrays
    ' not initialised means
    ' an object that is nothing
    ' an array that gives a ubound error
    ' and array that gives a ubound of -1
    
    HasItems = False
    
    If Types.Group.IsNotIterable(ipArg) Then Exit Function

    If VBA.IsArray(ipArg) Then
       
       If Arrays.HasNoItems(ipArg) Then Exit Function
        
    Else
    
        
        'Guard IsNotArray, ipRank <> 1, mcLocation & "HasItems", Array(TypeName(ipArg))
        
        If ipArg.Count <= 0 Then Exit Function

    End If
    
    HasItems = True
    
End Function

Public Function ToArray(ParamArray ipArgs() As Variant) As Variant

    ToArray = Empty
    
    If Arrays.HasNoItems(ipArgs) Then Exit Function
        
    Dim myArrayList As ArrayList
    Set myArrayList = New ArrayList
    Dim myItem As Variant
    
    If UBound(ipArgs) > 0 Then
        
        For Each myItem In ipArgs
        
            myArrayList.Add myItem
            
        Next
        
        
    ElseIf VBA.IsArray(ipArgs(0)) Then
        
        ToArray = ipArgs(0)
        Exit Function
        
        
    ElseIf Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
    
        Dim myargs As Variant
        myargs = CVar(ipArgs(0))
        For Each myItem In myargs
            
            myArrayList.Add myItem
            
        Next
        
        
    Else
    
        myArrayList.Add ipArgs(0)
        
        
    End If
    
    '@Ignore MemberNotOnInterface
    ToArray = myArrayList.ToArray
    
End Function



Public Function ToQueue(ParamArray ipArgs() As Variant) As Queue
    
    Dim myQueue As Queue
    Set myQueue = New Queue
    Set ToQueue = myQueue

    If Arrays.HasNoItems(ipArgs) Then Exit Function
    
        Dim myItem As Variant
        If UBound(ipArgs) > 0 Then
            
            For Each myItem In ipArgs
            
                '@Ignore MemberNotOnInterface
                myQueue.EnQueue myItem
                
            Next
            
            Exit Function
            
        End If
         
        ' There is only one argument in ipargs
        If Types.IsTypeByName(Types.Name.OfQueue, ipArgs(0)) Then
            
            Set ToQueue = ipArgs(0).Clone
            Exit Function
            
        End If
        
        If Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
        
            Dim myargs As Variant
            myargs = CVar(ipArgs(0))
            For Each myItem In myargs
                
                '@Ignore MemberNotOnInterface
                myQueue.EnQueue myItem
                
            Next
            
            Exit Function
            
        End If
        
        '@Ignore MemberNotOnInterface
        myQueue.EnQueue ipArgs(0)
            
End Function


Public Function ToStack(ParamArray ipArgs() As Variant) As Stack
    
    Dim myStack As Stack
    Set myStack = New Stack
    Set ToStack = myStack

    If Arrays.HasNoItems(ipArgs) Then Exit Function
    
        Dim myItem As Variant
        If UBound(ipArgs) > 0 Then
            
            For Each myItem In ipArgs
            
                '@Ignore MemberNotOnInterface
                myStack.Push myItem
                
            Next
            
            Exit Function
            
        End If
         
        ' There is only one argument in ipargs
        If Types.IsTypeByName(Types.Name.OfStack, ipArgs(0)) Then
            
            Set ToStack = ipArgs(0).Clone
            Exit Function
            
        End If
        
        If Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
        
            Dim myargs As Variant
            myargs = CVar(ipArgs(0))
            For Each myItem In myargs
                
                '@Ignore MemberNotOnInterface
                myStack.Push myItem
                
            Next
            
            Exit Function
            
        End If
        
        '@Ignore MemberNotOnInterface
        myStack.Push ipArgs(0)
            
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

    Dim myAL As ArrayList
    Set myAL = New ArrayList
    Set ToArrayList = myAL
   
    If Arrays.HasNoItems(ipArgs) Then Exit Function
     
    Dim myInterim As ArrayList
    Set myInterim = New ArrayList
    
    Select Case True
        
        Case Arrays.IsNotSingleItem(ipArgs)
        
            myInterim.Add CVar(ipArgs)
    
    
        Case Types.Group.IsNotIterable(ipArgs(0))
        
            myInterim.Add CVar(ipArgs)
            
            
        Case Arrays.IsTableArray(ipArgs(0))
                
                myInterim.Add Arrays.ToLystOfRankLysts(ipArgs(0), RankIsRowFirstItemActionIsNoAction)
                
            
         Case Arrays.IsMDArray(ipArgs(0))
                
                myInterim.Add ipArgs(0)
                     
        
        Case Types.Group.IsIterableItemsByEnum(ipArgs(0))
        
            myInterim.Add ipArgs(0)
            
            
        Case Types.Group.IsIterableKeysByEnum(ipArgs(0))
        
            myInterim.Add ipArgs(0).Items
            myInterim.Add ipArgs(0).Keys
        
        
        Case Else
        
            Guard ResultCode.UnexpectedItemInParamArray, True, mcLocation & "TryParseParamArray"
            Exit Function
            
        
    End Select
    
    Dim myItem As Variant
    For Each myItem In myInterim.Item(0)
        
        myAL.Add myItem
        
    Next
            
End Function

'@Description("Converts input to Lyst. If ipArgs is one iterable then the individual items of the iterable are added. Thus to add one iterable as an iterable, encapsulate it in an array")
Public Function ToLyst(ParamArray ipArgs() As Variant) As Lyst
Attribute ToLyst.VB_Description = "Converts input to Lyst. If ipArgs is one iterable then the individual items of the iterable are added. Thus to add one iterable as an iterable, encapsulate it in an array"
    
    Dim myresult As ParseResult
    Set myresult = ParamArrayParser.Parse(ipArgs)
        
    Set ToLyst = myresult.Items
     
End Function


Public Function ToCollection(ParamArray ipArgs() As Variant) As Collection
    
    Dim myCollection As Collection
    Set myCollection = New Collection
    Set ToCollection = myCollection

    If Arrays.HasNoItems(ipArgs) Then Exit Function
    
        Dim myItem As Variant
        If UBound(ipArgs) > 0 Then
            
            For Each myItem In ipArgs
            
                myCollection.Add myItem
                
            Next
            
            Exit Function
            
        End If
         
        ' There is only one argument in ipargs
        ' collection doesn't offer the clone method so no quick copy here
        If Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
        
            Dim myargs As Variant
            myargs = CVar(ipArgs(0))
            For Each myItem In myargs
                
                myCollection.Add myItem
                
            Next
            
            Exit Function
            
        End If
        
        myCollection.Add ipArgs(0)
    
End Function


' Public Function pvGetToThingArray(ByVal ipArgs As Variant) As Variant
    
'     ' ipArgs has already been triaged for empty
'     ' ipArgs is a parameter list from a 'ToThing' method where thing is an iterable array/object
'     ' the intent of this function is to process ipArgs to an array
'     ' in item 1 of the list.
'     ' In the case of a single item being provided to ipargs, and the single item is an
'     ' iterable type, the contents of the iterable type are added item by item
'     '
'     ' In the odd case where a single item is an iterable, which we want to add as
'     ' as single item, then that item should be encapsulated into an array
'     ' when provided as a parameter
    
'     Dim myList As ArrayList
'     Set myList = New ArrayList
    
'     ' if ipArgs is a single item
'     If UBound(ipArgs) = 0 Then
    
'         If Types.Group.IsPrimitive(ipArgs(0)) Then
        
'             myList.Add Array(ipArgs(0))
            
'         ElseIf Types.Group.IsIterableByMethod(ipArgs(0)) Then
'             ' captures object like stack and queue
'             Dim myarray As Variant
'             myarray = ipArgs(0).ToArray
'             myList.Add myarray
            
'         ElseIf Types.Group.IsIterableKeysByEnum(ipArgs(0)) Then
'             ' captures objects like Scripting.Dictionary and Kvp
'             myList.Add ipArgs(0).Items
            
'         ElseIf Types.Group.IsIterableItemsByEnum(ipArgs(0)) Then
'             ' Captures arrays and other object that can be enumerated
'             myList.Add ipArgs(0)
            
'         End If
        
'     Else
    
'         myList.Add ipArgs
        
'     End If
    
'     pvGetToThingArray = myList.ToArray

' End Function
