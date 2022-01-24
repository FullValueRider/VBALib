VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypesIterables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclarediD
'@eXPOSED
'@Folder("VBALib.Types")


'@Description("Returns the firstindex,lastindex and count in Items.  If the iterable has no items status is false.")
Public Function TryExtent _
( _
    ByVal ipIterable As Variant, _
    Optional ByVal ipRank As Variant _
) As Result
Attribute TryExtent.VB_Description = "Returns the firstindex,lastindex and count in Items.  If the iterable has no items status is false."
    
    'Items should be accessed using Item(ResultItemsEnums.ItemExtent(enum))
    Set TryExtent = Globals.Res.SetNotOkay
    Dim myRank As Long
    If VBA.IsMissing(ipRank) Then myRank = enums.Preset.Value(PresetId.ArrayFirstRank): Else myRank = ipRank
    If BailOut.When(Types.Group.IsNotIterable(ipIterable), Msg.IsNotIterable) Then Exit Function
        
    Select Case True
      
        Case Types.IsString(ipIterable)
        
            Strs.TryExtent ipIterable, myRank
                

        Case VBA.IsArray(ipIterable)
            'ToDo: Does not yet deal with dimensions of more than one.
            Arrays.TryExtent ipIterable, myRank
            

        Case Types.Group.IsIterableObject(ipIterable)
            ' Queues and stack type iterables don't allow indexes
                    
            If BailOut.When(Types.Iterable.HasNoItems(ipIterable), Msg.IterableHasNoItems) Then Exit Function
            If BailOut.When(myRank <> 1, Msg.ItemDoesNotSupportRanks) Then Exit Function
            
            If Types.Group.IsIterableByMethod(ipIterable) Then
                
                Globals.Res.SetOkay(Msg.DoesNotIterateByIndex).SetItems Empty, Empty, ipIterable.Count
                
            ElseIf Types.Group.IsIterableOneBased(ipIterable) Then
                    
                Globals.Res.SetOkay.SetItems 1, ipIterable.Count, ipIterable.Count
                
                
            ElseIf Types.Group.IsIterableItemsByEnum(ipIterable) Then
                            
                Globals.Res.SetOkay.SetItems 0, ipIterable.Count - 1, ipIterable.Count
                
                    
            End If
            
        Case Else
        
            Globals.Res.SetNotOkay enums.Message.AsEnum(Msg.InvalidType)
            
            
    End Select
        
Set TryExtent = Globals.Res
    
End Function

'@Description("Returns the number of items in a valid array or iterable object. An unitialised array returns -1. A collection object returns -1 if nothing and 0 if unpoulated0")")
Public Function Count(ByVal ipIterable As Variant, Optional ByVal ipRank As Variant) As Long

    Count = -1
    Dim myRank As Long
    If VBA.IsMissing(ipRank) Then myRank = enums.Preset.Value(PresetId.ArrayFirstRank): Else myRank = ipRank
    
    If TryExtent(ipIterable, myRank).IsNotOkay Then Exit Function
        
    Count = Globals.Res.Item(ResultItemsEnums.ItemExtent(ieCount))

End Function

'@Description("Gets  items from an Iterable that does not support itembyindex (e.g. queues and stacks")
Public Function GetByForEach(ByVal ipIterable As Variant, Optional ByVal ipCount As Long = 0) As Lyst
Attribute GetByForEach.VB_Description = "Gets  items from an Iterable that does not support itembyindex (e.g. queues and stacks"
    
    
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    Dim myCount As Long
    Dim myItem As Variant
    For Each myItem In ipIterable
    
        myLyst.Add myItem
        myCount = myCount + 1
        If (ipCount > 0) And (myCount > ipCount) Then Exit For
        
    Next
    
    Set GetByForEach = myLyst
    
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


'Public Function OfNoUse(ByVal ipForwardedIterable As Variant) As Boolean
'
'    If BailOut.When(Arays.IsNotValid(ipForwardedIterable), Msg.IterableHasNoItems) Then Exit Function
'    If BailOut.When(VBA.IsEmpty(ipForwardedIterable), Msg.IterableIsEmpty) Then Exit Function
'    If BailOut.When(VBA.IsArray(ipForwardedIterable) And (UBound(ipForwardedIterable) < LBound(ipForwardedIterable)), Msg.ArrayHasNoItems) Then Exit Function
'    If BailOut.When(Types.IsNotArray(ipForwardedIterable) And (ipForwardedIterable.Count <= 1), Msg.IterableHasNoItems) Then Exit Function
'    OfNoUse = Not Globals.Res.Init(Msg.Okay).Status
'
'End Function


'@Description("True if the count of items for each iterable is the same, the iterables do not have to be of the same type")
Public Function IsSameSize _
( _
    ByRef ipTestIterable As Variant, _
    ByRef ipRefIterable As Variant, _
    Optional ByVal ipRefRank As Long = 1, _
    Optional ByVal ipTestRank As Long = 1 _
) As Boolean
Attribute IsSameSize.VB_Description = "True if the count of items for each iterable is the same, the iterables do not have to be of the same type"
    IsSameSize = False
    
    If TryExtent(ipTestIterable, ipTestRank).IsNotOkay Then Exit Function
        
    Dim myTestCount As Long
    myTestCount = Globals.Res.Item(ResultItemsEnums.ItemExtent(ieCount))
    
    If TryExtent(ipRefIterable, ipRefRank).IsNotOkay Then Exit Function
    
    Dim myRefCount As Long
    myRefCount = Globals.Res.Item(ResultItemsEnums.ItemExtent(ieCount))
    
    If myTestCount = myRefCount Then
        
        Globals.Res.Init(Msg.okay).SetItems (Array(myTestCount, myRefCount))
    
        
    Else
        
        Globals.Res.SetNotOkay (Msg.notequal)
        
        
    End If
    
    IsSameSize = Globals.Res.Status
End Function


Public Function IsNotSameSize _
       ( _
       ipTestIterable As Variant, _
       ipRefIterable As Variant, _
       Optional ipTestRank As Long = 1, _
       Optional ipRefRank As Long = 1 _
       ) As Boolean
    'ipTestRank = IIf(VBA.IsMissing(ipTestRank), 1, ipTestRank)
    'ipRefRank = IIf(VBA.IsMissing(ipRefRank), 1, ipRefRank)
    IsNotSameSize = Not IsSameSize(ipTestIterable, ipRefIterable, ipTestRank, ipRefRank)
End Function

'@Description("Returns true if each item in the iterable is the same type")
Public Function IsHomogeneous(ByVal ipIterable As Variant) As Boolean
Attribute IsHomogeneous.VB_Description = "Returns true if each item in the iterable is the same type"

    IsHomogeneous = False
        
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
        
    IsHomogeneous = True

End Function

Public Function IsNotHomogeneous(ByVal ipArray As Variant) As Boolean
    IsNotHomogeneous = Not IsHomogeneous(ipArray)
End Function

Public Function HasNoItems(ByVal ipArg As Variant, Optional ByVal ipRank As Variant) As Boolean
    Dim myRank As Long
    If VBA.IsMissing(ipRank) Then myRank = enums.Preset.Value(PresetId.ArrayFirstRank): Else myRank = ipRank
    HasNoItems = Not HasItems(ipArg, myRank)
End Function

'@Description("Returns true if arg is an array and has ubound of 0 or more, or is an object with a count greater than 0")
Public Function HasItems(ByVal ipArg As Variant, Optional ByVal ipRank As Variant) As Boolean
Attribute HasItems.VB_Description = "Returns true if arg is an array and has ubound of 0 or more, or is an object with a count greater than 0"
    
    ' Uninitialised variables will be objects or arrays
    ' not initialised means
    ' an object that is nothing
    ' an array that gives a ubound error
    ' and array that gives a ubound of -1
        
    HasItems = False
    Dim myRank As Long
    If VBA.IsMissing(ipRank) Then myRank = enums.Preset.Value(PresetId.ArrayFirstRank): Else myRank = ipRank
        
    If Types.Group.IsNotIterable(ipArg) Then Exit Function

    If VBA.IsArray(ipArg) Then
        
        If Arrays.HasNoItems(ipArg) Then Exit Function  '  Need to something about myRankshere
            
    Else
            
        If ipArg.Count <= 0 Then Exit Function

    End If
        
    HasItems = True
        
End Function

'@Description("Takes a list of values or an iterable and returns a variant containing an Array")
Public Function ToArray(ParamArray ipArgArray() As Variant) As Variant
Attribute ToArray.VB_Description = "Takes a list of values or an iterable and returns a variant containing an Array"
    ToArray = ToArraybyFPA(ipArgArray)
End Function


'Description("Performs the ToArray method but taking forwarded paramarray")
Friend Function ToArraybyFPA(ByVal ipArgArray As Variant) As Variant

    ToArraybyFPA = Empty
        
    If BailOut.When(Arrays.HasNoItems(ipArgArray), Msg.ArrayHasNoItems) Then Exit Function
            
    Dim myArrayList As ArrayList
    Set myArrayList = New ArrayList
    Dim myItem As Variant
        
    If Arrays.HasMoreThanOneItem(ipArgArray) Then
            
        For Each myItem In ipArgArray
            
            myArrayList.Add myItem
                
        Next
        
        ToArraybyFPA = myArrayList.ToArray
        Exit Function
        
    End If
            
    ' There is only one item in ipArgsArray
    Select Case True
    
        Case Types.Group.IsNotIterable(ipArgArray(0))
        
            ToArraybyFPA = ipArgArray
            
            
        Case Types.IsTypeByName(Types.Name.OfArrayList, ipArgArray(0))
        
            ToArraybyFPA = ipArgArray(0).ToArray
            
    
        Case VBA.IsArray(ipArgArray(0))
        
            ToArraybyFPA = ipArgArray(0)
            
        
        Case Types.Group.IsIterableItemsByEnum(ipArgArray(0))
        
            For Each myItem In ipArgArray(0)
                    
                myArrayList.Add myItem
                    
            Next
            
            ToArraybyFPA = myArrayList.ToArray
            
            
        Case Types.Group.IsIterableKeysByEnum(ipArgArray(0))
        
            ToArraybyFPA = ipArgArray(0).Items
            
            
        Case Else
        
            If BailOut.When(True, Msg.UnexpectedItemInParamArray) Then Exit Function
            
            
    End Select
        
 
End Function


'@Description("Takes a list of values or an iterable and returns a Queue")
Public Function ToQueue(ParamArray ipArgArray() As Variant) As Queue
Attribute ToQueue.VB_Description = "Takes a list of values or an iterable and returns a Queue"
        
    Dim myArray As Variant
    myArray = ToArraybyFPA(ipArgArray)
    If BailOut.When(VBA.IsEmpty(myArray), Msg.ParamArrayIsNotValid) Then Exit Function
    
    Dim myQueue As Queue
    Set myQueue = New Queue
    
    Dim myItem As Variant
    For Each myItem In myArray
            
        myQueue.Enqueue myItem
                
    Next
            
    Set ToQueue = myQueue
                
End Function


'@Description("Takes a list of values or an iterable and returns a Stack")
Public Function ToStack(ParamArray ipArgArray() As Variant) As Stack
Attribute ToStack.VB_Description = "Takes a list of values or an iterable and returns a Stack"
        
    Dim myArray As Variant
    myArray = ToArraybyFPA(ipArgArray)
    If BailOut.When(VBA.IsEmpty(myArray), Msg.ParamArrayIsNotValid) Then Exit Function
    
    Dim myStack As Stack
    Set myStack = New Stack
    
    ' For each on a stack returns numbers in the opposite order to which we
    ' added them to the stack, so a strightforward for each here
    ' actually reverses the number on the stack.
    ' This is why the myIndex with step -1 is required
    
    Dim myIndex As Long
    For myIndex = UBound(myArray) To LBound(myArray) Step -1

        
        myStack.Push CVar(myArray(myIndex))
                
    Next
            
    Set ToStack = myStack
                
End Function

'@Description("Takes a list of values or an iterable and returns an ArrayList")
Public Function ToArrayList(ParamArray ipArgArray() As Variant) As ArrayList
Attribute ToArrayList.VB_Description = "Takes a list of values or an iterable and returns an ArrayList"

    Dim myArray As Variant
    myArray = ToArraybyFPA(ipArgArray)
    If BailOut.When(VBA.IsEmpty(myArray), Msg.ParamArrayIsNotValid) Then Exit Function
    
    Dim myAL As ArrayList
    Set myAL = New ArrayList
    
    Dim myItem As Variant
    For Each myItem In myArray
            
        myAL.Add myItem
                
    Next
            
    Set ToArrayList = myAL
    
End Function

'@Description("Takes a list of values or an iterable and returns a Lyst")
Public Function ToLyst(ParamArray ipArgArray() As Variant) As Lyst
Attribute ToLyst.VB_Description = "Takes a list of values or an iterable and returns a Lyst"
        
    Dim myArray As Variant
    myArray = ToArraybyFPA(ipArgArray)
    If BailOut.When(VBA.IsEmpty(myArray), Msg.ParamArrayIsNotValid) Then Exit Function
    
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    myLyst.AddRange myArray
            
    Set ToLyst = myLyst
        
End Function


'@Description("Takes a list of values or an iterable and returns a Collection)
Public Function ToCollection(ParamArray ipArgArray() As Variant) As Collection

    Dim myArray As Variant
    myArray = ToArraybyFPA(ipArgArray)
    If BailOut.When(VBA.IsEmpty(myArray), Msg.ParamArrayIsNotValid) Then Exit Function
    
    Dim myColl As Collection
    Set myColl = New Collection
    
    Dim myItem As Variant
    For Each myItem In myArray
    
        myColl.Add myItem
        
    Next
            
    Set ToCollection = myColl
        
End Function

'@Description("Takes a list of values or an iterable and returns a Scriting with number keys starting at 0")
Public Function ToDictionary(ParamArray ipArgArray() As Variant) As Scripting.Dictionary
Attribute ToDictionary.VB_Description = "Takes a list of values or an iterable and returns a Scriting with number keys starting at 0"
        
    Dim myArray As Variant
    myArray = ToArraybyFPA(ipArgArray)
    If BailOut.When(VBA.IsEmpty(myArray), Msg.ParamArrayIsNotValid) Then Exit Function
    
    Dim myDic As Scripting.Dictionary
    Set myDic = New Scripting.Dictionary
    
    Dim myItem As Variant
    For Each myItem In myArray
    
        myDic.Add myDic.Count, myItem
        
    Next
    
    Set ToDictionary = myDic
        
End Function

'@Description("Takes a list of values or an iterable and returns a Kvp with Long keys starting at 0&")
Public Function ToKvp(ParamArray ipArgArray() As Variant) As Kvp
Attribute ToKvp.VB_Description = "Takes a list of values or an iterable and returns a Kvp with Long keys starting at 0&"

    Dim myArray As Variant
    myArray = ToArraybyFPA(ipArgArray)
    If BailOut.When(VBA.IsEmpty(myArray), Msg.ParamArrayIsNotValid) Then Exit Function
    
    Dim myKvp As Kvp
    Set myKvp = Kvp.Deb
    
    myKvp.Add myArray
        
    Set ToKvp = myKvp
    
End Function
