VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Kvp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Kvp is a flexible Key/Value pair dictionary targeted at VB/VBA. It facilitates data input, direct content manipulation of items and keys, and indirect manipulation of data via Map and Comparer objects"
Option Explicit
''@IgnoreModule
'@PredeclaredId
'@Exposed
'@Folder("Kvp")
'@ModuleDescription("Kvp is a flexible Key/Value pair dictionary targeted at VB/VBA. It facilitates data input, direct content manipulation of items and keys, and indirect manipulation of data via Map and Comparer objects")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Keys and Indexing:
' Items in a Kvp may be accessed by Key or by Indexing.  Keys may have any legal value for the type of key in use.
' Indexing for Kvp's is 1 based.  Negative idexes and runs are allowed. Specifying an index of 0 for an insert related
' operation will result in an add operation.
' AutoKeys:
' Three types of autokey are supported
' 1. Number based
' 2. String based: A listt of characters allowed in a key can be specified in a string.
'    The index position in the string is used to determine precedence and not the ascii value.
'    If no character string is provided the string "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
'    is used and the initial key is set to "0000"
' 3. List based: Allowed values are provided to the Autokey as a Lyst of values.  The first key index may be specified
'    so that it does not need to be the index in the first item position of the key list.
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
 
Public Enum KvpAdmin
    noindex = 0          '
    inputisadd          'Used as a flag to the InputItems private method rather than overloading any Index for insert requests
    InputIsInsert       'Used as a flag to the InputItems private method rather than overloading any Index for insert requests
    SelectByIndex       ' Where a Kvp is being returned as multiple Kvp, set the keys of the sub kvp to be the index from the host Kvp
    SelectByKey         ' Where a Kvp is being returned as multiple Kvp, set the keys of the sub kvp to be the key from the host Kvp
    'AfterKey = 1
    'BeforeKey = 0
    
End Enum

' Public Enum RemoveType
    
'     AsArray
'     AsLyst
'     AsKvp
    
' End Enum

' Add and Insert methods have a lot of commonality in the preprocessing of the input param array.
' To simplify coding such preprocessing is consolidated into the InputItems method which takes
' a forwarded ParamArray and a long value to indicate an add or insert.
' Indexing for Kvp (and Lyst) objects is 1 based.  This a long value of 0 is used to indicate an add
' operation.  Values other than 0 indicate the index at which to start the insert operation
' i.e. the value at index will be the first item to appear after the insert


Private Type ReadOnly

    DefaultSeparator                            As String 'char.comma
    DefaultAdjustAmount                         As Long
    CountIsInvalid                              As Long ' -1
    'Const i.KvpError                           As Long = 17 + vbObjectError
    'Const i.ParamArrayIsEmpty                  As Long = -1
    'Const OptionalDefault                      As Long = -1
    'FlagAddItems                                As Long '-1
    

End Type

Private r                       As ReadOnly


Private Type Properties
  
  Keys                            As Lyst
  Items                           As Lyst
  AutoKey                         As IAutoKey
  Separator                       As String
  
End Type

Private p                         As Properties


Private Type State
  
    FirstUse                        As Boolean
    
End Type

Private s                           As State


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class admin
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Sub Class_Initialize()

        'Guard InvalidNew, Types.IsNotSameInstance(Me, Kvp), r.Location & "Class_Initialize"
    
  With r
    
    .DefaultSeparator = Char.twComma
    '.CountIsInvalid = -1
    '.FlagAddItems = -1
      
  End With
    
    
End Sub

'@Description("Returns a new instnce of a Kvp Class")
Public Function Deb(Optional ByVal ipKvp As Kvp = Nothing) As Kvp
Attribute Deb.VB_Description = "Returns a new instnce of a Kvp Class"

    'Guard UsePredeclaredIdInstance, Types.IsNotSameInstance(Me, Kvp), i.Location & "Debe"
    
    With New Kvp
    
        If ipKvp Is Nothing Then
            
            Set Deb = .ReadyToUseInstance
            
            
        Else
            
            Set Deb = .ReadyToUseClone(ipKvp)
            
            
        End If
        
    End With
    
End Function


Public Function ReadyToUseInstance() As Kvp

    'Guard UsePredeclaredIdInstance, Types.IsNotSameInstance(Me, Kvp), i.Location & "Deb"
    
    
    Set p.Keys = Lyst.Deb
    Set p.Items = Lyst.Deb
    Set p.AutoKey = AutoKeyByNumber.Deb
    p.Separator = r.DefaultSeparator
    's.FirstUse = True

    Set ReadyToUseInstance = Me
        
End Function


Public Function ReadyToUseClone(ByVal ipKvp As Kvp) As Kvp
    
    Set p.Keys = ipKvp.GetKeys.Clone
    Set p.Items = ipKvp.ItemsRef.Clone
    Set p.AutoKey = ipKvp.AutoKey
    p.Separator = ipKvp.Separator
    's.FirstUse = IIf(ipKvp.AutoKey Is Nothing, True, False)
    Set ReadyToUseClone = Me

End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class public members
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'#Region "Adding stuff"
''@Description("Adds items in the ParamArray and then returns itself. If ParamArray holds only 1 or 2 items the items are interpreted differently to  when more than two items are added")
'
' Case of 1 Item
'
'   NonIterable         autokey vs item
'   ListArray           autokeysequence vs ListArray
'   TableArray          autokeysequence vs table rows
'   MSArray             Invalid
'   Collection,ArrayList,Lyst,Queue,Stack  autokeysequence vs iterable
'   Scripting.Dictionary, Kvp               autokeysequence vs items
'
'Case of 2 Items
'
'   Item 1: Non Iterable                        Key:Item1, Item:Item2
'   Item 1: ListArray :Item 2: NonIterable      Invalid
'   Item 1: Tablearray :Item2: NonIterable TableToLystAction
'  ListArray or IsIterableItemsByEnum:   Lyst of Autokeys , ListArray(sic) or IsIterableItemsByEnum Items
' IsIterableKeysByEnum:                 iparg.Keys, iparg.Items
' TableArray:                           LystOfAutoKeys, rows as Kvps
' KvPair:                               array(iparg.Key),arrray(iparg.Item)
' To add an iterable as 1 item, encapsulate the iterable in an array e.g. .add array(Iteable)
' 2 items
'
' IsPrimitive, Anything:                array(ipParamOne), array(ipParamTwo)
' Is
'
'
' More than 2 items                     Lyst(AutoKeys), ParamArray
'
'
'@Description("Adds items in the ParamArray, returns the updated kvp. If ParamArray is 1 or 2 items the items are parsed as potential key/value pairs.")
Public Function Add(ParamArray ipArgArray() As Variant) As Kvp
Attribute Add.VB_Description = "Adds items in the ParamArray, returns the updated kvp. If ParamArray is 1 or 2 items the items are parsed as potential key/value pairs."
    InputItems KvpAdmin.inputisadd, ipArgArray
    Set Add = Me
End Function


'@Description("Sugar for Add")
Public Function Enqueue(ParamArray ipArgArray() As Variant) As Kvp
Attribute Enqueue.VB_Description = "Sugar for Add"
    InputItems KvpAdmin.inputisadd, ipArgArray
    Set Enqueue = Me
End Function


'@Description("Returns the updated Kvp. All Items are inserted Before the specified Index. Order of the ParamArray is preserved.)
Public Function InsertAt(ByVal ipInsertIndex As Long, ParamArray ipArgArray() As Variant) As Kvp
    InputItems KvpAdmin.InputIsInsert, ipArgArray, ipInsertIndex
    Set InsertAt = Me
End Function


'@Description("Returns the updated Kvp.  All Items are inserted After the specified key. The order of the ParamArray is preserved.")
Public Function InsertAfterKey(ByVal ipKey As Variant, ParamArray ipArgArray() As Variant) As Kvp
Attribute InsertAfterKey.VB_Description = "Returns the updated Kvp.  All Items are inserted After the specified key. The order of the ParamArray is preserved."
    
    If BailOut.When(p.Keys.LacksItem(ipKey), Msg.KeyNotFound) Then Exit Function
        
    Dim myIndex As Long
    myIndex = p.Keys.IndexOf(ipKey)
    
    If myIndex = p.Keys.LastIndex Then
        
        InputItems KvpAdmin.inputisadd, ipArgArray
        
    Else
    
        InputItems KvpAdmin.InputIsInsert, ipArgArray, myIndex
        
    End If
    
    Set InsertAfterKey = Me
End Function

'@Description("Returns the Modified Kvp.  All Items are inserted Before the specified Key.  The Order of the Paramarray is preserved")
Public Function InsertBeforeKey(ByVal ipKey As Variant, ParamArray ipArgArray() As Variant) As Kvp
Attribute InsertBeforeKey.VB_Description = "Returns the Modified Kvp.  All Items are inserted Before the specified Key.  The Order of the Paramarray is preserved"
    
    If BailOut.When(p.Keys.LacksItem(ipKey), Msg.KeyNotFound) Then Exit Function
        
    Dim myIndex As Long
    myIndex = p.Keys.IndexOf(ipKey)
    
    InputItems KvpAdmin.InputIsInsert, ipArgArray, myIndex
        
    Set InsertBeforeKey = Me
End Function


Private Function InsertByKey(ByVal ipArgArray As Variant, ByRef ipKey As Variant, Optional ByVal ipOffset As Long) As Kvp
    
    Set InsertByKey = Me
    Globals.Res.SetNotOkay
    
    If BailOut.When(Arrays.IsNotValid(ipArgArray), Msg.ParamArrayIsEmpty) Then Exit Function
        
        
    If VBA.IsEmpty(ipKey) Then Exit Function
    If Types.Group.IsIterable(ipArgArray) Then
      If Types.Iterable.HasNoItems(ipArgArray) Then
        
        Globals.Res.SetNotOkay Msg.InvalidKey
        
      End If
    End If
    If p.Keys.LacksItem(ipKey) Then Exit Function
    If Types.Group.IsIterable(ipKey) Then Exit Function
        
    InputItems ipArgArray, p.Keys.IndexOf(ipKey) + ipOffset
    
    Set InsertByKey = Me
End Function

    
'@Description("Returns the the modified Kvp. Items are added at LastIndex. The order in ParamArray is preserved")
Public Function Push(ParamArray ipArgArray() As Variant) As Kvp
Attribute Push.VB_Description = "Returns the the modified Kvp. Items are added at LastIndex. The order in ParamArray is preserved"
    InputItems KvpAdmin.inputisadd, ipArgArray
    Set Push = Me
End Function
'#End Region


'@Description("Keys/Value pairs from the input Kvp are added if missing from the host kvp")
Public Function MergeByKey(ByRef ipKvp As Kvp) As Kvp
Attribute MergeByKey.VB_Description = "Keys/Value pairs from the input Kvp are added if missing from the host kvp"
    
    If (ipKvp Is Nothing) Or ipKvp.HasNoItems Then
        
        Set MergeByKey = Me
        Exit Function
        
    End If
    
    Dim myKey As Variant
    For Each myKey In ipKvp.Keys
        
        If p.Keys.LacksItem(myKey) Then
            
            p.Keys.Add myKey
            p.Items.Add ipKvp.Item(myKey)
            
        End If
        
    Next
    
    Set MergeByKey = Me
    
End Function

'@Description("Keys/Value pairs from the input Kvp are added to the host kvp based on the Key. If the key exits the value is overwritten, if the key does not exist the pair is added")
Public Function MergeByValue(ByRef ipKvp As Kvp) As Kvp
Attribute MergeByValue.VB_Description = "Keys/Value pairs from the input Kvp are added to the host kvp based on the Key. If the key exits the value is overwritten, if the key does not exist the pair is added"
    
    If (ipKvp Is Nothing) Or ipKvp.HasNoItems Then
        
        Set MergeByValue = Me
        Exit Function
        
    End If
    
    Dim myKey As Variant
    For Each myKey In ipKvp.Keys
        
        If p.Keys.HoldsItem(myKey) Then
            
            p.Items.Item(p.Keys.IndexOf(myKey)) = ipKvp.Item(myKey)
            
        End If
        
    Next
    
    Set MergeByValue = Me
    
End Function

'#Region "Removing Stuff"

'@Description("Items are popped from Last index of the host.  Returns a Kvp of two Kvp.  Item(0) popped items. Item(1) is the modified host kvp.")
Public Function Pop(Optional ByVal ipCount As Long = 1) As Kvp
Attribute Pop.VB_Description = "Items are popped from Last index of the host.  Returns a Kvp of two Kvp.  Item(0) popped items. Item(1) is the modified host kvp."
    ' ToDo: Broken - needs refactoring
    Dim myResult As Kvp
    Set myResult = Kvp.Deb
'
'
'    If ipCount = 0 Then
'
'        myResult.Add Kvp.Deb
'        myResult.Add Me
'        Set Pop = myResult
'        Exit Function
'
'
'    End If
'
    If ipCount > p.Keys.Count Then
''
''        Reverse
''        myResult.Add Me.Clone
''
''        Remove
''        myResult.Add Me
''        Set Pop = myResult
''        Exit Function
''
'
   End If
'
'    Dim mystartrun As Result
'    If Ranges.TryStartRunFromAnyStartRunEnd(p.Keys, ipRun:=-ipCount).IsNotOkay Then
'
'
'        myResult.Add Kvp.Deb
'        myResult.Add Me
'        Set Pop = myResult
'        Exit Function
'
'    End If
'
'    Set mystartrun = Globals.Res.GetResultStartRun
'
'
'    Dim myStartIndex As Long
'    myStartIndex = mystartrun.
'    Dim myRun As Long
'    myRun = mystartrun.Run
'
'    myResult.Add Kvp.Deb.Add(p.Keys.GetRange(myStartIndex, myRun).Reverse, p.Items.GetRange(myStartIndex, myRun).Reverse)
'
'    p.Keys.RemoveRange myStartIndex, myRun
'    p.Items.RemoveRange myStartIndex, myRun
'
'    myResult.Add Me
   Set Pop = myResult
            
End Function


'@Description("Returns the modified Kvp. Removes the Key/Item pairs specified. If no argument is provided the whole Kvp is cleared")
Public Function RemoveKeys(ParamArray ipKeys() As Variant) As Kvp
Attribute RemoveKeys.VB_Description = "Returns the modified Kvp. Removes the Key/Item pairs specified. If no argument is provided the whole Kvp is cleared"
    
    If Arrays.HasNoItems(ipKeys) Then
        
        p.Keys.Clear
        p.Items.Clear
        Set RemoveKeys = Me
        Exit Function
        
        
    End If
    
    'Dim myResult As Result
    Dim myKeys As Variant
    If Arrays.HasOneItem(ipKeys) Then
        
        If Parser.ParseVariantWithOneItem(ipKeys).IsNotOkay Then
            
            Set RemoveKeys = Me
            Exit Function
            
        Else
            
            Set myKeys = Globals.Res.Item(ResultItemsEnums.ParsedVariant(pvItems))
            
        End If
        
    Else
        
        myKeys = ipKeys
        
    End If
    
    Dim myKey As Variant
    For Each myKey In myKeys
        
        Dim myIndex As Long
        If p.Keys.HoldsItem(myKey) Then
            
            myIndex = p.Keys.IndexOf(myKey)
            p.Keys.RemoveAt myIndex
            p.Items.RemoveAt myIndex
            
        End If
        
    Next
    
    Set RemoveKeys = Me
   
End Function


'@Description("Returns a Kvp removed items.  Items are removed from FirstIndex.")
Public Function Dequeue(Optional ByVal ipCount As Long = 1) As Kvp
Attribute Dequeue.VB_Description = "Returns a Kvp removed items.  Items are removed from FirstIndex."
    
    Dim myKvp As Kvp
    Set myKvp = Kvp.Deb
    Set Dequeue = myKvp
    If ipCount < 1 Then Exit Function
        
    If ipCount > Count Then
        
        Set myKvp.Keys = Keys.Clone
        Set myKvp.Items = Items.Clone
        Set Dequeue = myKvp
       
    End If
    
    Dim myIndex As Long
    For myIndex = 1 To ipCount
        
        myKvp.Add p.Keys.Dequeue.Item(0), p.Items.Dequeue.Item(0)
        
    Next
    
End Function


'@Description("Returns a KVPair of the Key and Item from the specified Key. If Key is not specified uses LastIndex.Deletes the Key and Item from the host Kvp")
Public Function Pull(Optional ByVal ipKey As Variant) As KVPair
Attribute Pull.VB_Description = "Returns a KVPair of the Key and Item from the specified Key. If Key is not specified uses LastIndex.Deletes the Key and Item from the host Kvp"

    Dim myIndex As Long
    If (VBA.IsMissing(ipKey) Or VBA.IsEmpty(ipKey)) Then
        
        myIndex = LastIndex
        
        
    ElseIf Types.Group.IsIterable(ipKey) Then
        
        If Types.Group.IsIterableByIndex Then
            
            myIndex = ipKey(0)
            
        Else
            
            
            myIndex = Types.Iterable.GetByForEach(ipKey, 1).Item(0)
            
        End If
        
    End If
        
    If (VBA.IsMissing(ipKey) Or VBA.IsEmpty(ipKey)) Then
        
        myIndex = LastIndex
        
    Else
        
        myIndex = p.Keys.IndexOf(ipKey)
        
    End If
    

    Set Pull = KVPair.Deb(ipKey, p.Items.Item(myIndex))
    p.Keys.RemoveAt myIndex
    p.Items.RemoveAt myIndex

End Function

'@Description("Returns a KVPair of the Key and Item from the specified index. Deletes the Key and Item from the host Kvp")
Public Function PullAt(ByVal ipIndex As Long) As KVPair
Attribute PullAt.VB_Description = "Returns a KVPair of the Key and Item from the specified index. Deletes the Key and Item from the host Kvp"

    If ipIndex < 1 Or ipIndex > Count Then
        
        Set PullAt = KVPair.Deb
        Exit Function
        
        
    End If
    
    Set PullAt = KVPair.Deb(p.Keys.Item(ipIndex), p.Items.Item(ipIndex))
    p.Keys.RemoveAt ipIndex
    p.Items.RemoveAt ipIndex

End Function

'@Description("Returns a KVPair of the Key and Item at Index 0. Deletes the Key and Item from the host Kvp")
Public Function PullFirst() As KVPair
Attribute PullFirst.VB_Description = "Returns a KVPair of the Key and Item at Index 0. Deletes the Key and Item from the host Kvp"
    
    Set PullFirst = PullAt(FirstIndex)
    
End Function

'@Description("Returns a KVPair of the Key and Item at Last Index. Deletes the Key and Item from the host Kvp")
Public Function PullLast() As KVPair
Attribute PullLast.VB_Description = "Returns a KVPair of the Key and Item at Last Index. Deletes the Key and Item from the host Kvp"
    
    Set PullLast = PullAt(LastIndex)
    
End Function

Public Function ReduceIt(ByVal ipReduce As IReduce) As Variant
    
    Dim myV As Variant
    myV = Array(ipReduce.ExecRed(Me.Items))
    
    If VBA.IsObject(myV(0)) Then
    
        Set ReduceIt = myV(0)
        
        
    Else
        
        ReduceIt = myV(0)
        
        
    End If
     
 End Function

'@Description("Returns the Result object which indicates the status of the last method call")
Public Function Res() As Result
Attribute Res.VB_Description = "Returns the Result object which indicates the status of the last method call"
  Set Res = Globals.Res
End Function


'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function Remove(Optional ByVal ipKey As Variant) As Kvp
    
    Dim myIndex As Long
    Dim myKeys As Variant
    If VBA.IsMissing(ipKey) Then
    
        p.Keys.Clear
        p.Items.Clear
        Set Remove = Me
        Exit Function
        
    ElseIf Types.Group.IsNotIterable(ipKey) Then
        
        myIndex = p.Keys.IndexOf(ipKey)
        p.Keys.RemoveAt (myIndex)
        p.Items.RemoveAt (myIndex)
        Set Remove = Me
        Exit Function
        
    ElseIf Types.Group.IsNotIterableByIndex Then
        
        Set myKeys = Types.Iterable.GetByForEach(ipKey)
        
        
    Else
        
        Set myKeys = ipKey
        
        
    End If
    
    Dim myKey As Variant
    For Each myKey In myKeys
    
        myIndex = p.Keys.IndexOf(myKey)
        p.Keys.RemoveAt myIndex
        p.Items.RemoveAt myIndex
        
    Next
    
    Set Remove = Me
    
End Function


Public Function RemoveAt(ByVal ipIndex As Long) As Kvp
    
    
    p.Items.RemoveAt ipIndex
    p.Keys.RemoveAt ipIndex
    
    Set RemoveAt = Me
    
End Function


'@Description("Returns the modified host Kvp after removing the Key and Item at Index 0")
Public Function RemoveFirst() As Kvp
Attribute RemoveFirst.VB_Description = "Returns the modified host Kvp after removing the Key and Item at Index 0"
    Set RemoveFirst = RemoveAt(FirstIndex)
End Function

'@Description("Returns the modified host Kvp after removing the Key and Item at LastIndex")
Public Function RemoveLast() As Kvp
Attribute RemoveLast.VB_Description = "Returns the modified host Kvp after removing the Key and Item at LastIndex"
    Set RemoveLast = RemoveAt(LastIndex)
End Function

'#End Region

Public Function Clear() As Kvp
    p.Keys.Clear
    p.Items.Clear
    p.AutoKey.CurrentKey = Empty
    Set Clear = Me
    
End Function

' The Kvp (and Lyst) classes use 1 based indexing
' To differentiate between a count of items and an object having no items
' Count returns a variant and uses Boolean False to indicate that the object cannot be counted
' as it has no Items.
' A bettter test for a count of 0 is Types.Iterable.HasNoItems
' as this will succeed with any iterable type
'@Description("Returns the number of key/item pairs in the Kvp. If there are no items the value returned is -1")
Public Function Count() As Variant
Attribute Count.VB_Description = "Returns the number of key/item pairs in the Kvp. If there are no items the value returned is -1"
    Count = p.Keys.Count
End Function


'@Description("Returns a copy of the Kvp. If Items are objects then a shallow copy is made unless the object supports a clone Method.")
Public Function Clone() As Kvp
Attribute Clone.VB_Description = "Returns a copy of the Kvp. If Items are objects then a shallow copy is made unless the object supports a clone Method."
    
    Set Clone = Me.Deb(Me)
    
End Function

'@Description("Decrements the Items  at the given Keys by the specified amount if the Items are Numbers")
Public Function Dec(Optional ByVal ipKeys As Variant = Empty, Optional ByVal ipDecrement As Variant = Constant.DefaultDec) As Kvp
Attribute Dec.VB_Description = "Decrements the Items  at the given Keys by the specified amount if the Items are Numbers"

    Globals.Res.Init (Msg.Okay)
    
    If BailOut.When(Types.Group.IsNotNumber(ipDecrement), Msg.IsNotNumber, ipAltMessage:="The increment variable is not a number") Then Exit Function
    If BailOut.When(ipDecrement = 0, Msg.ZeroIsInvalidParameter) Then Exit Function
        
    AdjustbyKey ipKeys, -ipDecrement
    Set Dec = Me
    
End Function

'@Description("Decrements the Items at the given indexes by the specified amount if the Items are Numbers")
Public Function DecAt(ByVal ipIndexes As Variant, Optional ByVal ipDecrement As Single) As Kvp
Attribute DecAt.VB_Description = "Decrements the Items at the given indexes by the specified amount if the Items are Numbers"
    If ipDecrement = 0 Then ipDecrement = r.DefaultAdjustAmount
    AdjustByIndex ipIndexes, -ipDecrement
    Set DecAt = Me
End Function

Public Function FirstIndex() As Long
    FirstIndex = 1
End Function

'@Description("Returns a Kvp comprised of Key/Item pairs where the Key appears in both the Host Kvp and the input parameter")
Public Function FilterByKeys(ByVal ipKeysIterable As Variant) As Kvp
Attribute FilterByKeys.VB_Description = "Returns a Kvp comprised of Key/Item pairs where the Key appears in both the Host Kvp and the input parameter"
    ' Preference should be given to using the specific KeysIn<xxx> method
    Dim myResult As Kvp
    Set myResult = Me.Deb
    Set FilterByKeys = myResult

    If VBA.IsEmpty(ipKeysIterable) Then Exit Function

    Dim myKeys As Lyst
    If Types.IsTypeByName(Types.Name.OfLyst, ipKeysIterable) Then

        Set myKeys = ipKeysIterable


    ElseIf Types.Group.IsIterableKeysByEnum(ipKeysIterable) Then

        Set myKeys = Types.Iterable.ToLyst(ipKeysIterable.Keys)


    Else

        Set myKeys = Types.Iterable.ToLyst(ipKeysIterable)


    End If

    myResult.Add 0&, Me.KeysInAandB(myKeys)
    myResult.Add 1&, Me.KeysInAOnly(myKeys)
    myResult.Add 2&, Me.KeysInBOnly(myKeys)

    Set FilterByKeys = myResult
    
End Function

'@Description("Returns a KVPair of the Key and Item at Index 0")
Public Function GetFirst() As KVPair
Attribute GetFirst.VB_Description = "Returns a KVPair of the Key and Item at Index 0"
    
    If p.Keys.Count = 0 Then
        
        Set GetFirst = KVPair.Deb
        
    Else
        
        Set GetFirst = KVPair.Deb(p.Keys.Item(0), p.Items.Item(0))
        
    End If
    
End Function

'@Description("Returns a Kvp of Key vs Index at which the key is found")
Public Function GetKeysWithValue(ByVal ipValue As Variant, Optional ByVal ipIndexer As KvpAdmin = KvpAdmin.SelectByIndex) As Kvp
Attribute GetKeysWithValue.VB_Description = "Returns a Kvp of Key vs Index at which the key is found"
    
    If Not p.Items.HoldsItem(ipValue) Then
    
        Set GetKeysWithValue = Nothing
        Exit Function
        
    End If
    
    Dim myResult As Kvp
    Set myResult = Me.Deb

    Dim myIndex As Long
    For myIndex = 0 To Me.LastIndex

        If p.Items.Item(myIndex) = ipValue Then
        
            If ipIndexer - KvpAdmin.SelectByIndex Then
            
                myResult.Add p.Keys.Item(myIndex), myIndex
                
            Else
                
                myResult.Add myIndex, p.Keys.Item(myIndex)
                
            End If
            
        End If
        
    Next

    Set GetKeysWithValue = myResult
    
End Function

'@Description("Returns the zero based index of the Key")
Public Function GetIndexOfKey(ByVal ipKey As Variant) As Long
Attribute GetIndexOfKey.VB_Description = "Returns the zero based index of the Key"

    If Types.IsEmpty(ipKey) Then
    
        GetIndexOfKey = -1
        
    Else
        
        GetIndexOfKey = p.Keys.IndexOf(ipKey)
        
    End If
    
End Function


'@Description("Returns a kvp of Index vs Key")
Public Function GetIndexOfValue(ByVal ipValue As Variant, Optional ByVal ipOrganiseby As KvpAdmin = KvpAdmin.SelectByIndex) As Kvp
Attribute GetIndexOfValue.VB_Description = "Returns a kvp of Index vs Key"
    's.values may hold more than one item with the same value so we return a Kvp
    Set GetIndexOfValue = Me.MirrorByValue(ipOrganiseby).Item(ipValue)
End Function

'@Description("Returns a shallow copy of the Keys")
Public Function GetKeys() As Lyst
Attribute GetKeys.VB_Description = "Returns a shallow copy of the Keys"
    
    If p.Keys Is Nothing Then
    
        Set GetKeys = Lyst.Deb
        
    Else
    
        Set GetKeys = p.Keys
        
    End If
    
End Function

'@Description("Returns a KVPair of the Key and Item at LastIndex")
Public Function GetLast() As KVPair
Attribute GetLast.VB_Description = "Returns a KVPair of the Key and Item at LastIndex"
    Set GetLast = KVPair.Deb(p.Keys.Item(Me.LastIndex), p.Items.Item(Me.LastIndex))
End Function

'@Description("Returns a Lyst of the specified Items. An empty Paramarray will return all Iems. If the item is an object the Clone methos will be used if available. ")
' Public Function GetItems(ParamArray ipKeys() As Variant) As Lyst

'     Dim myLyst As Lyst
'     Set myLyst = Lyst.Deb
'     Set GetItems = myLyst
'     Dim myKeys As Variant
    
'     Select Case True
        
'         ' In the case of an empty paramarray we return all items
'         Case Arrays.HasNoItems(ipKeys)

'             Set myKeys = p.Keys
            
    
'         Case Arrays.IsNotSingleItem(ipKeys)
        
'             myKeys = ipKeys
        
'         ' From here onwards we have a single item paramarray
'         Case Types.Group.IsNotIterable(ipKeys(0))
            
'             myKeys = Array(ipKeys(0))
            
'         ' If the single item is an array, only accept a ListArray
'         Case VBA.IsArray(ipKeys(0))
                
'             Guard Id.IsNotAListArray, Arrays.IsNotListArray(ipKeys(0)), r.Location & "GetItems"
'             myKeys = ipKeys(0)
            
        
'         Case Types.Group.IsIterableItemsByEnum(ipKeys(0))
            
'                 Set myKeys = ipKeys(0)
                
                
'         Case Types.Group.IsIterableKeysByEnum(ipKeys(0))
            
'             Set myKeys = ipKeys(0).Keys
            
            
'         Case Else
        
'             Guard _
'                 Id.InvalidKey, _
'                 True, _
'                 r.Location & "Items", _
'                 ipAltMessage:=Fmt.Txt("The Key '{0}' is unexpected", TypeName(ipKeys(0)))
            
'     End Select
    
'     ' if items have a clone method then we call the clone method
    
'     Dim myKey As Variant
'     For Each myKey In myKeys
        
'         If Types.HasCloneMethod(myKey) Then
            
'             myLyst.Add Me.Item(myKey).Clone
            
'         Else
            
'             myLyst.Add Me.Item(myKey)
            
'         End If
        
'     Next
    
' End Function



' <summary> NextPair/PrevPair get the indicated KVPair from the Host
' Dictionary. These methods are not intended to be used as iterators
' but as ad hoc requests for the next/prev KVPair in the Kvp. If the
' supplied key is the last/first key in the Dictionary then ipLoop
' will determine if the KVPair returned will be null:null (ipLoop =
' false) or the first/last KVPair in the Dictionary (ipLoop=True)
' </summary> <paramname="ipKey">The Key for which the next/prev
' KVPair should be returned. <paramname="ipLoop">Wraps around the
' Dictionary if True </param> <returns>The Indicated KVPair</returns></param>
Public Function NextPair(ByVal ipKey As Variant, Optional ByVal ipTreatKeysArrayAsLoop As Boolean = True) As KVPair
Attribute NextPair.VB_Description = "Returns a Lyst of the specified Items. An empty Paramarray will return all Iems. If the item is an object the Clone methos will be used if available. "

    Dim myIndex As Long
    myIndex = p.Keys.IndexOf(ipKey) + 1

    If myIndex >= p.Keys.Count Then
    
        If ipTreatKeysArrayAsLoop Then
        
            Set NextPair = KVPair.Deb(p.Keys.Item(0), p.Items.Item(0))
            
        Else
        
            Set NextPair = Nothing
            
        End If
        
    Else
    
        Set NextPair = KVPair.Deb(p.Keys.Item(myIndex), p.Items.Item(myIndex))
        
    End If
    
End Function


Public Function PrevPair(ByVal ipKey As Variant, Optional ByVal ipTreatKeysArrayAsLoop As Boolean = True) As KVPair

    Dim myIndex As Long
    myIndex = p.Keys.IndexOf(ipKey) - 1

    If myIndex < 0 Then
    
        If ipTreatKeysArrayAsLoop Then
        
            Set PrevPair = KVPair.Deb(p.Keys.Item(Me.LastIndex), p.Items.Item(Me.LastIndex))
            
        Else
        
            Set PrevPair = Nothing
            
        End If
        
    Else
    
        Set PrevPair = KVPair.Deb(p.Keys.Item(myIndex), p.Items.Item(myIndex))
        
    End If
    
End Function

'@Description("Returns True if the value is a Key of the Kvp")
Public Function HoldsKey(ByVal Key As Variant) As Boolean
Attribute HoldsKey.VB_Description = "Returns True if the value is a Key of the Kvp"
    HoldsKey = p.Keys.HoldsItem(Key)
End Function

'@Description("Returns True if the value is an Item of the Kvp")
Public Function HoldsValue(ByVal ipValue As Variant) As Boolean
Attribute HoldsValue.VB_Description = "Returns True if the value is an Item of the Kvp"
    HoldsValue = p.Items.HoldsItem(ipValue)
End Function

'@Description("Increments the Items  at the given Keys by the specified amount if the Items are Numbers")
Public Function Inc(Optional ByVal ipKeys As Variant = Empty, Optional ByVal ipIncrement As Double) As Kvp
Attribute Inc.VB_Description = "Increments the Items  at the given Keys by the specified amount if the Items are Numbers"
    If ipIncrement = 0 Then ipIncrement = r.DefaultAdjustAmount
    
    AdjustbyKey ipKeys, ipIncrement
    Set Inc = Me
End Function


Public Function IncAt(Optional ByVal ipIndexes As Variant = Empty, Optional ByVal ipIncrement As Double) As Kvp
    If ipIncrement = 0 Then ipIncrement = r.DefaultAdjustAmount
   
    AdjustByIndex ipIndexes, ipIncrement
    Set IncAt = Me
End Function

'@Description("True if Count = 0, added to match terminology in Lyst and Group Iterable/Arrays")
Public Function HasNoItems() As Boolean
Attribute HasNoItems.VB_Description = "True if Count = 0, added to match terminology in Lyst and Group Iterable/Arrays"
    HasNoItems = p.Keys.Count = 0
End Function


'@Description("True if Count > 0, added to match terminology in Lyst and Group Iterable/Arrays")
Public Function HasItems() As Boolean
Attribute HasItems.VB_Description = "True if Count > 0, added to match terminology in Lyst and Group Iterable/Arrays"
    HasItems = Not HasNoItems
End Function

'@Description("Returns true if the items values onlu occur once")
Public Function IsUnique() As Boolean
Attribute IsUnique.VB_Description = "Returns true if the items values onlu occur once"
    IsUnique = Me.MirrorFirstValues.Item(1&).Count = 0
End Function

'@Description("Returns True is Item values occur more than once")
Public Function IsNotUnique() As Boolean
Attribute IsNotUnique.VB_Description = "Returns True is Item values occur more than once"
    IsNotUnique = Not IsUnique
End Function



'@DefaultMember
Public Property Get Item(ByVal ipKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    
    Dim myIndex As Long
    myIndex = p.Keys.IndexOf(ipKey)
    If IsObject(p.Items.Item(myIndex)) Then
    
        Set Item = p.Items.Item(myIndex)
        
    Else
    
        Item = p.Items.Item(myIndex)
    
    End If
    
End Property

Public Property Let Item(ByVal ipKey As Variant, ByVal ipValue As Variant)
    
    If BailOut.When(VBA.IsEmpty(ipKey), Msg.HasNoItems) Then Exit Property
    
    Dim myIndex As Long
    myIndex = p.Keys.IndexOf(ipKey)
    If VBA.IsObject(ipValue) Then
    
        Set p.Items.Item(myIndex) = ipValue
        
    Else
    
        p.Items.Item(myIndex) = ipValue
        
    End If
    
End Property


Public Property Get ItemAt(ByVal ipIndex As Long) As Variant
    
    If BailOut.When(ipIndex < 0, Msg.IndexBeforeFirstIndex) Then Exit Property
    If BailOut.When(ipIndex > Me.LastIndex, Msg.IndexAfterLastIndex) Then Exit Property
    
    If IsObject(p.Items.Item(ipIndex)) Then
    
        Set ItemAt = p.Items.Item(ipIndex)
        
    Else
    
        ItemAt = p.Items.Item(ipIndex)
        
    End If
    
End Property

Public Property Set ItemAt(ByVal ipIndex As Long, ByVal ipValue As Variant)
    
    If BailOut.When(ipIndex < 0, Msg.IndexBeforeFirstIndex) Then Exit Property
    If BailOut.When(ipIndex > Me.LastIndex, Msg.IndexAfterLastIndex) Then Exit Property

    If IsObject(ipValue) Then
    
        Set p.Items.Item(ipIndex) = ipValue
        
    Else
    
        p.Items.Item(ipIndex) = ipValue
        
    End If
    
End Property


'@Description("Returns a variant containing an array of the Items For alignment with Scripting.Dictionaries. See Method ItemsRef for a reference to the Items lyst")
Public Function Items() As Variant
Attribute Items.VB_Description = "Returns a variant containing an array of the Items For alignment with Scripting.Dictionaries. See Method ItemsRef for a reference to the Items lyst"
  Items = p.Items.ToArray
End Function


'@Description("Returns a reference to the Kvp Items Lyst")
Public Function ItemsRef() As Lyst
Attribute ItemsRef.VB_Description = "Returns a reference to the Kvp Items Lyst"
    Set ItemsRef = p.Items
End Function


'@Description("Returns a variant containing an array of the keys. For alignment with Scripting.Dictionaries. See Method KeysRef for a reference to the Keys lyst")
Public Function Keys() As Variant
Attribute Keys.VB_Description = "Returns a variant containing an array of the keys. For alignment with Scripting.Dictionaries. See Method KeysRef for a reference to the Keys lyst"

    Keys = p.Keys.ToArray               'this works fine
    'Return p.Keys.ToArray              'this works fine
   ' Return (p.Keys.ToArray)            'this works fine
    
End Function 'fails at this step on second iteration

'@Description("Returns a reference to the Host Keys")
Public Function KeysRef() As Lyst
Attribute KeysRef.VB_Description = "Returns a reference to the Host Keys"

    Set KeysRef = p.Keys
End Function

Public Function KeysInAOnly(ByVal ipB As Variant) As Kvp

    Dim myResult As Kvp
    Set myResult = Me.Deb

    Dim myKey As Variant
    For Each myKey In p.Keys

        If ipB.LacksKey(myKey) Then
        
            myResult.Add myKey, Me.Item(myKey)
            
        End If
        
    Next

    Set KeysInAOnly = myResult
    
End Function


Public Function KeysInBOnly(ByVal ipB As Variant) As Kvp

    Dim myResult As Kvp
    Set myResult = Me.Deb

    Dim myKey As Variant
    For Each myKey In ipB.GetKeys

        If Me.LacksKey(myKey) Then
        
            myResult.Add myKey, ipB.Item(myKey)
            
        End If
        
    Next

    Set KeysInBOnly = myResult
    
End Function


Public Function KeysAllAandOnlyB(ByVal ipB As Variant) As Kvp

    Dim myResult As Kvp
    Set myResult = Me.Clone
    
    Dim myResultB As Kvp
    Set myResultB = Me.KeysInBOnly(ipB)
    myResult.Add myResultB.GetKeys, myResultB.ItemsRef
    
    Set KeysAllAandOnlyB = myResult
    
End Function


Public Function KeysInAandB(ByVal ipB As Variant) As Kvp
    
    Dim myResult As Kvp
    Set myResult = Me.Deb
    
    Dim myKey As Variant
    For Each myKey In p.Keys
    
        If ipB.HoldsKey(myKey) Then
    
            myResult.Add myKey, Me.Item(myKey)
    
        End If
        
    Next
    
    Set KeysInAandB = myResult
    
End Function


Public Function KVPairsInAandB(ByVal ipB As Variant) As Kvp
    ' myInB is used for keys in B that have a different value to A
    Dim myResult As Kvp
    Set myResult = Me.Deb
    
    Dim myKey As Variant
    For Each myKey In p.Keys
    
        If ipB.HoldsKey(myKey) Then
    
            If Me.Item(myKey) = ipB.Item(myKey) Then

                myResult.Add myKey, Me.Item(myKey)
                
            End If
    
        End If
        
    Next
    
    Set KVPairsInAandB = myResult
    
End Function


Public Function KeysNotInBothAandB(ByVal ipB As Variant) As Kvp

    Dim myResult As Kvp
    Set myResult = ipB.Clone
    
    Dim myKey As Variant
    For Each myKey In p.Keys
    
        If myResult.HoldsKey(myKey) Then
        
            myResult.Remove myKey
            
        Else
        
            myResult.Add myKey, Me.Item(myKey)
            
        End If
    
    Next
    
    Set KeysNotInBothAandB = myResult
    
End Function


Public Function LacksKey(ByVal ipKey As Variant) As Boolean
    LacksKey = Not HoldsKey(ipKey)
End Function


Public Function LacksValue(ByVal ipValue As Variant) As Boolean
    LacksValue = Not HoldsValue(ipValue)
End Function


Public Function LastIndex() As Long

    LastIndex = -1
    
    If p.Keys Is Nothing Then Exit Function
    
    LastIndex = p.Keys.LastIndex
    
End Function


'returns a Kvp with values as keys.  Each Value is linked
' to a Kvp of Index vs Key
Public Function MirrorByValue(Optional ByVal ipOrganiseby As KvpAdmin = KvpAdmin.SelectByIndex) As Kvp
    
    Dim myResult As Kvp
    Set myResult = Me.Deb
    
    Dim myKey  As Variant
    For Each myKey In p.Keys
    
        Dim myIndex As Variant
        myIndex = p.Keys.IndexOf(myKey)
        If myResult.LacksKey(p.Items.Item(myIndex)) Then
        
            myResult.Add p.Items.Item(myIndex), Me.Deb
            
        End If
        
        If ipOrganiseby = KvpAdmin.SelectByIndex Then
        
            myResult.Item(p.Items.Item(myIndex)).Add myIndex, myKey
            
        Else
        
            myResult.Item(p.Items.Item(myIndex)).Add myKey, myIndex
            
        End If
        
    Next
    
    Set MirrorByValue = myResult

End Function

'@Description("Returns a new Kvp where each value has been transformed by the IMap Object")
Public Function MapIt(ByVal ipMapper As IMapper) As Kvp
Attribute MapIt.VB_Description = "Returns a new Kvp where each value has been transformed by the IMap Object"

    Dim myKvp As Kvp
    Set myKvp = Kvp.Deb

    Dim myIndex As Long
    For myIndex = p.Keys.FirstIndex To p.Keys.LastIndex

        myKvp.Add p.Keys.Item(myIndex), ipMapper.ExecMapper(p.Items.Item(myIndex))

    Next

    Set MapIt = myKvp

End Function

'todo: need to add byindex of bykey
Public Function MirrorFirstValues(Optional ByVal ipOrganiseby As KvpAdmin = KvpAdmin.SelectByIndex) As Kvp

    Dim myFirstValues As Kvp
    Set myFirstValues = Me.Deb
    
    Dim myDupValues As Kvp
    Set myDupValues = Me.Deb

    Dim myIndex As Long
    For myIndex = 0 To p.Keys.Count - 1
    
        If myFirstValues.LacksKey(p.Items.Item(myIndex)) Then
        
            myFirstValues.Add p.Keys.Item(myIndex), p.Items.Item(myIndex)
            
        Else
        
            myDupValues.Add p.Items.Item(myIndex), p.Keys.Item(myIndex)
            
        End If
        
    Next
        
    Set MirrorFirstValues = Me.Deb.Add(Array(myFirstValues, myDupValues))
    
End Function


'@Enumerator
'@Ignore EmptyMethod
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Dim enumerator As IUnknown
    '@Ignore DefaultMemberRequired
    Set enumerator = p.Keys.Host.GetEnumerator(0, p.Keys.Count - 1)
    Set NewEnum = enumerator
    
End Function

Public Function GetEnumerator() As Variant
    GetEnumerator = p.Keys.GetEnumerator
End Function
'Public Function MissingKeys(ByVal ipKeysIterable As Variant) As Kvp
'
'    Dim myKeysList As ArrayList
'    Set myKeysList = GetArrayAsList(ipKeysIterable)
'
'    Dim myMissingInHost As Kvp
'    Set myMissingInHost = Kvp.Deb
'
'    Dim myKey As Variant
'    For Each myKey In myKeysList
'
'        If Not p.Keys.ContainsKey(myKey) Then
'
'            myMissingInHost.Add myKey
'
'        End If
'
'    Next
'
'    Dim myMissingInArray As Kvp
'    Set myMissingInArray = Kvp.Deb
'
'    For Each myKey In p.Keys
'
'        If Not myKeysList.Contains(myKey) Then
'
'            myMissingInArray.Add myKey
'
'        End If
'    Next
'
'    Set MissingKeys = Kvp.Deb.Add(Array(myMissingInHost, myMissingInArray))
'
'End Function

'@Description("Returns the string 'Kvp'")
Public Function Name() As String
Attribute Name.VB_Description = "Returns the string 'Kvp'"
    Name = TypeName(Me)
End Function

'@Description("Returns a Kvp of maximum and minimum values.  Keys are the max and min value, associated with the keys is a Kvp of Index vs Key for each value corresponding to the parent key.")
Public Function NumMaxMin() As Kvp
Attribute NumMaxMin.VB_Description = "Returns a Kvp of maximum and minimum values.  Keys are the max and min value, associated with the keys is a Kvp of Index vs Key for each value corresponding to the parent key."
    
    If BailOut.When(Types.Group.IsNotNumber(p.Items.First), Msg.IsNotNumber) Then Exit Function
    
    Dim myResult As Kvp
    Set myResult = Kvp.Deb
    
    Dim myValues As Lyst
    Set myValues = p.Items.Clone
    myValues.Sort
    
    'Item 0 of the KVp is the minimum values
    Const maMin As Long = 0
    myResult.Add maMin, Me.Deb
    
    'Item 1 of the Kvp is the maximum values
    Const maMax As Long = 1
    myResult.Add maMax, Me.Deb

    ' A maximum or minimum value may occur more than once
    ' so we populate the max and min kvp  with the key vs value
    ' of each found max or min value
    
    Dim myLbound As Long
    myLbound = 0
    
    Dim myUBound As Long
    myUBound = Me.LastIndex
    
    Dim myIndex As Long
    For myIndex = myLbound To myUBound

        Select Case p.Items.Item(myIndex)
        
            Case myValues.Item(myLbound)
                
                myResult.Item(maMin).Add p.Items.Item(myIndex), p.Keys.Item(myIndex)
                
        
            Case myValues.Item(myUBound)
                
                myResult.Item(maMax).Add p.Items.Item(myIndex), p.Keys.Item(myIndex)
                
        End Select
        
    Next

    Set NumMaxMin = myResult
    
    
    
End Function


'Public Function PushAt(ByVal ipIndex As Long, ByVal ipValue As Variant, Optional ByVal ipKey As Variant = Empty)
'
'
'End Function
'
'public function PushByKey( byval ipKey as Variant, byval ipValue)

'@Description("Sugar for InsertFirst")



Public Function Reverse() As Kvp
Attribute Reverse.VB_Description = "Sugar for InsertFirst"

    p.Keys.Reverse
    p.Items.Reverse
    
    Set Reverse = Me

End Function


Public Property Get Separator() As String
    Separator = p.Separator
End Property

Public Property Let Separator(ByVal ipSeparator As String)
    p.Separator = ipSeparator
End Property


Public Function SetAutoKeyToAutoKeyByNumber(Optional ByVal ipFirstKey As Variant = 0&) As Kvp

    'Guard AutoKeyResetNotAllowed, Not s.FirstUse, r.Location & "SetAutoKeyByNumber"
    'Guard IsNotNumber, Types.Group.IsNotNumber(ipFirstKey), r.Location & "SetAutoKeyByNumvber"
    
    
    Set p.AutoKey = AutoKeyByNumber.Deb(ipFirstKey)
    s.FirstUse = False
    Set SetAutoKeyToAutoKeyByNumber = Me
    
End Function


Public Function SetAutoKeyToAutoKeyByString _
( _
    Optional ByVal ipFirstKey As String, _
    Optional ByVal ipCharSequence As String _
) As Kvp
    
    'Guard AutoKeyResetNotAllowed, Not s.FirstUse, r.Location & "SetAutoKeyByString"
    'Guard IsNotString, Types.IsNotString(ipFirstKey), r.Location & "SetAutoKeyByString"
        
    Set p.AutoKey = AutoKeyByString.Deb(ipFirstKey, ipCharSequence)
    s.FirstUse = False
    Set SetAutoKeyToAutoKeyByString = Me
    
End Function


Public Function SetAutoKeyToAutoKeyByIndex(ByVal ipKeysIterable As Variant, Optional ByVal ipStartIndex As Variant) As Kvp
    
    Set SetAutoKeyToAutoKeyByIndex = Me
    
    If BailOut.When(Not s.FirstUse, Msg.AutoKeyResetNotAllowed) Then Exit Function
    If BailOut.When(Types.Group.IsNotIterable(ipKeysIterable), Msg.IsNotIterable) Then Exit Function
    If BailOut.When(Types.Iterable.HasNoItems(ipKeysIterable), HasNoItems) Then Exit Function


    Dim myLyst As Lyst
    Set myLyst = Types.Iterable.ToLyst(ipKeysIterable)
    
    'Dim myResult As Result
    'Guard GuardClause.IndexLessThanFirstIndex, VBA.IsEmpty(Types.Iterable.TryExtent(ipKeysIterable, myResult).Item(Extent.IsFirstIndex)), myPlace
    Dim myKeysFirstIndex As Long
    If Types.Iterable.TryExtent(ipKeysIterable).IsNotOkay Then Exit Function
    myKeysFirstIndex = Globals.Res.Item(ResultItemsEnums.ItemExtent(ieFirstIndex))
    ' Normalise ipStartIndex for use when ipKeysIterable is converted to an ArrayList
    ' An arraylist is required to get access to the 'IndexOf' method
    ' ipStartIndex is the first index to use with ipKeysIterable and is not necessarily
    ' the first item in the iterable
    
    Dim myStartIndex As Long
    Select Case True
            
        Case Types.Group.IsNotNumber(ipStartIndex)
        
            ' case where a value has been provided rather than an index
            ' only valid for when ipFirstIndex is not a number
            
            myStartIndex = myLyst.IndexOf(ipStartIndex) - myKeysFirstIndex + 1
            
        Case Else
        
            myStartIndex = ipStartIndex - myKeysFirstIndex ' + 1
        
    End Select
    
    Set p.AutoKey = AutoKeyByIndex.Deb(myLyst, myStartIndex)
    s.FirstUse = False
    Set SetAutoKeyToAutoKeyByIndex = Me
        
        
End Function


Public Function SumKeys() As LongLong

    If BailOut.When(Types.Group.IsNotNumber(p.Keys.First), Msg.IsNotNumber) Then Exit Function

    Dim mySum As LongLong
    Dim myKey As Variant
    For Each myKey In p.Keys
    
        mySum = mySum + myKey
        
    Next

    SumKeys = mySum
    
End Function


Public Function SumValues() As LongLong
    
    If BailOut.When(Types.Group.IsNotNumber(p.Items.First), Msg.IsNotNumber) Then Exit Function

    Dim mySum As LongLong
    Dim myValue As Variant
    For Each myValue In p.Items
    
        mySum = mySum + myValue
        
    Next

    SumValues = mySum
    
End Function


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' class private methods
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Add methods which specify a Key MUST ensure that they set the current key in the AutoKey
' to the last key specified after completing all adds to ensure continuity of keys
' from the last key used by the method

'Description("Adds.Inserts the forwarded ParamArray based on the ParamArray having 1,2, or more tthan 2 items. If Input action is 0 the paramarray is added otherwise inserted ))")
Friend Sub InputItems(ByVal ipInputType As Long, ByVal ipFPA As Variant, Optional ipIndex As Long = KvpAdmin.noindex)
    'ipIFP is Forwarded ParamArray
   ' do we need a first use test
    If BailOut.When(VBA.Abs(ipIndex) > p.Keys.LastIndex + 1, Msg.IndexAfterLastIndex) Then Exit Sub
    If BailOut.When((ipIndex = 0) And (ipInputType = InputIsInsert), Msg.IndexBeforeFirstIndex) Then Exit Sub
        
    
        Select Case Arrays.Count(ipFPA)
    
            Case Is < 1: Globals.Res.SetNotOkay
            Case 1: InputWhenFPAHasOneItem ipInputType, ipFPA(0), ipIndex
            Case 2: InputWhenFPAHasTwoItems ipInputType, ipFPA(0), ipFPA(1), ipIndex
            Case Is > 2: IputWhenFPAIsListOfItems ipInputType, ipFPA, ipIndex
    
        End Select
'
End Sub

    
Friend Sub InputWhenFPAHasOneItem(ByVal ipInputAction As Long, ByVal ipItem As Variant, ByVal ipIndex As Long)
    
    If Parser.ParseVariantWithOneItem(ipItem).IsNotOkay Then Exit Sub
       
    Dim myParsed As ResultParsedVariant
    Set myParsed = Globals.Res.GetResultParsedVariant
    
    Dim myKeys As Lyst
    Set myKeys = p.AutoKey.GetKeySequence(myParsed.GetItems.Count)
        
    If ipInputAction = KvpAdmin.inputisadd Then
    
        p.Keys.Host.AddRange myKeys.ToArrayList
        p.Items.Host.AddRange myParsed.GetItems.ToArrayList
        
    Else
        
        Dim mystartrun As ResultStartRun
        ' remember that ipindex may be negative
       If Ranges.TryStartRunFromAnyStartRunEnd(p.Keys, ipIndex).IsNotOkay Then Exit Sub
        
        '@Ignore DefaultMemberRequired
        Set mystartrun = Globals.Res.GetResultStartRun
        p.Keys.InsertRange mystartrun.StartIndex, myKeys
        p.Items.InsertRange mystartrun.StartIndex, myParsed.GetItems
        
    End If
        
End Sub

Friend Sub InputWhenFPAHasTwoItems(ByVal ipInputAction As KvpAdmin, ByVal ipfirst As Variant, ByVal ipsecond As Variant, ByVal ipIndex As Long)
    
    If Parser.ParseVariantWithTwoItems(ipfirst, ipsecond).IsNotOkay Then Exit Sub
        
    Dim myParsed As ResultParsedVariant
    Set myParsed = Globals.Res.GetResultParsedVariant
        
    If ipInputAction = KvpAdmin.inputisadd Then
    
        p.Keys.AddRange myParsed.GetKeys
        p.Items.AddRange myParsed.GetItems
        
    Else
        
        Dim mystartrun As ResultStartRun
        ' remember that ipindex may be negative
        If Ranges.TryStartRunFromAnyStartRunEnd(p.Keys, ipIndex).IsNotOkay Then Exit Sub
        
        '@Ignore DefaultMemberRequired
        Set mystartrun = Globals.Res.GetResultStartRun
        p.Keys.InsertRange mystartrun.StartIndex, myParsed.GetKeys
        p.Items.InsertRange mystartrun.StartIndex, myParsed.GetItems
        
    End If
    
End Sub


Friend Sub IputWhenFPAIsListOfItems(ByVal ipInsertAt As Long, ByVal ipList As Variant, ByVal ipIndex As Long)
    
    Dim myCount As Long
    myCount = Arrays.Count(ipList)
    If BailOut.When(VBA.Abs(ipInsertAt) > myCount, Msg.IndexAfterLastIndex) Then Exit Sub
        
    Dim myKeys As Lyst
    Set myKeys = p.AutoKey.GetKeySequence(myCount)
    
    If ipInsertAt = KvpAdmin.inputisadd Then
          
        p.Keys.AddRange myKeys
        p.Items.AddRange ipList
        
    Else
    
        If Ranges.TryStartRunFromAnyStartRunEnd(p.Keys, ipIndex).IsNotOkay Then Exit Sub
        Dim myIndex As Long
        myIndex = Globals.Res.Item(ResultItemsEnums.StartRun(srStartindex))
        
        p.Keys.InsertAt myIndex, myKeys
        p.Items.InsertAt myIndex, ipList
            
        
    End If
    
End Sub

' Friend Sub InputItemsTwoParams(ByVal ipParamOne As Variant, ByVal ipParamTwo As Variant, ByVal ipStartIndex As Long)

' ' To add two iterables as objects, rather than as their component items,
' ' combine the two iterables into an array so that the ParamArray
' ' consists of a single argument which is processed by the else clause
' ' in the Case 1 branch of pvAddorInsert
'     Guard Id.InvalidType, Types.Group.IsAdmin(ipParamOne), r.Location & "pvAddOrInsertTwoParams", Array(TypeName(ipParamOne)), "Admin types cannot be keys '{0}'"
    
    
'     Select Case True
                
'         Case Types.Group.IsPrimitive(ipParamOne) 'Imlemented
    
'             ' for two parameters when ipParamOne is a primitive
'             ' a Key:Value pair is being presented
'             InputItemsAsIterableIterable Array(ipParamOne), Array(ipParamTwo), ipStartIndex
        
            
'         Case Arrays.IsTableArray(ipParamOne)
            
'             CaseTwoParamsAsTableArrayvsParamTwo ipParamOne, ipParamTwo, ipStartIndex
        
        
'         Case Types.Group.IsIterableKeysByEnum(ipParamOne) ' Implemented
        
'             CaseTwoParamsAsIterableKeysByEnumvsParamTwo ipParamOne, ipParamTwo, ipStartIndex
            
        
'         Case Types.Group.IsIterableItemsByEnum(ipParamOne) 'Implemented
        
'             CaseTwoParamsAsIterableItemsByEnumVsParamTwo ipParamOne, ipParamTwo, ipStartIndex
    
        
'         Case Else
        
'             Guard Id.InvalidType, True, r.Location & "pvAddOrInsertTwoParams", TypeName(ipParamOne)

        
'     End Select

' End Sub


' '@Description("Adds the items of ipArgArray as individual items")
'  Friend Sub InputItemsThreeOrMoreParams(ByVal ipArgArray As Variant, ByVal ipStartIndex As Long)
    
'     If s.FirstUse Then InjectAutoKey
    
'     InputItemsAsIterableIterable GetAutoKeySequence(UBound(ipArgArray) - LBound(ipArgArray) + 1), ipArgArray, ipStartIndex
    
' End Sub

'@Description("Adds or Inserts the items and keys Lysts contained in Globals.Res")
' Private Sub InputItemsInGlobalRes(ByVal ipInsertAtIndex As Long)
' Attribute InputItemsInGlobalRes.VB_Description = "Adds or Inserts the items and keys Lysts contained in Globals.Res"

'   '@Ignore DefaultMemberRequired
'   Dim myParsedVariant As ResultParsedVariant
'   Set myParsedVariant = Globals.Res.GetResultParsedVariant
  
'   Dim myKeys As Lyst
'   If myParsedVariant.ResultItems.Count = 1 Then
'     ' myParsedVariant only contains an items iterable
'     ' so we need to generate a seuqence of keys from the autokey
'     '@Ignore ValueRequired
    
'     Set myKeys = p.AutoKey.GetKeySequence(myParsedVariant.GetItems.Count)
    
'   Else
    
'     Set myKeys = Lyst.Deb.AddKnownIterableAsItems(myParsedVariant.GetKeys)
    
'   End If
  
'   Dim myItems As Lyst
'   Set myItems = Lyst.Deb.AddKnownIterableAsItems(myParsedVariant.GetItems)
  
'   If ipInsertAtIndex = KvpAdmin.AddInput Then
    
'         '@Ignore DefaultMemberRequired
'         p.Keys.Host.AddRange myKeys.ToArrayList
'         p.Items.Host.AddRange myItems.ToArrayList
        
'     Else
        
'         Dim mystartrun As ResultStartRun
'         ' remember that ipstartindexindex may be negative
'         Ranges.TryStartRunFromAnyStartRunEnd p.Keys, ipInsertAtIndex
'         '@Ignore DefaultMemberRequired
'         Set mystartrun = Globals.Res.GetResultStartRun
'         p.Keys.Host.InsertRange mystartrun.StartIndex - 1, myKeys.ToArrayList
'         p.Items.Host.InsertRange mystartrun.StartIndex - 1, myItems.ToArrayList
        
'     End If
    
' End Sub



Friend Sub AdjustAllItems(ByVal ipAdjustAmount As Variant)
Attribute AdjustAllItems.VB_Description = "Adds or Inserts the items and keys Lysts contained in Globals.Res"
    
    With p.Items
    
        Dim myIndex As Long
        For myIndex = 0 To LastIndex
            
            .Item(myIndex) = .Item(myIndex) + ipAdjustAmount
            
        Next
    
    End With
    
End Sub
' Private Sub pvCase2Arg0IsPrimitive(ByVal ipArgArray As Variant, ByVal ipIndex As Long)

'     ' If two parameters and ipParamOne is a primitive we are adding
'     ' a Key:Value pair
    
'     pvAddOrInsertIterableIterable Array(ipArgArray(0)), Array(ipArgArray(1)), ipIndex
                        
' End Sub

Public Property Get AutoKey() As IAutoKey
    Set AutoKey = p.AutoKey
End Property

' Public Property Set AutoKey(ByVal ipAutoKey As IAutoKey)
    
'     Guard DisallowedOverwrite, Types.IsNotSameInstance(p.AutoKey, Nothing), i.Location & "Set pvAutoKey"
'     Set p.AutoKey = ipAutoKey
    
' End Property

' Private Sub CaseTwoParamsAsIterableKeysByEnumvsParamTwo(ByVal ipParamOne As Variant, ByVal ipParamTwo As Variant, ByVal ipStartIndex As Long)

'     Select Case True
                                    
'         Case Types.Group.IsPrimitive(ipParamTwo), Types.Group.IsAdmin(ipParamTwo)
        
'             'add row first items vs primitive in ipParamTwo
'             InputItemsAsIterableIterable _
'                 ipParamOne.Keys, _
'                 Ranges.GetConstSeries(Types.Iterable.Count(ipParamOne), ipParamTwo), _
'                 ipStartIndex
            
            
'         Case Types.Group.IsIterableItemsByEnum(ipParamTwo)
        
'             ' arg0 first value of rows pairs with  arg1 item
'             InputItemsAsIterableIterable _
'                 ipParamOne.Keys, _
'                 ipParamTwo, _
'                 ipStartIndex
            
                
'         Case Arrays.IsTableArray(ipParamTwo)
        
'             'Pair arg0 firstvalues with split first rank values
'             Dim myArg1RankLysts As Lyst
'             Set myArg1RankLysts = Arrays.ToLystOfRanksAsLyst(ipParamTwo, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsSplit))
    
'             InputItemsAsIterableIterable _
'                 ipParamOne.Keys, _
'                 myArg1RankLysts.Item(Enums.TableToLystActions..RankLysts), _
'                 ipStartIndex
            
            
'         Case Types.Group.IsIterableKeysByEnum(ipParamTwo)
'             ' first value of arg0 rows with arg1 items
'             InputItemsAsIterableIterable _
'             ipParamOne.Keys, _
'             ipParamTwo.Items, _
'             ipStartIndex
            
            
'         Case Else
'             'todo we can't just add object as each item will point to the same object
'             ' so rethink this code
'             InputItemsAsIterableIterable _
'                 ipParamOne, _
'                 Ranges.GetConstSeries(TypesIterables.Count(ipParamOne), ipParamTwo), _
'                 ipStartIndex
'             ' CustomErrors.Throw _
'             '     InvalidType, _
'             '     i.Location & "pvCase2TableArrayvsArg1", _
'             '     Fmt.Txt(Enums.TableToLystActions.ToString(GuardClause.IsInvalidType), "Iterable", TypeName(ipParamTwo))
            
                
'     End Select
    
' End Sub


' Private Sub CaseTwoParamsAsTableArrayvsParamTwo(ByVal iparray As Variant, ByVal ipTableToLystAction As Variant, ByVal ipStartIndex As Long)
    
'     'For use when ipParamTwo is not a primitive number
'     Dim myArg0RankLysts As Lyst
'     Set myArg0RankLysts = Arrays.ToLystOfRanksAsLyst(iparray, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsSplit))
    
'     Select Case True
                        
'         Case Enums.TableToLystActions.HoldsMember(ipTableToLystAction)
        
'             'Add table according to TableLystAction enum in ipParamTwo
'             InputItemsAsTableArray iparray, ipTableToLystAction, ipStartIndex
            
        
'         Case Types.Group.IsPrimitive(ipTableToLystAction), Types.Group.IsAdmin(ipTableToLystAction)
        
'             'add row first items vs primitive in ipParamTwo
'             InputItemsAsIterableIterable _
'                 myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
'                 Ranges.GetConstSeries(myArg0RankLysts.Count, ipTableToLystAction), _
'                 ipStartIndex
            
            
'         Case Types.Group.IsIterableItemsByEnum(ipTableToLystAction)
        
'             ' arg0 first value of rows pairs with  arg1 item
'             InputItemsAsIterableIterable _
'                 myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
'                 ipTableToLystAction, _
'                 ipStartIndex
            
                
'         Case Arrays.IsTableArray(ipTableToLystAction)
        
'             'Pair arg0 firstvalues with split first rank values
'             Dim myArg1RankLysts As Lyst
'             Set myArg1RankLysts = Arrays.ToLystOfRanksAsLyst(ipTableToLystAction, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsSplit))
    
'             InputItemsAsIterableIterable _
'                 myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
'                 myArg1RankLysts.Item(Enums.TableToLystActions.RankLysts), _
'                 ipStartIndex
            
            
'         Case Types.Group.IsIterableKeysByEnum(ipTableToLystAction)
'             ' first value of arg0 rows with arg1 items
'             InputItemsAsIterableIterable _
'             myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
'             ipTableToLystAction.Items, _
'             ipStartIndex
            
            
'         Case Else
'             'todo we can't just add object as each item will point to the same object
'             ' so rethink this code
'             'Assume we have a. n. other object
    
'             InputItemsAsIterableIterable _
'             myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
'             Ranges.GetConstSeries(myArg0RankLysts.Count, ipTableToLystAction), _
'             ipStartIndex
'             ' CustomErrors.Throw _
'             '     InvalidType, _
'             '     i.Location & "pvCase2TableArrayvsArg1", _
'             '     Fmt.Txt(Enums.TableToLystActions.ToString(GuardClause.IsInvalidType), "Iterable", TypeName(ipParamTwo))
            
                
'     End Select
    
' End Sub


' Private Sub CaseTwoParamsAsIterableItemsByEnumVsParamTwo(ByVal ipParamOne As Variant, ByVal ipParamTwo As Variant, ByVal ipStartIndex As Long)
    
'     Select Case True
                        
'         ' Case Types.Group.IsNumber(ipParamTwo)
        
'         '     'Add table according to TableLystAction enum in ipParamTwo
'         '     pvAddOrInsertIterableIterable _
'         '         ipParamOne, _
'         '         Ranges.GetConstSeries(Types.Iterable.GetSize(ipParamOne), ipParamTwo), _
'         '         ipStartIndex
            
        
'         Case Types.Group.IsPrimitive(ipParamTwo), Types.Group.IsAdmin(ipParamTwo)
        
'             'add row first items vs primitive in ipParamTwo
'             InputItemsAsIterableIterable _
'                 ipParamOne, _
'                 Ranges.GetConstSeries(Types.Iterable.Count(ipParamOne), ipParamTwo), _
'                 ipStartIndex
            
            
'         Case Types.Group.IsIterableItemsByEnum(ipParamTwo)
        
'             ' arg0 first value of rows pairs with  arg1 item
'             InputItemsAsIterableIterable _
'                 ipParamOne, _
'                 ipParamTwo, _
'                 ipStartIndex
            
                
'         Case Arrays.IsTableArray(ipParamTwo)
        
'             'Pair arg0 firstvalues with split first rank values
'             Dim myArg1RankLysts As Lyst
'             Set myArg1RankLysts = Arrays.ToLystOfRanksAsLyst(ipParamTwo, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsSplit))
    
'             InputItemsAsIterableIterable _
'                 ipParamOne, _
'                 myArg1RankLysts.Item(Enums.TableToLystActions.RankLysts), _
'                 ipStartIndex
            
            
'         Case Types.Group.IsIterableKeysByEnum(ipParamTwo)
'             ' first value of arg0 rows with arg1 items
'             InputItemsAsIterableIterable _
'             ipParamOne, _
'             ipParamTwo.Items, _
'             ipStartIndex
            
            
'         Case Else
'             'todo we can't just add object as each item will point to the same object
'             ' so rethink this code
'             ' A. n. other object
'             InputItemsAsIterableIterable _
'                 ipParamOne, _
'                 Ranges.GetConstSeries(TypesIterables.Count(ipParamOne), ipParamTwo), _
'                 ipStartIndex
            
                
'     End Select
    
' End Sub




Friend Sub AdjustByIndex(ByVal ipIndexes As Variant, ByVal ipAmount As Single)
    
    If ipAmount = 0 Then Exit Sub
        
    If BailOut.When(Types.IsNotSameInstance(p.Keys, Nothing), Msg.NullReference) Then Exit Sub
    If BailOut.When(Types.Group.IsNotNumber(ipAmount), Msg.InvalidType) Then Exit Sub
    If BailOut.When(Types.Group.IsNotNumber(p.Items.First), Msg.InvalidType) Then Exit Sub
    
    If VBA.IsEmpty(ipIndexes) Then
        
        AdjustbyKey p.Keys, ipAmount
        Exit Sub
        
    End If

    ' Guard InvalidType, Types.IsTypeByName(Types.Name.OfLong, ipIndexes), myLocation, Array("Long", TypeName(s.Values.First), "ipIndexes")
    
    Dim myIndexes As Lyst
    Set myIndexes = Types.Iterable.ToLyst(ipIndexes)
    Dim myIndex As Variant
    For Each myIndex In myIndexes

        ' Ignore any index outside current bounds
        If myIndex > -1 And myIndex < p.Keys.Count Then
            
            ' preserve the type in s.values
            ' we can use vartype as we are only operating on numbers
            Dim myType As VbVarType
            myType = VarType(p.Items.Item(myIndex))
            p.Items.Item(myIndex) = Types.ToVarTypePrimitive(myType, p.Items.Item(myIndex) + ipAmount)
            
        End If
        
    Next
    
            
End Sub


'@Description("Common code for IncByKey and DecByKey")
Friend Sub AdjustbyKey(ByVal ipKeys As Variant, ByVal ipAmount As Variant)
Attribute AdjustbyKey.VB_Description = "Common code for IncByKey and DecByKey"

    'If Guard.IsNothing(p.Keys, Msg.IterableHasNoItems).NotOkay Then Exit Sub
    If BailOut.When(Types.Iterable.HasNoItems(p.Keys), Msg.HasNoItems) Then Exit Sub
    If BailOut.When(Types.Group.IsNotNumber(ipAmount, p.Items.First), Msg.IsNotNumber) Then Exit Sub

    
    Dim myKeys As Variant
    If Arrays.IsNotValid(ipKeys) Then
        
        Set myKeys = p.Keys
        
    ElseIf Types.Group.IsIterable(ipKeys) Then
    
        Set myKeys = p.Keys
        
    Else
    
        Set myKeys = Array(ipKeys)
        
    End If
        
    

    Dim myitem As Variant
    For Each myitem In myKeys

        ' We choose to ignore any key that is not in p.keys
        If p.Keys.HoldsItem(myitem) Then
            
            Dim myIndex As Long
            myIndex = p.Keys.IndexOf(myitem)
            
            ' The code below preserves the Type of the variable stored at ipIndex
            ' e.g. a Decimal Type is classed as a number But decimal +1 yeilds Double Type (or long Type)
            ' so we need jiggery pokery to make sure that the Type stored at ipIndex is not
            ' when incremented or decremented.
            Dim myType As VbVarType
            myType = VarType(p.Items.Item(myIndex))
            p.Items.Item(myIndex) = Types.ToVarTypePrimitive(myType, p.Items.Item(myIndex) + ipAmount)
            
        End If
        
    Next
        
End Sub


Private Sub DoNothing()

End Sub


Friend Function GetAutoKeySequence(ByVal ipCount As Long) As Lyst
    
    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    Dim myCount As Long
    For myCount = 1 To ipCount
    
        myLyst.Add p.AutoKey.GetNextKey
        
    Next
    
    Set GetAutoKeySequence = myLyst
        
End Function


Private Function pvGetStringKeys() As ArrayList

    Dim myKeys As ArrayList
    Set myKeys = New ArrayList

    Dim myitem As Variant
    For Each myitem In p.Keys

        myKeys.Add CStr(myitem)

    Next

    Set pvGetStringKeys = myKeys

End Function


Public Sub InjectAutoKey(ByRef ipAutokey As IAutoKey)
    
    Globals.Res.Init (Msg.Okay)
  'Debug.Print ipAutokey.Name
  Select Case ipAutokey.Name
    'Debug.Print Types.Name.OfAutoKeyByIndex, Types.Name.OfAutoKeyByNumber, Types.Name.OfAutoKeybyString
    Case Types.Name.OfAutoKeyByIndex, Types.Name.OfAutoKeyByNumber, Types.Name.OfAutoKeybyString
    
      Set p.AutoKey = ipAutokey
      
    Case Else
    
        Globals.Res.SetNotOkay
        Globals.Res.MsgEnum = enums.Message.AsEnum(Msg.AutoKeyIsNotValid)
        
  End Select
        
End Sub
