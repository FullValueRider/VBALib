VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Kvp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Kvp is a flexible Key/Value pair dictionary that retains the order of addition and allows values to be accessed by Key or by Index"
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@PredeclaredId
'@Exposed
'@ModuleDescription("Kvp is a flexible Key/Value pair dictionary targeted at VB/VBA. It facilitates data input, direct content manipulation of items and keys, and indirect manipulation of data via Map and Comparer objects")
    'End Enum

    Const mcLocation                    As String = "VBALib.Kvp."
    Const mcDefaultSeparator            As String = Char.comma
    Const mcClassName                   As String = "Kvp"
   ' Const mcKvpError                    As Long = 17 + vbObjectError
    'Const mcParamArrayIsEmpty           As Long = -1
    'Const OptionalDefault               As Long = -1
    Const mcFlagAddItems                As Long = -1

    Private Type State

        Keys                            As Lyst
        Values                          As Lyst
        AutoKey                         As IAutoKey
        Separator                       As String
        FirstUse                        As Boolean
        
    End Type

    Private s                           As State

    '========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
    '
    ' class admin
    '
    '========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
    Public Sub Class_Initialize()
    
            'Guard InvalidNew, Types.IsNotSameInstance(Me, Kvp), mcLocation & "Class_Initialize"
    
    End Sub
    
    
    Public Function Deb(Optional ipKvp As Kvp = Nothing) As Kvp

        'Guard UsePredeclaredIdInstance, Types.IsNotSameInstance(Me, Kvp), mcLocation & "Debe"
        
        With New Kvp
        
            If ipKvp Is Nothing Then
                
                Set Deb = .ReadyToUseInstance
                
                
            Else
                
                Set Deb = .ReadyToUseClone(ipKvp)
                
            End If
            
        End With
        
    End Function


    Public Function ReadyToUseInstance() As Kvp

        'Guard UsePredeclaredIdInstance, Types.IsNotSameInstance(Me, Kvp), mcLocation & "Deb"
        
        Set s.Keys = Lyst.Deb
        Set s.Values = Lyst.Deb
        Set s.AutoKey = Nothing
        s.Separator = mcDefaultSeparator
        s.FirstUse = True
    
        Set ReadyToUseInstance = Me
           
    End Function
    
    
    Public Function ReadyToUseClone(ByVal ipKvp As Kvp) As Kvp
        
        Set s.Keys = ipKvp.GetKeys.Clone
        Set s.Values = ipKvp.GetItems.Clone
        Set s.AutoKey = ipKvp.AutoKey
        s.Separator = ipKvp.Separator
        s.FirstUse = IIf(ipKvp.AutoKey Is Nothing, True, False)
        Set ReadyToUseClone = Me
        
    End Function
    
    '========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
    '
    ' class public members
    '
    '========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
    Public Function Add(ParamArray ipArgs() As Variant) As Kvp

        Set Add = Me
        If Arrays.IsNotValidArray(ipArgs) Then Exit Function
            
        pvAddOrInsert ipArgs, mcFlagAddItems

    End Function


    Public Function Count() As Long
        Count = s.Keys.Count
    End Function


    Public Function Clone() As Kvp
        
        Set Clone = Kvp.Deb(Me)
        
    End Function

    
    
    '@Description("Decrements the values specified by one or more keys or all keys")
    Public Function Dec(Optional ByVal ipKeys As Variant = Empty, Optional ByVal ipDecrement As Variant = DefaultAdjustAmount) As Kvp

        Set Dec = Me
        If ipDecrement = 0 Then Exit Function
        If s.Keys Is Nothing Then Exit Function
            
        Dim myLocation As String
        myLocation = mcLocation & "Dec"
            
        Guard InvalidType, Types.Group.IsNotNumber(ipDecrement), myLocation, Array("Number", TypeName(ipDecrement), "ipDecrement")
        Guard InvalidType, Types.Group.IsNotNumber(s.Values.First), myLocation, Array("Number", TypeName(s.Values.First), "s.values")
        
        
        Dim myKeys As Lyst
        If VBA.IsEmpty(ipKeys) Then
            
            Set myKeys = s.Keys
                
        ElseIf VBA.IsArray(ipKeys) Then
            
            Set myKeys = Types.Iterable.ToLyst(ipKeys)
            
        Else
            
            Guard GuardClause.InvalidIterable, Types.Group.IsNotIterable(ipKeys), myLocation
            Set myKeys = ipKeys
            
        End If
        
        
        ' Debug.Print TypeName(ipKeys(0)), TypeName(s.Keys.First)
        ' Debug.Print Types.IsNotSameType(ipKeys(0), s.Keys.First)
        Guard InvalidType, TypeName(myKeys(0)) <> TypeName(s.Keys.First), myLocation, Array(TypeName(s.Keys.First), TypeName(myKeys(0)), "ipKeys")
        
        pvAdjustbyKey myKeys, -ipDecrement
        
    End Function


    Public Function DecAt(ByVal ipIndexes As Variant, Optional ByVal ipDecrement As Single = DefaultAdjustAmount) As Kvp

        Set DecAt = Me
        If ipDecrement = 0 Then Exit Function
            
        Dim myLocation As String
        myLocation = mcLocation & "Dec"
        Guard NullReference, Types.IsNotSameInstance(s.Keys, Nothing), myLocation
        Guard InvalidType, Types.Group.IsNotNumber(ipDecrement), myLocation, Array("Number", TypeName(ipDecrement), "ipDecrement")
        Guard InvalidType, Types.Group.IsNotNumber(s.Values.First), myLocation, Array("Number", TypeName(s.Values.First), "s.values")
        
        If VBA.IsEmpty(ipIndexes) Then
            
            pvAdjustbyKey s.Keys, -ipDecrement
            Exit Function
            
        End If
    
        Guard InvalidType, Types.IsTypeByName(Types.Name.OfLong, ipIndexes), myLocation, Array("Long", TypeName(s.Values.First), "ipIndexes")
        
        Dim myIndexes As Lyst
        Set myIndexes = Types.Iterable.ToLyst(ipIndexes)

         pvAdjustByIndex myIndexes, -ipDecrement
         
    End Function


    Public Function DeQueue() As KVPair
        Set DeQueue = PullFirst
    End Function


    ' Public Property Get ErrorOutPut() As ErrorReporter
    '     ErrorOutPut = s.ErrorOutPut
    ' End Property

    ' Public Property Let ErrorOutPut(ByVal ipValue As ErrorReporter)
    '     s.ErrorOutPut = ipValue
    ' End Property


    Public Function EnQueue(ParamArray ipargs() As Variant) As Kvp
        
        Set EnQueue = Me
        If Arrays.IsNotValidArray(ipargs) Then Exit Function
            
        pvAddOrInsert ipargs, mcFlagAddItems
        
    End Function

    Public Function FilterByKeys(ByVal ipKeysIterable As Variant) As Kvp
       ' Preference should be given to using the specific KeysIn<xxx> method
       Dim myResult As Kvp
       Set myResult = Kvp.Deb
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


    Public Function GetFirst() As KVPair
        
        Guard HasNoItems, s.Keys.Count = 0, mcLocation & "Getfrst"
        Set GetFirst = KVPair.Deb(s.Keys.Item(0), s.Values.Item(0))
        
    End Function

    '@Description("Returns a Kvp of Key vs Index at which the key is found")
    Public Function GetKeysWithValue(ByVal ipValue As Variant, Optional ByVal ipIndexer As Indexer = Indexer.UseIndex) As Kvp
Attribute GetKeysWithValue.VB_Description = "Returns a Kvp of Key vs Index at which the key is found"
    
        If Not s.Values.HoldsValue(ipValue) Then
        
            Set GetKeysWithValue = Nothing
            Exit Function
            
        End If
        
        Dim myResult As Kvp
        Set myResult = Kvp.Deb

        Dim myIndex As Long
        For myIndex = 0 To Me.LastIndex

            If s.Values.Item(myIndex) = ipValue Then
            
                If ipIndexer - Indexer.UseIndex Then
                
                    myResult.Add s.Keys.Item(myIndex), myIndex
                    
                Else
                    
                    myResult.Add myIndex, s.Keys.Item(myIndex)
                    
                End If
                
            End If
            
        Next

        Set GetKeysWithValue = myResult
        
    End Function


    Public Function GetIndexOfKey(ByVal ipKey As Variant) As Long

        If Types.IsEmpty(ipKey) Then
        
            GetIndexOfKey = -1
            
        Else
            
            GetIndexOfKey = s.Keys.IndexOf(ipKey)
            
        End If
        
    End Function


    '@Description("Returns a kvp of Index vs Key")
    Public Function GetIndexOfValue(ByVal ipValue As Variant, Optional ByVal ipOrganiseBy As Indexer = Indexer.UseIndex) As Kvp
Attribute GetIndexOfValue.VB_Description = "Returns a kvp of Index vs Key"
        's.values may hold more than one item with the same value so we return a Kvp
        Set GetIndexOfValue = Me.MirrorByValue(ipOrganiseBy).Item(ipValue)
    End Function


    Public Function GetKeys() As Lyst
        
        If s.Keys Is Nothing Then
        
            Set GetKeys = Lyst.Deb
            
        Else
        
            Set GetKeys = s.Keys
            
        End If
        
    End Function


    Public Function GetLast() As KVPair
        Set GetLast = KVPair.Deb(s.Keys.Item(Me.LastIndex), s.Values.Item(Me.LastIndex))
    End Function
    

    Public Function GetItems() As Lyst

        If s.Values Is Nothing Then
        
            Set GetItems = Lyst.Deb
            
        Else
        
            Set GetItems = s.Values
            
        End If
        
    End Function

    
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

        Dim myIndex As Long
        myIndex = s.Keys.IndexOf(ipKey) + 1

        If myIndex >= s.Keys.Count Then
        
            If ipTreatKeysArrayAsLoop Then
            
                Set NextPair = KVPair.Deb(s.Keys.Item(0), s.Values.Item(0))
                
            Else
            
                Set NextPair = Nothing
                
            End If
            
        Else
        
            Set NextPair = KVPair.Deb(s.Keys.Item(myIndex), s.Values.Item(myIndex))
            
        End If
        
    End Function


    Public Function PrevPair(ByVal ipKey As Variant, Optional ByVal ipTreatKeysArrayAsLoop As Boolean = True) As KVPair

        Dim myIndex As Long
        myIndex = s.Keys.IndexOf(ipKey) - 1
    
        If myIndex < 0 Then
        
            If ipTreatKeysArrayAsLoop Then
            
                Set PrevPair = KVPair.Deb(s.Keys.Item(Me.LastIndex), s.Values.Item(Me.LastIndex))
                
            Else
            
                Set PrevPair = Nothing
                
            End If
            
        Else
        
            Set PrevPair = KVPair.Deb(s.Keys.Item(myIndex), s.Values.Item(myIndex))
            
        End If
        
    End Function


    Public Function HoldsKey(ByVal Key As Variant) As Boolean
        HoldsKey = s.Keys.HoldsValue(Key)
    End Function


    Public Function HoldsValue(ByVal ipValue As Variant) As Boolean
        HoldsValue = s.Values.HoldsValue(ipValue)
    End Function
    

    Public Function Inc(Optional ByVal ipKeys As Variant = Empty, Optional ByVal ipIncrement As Single = DefaultAdjustAmount) As Kvp

        Set Inc = Me
        If ipIncrement = 0 Then Exit Function
            
        Dim myLocation As String
        myLocation = mcLocation & "Dec"
        Guard NullReference, Types.IsNotSameInstance(s.Keys, Nothing), myLocation
        Guard InvalidType, Types.Group.IsNotNumber(ipIncrement), myLocation, Array("Number", TypeName(ipIncrement), "ipDecrement")
        Guard InvalidType, Types.Group.IsNotNumber(s.Values.First), myLocation, Array("Number", TypeName(s.Values.First), "s.values")
       
        Dim myKeys As Lyst
        If VBA.IsEmpty(ipKeys) Then
            
            Set myKeys = s.Keys
                
        Else
        
            Guard InvalidType, Types.IsNotSameType(ipKeys(0), s.Keys.First), myLocation, Array(TypeName(s.Keys.First), TypeName(myKeys.Item(0)), "ipKeys")
            Set myKeys = Types.Iterable.ToLyst(ipKeys)
            
        End If
        
        pvAdjustbyKey myKeys, ipIncrement
    
    End Function

    
    Public Function IncAt(Optional ByVal ipIndexes As Variant = Empty, Optional ByVal ipIncrement As Single = DefaultAdjustAmount) As Kvp

        Set IncAt = Me
        If ipIncrement = 0 Then Exit Function
            
        Dim myLocation As String
        myLocation = mcLocation & "Dec"
        Guard NullReference, Types.IsNotSameInstance(s.Keys, Nothing), myLocation
        Guard InvalidType, Types.Group.IsNotNumber(ipIncrement), myLocation, Array("Number", TypeName(ipIncrement), "ipDecrement")
        Guard InvalidType, Types.Group.IsNotNumber(s.Values.First), myLocation, Array("Number", TypeName(s.Values.First), "s.values")
        
        If VBA.IsEmpty(ipIndexes) Then
            
            pvAdjustbyKey s.Keys, -ipIncrement
            Exit Function
            
        End If
    
        Guard InvalidType, Types.IsTypeByName(Types.Name.OfLong, ipIndexes), myLocation, Array("Long", TypeName(s.Values.First), "ipIndexes")
        
        Dim myIndexes As Lyst
        Set myIndexes = Types.Iterable.ToLyst(ipIndexes)

         pvAdjustByIndex myIndexes, -ipIncrement
        
    End Function
    
    Public Function InsertAt(ByVal ipStartIndex As Long, ParamArray ipArgs() As Variant) As Kvp

        Set InsertAt = Me
        If Arrays.IsValidArray(ipArgs) Then Exit Function
            
        Dim myLocation As String
        myLocation = mcLocation & "InsertAt"
        Guard IndexOutOfBounds, ipStartIndex < mcFlagAddItems, myLocation, Array("ipstartindex")
        Guard IndexOutOfBounds, ipStartIndex > s.Keys.LastIndex, myLocation, Array("ipStartIndex")
        
        pvAddOrInsert CVar(ipArgs), IIf(ipStartIndex = s.Keys.Count, mcFlagAddItems, ipStartIndex)
        
    End Function


    Public Function InsertAfterKey(ByVal ipAfterKey As Variant, ParamArray ipArgs() As Variant) As Kvp
        
        Set InsertAfterKey = Me
        If Arrays.IsValidArray(ipArgs) Then Exit Function
        
        Dim myLocation As String
        myLocation = mcLocation & "InsertAfterKey"
        Guard UnexpectedEmpty, VBA.IsEmpty(ipAfterKey), myLocation, Array("ipAfterKey")
        Guard InvalidInsertCountZero, s.Keys.Count = 0, myLocation
        
        pvAddOrInsert ipArgs, IIf(ipAfterKey = s.Keys.Last, mcFlagAddItems, s.Keys.IndexOf(ipAfterKey) + 1)
            
    End Function


    Public Function InsertBeforeKey(ByVal ipBeforeKey As Variant, ParamArray ipArgs() As Variant) As Kvp
    
        Set InsertBeforeKey = Me
        If Arrays.IsValidArray(ipArgs) Then Exit Function
        
        Dim myLocation As String
        myLocation = mcLocation & "InsertAfterKey"
        Guard UnexpectedEmpty, VBA.IsEmpty(ipBeforeKey), myLocation, Array("ipBeforeKey")
        Guard InvalidInsertCountZero, s.Keys.Count = 0, myLocation & "InsertBeforeKey"
        
        pvAddOrInsert ipArgs, s.Keys.IndexOf(ipBeforeKey)
        
    End Function


    Public Function InsertFirst(ParamArray ipArgs() As Variant) As Kvp
        
        Set InsertFirst = Me
        If Types.Iterable.LacksItems(ipArgs) Then Exit Function
        pvAddOrInsert ipArgs, 0
        
    End Function


    Public Function InsertLast(ParamArray ipArgs() As Variant) As Kvp
        
        Set InsertLast = Me
        If Types.Iterable.LacksItems(ipArgs) Then Exit Function
        pvAddOrInsert ipArgs, s.Keys.LastIndex
        
    End Function

    Public Function LacksItems() As Boolean
        LacksItems = Count = 0
    End Function

    Public Function HoldsItems() As Boolean
        HoldsItems = Not (Count = 0)
    End Function

    Public Function IsUnique() As Boolean
        IsUnique = Me.MirrorFirstValues.Item(1&).Count = 0
    End Function


    Public Function IsNotUnique() As Boolean
        IsNotUnique = Not IsUnique
    End Function



    '@DefaultMember
    Public Property Get Item(ByVal ipKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
        
        Dim myIndex As Long
        myIndex = s.Keys.IndexOf(ipKey)
        If IsObject(s.Values.Item(myIndex)) Then
        
            Set Item = s.Values.Item(myIndex)
            
        Else
        
            Item = s.Values.Item(myIndex)
        
        End If
        
    End Property

    Public Property Let Item(ByVal ipKey As Variant, ByVal ipValue As Variant)
        
        Guard UnexpectedEmpty, VBA.IsEmpty(ipKey), mcLocation & "Let Item"
        
        Dim myIndex As Long
        myIndex = s.Keys.IndexOf(ipKey)
        If VBA.IsObject(ipValue) Then
        
            Set s.Values.Item(myIndex) = ipValue
            
        Else
        
            s.Values.Item(myIndex) = ipValue
            
        End If
        
    End Property


    Public Property Get ItemAt(ByVal ipIndex As Long) As Variant
        
        Guard GuardClause.IndexLessThanFirstIndex, ipIndex < 0, mcLocation & "ItemAt"
        Guard IndexMoreThanLastIndex, ipIndex > Me.LastIndex, mcLocation & "ItemAt"
        
        If IsObject(s.Values.Item(ipIndex)) Then
        
            Set ItemAt = s.Values.Item(ipIndex)
            
        Else
        
            ItemAt = s.Values.Item(ipIndex)
            
        End If
        
    End Property

    Public Property Set ItemAt(ByVal ipIndex As Long, ByVal ipValue As Variant)
        
        Guard IndexLessThanFirstIndex, ipIndex < 0, mcLocation & "ItemAt"
        Guard IndexMoreThanLastIndex, ipIndex > Me.LastIndex, mcLocation & "ItemAt"

        If IsObject(ipValue) Then
        
            Set s.Values.Item(ipIndex) = ipValue
            
        Else
        
            s.Values.Item(ipIndex) = ipValue
            
        End If
        
    End Property
    
    '@Description("Wrapper for GetValues to allow methods compatibility with Scripting.Dictionaries")
    Public Function Items() As Variant
Attribute Items.VB_Description = "Wrapper for GetValues to allow methods compatibility with Scripting.Dictionaries"
        Items = s.Values.ToArray
    End Function


    '@Description("Wrapper for GetKeys to allow compatibility with Scripting.Dictionaries")
    Public Function Keys() As Variant
Attribute Keys.VB_Description = "Wrapper for GetKeys to allow compatibility with Scripting.Dictionaries"
        Keys = s.Keys.ToArray
    End Function


    Public Function KeysInAOnly(ByVal ipB As Kvp) As Kvp

        Dim myResult As Kvp
        Set myResult = Kvp.Deb

        Dim myKey As Variant
        For Each myKey In s.Keys

            If ipB.LacksKey(myKey) Then
            
                myResult.Add Me.Item(myKey), myKey
                
            End If
            
        Next

        Set KeysInAOnly = myResult
        
    End Function


    Public Function KeysInBOnly(ByVal ipB As Kvp) As Kvp

        Dim myResult As Kvp
        Set myResult = Kvp.Deb

        Dim myKey As Variant
        For Each myKey In ipB.GetKeys

            If Me.LacksKey(myKey) Then
            
                myResult.Add ipB.Item(myKey), myKey
                
            End If
            
        Next

        Set KeysInBOnly = myResult
        
    End Function


    Public Function KeysAllAandOnlyB(ByVal ipB As Kvp) As Kvp

        Dim myResult As Kvp
        Set myResult = Me.Clone
        
        Dim myResultB As Kvp
        Set myResultB = Me.KeysInBOnly(ipB)
        myResult.Add myResultB.GetItems, myResultB.GetKeys
        
        Set KeysAllAandOnlyB = myResult
        
    End Function


    Public Function KeysInAandB(ByVal ipB As Kvp) As Kvp
        
        Dim myResult As Kvp
        Set myResult = Kvp.Deb
        
        Dim myKey As Variant
        For Each myKey In s.Keys
        
            If ipB.HoldsKey(myKey) Then
        
                myResult.Add Me.Item(myKey), myKey
        
            End If
            
        Next
        
        Set KeysInAandB = myResult
        
    End Function


    Public Function KVPairsInAandB(ByVal ipB As Kvp) As Kvp
        ' myInB is used for keys in B that have a different value to A
        Dim myResult As Kvp
        Set myResult = Kvp.Deb
        
        Dim myKey As Variant
        For Each myKey In s.Keys
        
            If ipB.HoldsKey(myKey) Then
        
                If Me.Item(myKey) = ipB.Item(myKey) Then

                    myResult.Add Me.Item(myKey), myKey
                    
                End If
        
            End If
            
        Next
        
        Set KVPairsInAandB = myResult
        
    End Function


    Public Function KeysNotInBothAandB(ByVal ipB As Kvp) As Kvp

        Dim myResult As Kvp
        Set myResult = ipB.Clone
        
        Dim myKey As Variant
        For Each myKey In s.Keys
        
            If myResult.HoldsKey(myKey) Then
            
                myResult.Remove myKey
                
            Else
            
                myResult.Add Me.Item(myKey), myKey
                
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
        
        If s.Keys Is Nothing Then Exit Function
        
        LastIndex = s.Keys.LastIndex
        
    End Function


    'returns a Kvp with values as keys.  Each Value is linked
    ' to a Kvp of Index vs Key
    Public Function MirrorByValue(Optional ByVal ipOrganiseBy As Indexer = Indexer.UseIndex) As Kvp
        
        Dim myResult As Kvp
        Set myResult = Kvp.Deb
        
        Dim myKey  As Variant
        For Each myKey In s.Keys
        
            Dim myIndex As Variant
            myIndex = s.Keys.IndexOf(myKey)
            If myResult.LacksKey(s.Values.Item(myIndex)) Then
            
                myResult.Add Kvp.Deb, s.Values.Item(myIndex)
                
            End If
            
            If ipOrganiseBy = Indexer.UseIndex Then
            
                myResult.Item(s.Values.Item(myIndex)).Add myKey, myIndex
                
            Else
            
                myResult.Item(s.Values.Item(myIndex)).Add myIndex, myKey
                
            End If
            
        Next
        
        Set MirrorByValue = myResult

    End Function

    '@Description("Returns a new Kvp where each value has been transformed by the IMap Object")
'    Public Function Map(ByVal ipMapper As IMap) As Kvp
'
'        Dim myKvp As Kvp
'        Set myKvp = Kvp.Deb
'
'        Dim myIndex As Long
'        Dim myMaxIndex As Long
'        myMaxIndex = Me.LastIndex
'        For myIndex = 0 To myMaxIndex
'
'            myKvp.Add ipMapper.ExecMap(s.Values.Item(myIndex)), s.Keys.Item(myIndex)
'
'        Next
'
'        Set Map = myKvp
'
'    End Function

    Public Function MirrorFirstValues() As Kvp
Attribute MirrorFirstValues.VB_Description = "Returns a new Kvp where each value has been transformed by the IMap Object"

        Dim myFirstValues As Kvp
        Set myFirstValues = Kvp.Deb
        
        Dim myDupValues As Kvp
        Set myDupValues = Kvp.Deb

        Dim myIndex As Long
        For myIndex = 0 To s.Keys.Count - 1
        
            If myFirstValues.LacksKey(s.Values.Item(myIndex)) Then
            
                myFirstValues.Add s.Keys.Item(myIndex), s.Values.Item(myIndex)
                
            Else
            
                myDupValues.Add s.Values.Item(myIndex), s.Keys.Item(myIndex)
                
            End If
            
        Next
            
        Set MirrorFirstValues = Kvp.Deb.Add(Array(myFirstValues, myDupValues))
        
    End Function


    '@Enumerator
    '@Ignore EmptyMethod
    Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    '    Static myKeysIndex As Long
    '    Static myValuesIndex As Long
    '
    '    If myKeysIndex >= s.Keys.Count Then
        
        '  Set NewEnum = KVPair.Deb(s.Keys.[_NewEnum], s.Values.[_NewEnum])
    '        myKeysIndex = 0
    '        myValuesIndex = 0
    '
    '    Else
    '
    '        Set NewEnum = GetPair(myKeysIndex)
    '        myKeysIndex = myKeysIndex + 1
    '        myValuesIndex = myValuesIndex + 1
    '
    '    End If
        
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
    '        If Not s.Keys.ContainsKey(myKey) Then
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
    '    For Each myKey In s.Keys
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
        Name = mcClassName
    End Function

    '@Description("Returns a Kvp of maximum and minimum values.  Keys are the max and min value, associated with the keys is a Kvp of Index vs Key for each value corresponding to the parent key.")
    Public Function NumMaxMin() As Kvp
Attribute NumMaxMin.VB_Description = "Returns a Kvp of maximum and minimum values.  Keys are the max and min value, associated with the keys is a Kvp of Index vs Key for each value corresponding to the parent key."
        
        Guard IsNotNumber, Types.Group.IsNotNumber(s.Values.Item(0)), mcLocation & "NumMaxMin", Array(TypeName(s.Values.Item(0)), Join(Types.Group.Numbers.ToArray, ","))
        
        Dim myResult As Kvp
        Set myResult = Kvp.Deb
        
        Dim myValues As VBAlib.Lyst
        Set myValues = s.Values.Clone
        myValues.Sort
        
        'Item 0 of the KVp is the minimum values
        Const maMin As Long = 0
        myResult.Add Kvp.Deb, maMin
        
        'Item 1 of the Kvp is the maximum values
        Const maMax As Long = 1
        myResult.Add Kvp.Deb, maMax

        ' A maximum or minimum value may occur more than once
        ' so we populate the max and min kvp  with the key vs value
        ' of each found max or min value
        
        Dim myLbound As Long
        myLbound = 0
        
        Dim myUbound As Long
        myUbound = Me.LastIndex
        
        Dim myIndex As Long
        For myIndex = myLbound To myUbound

            Select Case s.Values.Item(myIndex)
            
                Case myValues.Item(myLbound)
                    
                    myResult.Item(maMin).Add s.Values.Item(myIndex), s.Keys.Item(myIndex)
                    
            
                Case myValues.Item(myUbound)
                    
                    myResult.Item(maMax).Add s.Values.Item(myIndex), s.Keys.Item(myIndex)
                    
            End Select
            
        Next

        Set NumMaxMin = myResult
        
        
        
    End Function


    Public Function PullByKey(ByVal ipKey As Variant) As KVPair

        Dim myIndex As Long
        myIndex = s.Keys.IndexOf(ipKey)

        Set PullByKey = KVPair.Deb(ipKey, s.Values.Item(myIndex))
        s.Keys.RemoveAt myIndex
        s.Values.RemoveAt myIndex

    End Function


    Public Function PullAt(ByVal ipIndex As Long) As KVPair

        Set PullAt = KVPair.Deb(s.Keys.Item(ipIndex), s.Values.Item(ipIndex))
        s.Keys.RemoveAt ipIndex
        s.Values.RemoveAt ipIndex

    End Function


    Public Function PullFirst() As KVPair
        
        Set PullFirst = PullAt(0)
        
    End Function


    Public Function PullLast() As KVPair
        
        Set PullLast = PullAt(s.Keys.Count - 1)
        
    End Function


    'Public Function PushAt(ByVal ipIndex As Long, ByVal ipValue As Variant, Optional ByVal ipKey As Variant = Empty)
    '
    '
    'End Function
    '
    'public function PushByKey( byval ipKey as Variant, byval ipValue

    '@Description("Sugar for InsertFirst")
    Public Function PushFirst(ByVal ipValue As Variant, Optional ByVal ipKey As Variant = Empty) As Kvp
Attribute PushFirst.VB_Description = "Sugar for InsertFirst"

        Set PushFirst = InsertFirst(ipValue, ipKey)

    End Function

    '@Description("Sugar for InsertLast")
    Public Function PushLast(ByVal ipValue As Variant, Optional ByVal ipKey As Variant = Empty) As Kvp
Attribute PushLast.VB_Description = "Sugar for InsertLast"

        Set PushLast = InsertLast(ipValue, ipKey)

    End Function


    '@Ignore FunctionReturnValueAlwaysDiscarded
    Public Function Remove(Optional ByVal ipKey As Variant) As Kvp

        Set Remove = Me
        
        If VBA.IsMissing(ipKey) Then
        
            s.Keys.Clear
            s.Values.Clear
            
        Else
        
            Dim myKey As Variant
            For Each myKey In ipKey
            
                Dim myIndex As Long
                myIndex = s.Keys.IndexOf(myKey)
                s.Keys.RemoveAt myIndex
                s.Values.RemoveAt myIndex
                
            Next
            
        End If
        
    End Function


    Public Function RemoveAt(ByVal ipIndex As Long) As Kvp
        
        
        s.Values.RemoveAt ipIndex
        s.Keys.RemoveAt ipIndex
        Set RemoveAt = Me
        
    End Function


    Public Function RemoveAll() As Kvp
        
        s.Keys.Clear
        s.Values.Clear
        
        Set RemoveAll = Me
        
    End Function


    Public Function RemoveFirst() As Kvp
        Set RemoveFirst = RemoveAt(0)
    End Function


    Public Function RemoveLast() As Kvp
        Set RemoveLast = RemoveAt(s.Keys.Count - 1)
    End Function


    Public Function Reverse() As Kvp

        s.Keys.Reverse
        s.Values.Reverse
        
        Set Reverse = Me

    End Function


    Public Property Get Separator() As String
        Separator = s.Separator
    End Property

    Public Property Let Separator(ByVal ipSeparator As String)
        s.Separator = ipSeparator
    End Property


    Public Function SetAutoKeyToByNumber(Optional ByVal ipFirstKey As Variant = 0&) As Kvp

        Guard AutoKeyResetNotAllowed, Not s.FirstUse, mcLocation & "SetAutoKeyByNumber"
        Guard IsNotNumber, Types.Group.IsNotNumber(ipFirstKey), mcLocation & "SetAutoKeyByNumvber"
        
        
        Set s.AutoKey = AutoKeyByNumber.Deb(ipFirstKey)
        s.FirstUse = False
        Set SetAutoKeyToByNumber = Me
        
    End Function


    Public Function SetAutoKeyToByString _
    ( _
        Optional ByVal ipFirstKey As String = AutoKeyAid.DefaultKeyForStringKey, _
        Optional ByVal ipCharSequence As String = AutoKeyAid.DefaultKeyCharSequence _
    ) As Kvp
        
        Guard AutoKeyResetNotAllowed, Not s.FirstUse, mcLocation & "SetAutoKeyByString"
        Guard IsNotString, Types.IsNotString(ipFirstKey), mcLocation & "SetAutoKeyByString"
            
        Set s.AutoKey = AutoKeyByString.Deb(ipFirstKey, ipCharSequence)
        s.FirstUse = False
        Set SetAutoKeyToByString = Me
        
    End Function


    Public Function SetAutoKeyToByIndex(ByVal ipKeysIterable As Variant, Optional ByVal ipStartIndex As Variant = AutoKeyAid.DefaultKeyIndexForArrayList) As Kvp
        
        Set SetAutoKeyToByIndex = Me
        
        Dim myPlace As String
        myPlace = mcLocation & "SetToAutoKeyByIndex"
        Guard AutoKeyResetNotAllowed, Not s.FirstUse, myPlace
        Guard NotIterableType, Types.Group.IsNotIterable(ipKeysIterable), myPlace
        Guard HasNoItems, Types.Iterable.LacksItems(ipKeysIterable), myPlace

    
        Dim myLyst As Lyst
        Set myLyst = Types.Iterable.ToLyst(ipKeysIterable)
        
        Dim myKeysFirstIndex As Long
        Guard GuardClause.IndexLessThanFirstIndex, Not Types.Iterable.TryFirstIndex(ipKeysIterable, myKeysFirstIndex), myPlace
            
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
        
        Set s.AutoKey = AutoKeyByIndex.Deb(myLyst, myStartIndex)
        s.FirstUse = False
        Set SetAutoKeyToByIndex = Me
            
            
    End Function
    
    Public Function SumKeys() As LongLong

        Guard GuardClause.IsNotNumber, Types.Group.IsNotNumber(s.Keys.Item(0)), mcLocation & "SumKeys"

        Dim mySum As LongLong
        Dim myKey As Variant
        For Each myKey In s.Keys
        
            mySum = mySum + myKey
            
        Next

        SumKeys = mySum
        
    End Function


    Public Function SumValues() As LongLong
        
        Guard GuardClause.IsNotNumber, Types.Group.IsNotNumber(s.Values.Item(0)), mcLocation & "SumValues"

        Dim mySum As LongLong
        Dim myValue As Variant
        For Each myValue In s.Values
        
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

    ' Private Sub pvAddIterable(ByVal ipIterable As Variant)
    ' ' The iterable may be a variant array, a collection an arraylist or other
    ' ' one dimensional object that can be enumerate using for each
    '     s.Logger.Push "Private AddIterable"
        
    '     If s.FirstUse Then pvInitialiseAutoKey 0&

    '     Dim myKey As Variant
    '     Dim myValue As Variant
    '     For Each myValue In ipIterable
        
    '         myKey = s.AutoKey.GetNextKey
    '         pvDoAdd myKey, myValue
        
    '     Next
        
    '     s.AutoKey.CurrentKey = myKey
    '     s.Logger.Pop
            
    ' End Sub


    Private Sub pvAddOrInsert(ByVal ipArgsArray As Variant, Optional ByVal ipStartIndex As Long = -1)
        
        
        ' Insertion is defined by ipstartindex being 0 or larger
        
        Select Case UBound(ipArgsArray, 1) + 1
        
    '       Case 0 - Not required as exluded before pvAddOrInsert is called
                
            Case 1
            
                pvAddOrInsertOneParam CVar(ipArgsArray(0)), ipStartIndex 'complete
                
                
            Case 2
            
                pvAddOrInsertTwoParams CVar(ipArgsArray(0)), CVar(ipArgsArray(1)), ipStartIndex
                
                    
            Case Is > 2
                    
                pvAddOrInsertThreeOrMoreParams CVar(ipArgsArray), ipStartIndex ' complete
                
            
        End Select
        
        
    End Sub
            

    Private Sub pvAddOrInsertOneParam(ByVal ipArg As Variant, ByVal ipStartIndex As Long)
        
        ' s.firstuse is tested here becauase some of the cases use s.autokey
        If s.FirstUse Then pvInitialiseAutoKey AutoKeyAid.DefaultKeyForNumberTypes
            
        ' primitive, Kvp and 2d arrays need special treatment
        Select Case True
                
            Case Types.Group.IsIterableKeysByEnum(ipArg)
        
                pvAddOrInsertIterableIterable ipArg.Keys, ipArg.Items, ipStartIndex
                
            
            Case Arrays.IsTableArray(ipArg)
            
                ' No KeysOrientation is specified with one parameter default to KeysOrientation.RankIsColumnCopyFirstItem
                pvAddOrInsertTableArray ipArg, TableToLystAction.RankIsRowFirstItemActionIsNoAction, ipStartIndex
                
                
            Case Types.Group.IsPrimitive(ipArg)
            
                Dim myKey As Variant
                myKey = s.AutoKey.GetNextKey
                
                Dim myKeysArray As Variant
                myKeysArray = Array(myKey)
                
                
                ' Dim myItem As Variant
                ' myItem = CVar(ipArg)
                Dim myItemsArray As Variant
                
                myItemsArray = Array(ipArg)
                'Debug.Print Err.Number, Err.Description
                pvAddOrInsertIterableIterable myKeysArray, myItemsArray, ipStartIndex
                
                
            Case Types.Group.IsIterableItemsByEnum(ipArg)
            
                Dim myKeysSize As Long
                myKeysSize = Types.Iterable.Count(ipArg)
                
                Dim myKeys As Lyst
                Set myKeys = pvGetAutoKeySequence(myKeysSize)
                pvAddOrInsertIterableIterable myKeys, ipArg, ipStartIndex
                
                
            Case Types.IsKVPair(ipArg(0))
            
                Dim myPair As KVPair
                Set myPair = ipArg(0)
                pvAddOrInsertIterableIterable Array(myPair.Key), Array(myPair.Item), ipStartIndex
                
                
            Case Else
                            Guard GuardClause.InvalidIterable, mcLocation & "pvAddOrInsertOneParam", TypeName(ipArg)
                Exit Sub
            
        End Select
        
    End Sub


    Private Sub pvAddOrInsertTwoParams(ByVal ipArg0 As Variant, ByVal ipArg1 As Variant, ByVal ipStartIndex As Long)

    ' To add two iterables as objects, rather than as their component items,
    ' combine the two iterables into an array so that the ParamArray
    ' consists of a single argument which is processed by the else clause
    ' in the Case 1 branch of pvAddorInsert
        Guard GuardClause.InvalidType, Types.Group.IsAdmin(ipArg0), mcLocation & "pvAddOrInsertTwoParams", Array(TypeName(ipArg0)), "Admin types cannot be keys '{0}'"
        
        Select Case True
                    
            Case Types.Group.IsPrimitive(ipArg0) 'Imlemented
        
                ' for two parameters when ipArg0 is a primitive
                ' a Key:Value pair is being presented
                pvAddOrInsertIterableIterable Array(ipArg0), Array(ipArg1), ipStartIndex
            
                
            Case Arrays.IsTableArray(ipArg0)
                
                pvCase2TableArrayvsArg1 ipArg0, ipArg1, ipStartIndex
            
            
            Case Types.Group.IsIterableKeysByEnum(ipArg0) ' Implemented
            
                pvCase2IterableKeysByEnumvsArg1 ipArg0, ipArg1, ipStartIndex
                
            
            Case Types.Group.IsIterableItemsByEnum(ipArg0) 'Implemented
            
                pvCase2IterableItemsByEnumVsArg1 ipArg0, ipArg1, ipStartIndex
        
            
            Case Else
            
                Guard GuardClause.InvalidType, True, mcLocation & "pvAddOrInsertTwoParams", TypeName(ipArg0)

            
    End Select
    
    End Sub

    Public Sub pvAddOrInsertThreeOrMoreParams(ByVal ipArgs As Variant, ByVal ipStartIndex As Long)
        
    ' Parameters in ipArgs are added as individual items
        
        Dim myKeysSize As Long
        myKeysSize = UBound(ipArgs) - LBound(ipArgs) + 1
        
        If s.FirstUse Then pvInitialiseAutoKey
            
        Dim myKeys As Lyst
        Set myKeys = pvGetAutoKeySequence(myKeysSize)
        
        pvAddOrInsertIterableIterable myKeys, ipArgs, ipStartIndex
        
    End Sub

    'Private Sub pvAddOrInsertAutoKeyIterable(ByVal ipIterable As Variant, ByVal ipStartindex As Long)
    '
    '    Dim myAutoKeyLyst As Lyst
    '    Set Autokeylyst = GetAutoKeySequence(iterables.GetSize(ipIterable))
    '
    '    If ipStartindex = -1 Then
    '
    '        s.Keys.AddRange myAutoKeyLyst
    '        s.Values.AddRange ipIterable
    '
    '    Else
    '
    '        s.Keys.InsertRange ipStartindex, myAutoKeyLyst
    '        s.Values.InsertRange ipStartindex, ipIterable
    '
    '    End If
    '
    'End Sub

    Public Sub pvAddOrInsertIterableIterable(ByVal ipIterableKeys As Variant, ByVal ipIterableItems As Variant, ByVal ipStartIndex As Long)

        'Dim myFirstIndex As Long
        ' If s.FirstUse Then
            
        '     pvInitialiseAutoKey IIf(Types.IsTypeByName(Types.Name.OfCollection, ipIterableKeys), 1, 0)
                
        ' End If
    
        If ipStartIndex = -1 Then
        
            s.Keys.AddRange ipIterableKeys
            s.Values.AddRange ipIterableItems
            
        Else
        
            s.Keys.InsertRange ipStartIndex, ipIterableKeys
            s.Values.InsertRange ipStartIndex, ipIterableItems
            
        End If
        
    End Sub

    Public Sub pvAddOrInsertTableArray(ByVal ipTableArray As Variant, ByVal ipAction As TableToLystAction, ByVal ipStartIndex As Long)
        'This method has significant reliance of arrays.trytolystoflysts to preprocess the table
        
        'If ipAction Is Not a member of the TabletoLystAction enum, 
        'then the parameter pair Is treated as nonsensical
        Guard _
            NotEnumerationMember, _
            Enums.TableToLystActions.LacksMember(ipAction), _
            mcLocation + "pvpvAddOrInsertTableArray", _
            Array(TypeName(Enums.TableToLystActions))
        
        'convert the table to a lyst of lysts
        Dim myPreKvp As Lyst
        Set myPreKvp = Enums.TableToLystActions.Exec(ipAction, ipTableArray)
        
        Dim myPreKvpKeys As Lyst
        Dim myPreKvpItems As Lyst
        
        If myPreKvp.Count = 1 Then
            
                ' myPreKvp is a single lyst of rows or columns
                Set myPreKvpKeys = s.AutoKey.GetKeySequence(myPreKvp.First.Count)
                Set myPreKvpItems = myPreKvp
                
                
        Else
            
                ' myPreKvp contains two lysts.  
                ' The first is a list of first row/column items
                ' the second is a list of the rows/columns adjusted as per the split/copy first item requirement
                Set myPreKvpKeys = myPreKvp.First
                Set myPreKvpItems = myPreKvp.Item(1)
                
                
        End If
        
        Dim myIndex As Long
        For myIndex = 0 To myPreKvpKeys.LastIndex
            
            pvAddOrInsertTwoParams myPreKvpKeys.Item(myIndex), Kvp.Deb.Add(myPreKvpItems.Item(myIndex)), ipStartIndex
            
        Next
                    
    End Sub

    Public Sub pvAdjustAll(ByVal ipAdjustAmount As Variant)
        
        With s.Values
        
            Dim myIndex As Long
            For myIndex = 0 To LastIndex
                
                .Item(myIndex) = .Item(myIndex) + ipAdjustAmount
                
            Next
        
        End With
        
    End Sub
    ' Private Sub pvCase2Arg0IsPrimitive(ByVal ipArgs As Variant, ByVal ipIndex As Long)

    '     ' If two parameters and ipArg0 is a primitive we are adding
    '     ' a Key:Value pair
        
    '     pvAddOrInsertIterableIterable Array(ipArgs(0)), Array(ipArgs(1)), ipIndex
                            
    ' End Sub
    
    Public Property Get AutoKey() As IAutoKey
        Set AutoKey = s.AutoKey
    End Property
    
    ' Public Property Set AutoKey(ByVal ipAutoKey As IAutoKey)
        
    '     Guard DisallowedOverwrite, Types.IsNotSameInstance(s.AutoKey, Nothing), mcLocation & "Set pvAutoKey"
    '     Set s.AutoKey = ipAutoKey
        
    ' End Property

    Private Sub pvCase2IterableKeysByEnumvsArg1(ByVal ipArg0 As Variant, ByVal ipArg1 As Variant, ByVal ipStartIndex As Long)

        Select Case True
                                      
            Case Types.Group.IsPrimitive(ipArg1), Types.Group.IsAdmin(ipArg1)
            
                'add row first items vs primitive in ipArg1
                pvAddOrInsertIterableIterable _
                    ipArg0.Keys, _
                    Ranges.GetConstSeries(Types.Iterable.Count(ipArg0), ipArg1), _
                    ipStartIndex
             
                
            Case Types.Group.IsIterableItemsByEnum(ipArg1)
            
                ' arg0 first value of rows pairs with  arg1 item
                pvAddOrInsertIterableIterable _
                    ipArg0.Keys, _
                    ipArg1, _
                    ipStartIndex
                
                 
            Case Arrays.IsTableArray(ipArg1)
            
                'Pair arg0 firstvalues with split first rank values
                Dim myArg1RankLysts As Lyst
                Set myArg1RankLysts = Arrays.ToLystOfRankLysts(ipArg1, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsSplit))
        
                pvAddOrInsertIterableIterable _
                    ipArg0.Keys, _
                    myArg1RankLysts.Item(Enums.TableToLystActions.RankLysts), _
                    ipStartIndex
                
                
            Case Types.Group.IsIterableKeysByEnum(ipArg1)
                ' first value of arg0 rows with arg1 items
                pvAddOrInsertIterableIterable _
                ipArg0.Keys, _
                ipArg1.Items, _
                ipStartIndex
                
                
            Case Else
                'todo we can't just add object as each item will point to the same object
                ' so rethink this code
                pvAddOrInsertIterableIterable _
                    ipArg0, _
                    Ranges.GetConstSeries(TypesIterables.Count(ipArg0), ipArg1), _
                    ipStartIndex
                ' CustomErrors.Throw _
                '     InvalidType, _
                '     mcLocation & "pvCase2TableArrayvsArg1", _
                '     Fmt.Txt(Enums.TableToLystActions.ToString(GuardClause.IsInvalidType), "Iterable", TypeName(ipArg1))
                
                    
        End Select
        
    End Sub


    Private Sub pvCase2TableArrayvsArg1(ByVal ipArray As Variant, ByVal ipTableToLystAction As Variant, ByVal ipStartIndex As Long)
        
        'For use when iparg1 is not a primitive number
        Dim myArg0RankLysts As Lyst
        Set myArg0RankLysts = Arrays.ToLystOfRankLysts(ipArray, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsSplit))
        
        Select Case True
                            
            Case Enums.TableToLystActions.HoldsMember(ipTableToLystAction)
            
                'Add table according to TableLystAction enum in ipArg1
                pvAddOrInsertTableArray ipArray, ipTableToLystAction, ipStartIndex
                
            
            Case Types.Group.IsPrimitive(ipTableToLystAction), Types.Group.IsAdmin(ipTableToLystAction)
            
                'add row first items vs primitive in ipArg1
                pvAddOrInsertIterableIterable _
                    myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
                    Ranges.GetConstSeries(myArg0RankLysts.Count, ipTableToLystAction), _
                    ipStartIndex
             
                
            Case Types.Group.IsIterableItemsByEnum(ipTableToLystAction)
            
                ' arg0 first value of rows pairs with  arg1 item
                pvAddOrInsertIterableIterable _
                    myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
                    ipTableToLystAction, _
                    ipStartIndex
                
                 
            Case Arrays.IsTableArray(ipTableToLystAction)
            
                'Pair arg0 firstvalues with split first rank values
                Dim myArg1RankLysts As Lyst
                Set myArg1RankLysts = Arrays.ToLystOfRankLysts(ipTableToLystAction, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsSplit))
        
                pvAddOrInsertIterableIterable _
                    myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
                    myArg1RankLysts.Item(Enums.TableToLystActions.RankLysts), _
                    ipStartIndex
                
                
            Case Types.Group.IsIterableKeysByEnum(ipTableToLystAction)
                ' first value of arg0 rows with arg1 items
                pvAddOrInsertIterableIterable _
                myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
                ipTableToLystAction.Items, _
                ipStartIndex
                
                
            Case Else
                'todo we can't just add object as each item will point to the same object
                ' so rethink this code
                'Assume we have a. n. other object
        
                pvAddOrInsertIterableIterable _
                myArg0RankLysts.Item(Enums.TableToLystActions.FirstItems), _
                Ranges.GetConstSeries(myArg0RankLysts.Count, ipTableToLystAction), _
                ipStartIndex
                ' CustomErrors.Throw _
                '     InvalidType, _
                '     mcLocation & "pvCase2TableArrayvsArg1", _
                '     Fmt.Txt(Enums.TableToLystActions.ToString(GuardClause.IsInvalidType), "Iterable", TypeName(ipArg1))
                
                    
        End Select
        
    End Sub

    
    Private Sub pvCase2IterableItemsByEnumVsArg1(ByVal ipArg0 As Variant, ByVal ipArg1 As Variant, ByVal ipStartIndex As Long)
        
        Select Case True
                            
            ' Case Types.Group.IsNumber(ipArg1)
            
            '     'Add table according to TableLystAction enum in ipArg1
            '     pvAddOrInsertIterableIterable _
            '         ipArg0, _
            '         Ranges.GetConstSeries(Types.Iterable.GetSize(ipArg0), ipArg1), _
            '         ipStartIndex
                
            
            Case Types.Group.IsPrimitive(ipArg1), Types.Group.IsAdmin(ipArg1)
            
                'add row first items vs primitive in ipArg1
                pvAddOrInsertIterableIterable _
                    ipArg0, _
                    Ranges.GetConstSeries(Types.Iterable.Count(ipArg0), ipArg1), _
                    ipStartIndex
             
                
            Case Types.Group.IsIterableItemsByEnum(ipArg1)
            
                ' arg0 first value of rows pairs with  arg1 item
                pvAddOrInsertIterableIterable _
                    ipArg0, _
                    ipArg1, _
                    ipStartIndex
                
                 
            Case Arrays.IsTableArray(ipArg1)
            
                'Pair arg0 firstvalues with split first rank values
                Dim myArg1RankLysts As Lyst
                Set myArg1RankLysts = Arrays.ToLystOfRankLysts(ipArg1, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsSplit))
        
                pvAddOrInsertIterableIterable _
                    ipArg0, _
                    myArg1RankLysts.Item(Enums.TableToLystActions.RankLysts), _
                    ipStartIndex
                
                
            Case Types.Group.IsIterableKeysByEnum(ipArg1)
                ' first value of arg0 rows with arg1 items
                pvAddOrInsertIterableIterable _
                ipArg0, _
                ipArg1.Items, _
                ipStartIndex
                
                
            Case Else
                'todo we can't just add object as each item will point to the same object
                ' so rethink this code
                ' A. n. other object
                pvAddOrInsertIterableIterable _
                    ipArg0, _
                    Ranges.GetConstSeries(TypesIterables.Count(ipArg0), ipArg1), _
                    ipStartIndex
                
                    
        End Select
        
    End Sub



    ' Private Sub pvAddPairOfIterables(ByVal ipKeysList As VBAlib.Lyst, ByVal ipValuesList As VBAlib.Lyst)

    '     s.Logger.Push "pvAddPairs"
        
    '     If ipKeysList.Count <> ipValuesList.Count Then
        
    '         pvReport IterableSizeMismatch, ipKeysList.Count, ipValuesList.Count
    '         Exit Sub
            
    '     End If
        
    '     If s.FirstUse Then
        
    '         pvInitialiseAutoKey ipKeysList(0)
        
    '     ElseIf Not VBA.IsEmpty(ipKeysList(0)) Then
        
    '         If Types.IsNotSameType(s.AutoKey.FirstKey, ipKeysList(0)) Then
        
    '             pvReport KeyTypeMismatch, s.AutoKey.FirstKeyType, TypeName(ipKeysList(0))
    '             Exit Sub
                
    '         End If
            
    '     End If

    '     ' Use an index as we are iterating two arraylists
    '     Dim myIndex As Long
        
    '     'Debug.Print Iterables.GetSize(ipKeysList)
    '     For myIndex = 0 To ipKeysList.Count - 1
        
    '         If VBA.IsEmpty(ipKeysList.Item(myIndex)) Then
        
    '             pvDoAdd s.AutoKey.GetNextKey, ipValuesList(myIndex)
                
    '         Else
            
    '             pvDoAdd ipKeysList(myIndex), ipValuesList(myIndex)
            
    '         End If
        
    '     Next
        
        'Adding pairs does not use the autokey
        ' the to ensure continuity the autokey.currentkey is set
        ' to the last key added so that the next autokey
        ' requested will be the next in sequence from the
    '     ' keys used in the aded pairs.
    '     s.AutoKey.CurrentKey = s.Keys.Item(s.Keys.Count - 1)
        
    '     s.Logger.Pop
        
    ' End Sub

            
    ' Private Sub pvAddOrInsertTableArray(ByVal ipTable As Variant, Optional ByVal ipEnum As Long = -1, Optional ByVal ipStartIndex As Long = -1)

    '     If ipEnum = Then ipEnum = Enums.TableArrayKeys(RankIsRowSplitFirstItem)
    '     s.Logger.Push "pvAddOrInsertTableArray"
    
        
    '     If s.FirstUse Then
        
    '         pvInitialiseAutoKey myTable.Alpha
        
    '     ElseIf Types.IsNotSameType(s.AutoKey.FirstKeyType, myTable.Alpha) Then
        
    '         pvReport KeyTypeMismatch, s.AutoKey.FirstKeyType, myTable.Alpha
    '         Exit Sub
        
    '     End If
        
        
        
    '     Select Case ipKeysElement
        
    '         Case Orientation.Col1
            
    '             pvAddTableArrayCopyByColumnSeparateKeys myTable
                
    '         Case Orientation.RankIsColumnCopyFirstItem
            
    '             pvAddTableArrayCopyByColCopyKeys myTable
                
    '         Case Orientation.Row1
            
    '             pvAddTableArrayCopyByRowSeparateKeys myTable
                
    '         Case Orientation.CopyRow1
            
    '             pvAddTableArrayCopyByRowCopyKeys myTable
                
    '         Case Else
            
    '             pvReport InvalidTableKeysEnum, ipKeysElement
                
    '     End Select

    '    s.Logger.Pop
        
    ' End Sub


    Private Sub pvAdjustByIndex(ByVal ipIndexes As Variant, ByVal ipAmount As Single)
      
        Dim myIndex As Variant
        For Each myIndex In ipIndexes

            ' Ignore any index outside current bounds
            If myIndex > -1 And myIndex < s.Keys.Count Then
                
                ' preserve the type in s.values
                ' we can use vartype as we are only operating on numbers
                Dim myType As VbVarType
                myType = VarType(s.Values.Item(myIndex))
                s.Values.Item(myIndex) = Types.ToVarTypePrimitive(myType, s.Values.Item(myIndex) + ipAmount)
                
            End If
            
        Next
        
             
    End Sub


    '@Description("Common code for IncByKey and DecByKey")
    Private Sub pvAdjustbyKey(ByVal ipKeys As Variant, ByVal ipAmount As Single)
Attribute pvAdjustbyKey.VB_Description = "Common code for IncByKey and DecByKey"
    
        Dim myItem As Variant
        For Each myItem In ipKeys

            ' We choose to ignore any key that is not in s.keys
            If s.Keys.HoldsValue(myItem) Then
                
                Dim myIndex As Long
                myIndex = s.Keys.IndexOf(myItem)
                
                ' preserve the type in s.values
                ' we can use vartype as we are only operating on numbers
                Dim myType As VbVarType
                myType = VarType(s.Values.Item(myIndex))
                s.Values.Item(myIndex) = Types.ToVarTypePrimitive(myType, s.Values.Item(myIndex) + ipAmount)
                
            End If
            
        Next
        
    End Sub


    ' Private Sub pvDoAdd(ByVal ipKey As Variant, ByVal ipValue As Variant)

    '     s.Logger.Push "pvDoAdd"
    '     If s.Keys.HoldsValue(ipKey) Then
        
    '         pvReport KeyExists, ipKey
    '         Exit Sub
            
    '     End If
        
    '     s.Keys.Add ipKey
    '     s.Values.Add ipValue
            
    ' End Sub


    ' Private Sub pvDoInsert(ByVal ipIndex As Long, ByVal ipKey As Variant, ByVal ipValue As Variant)

    '     If s.Keys.HoldsValue(ipKey) Then
        
    '         pvReport "pvDoInsert", KeyExists, ipKey.peek
    '         Exit Sub
            
    '     End If
        
    '     s.Keys.InsertAt ipIndex, ipKey
    '     s.Values.InsertAt ipIndex, ipValue
            
    ' End Sub


    '@Ignore EmptyMethod
    Private Sub DoNothing()

    End Sub
    
    ' VBA uses (col by row) for 2d arrays
    ' using excel
    '   Range("a1:e1") gives variant(1 to 1,1 to 5) i.e. 1 row
    '   Range("A2:A5") gives variant(1 to 4,1 to 1) i.e. 1 column

    ' Private Sub pvAddTableArrayCopyByColCopyKeys(ByVal ipTable As Variant)

    '     Dim myTableKvp As Kvp
    '     Set myTableKvp = Kvp.Deb
        
    '     Dim myCol As Variant
    '     For Each myCol In ipTable.Columns
        
    '         myTableKvp.Add myCol.Item(0), Kvp.Deb.Add(myCol)
            
    '     Next
        
    '     Me.Add myTableKvp
        
    ' End Sub


    ' Private Sub pvAddTableArrayCopyByRowCopyKeys(ByVal ipTable As Variant)

    '     Dim myTableKvp As Kvp
    '     Set myTableKvp = Kvp.Deb
        
    '     Dim myRows As VBAlib.Lyst
    '     Set myRows = ipTable.Rows
    '     Dim myRow As Variant
    '     For Each myRow In myRows
        
    '         Dim myKey As Collection
    '         Set myKey = New Collection
    '         myKey.Add myRow.Item(0)
            
    '         myTableKvp.Add myKey.Item(1), Kvp.Deb.Add(myRow)
            
    '     Next
        
    '     Me.Add myTableKvp
            
    ' End Sub

    ' Private Sub pvAddTableArrayCopyByRowSeparateKeys(ByVal ipTable As Variant)

    '     Dim myTableKvp As Kvp
    '     Set myTableKvp = Kvp.Deb
        
    '     Dim myRow As Variant
    '     For Each myRow In ipTable.Rows
        
    '         ' the following jiggery pokery is required bycause operations on
    '         ' an arraylist do not return a reference to self thus
    '         ' properties of arraylists cannot be chained.
    '         Dim myKey As Collection
    '         Set myKey = New Collection
    '         myKey.Add myRow.Item(0)
            
    '         myRow.RemoveAt 0
    '         myTableKvp.Add myKey.Item(1), Kvp.Deb.Add(myRow)
            
    '     Next
        
    '     Me.Add myTableKvp
        
    ' End Sub

    ' Private Sub pvAddTableArrayCopyByColumnSeparateKeys(ByVal ipTable As Variant)

    '     Dim myTableKvp As Kvp
    '     Set myTableKvp = Kvp.Deb
        
    '     Dim myCol As Variant
    '     For Each myCol In ipTable.Columns
        
    '         ' the following jiggery pokery is required bycause operations on
    '         ' an arraylist do not return a reference to self thus
    '         ' properties of arraylists cannot be chained.
    '         Dim myKey As Collection
    '         Set myKey = New Collection
    '         myKey.Add myCol.Item(0)
        
    '         myCol.RemoveAt 0
    '         myTableKvp.Add myKey.Item(1), Kvp.Deb.Add(myCol)
            
    '     Next
        
    '     Me.Add myTableKvp
        
    ' End Sub


    Private Function pvGetAutoKeySequence(ByVal ipCount As Long) As Lyst
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        
        Dim myCount As Long
        For myCount = 1 To ipCount
        
            myLyst.Add s.AutoKey.GetNextKey
            
        Next
        
        Set pvGetAutoKeySequence = myLyst
            
    End Function


    Private Function pvGetStringKeys() As ArrayList

        Dim myKeys As ArrayList
        Set myKeys = New ArrayList

        Dim myItem As Variant
        For Each myItem In s.Keys

            myKeys.Add CStr(myItem)

        Next

        Set pvGetStringKeys = myKeys

    End Function


    Public Sub pvInitialiseAutoKey _
    ( _
        Optional ByVal ipKey As Variant = Empty, _
        Optional ByVal ipKeyChars As String = vbNullString, _
        Optional ByVal ipKeysList As Lyst = Nothing _
    )
        s.FirstUse = False
        
        Select Case True
        
            Case VBA.IsEmpty(ipKey)
            
                Set s.AutoKey = AutoKeyByNumber.Deb(AutoKeyAid.DefaultKeyForNumberTypes)
        
            Case Types.Group.IsNumber(ipKey)
        
                If ipKeysList Is Nothing Then
                    
                    Set s.AutoKey = AutoKeyByNumber.Deb(ipKey)
                    
                Else
                    
                    Set s.AutoKey = AutoKeyByIndex.Deb(ipKeysList, ipKey)
                    
                End If
            
            Case Types.IsString(ipKey)
        
                Set s.AutoKey = AutoKeyByString.Deb(ipKey, ipKeyChars)
            
            Case Else
        
                Guard GuardClause.InvalidType, True, mcLocation & "pvInitialiseAutoKey", Array(TypeName(Me))
                Exit Sub
            
        End Select
        
       
        
    End Sub


    ' Private Sub pvInsertIterable(ByVal ipIndex As Long, ByVal ipIterable As Variant)

    '     Dim myIndex As Long
    '     myIndex = ipIndex
    '     s.Logger.Push "InsertIterable"
        
    '     If s.FirstUse Then pvInitialiseAutoKey 0&
        
    '     Dim myValue As Variant
    '     For Each myValue In ipIterable
        
    '         pvDoInsert myIndex, s.AutoKey.GetNextKey, myValue
    '         myIndex = myIndex + 1
            
    '     Next
        
        
    '     s.Logger.Pop
        
    ' End Sub
            
            
    ' Private Sub pvInsertPairOfIterables(ByVal ipIndex As Long, ByVal ipKeysList As VBAlib.Lyst, ByVal ipValuesList As VBAlib.Lyst)
    
    '     s.Logger.Push "Private InsertPairs"
        
    '     ' We have two inputs that will be iterated in a for each loop
    '     ' to avoid having to determin lbound and ubound if either or
    '     ' both iterables are an array its simpler to convert both inputs
    '     ' to an arraylist
        
    ' '    Dim myKeys As ArrayList
    ' '    Set myKeys = Iterables.ToLyst(ipKeysIterable)
    ' '
    ' '    Dim myValues As ArrayList
    ' '    Set myValues = Iterables.ToLyst(ipValuesIterable)
        
    ' '    If Not Arrays.Is1D(ipKeysIterable) Then
    ' '
    ' '        pvReport OneDimensionOnly
    ' '        Exit Sub
    ' '
    ' '    End If
    ' '
    ' '    If Not Arrays.Is1D(ipValuesIterable) Then
    ' '
    ' '        pvReport OneDimensionOnly
    ' '        Exit Sub
    ' '
    ' '    End If
        
    '     If ipKeysList.Count <> ipValuesList.Count Then
        
    '         pvReport IterableSizeMismatch, ipKeysList.Count, ipValuesList.Count
    '         Exit Sub
            
    '     End If
        
    '     If s.FirstUse Then
        
    '         pvInitialiseAutoKey ipKeysList(0)
        
    '     ElseIf ipKeysList.Count > 0 And Types.IsNotSameType(s.AutoKey.FirstKeyType, ipKeysList(0)) Then
        
    '         pvReport KeyTypeMismatch, s.AutoKey.FirstKeyType, TypeName(ipKeysList(0))
    '         Exit Sub
        
    '     End If
        
    '     Dim myIndex As Long
    '     For myIndex = 0 To ipKeysList.Count - 1
        
    '         'Dim myKey As Variant
            
    '         If VBA.IsEmpty(ipKeysList(myIndex)) Then
        
    '             pvDoInsert ipIndex + myIndex, s.AutoKey.GetNextKey, ipValuesList(myIndex)
                
    '         Else
            
    '             pvDoInsert ipIndex + myIndex, ipKeysList(myIndex), ipValuesList(myIndex)
            
    '         End If
            
    '     Next
        
    '     s.Logger.Pop

    ' End Sub
                
                
    ' Private Sub pvInsertTable(ByVal ipIndex As Long, ByVal ipTable As Variant, Optional ByVal ipKeysElement As Orientation = Orientation.Row1)

    '     'todo write code to insrt tables
    '     s.Logger.Push "Private InsertTable"
        
    '     If s.FirstUse Then
        
    '         pvInitialiseAutoKey ipTable(0, 0)
        
    '     ElseIf Types.IsNotSameType(s.AutoKey.FirstKeyType, ipTable(0, 0)) Then
        
    '         pvReport KeyTypeMismatch, s.AutoKey.FirstKeyType, TypeName(ipTable(0, 0))
    '         Exit Sub
        
    '     End If
        
    ' End Sub

    ' Private Sub pvReport(ByVal ipError As KvpError, ParamArray ipArgs() As Variant)

    '     Dim myMethod As String
    '     myMethod = s.Logger.Pop
        
    '     Dim myStack As Variant
    '     myStack = s.Logger.toarray
        
    '     If s.ErrorOutPut = ToDebugPrint Then
        
    '         Debug.Print _
    '             mcClassName _
    '             & Char.Period _
    '             & myMethod
    '         Debug.Print _
    '             Fmt.TxtArr(KvpAid.Item(ipError), CVar(ipArgs)) _
    '             & vbCrLf _
    '             & Stringifier.ToString(myStack, vbCrLf, ArrayMarkup.Deb(vbNullString, vbNullString, vbNullString))
            
    '     Else
        
    '         Err.Raise _
    '             mcKvpError, _
    '             mcClassName & Char.Period & myMethod, _
    '             Fmt.TxtxArr(KvpAid.Item(ipError), CVar(ipArgs)) _
    '                 & vbCrLf _
    '                 & Stringifier.ToString(myStack, vbCrLf, ArrayMarkup.Deb(vbNullString, vbNullString, vbNullString))
        
    '     End If
        
    ' End Sub


