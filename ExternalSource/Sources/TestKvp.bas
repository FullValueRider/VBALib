Attribute VB_Name = "TestKvp"
    '@IgnoreModule
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
    '@TestModule
    '@Folder("Tests")
     
   Public Sub KvpTests()

       myPlace = "TestKvp" & Char.Colon & CurrentComponentName & Char.Period
       'region autokeybynumber
       Test01_AutoKeyNumber_IsObject
       Test02_AutoKeyNumber_IsAutoKeyByNumber
       Test03_AutoKeyNumber_DefaultKey
       Test04_AutoKeyNumber_DefaultKeySequence
       Test05_AutoKeyNumber_StartAtFiveSequence
       Test06_AutoKeyNumber_ResetCurrentKeySequence

       ' region AutoKeyByString
       Test01_AutoKeyByString_IsObject
       Test02_AutoKeyByString_IsAutoKeyByString
       Test03_AutoKeyByString_DefaultKey
       Test04_AutoKeyByString_DefaultKeySequence
       Test05_AutoKeyByString_StartAtaaaaSequence
       Test06_AutoKeyByString_ResetCurrentKeySequence
       Test07_AutoKeyByString_RolloverKeySequence
       Test08_AutoKeyByString_RolloverWithFenceKeySequence
       Test09_AutoKeyByString_AltKeySequenceFirstKey

       ' region AutoKeyByIndex
       Test01_AutoKeyByIndex_IsObject
       Test02_AutoKeyByIndex_IsAutoKeyByIndex
       Test03_AutoKeyByIndex_InitialiseByDebArray
       Test04_AutoKeyByIndex_InitiliseByKeysList
       Test05_AutoKeyByIndex_StartAtIndexTwoWrapAround

       'Region NewKvp
       Test01_NewKvpIsObject
       Test02_NewKvpIsKvp
       Test03_NewKvpHasCountZero

       'regiona pvInitialiseAutoKey
       Test01_pvInitialiseAutoKey_Number
       Test02_pvInitialiseAutoKey_String
       Test03_pvInitialiseAutoKey_Index

       'region pvAddOrInsertIterableIterable
       Test01_pvAddOrInsertIterableIterable_ByAdd
       Test02_pvAddOrInsertIterableIterable_ByInsert

       'region Add
       Test01_Add_SingleAdd_GetKeys
       Test02_Add_SingleAdd_GetValues
       Test03_Add_MultiAdd_GetKeys
       Test04_Add_MultiAdd_GetValues
       Test05_Add_SingleAdd_Start0000_GetKeys
       Test06_Add_Single_Add_Start0000_GetValues
       Test07_Add_MultipleAdd_Start0000_GetKeys
       Test08_Add_MultipleAdd_Start0000_GetValues
       Test09_Add_SingleAdd_StartOneHundred_SingleAddGetKeys
       Test10_Add_SingleAdd_StartOneHundred_SingleAddGetValues
       Test11_Add_MultipleAdd_StartOneHundred_MultiAddGetKeys
       Test12_Add_MultipleAdd_StartOneHundred_MultiAddGetValues
       Test13_SingleAdd_Startaaaa_GetKeys
       Test14_Add_SingleAdd_Startaaaa_GetValues
       Test15_Add_MultipleAdd_Startaaaa_GetKeys
       Test16_Add_MultipleAdd_Startaaaa_GetValues
       Test17_AddArray_DefaultLongKeyZero_ArrayOfLong_GetKeys
       Test18_AddArray_DefaultLongKeyZero_ArrayOfLong_GetValues
       Test19_AddArrayDefaultStringKey0000_ArrayOfLong_GetKeys
       Test20_AddArray_DefaultStringKey0000_ArrayOfLong_GetValues
       Test21_AddArray_DefinedLongKeytwenty_ArrayOfLong_GetKeys
       Test22_AddArray_DefinedLongKeyTwenty_ArrayOfLong_GetValues
       Test23_AddArray_DefinedStringKeyaaaa_ArrayOfLong_GetKeys
       Test24_AddArray_DefinedStringKeyaaaa_ArrayOfLong_GetValues
       Test25_Add_DefaultLongKeyZero_ArrayOfLong_GetKeys
       Test26_Add_DefaultLongKeyZero_ArrayOfLong_GetValues
       Test27_Add_DefaultStringKey0000_ArrayOfLong_GetKeys
       Test28_Add_DefaultStringKey0000_ArrayOfLong_GetValues
       Test29_Add_DefinedLongKeytwenty_ArrayOfLong_GetKeys
       Test30_Add_DefinedLongKeyTwenty_ArrayOfLong_GetValues
       Test31_Add_DefinedStringKeyaaaa_ArrayOfLong_GetKeys
       Test32_Add_DefinedStringKeyaaaa_ArrayOfLong_GetValues
       Test33_Add_FourPairs_GetValues
       Test34_Add_FourPairs_GetKeys
       Test35_Add_AddArrayIterableArrayIterable_GetValues
       Test36_Add_AddIterableArrayIterableArray_GetKeys
       Test37_Add_ByRow_UseRow1AsKeysAndData
       Test38_Add_ByColumn_UseCol1AsKeysAndData
       Test39_Add_ByRows_UseRow1AsKeysOnly
       Test40_Add_byColumn_UseCol1AsKeysOnly



       'region Test01_Clone
       Test01_Clone_KeysAreSame
       Test02_Clone_ValuesAreSame
       Test03_Clone_OfNewKvpSucceeds

       'region Count
       Test01_Count_NewKvpIsZero
       Test02_Count_PopulatedKvpIsFive
       Test03_Count_EmptiedKvpIsZero


       'region DecAll
       Test01_DecAll_Item5DefaultOne
       Test02_DecAll_Item5SpecifyOne
       Test03_DecAll_Item5Specify3

       'region Dec ?duplicate?
       Test01_Dec_Items147DefaultOne
       Test02_Dec_Item147SpecifyOne
       Test03_Dec_Item147Specify3


       'region queue
       Test01_Enqueue_DefaultLongKey
       Test02_Enqueue_SpecifiedLongKey10
       Test03_Enqueue_DefaultStringKey
       Test04_Enqueue_SpecifiedStringKeyaaaa
       Test05_KeysAsKeysArrayDefaultStartIndex
       Test06_KeysAsKeysArraySpecifiedStartIndex4
       Test07_Dequeue_DefaultLongKey
       Test08_Dequeue_SpecifiedLongKey10
       Test09_Dequeue_DefaultStringKey
       Test10_Dequeue_SpecifiedStringKeyaaaa
       Test11_ArrayKey_DefaultKeyIndex
       Test12_ArrayKey_SpecifiedKeyIndex4

       'region FilterByKeys
       'Test01_FilterByKeys

       ' 'region Item
       Test01_Item_DefaultKey_Key3
       Test02_Item_LongKey_Key3
       Test03_Item_SpecifiedLongKey_Key13
       Test04_Item_DefaultStringKey_Key0003
       Test05_Item_SpecifiedStringKey_Keyaaad

       ' 'region GetFirst
       Test01_GetFirst_DefaultKey
       Test02_GetFirst_SpecifiedLongKey
       Test03_GetFirst_SpecifiedLongKey_Key10
       Test03_Get_First_DefaultStringKey
       Test04_GetFirst_SpecifiedStringKey_Keyaaaa
       Test05_GetFirst_KeysArray_DefaultStartIndex

       ' 'region getindexofkey
       Test01_GetIndexOfKey_DefaultKey_IndexOf3
       Test02_GetIndexOfKey_SpecifiedDefaultLongKey_IndexOf3
       Test03_GetIndexOfKey_SpecifiedLongKey10_IndexOf3
       Test05_GetIndexOfKey_SpecifiedStringKeyaaaa_IndexOfaaad
       Test05_GetIndexOfKey_SpecifiedStringKeyaaaa_IndexOfaaad
       Test06_GetIndexOfKey_SpecifiedKeysArray_IndexOf4point4

       ' 'region Test01_GetIndexOfValue_DefaultKey_IndexOf40
       Test01_GetIndexOfValue_DefaultKey_IndexOf40
       Test02_GetIndexOfValue_SpecifiedDefaultLongKey_IndexOf40
       Test03_GetIndexOfValue_SpecifiedLongKey10_IndexOf40
       Test04_GetIndexOfValue_DefaultStringKey_IndexOf40
       Test05_GetIndexOfValue_SpecifiedStringKeyaaaa_IndexOf40
       Test06_GetIndexOfValue_SpecifiedKeysArray_IndexOf40

       ' 'region getkeys 'Empty

       ' 'region GetKeysWithValue
       Test01_GetKeysWithValue_GetTwoValues

       ' 'region holdskey
       Test01_HoldsKey_NumberKeyIsPresent
       Test02_HoldsKey_NumberKeyIsMissing
       Test03_HoldsKey_StringKeyIsPresent
       Test04_HoldsKey_StringKeyIsMissing
       Test05_HoldsKey_ArrayKeyIsPresent
       Test06_HoldsKey_ArrayKeyIsMissing

       ' 'region holdsitems
       Test01_HoldsItem_NumberKeyValueIsPresent
       Test02_HoldsItem_NumberKeyValueIsMissing
       Test03_HoldsItem_StringKeyValueIsPresent
       Test04_HoldsItem_StringKeyValueIsMissing
       Test05_HoldsItem_ArrayKeyValueIsPresent
       Test06_HoldsItem_ArrayKeyValueIsMissing

       ' 'region inc
       Test01_Inc_Item5DefaultOne
       Test02_Inc_Item5SpecifyOne
       Test03_Inc_Item5Specify3

       ' 'region incall
       Test01_IncAll_Item5DefaultOne
       Test02_IncAll_Item5SpecifyOne
       Test03_IncAll_Item5Specify3

       ' 'region IncByKeys
       Test01_IncByKeys_Items147DefaultOne
       Test02_IncByKeys_Item147SpecifyOne
       Test03_IncByKeys_Item147Specify3

       ' 'region insertafterkey
       Test01_InsertAfterKey_LastItemLongKey
       Test02_InsertAfterKey_IncByKeysLastItemStringKey
       Test03_InsertAfterKey_FirstItemLongKey
       Test04_InsertAfterKey_FirstItemStringKey
       Test05_InsertAfterKey_MidItemLongKey
       Test06_InsertAfterKey_MidItemStringKey
       Test07_InsertAfterKey_MidItemUserLongKey
       Test08_InsertAfterKey_MidItemUserStringKey

       ' 'region insertbeforekey
       Test01_InsertBeforeKey_LastItemLongKey
       Test02_InsertBeforeKey_LastItemStringKey
       Test03_InsertBeforeKey_FirstItemLongKey
       Test04_InsertBeforeKey_FirstItemStringKey
       Test05_InsertBeforeKey_MidItemLongKey
       Test06_InsertBeforeKey_MidItemStringKey
       Test07_InsertBeforeKey_MidItemUserLongKey
       Test08_InsertBeforeKey_MidItemUserStringKey

       ' 'region insertfirst
       Test01_InsertFirst_DefaultKeys
       Test02_InsertFirst_DefaultNumberKey
       Test03_InsertFirst_SpecifiedNumberKey
       Test04_InsertFirst_SpecifiedNumberKeySpeciedKey
       Test05_InsertFirst_SpecifiedNumberKeySpeciedKeyTwoInserts
       Test06_InsertFirst_SpecifiedStringKeyDefaultKey
       Test06_InsertFirst_SpecifiedStringKeySpecifiedKey

       ' 'region insertlast
       Test01_InsertLast_DefaultKeys
       Test02_InsertLast_DefaultNumberKey
       Test03_InsertLast_SpecifiedNumberKey
       Test04_InsertLast_SpecifiedNumberKeySpeciedKey
       Test05_InsertLast_SpecifiedNumberKeySpeciedKeyTwoInserts
       Test06_InsertLast_SpecifiedStringKeyDefaultKey
       Test07_InsertLast_SpecifiedStringKeySpecifiedKey

       ' 'region intminmax
       Test01_IntMinMax_DefaultKeys

       ' 'region invert

        Test01_Invert_DefaultKeys

       ' 'region holdsitems
       Test01_LacksItems_IsTrue
       Test02_LacksItems_PopulatedKvpIsFalse
       Test03_LacksItems_EmptiedKvpIsTrue
       Test01_HoldsItems_NewKvpIsFalse
       Test02_HoldsItems_PopulatedKvpIsTrue
       Test03_HoldsItems_EmptiedKvpIsFalse

       ' 'region itemat
       Test01_ItemAt_DefaultKey_IndexAt3
       Test02_ItemAt_LongKey_IndexAt3
       Test03_ItemAt_SpecifiedLongKey_IndexAt3
       Test04_ItemAt_DefaultStringKey_IndexAt3
       Test05_ItemAt_SpecifiedStringKey_IndexAt3
       Test06_ItemAt_SpecifiedKeysArray_IndexAt3

       ' 'region keysallaandbonly
       Test01_KeysAllAandOnlyB_DefaultKeysLongKey
       Test02_KeysAllAandOnlyB_DefaultNumberKeysLongKey
       Test03_KeysAllAandOnlyB_SpecifiedNumberKeysLongKey

       ' 'region keysinaonly
       Test01_KeysInAOnly_DefaultKeys
       Test02_KeysInAOnly_DefaultNumberKeysLongKey
       Test03_KeysInAOnly_SpecifiedNumberKeysLongKey
       Test04_KeysInAOnly_SpecifiedStringKeysDefaultKey
       Test05_KeysInAOnly_SpecifiedStringKeysSpecifiedKey

       ' 'region keysinbonly
       Test01_KeysInBOnly_DefaultKeys
       Test02_KeysInBOnly_DefaultNumberKeysLongKey
       Test03_KeysInBOnly_SpecifiedNumberKeysLongKey
       Test04_KeysInBOnly_SpecifiedStringKeysDefaultKey
       Test05_KeysInBOnly_SpecifiedStringKeysSpecifiedKey

       ' 'region keysinbothandb
       Test01_KeysInBothAandB_DefaultKeys
       Test02_KeysInBothAandB_DefaultNumberKeysLongKey
       Test03_KeysInBothAandB_SpecifiedNumberKeysLongKey
       Test04_KeysInBothAandB_SpecifiedStringKeysDefaultKey
       Test05_KeysInBothAandB_SpecifiedStringKeysSpecifiedKey

       ' 'Region "KeysNotInBothAandB"
       Test01_KeysNotInBothAandB_DefaultKeys
       Test02_KeysNotInBothAandB_DefaultNumberKeysLongKey
       Test03_KeysNotInBothAandB_SpecifiedNumberKeysLongKey
       Test04_KeysNotInBothAandB_SpecifiedStringKeysDefaultKey
       Test05_KeysNotInBothAandB_SpecifiedStringKeysSpecifiedKey

       ' 'region lackskey
       Test01_LacksKey_NumberKeyIsPresent
       Test02_LacksKey_NumberKeyIsMissing
       Test03_LacksKey_StringKeyIsPresent
       Test04_LacksKey_StringKeyIsMissing
       Test05_LacksKey_ArrayKeyIsPresent
       Test06_LacksKey_ArrayKeyIsMissing

       ' 'region lacksitem
       Test01_LacksItem_umberKeyValueIsPresent
       Test02_LacksItem_NumberKeyValueIsMissing
       Test03_LacksItem_StringKeyValueIsPresent
       Test04_LacksItem_StringKeyValueIsMissing
       Test05_LacksItem_ArrayKeyValueIsPresent
       Test06_LacksItem_ArrayKeyValueIsMissing


       ' 'region mirrorbyitem
       Test01_MirrorByItem_DefaultKey_DefaultIntegerByIndex
       Test02_MirrorByItem_DefaultKey_DefaulStringByIndex

       ' 'region mirrorbyfirstitem
       Test01_MirrorByFirstItem_DefaultKey_DefaultIntegerByIndex
       Test01_MirrorByFirstItem_DefaultKey_DefaultStringByIndex

       ' 'region nextpair
       Test01_NextPair_DefaultNoSpecifiedKey
       Test02_NextPair_DefaultNumberIntegerKey
       Test03_NextPair_SpecifiedNumberIntegerKey
       Test04_NextPair_DefaultStringKey
       Test05_NextPair_SpecifiedStringKey

       ' 'region prevpair
       Test01_PrevPair_DefaultNoSpecifiedKey
       Test02_PrevPair_DefaultNumberIntegerKey
       Test03_PrevPair_SpecifiedNumberIntegerKey
       Test04_PrevPair_DefaultStringKey
       Test05_PrevPair_SpecifiedStringKey

       ' 'region pull
       Test01_Pull_DefaultKeys
       Test02_Pull_NumberKeysDefaultKeys
       Test03_Pull_NumberKeysSpecifiedKeys
       Test04_Pull_StringKeysDefaultKeys
       Test05_Pull_StringKeysSpecifiedKeys


       ' 'region pullat
       Test01_PullAt_DefaultKeys
       Test02_PullAt_NumberKeysDefaultKeys
       Test03_PullAt_NumberKeysSpecifiedKeys
       Test04_PullAt_StringKeysDefaultKeys
       Test05_PullAt_StringKeysSpecifiedKeys


       'region pullfirst
       Test01_PullFirst_DefaultKeys
       Test02_PullFirst_NumberKeysDefaultKeys
       Test03_PullFirst_NumberKeysSpecifiedKeys
       Test04_PullFirst_StringKeysDefaultKeys
       Test05_PullFirst_StringKeysSpecifiedKeys

       'region pulllast
       Test01_PullLast_DefaultKeys
       Test02_PullLast_NumberKeysDefaultKeys
       Test03_PullLast_NumberKeysSpecifiedKeys
       Test04_PullLast_StringKeysDefaultKeys
       Test05_PullLast_StringKeysSpecifiedKeys

       'region removeall
       Test01_RemoveAll_DefaultKeys
       Test02_RemoveAll_NumberKeysDefaultKeys
       Test03_RemoveAll_NumberKeysSpecifiedKeys
       Test04_RemoveAll_StringKeysDefaultKeys
       Test04_RemoveAll_StringKeysSpecifiedKeys

       'region remove
       Test01_Remove_DefaultKeys
       Test02_Remove_NumberKeysDefaultKeys
       Test03_Remove_NumberKeysSpecifiedKeys
       Test04_Remove_StringKeysDefaultKeys
       Test05_Remove_StringKeysSpecifiedKeys

       'region removeat
       Test01_RemoveAt_DefaultKeys
       Test02_RemoveAt_NumberKeysDefaultKeys
       Test03_RemoveAt_NumberKeysSpecifiedKeys
       Test04_RemoveAt_StringKeysDefaultKeys
       Test05_RemoveAt_StringKeysSpecifiedKeys

       'region removefirst
       Test01_RemoveFirst_DefaultKeys
       Test02_RemoveFirst_NumberKeysDefaultKeys
       Test03_RemoveFirst_NumberKeysSpecifiedKeys
       Test04_RemoveFirst_StringKeysDefaultKeys
       Test05_RemoveFirst_StringKeysSpecifiedKeys

       'region removelast
       Test01_RemoveLast_efaultKeys
       Test02_RemoveLast_NumberKeysDefaultKeys
       Test03_RemoveLast_NumberKeysSpecifiedKeys
       Test04_RemoveLast_StringKeysDefaultKeys
       Test05_RemoveLast_StringKeysSpecifiedKeys

       'region sumkeys 'empty
       Test01_SumKeys_IsObject

       'region sumvalues 'empty
       Test01_SumItems_IsObject

       'region uniqueitems
       Test01_UniqueItems_DefaultKeysTrue
       Test02_UniqueItems_DefaultKeysFalse
       Test01_NotUniqueItems_DefaultKeysTrue
       Test02_NotUniqueItems_DefaultKeysFalse

       Debug.Print CurrentComponentName & vbTab & vbTab & vbTab & " testing completed"

   End Sub
    
'#Region "AutoKeyByNumber"
    '@TestMethod("AutoKeyNumeric")
    Public Sub Test01_AutoKeyNumber_IsObject()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByNumber.Deb()
        
        Dim myExpected As Boolean
        myExpected = True
        
        Dim myResult As Variant
        
        'Act:
        myResult = VBA.IsObject(myAutoKey)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("AutoKeyNumeric")
    Public Sub Test02_AutoKeyNumber_IsAutoKeyByNumber()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByNumber.Deb()
        
        Dim myExpected As String
        myExpected = "AutoKeyByNumber"
        
        Dim myResult As String
        
        'Act:
        myResult = VBA.TypeName(myAutoKey)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("AutoKeyByNumber")
    Public Sub Test03_AutoKeyNumber_DefaultKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByNumber.Deb
       
        Dim myExpected As Long
        myExpected = 0&
        
         Dim myResult As Long
        'Act:
        myResult = myAutoKey.GetNextKey
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("AutoKeyByNumber")
    Public Sub Test04_AutoKeyNumber_DefaultKeySequence()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByNumber.Deb
       
        Dim myExpected As Variant
        myExpected = Array(0, 1, 2, 3, 4, 5)
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        
        Dim myResult As Variant
        'Act:
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
    '@TestMethod("AutoKeyByNumber")
    Public Sub Test05_AutoKeyNumber_StartAtFiveSequence()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByNumber.Deb
        myAutoKey.FirstKey = 5&
        Dim myExpected As Variant
        myExpected = Array(5&, 6&, 7&, 8&, 9&, 10&)
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        
        Dim myResult As Variant
        'Act:
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
    '@TestMethod("AutoKeyByNumber")
    Public Sub Test06_AutoKeyNumber_ResetCurrentKeySequence()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByNumber.Deb
        myAutoKey.FirstKey = 5&
        Dim myExpected As Variant
        myExpected = Array(5&, 6&, 7&, 101&, 102&, 103&)
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            myAutoKey.CurrentKey = 100&
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        
        Dim myResult As Variant
        'Act:
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
 '#End Region
    
'#Region "AutoKeyByString"
 '@TestMethod("AutoKeyByString")
    Public Sub Test01_AutoKeyByString_IsObject()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByString.Deb()
        
        Dim myExpected As Boolean
        myExpected = True
        
        Dim myResult As Variant
        
        'Act:
        myResult = VBA.IsObject(myAutoKey)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("AutoKeyByString")
    Public Sub Test02_AutoKeyByString_IsAutoKeyByString()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByString.Deb()
        
        Dim myExpected As String
        myExpected = "AutoKeyByString"
        
        Dim myResult As String
        
        'Act:
        myResult = VBA.TypeName(myAutoKey)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("AutoKeyByString")
    Public Sub Test03_AutoKeyByString_DefaultKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByString.Deb
       
        Dim myExpected As String
        myExpected = "0000"
        
         Dim myResult As String
        'Act:
        myResult = myAutoKey.GetNextKey
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("AutoKeyByString")
    Public Sub Test04_AutoKeyByString_DefaultKeySequence()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByString.Deb
       
        Dim myExpected As Variant
        myExpected = Array("0000", "0001", "0002", "0003", "0004", "0005")
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        
        Dim myResult As Variant
        'Act:
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
    '@TestMethod("AutoKeyByString")
    Public Sub Test05_AutoKeyByString_StartAtaaaaSequence()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByString.Deb
        myAutoKey.FirstKey = "aaaa"
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaab", "aaac", "aaad", "aaae", "aaaf")
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        
        Dim myResult As Variant
        'Act:
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
    '@TestMethod("AutoKeyByString")
    Public Sub Test06_AutoKeyByString_ResetCurrentKeySequence()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByString.Deb
        myAutoKey.FirstKey = "aaaa"
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaab", "aaac", "bbbe", "bbbf", "bbbg")
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            myAutoKey.CurrentKey = "bbbd"
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        
        Dim myResult As Variant
        'Act:
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
     
    '@TestMethod("AutoKeyByString")
    Public Sub Test07_AutoKeyByString_RolloverKeySequence()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByString.Deb
        myAutoKey.FirstKey = "zzzx"
        Dim myExpected As Variant
        myExpected = Array("zzzx", "zzzy", "zzzz", "10000", "10001", "10002")
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        
        Dim myResult As Variant
        'Act:
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
    '@TestMethod("AutoKeyByString")
    Public Sub Test08_AutoKeyByString_RolloverWithFenceKeySequence()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByString.Deb
        myAutoKey.FirstKey = "aaa/zzzx"
        Dim myExpected As Variant
        myExpected = Array("aaa/zzzx", "aaa/zzzy", "aaa/zzzz", "aaa/10000", "aaa/10001", "aaa/10002")
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        
        Dim myResult As Variant
        'Act:
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
    '@TestMethod("AutoKeyByString")
    Public Sub Test09_AutoKeyByString_AltKeySequenceFirstKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByString.Deb("kkkk", "k1z2m3n4o5p6")
       
        Dim myExpected As Variant
        myExpected = Array("kkkk", "kkk1", "kkkz", "kkk2", "kkkm", "kkk3")
    
        
        Dim myResult As Variant
        'Act:
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
   '#End Region
    
'#Region "AutoKeyByIndex"
   '@TestMethod("AutoKeyByIndex")
    Public Sub Test01_AutoKeyByIndex_IsObject()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByIndex.Deb()
        
        Dim myExpected As Boolean
        myExpected = True
        
        Dim myResult As Variant
        
        'Act:
        myResult = VBA.IsObject(myAutoKey)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("AutoKeyByIndex")
    Public Sub Test02_AutoKeyByIndex_IsAutoKeyByIndex()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByIndex.Deb()
        
        Dim myExpected As String
        myExpected = "AutoKeyByIndex"
        
        Dim myResult As String
        
        'Act:
        myResult = VBA.TypeName(myAutoKey)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("AutoKeArray")
    Public Sub Test03_AutoKeyByIndex_InitialiseByDebArray()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByIndex.Deb(Array("One", "Two", "Three", "Four", "Five", "Six"))
       
        Dim myExpected As String
        myExpected = "One"
        
         Dim myResult As String
        'Act:
        myResult = myAutoKey.GetNextKey
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("AutoKeArray")
    Public Sub Test04_AutoKeyByIndex_InitiliseByKeysList()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByIndex.Deb.SetKeysList(Lyst.Deb.Add("One", "Two", "Three", "Four", "Five", "Six"))
       
        Dim myExpected As String
        myExpected = "One"
        
        Dim myResult As String
        'Act:
        myResult = myAutoKey.GetNextKey
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
    '@TestMethod("AutoKeArray")
    Public Sub Test05_AutoKeyByIndex_StartAtIndexTwoWrapAround()
        'On Error GoTo TestFail
    
        'Arrange:
        Dim myAutoKey As IAutoKey
        Set myAutoKey = AutoKeyByIndex.Deb(Lyst.Deb.Add("One", "Two", "Three", "Four", "Five", "Six"), 2)
        myAutoKey.FirstKey = "Three"
        Dim myExpected As Variant
        myExpected = Array("Three", "Four", "Five", "Six", "One", "Two")
        
        Dim myLyst As Lyst
        Set myLyst = Lyst.Deb
        With myLyst
        
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            .Add myAutoKey.GetNextKey
            
        End With
        
        Dim myResult As Variant
        'Act:
        myResult = myLyst.ToArray
        
        'Assert:
        Assert.Strict.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
'#End Region
   
'#Region "NewKvp"
    
    Public Sub Test01_NewKvpIsObject()
        'On Error GoTo TestFail
        
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Boolean
        myExpected = True
        
       
        Dim myResult As Boolean
        
        'Act:
        myResult = VBA.IsObject(myKvp)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        
    End Sub
    
    Public Sub Test02_NewKvpIsKvp()
        'On Error GoTo TestFail
        
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As String
        myExpected = "Kvp"
        
       
        Dim myResult As String
        
        'Act:
        myResult = VBA.TypeName(myKvp)
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName & ":AddTable"

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        
    End Sub
    
    Public Sub Test03_NewKvpHasCountZero()
        'On Error GoTo TestFail
        
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Long
        myExpected = 0
        
       
        Dim myResult As Long
        
        'Act:
        myResult = myKvp.Count
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName & ":AddTable"

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
        
    End Sub
    
'#End Region

'#Region "pvInitialiseAutoKey"
    '@TestMethod("pvInitialiseAutoKey")
    Public Sub Test01_pvInitialiseAutoKey_Number()
        'On Error GoTo TestFail
        
        'Arrange:
        
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        ' the 5& is deliberate
        myKvp.InjectAutoKey 5&
        
        Dim myExpected As Long
        myExpected = 5
        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvp.AutoKey.GetNextKey
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
    '@TestMethod("pvInitialiseAutoKey")
    Public Sub Test02_pvInitialiseAutoKey_String()
        'On Error GoTo TestFail
        
        'Arrange:
        
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        ' the 5& is deliberate
        myKvp.InjectAutoKey "Hello World"
        
        Dim myExpected As String
        myExpected = "Hello World"
        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvp.AutoKey.GetNextKey
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
    
    '@TestMethod("pvInitialiseAutoKey")
    Public Sub Test03_pvInitialiseAutoKey_Index()
        'On Error GoTo TestFail
        
        'Arrange:
        
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        ' the 5& is deliberate
      '  myKvp.InjectAutoKey 2, ipKeysList:=Lyst.Deb.AddRange(Array("Hello World 0", "Hello World 1", "Hello World 2", "Hello World 3"))
        
        Dim myExpected As String
        myExpected = "Hello World 2"
        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvp.AutoKey.GetNextKey
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ' ' , myPlace & CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region

'#Region "pvAddOrInsertIterableIterable"
    Public Sub Test01_pvAddOrInsertIterableIterable_ByAdd()
        'On Error GoTo TestFail
        
        Dim myBaseItems As Variant
        myBaseItems = Array(10, 20, 30, 40, 50)
        
        Dim myBaseKeys As Variant
        myBaseKeys = Array(0&, 1&, 2&, 3&, 4&)
       
        Dim myExpectedKeys As Variant
        myExpectedKeys = _
            Array _
            ( _
                myBaseKeys, _
                Types.Iterable.ToCollection(myBaseKeys), _
                Types.Iterable.ToArrayList(myBaseKeys), _
                Types.Iterable.ToLyst(myBaseKeys) _
            )
        
        Dim myExpectedItems As Variant
        myExpectedItems = _
            Array _
            ( _
                myBaseItems, _
                Types.Iterable.ToCollection(myBaseItems), _
                Types.Iterable.ToArrayList(myBaseItems), _
                Types.Iterable.ToLyst(myBaseItems) _
            )
            
        Dim myResult As Variant
        Dim myResultItems As Variant
        
        'Act:
        Dim myKeys As Variant
        For Each myKeys In myExpectedKeys
        
            Dim myItems As Variant
            For Each myItems In myExpectedItems
                
                Dim myKvp As Kvp
                Set myKvp = Kvp.Deb
                myKvp.RemoveAll
                
                myKvp.InputItemsAsIterableIterable myKeys, myItems, -1
                myResult = myKvp.Keys
                myResultItems = myKvp.Items
                
                
                'Assert:
                Assert.SequenceEquals myBaseKeys, myResult  ', TypeName(myKeys) & Char.Space & myPlace & CurrentProcedureName & ":Keys"
                Assert.SequenceEquals myBaseItems, myResultItems  ', TypeName(myItems) & Char.Space & myPlace & CurrentProcedureName & ":Items"
                
            Next
            
        Next
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
     
    End Sub

    Public Sub Test02_pvAddOrInsertIterableIterable_ByInsert()
        'On Error GoTo TestFail
        
        Dim myBaseItems As Variant
        myBaseItems = Array(10, 20, 30, 40, 50)
        
        Dim myBaseKeys As Variant
        myBaseKeys = Array(0&, 1&, 2&, 3&, 4&)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(100&, 0&, 1&, 2&, 3&, 4&, 101&, 102&)
       
        Dim myExpectedItems As Variant
        myExpectedItems = Array(1000, 10, 20, 30, 40, 50, 2000, 3000)
        Dim myTestKeys As Variant
        myTestKeys = _
            Array _
            ( _
                myBaseKeys, _
                Types.Iterable.ToCollection(myBaseKeys), _
                Types.Iterable.ToArrayList(myBaseKeys), _
                Types.Iterable.ToLyst(myBaseKeys) _
            )
        
        Dim myTestItems As Variant
        myTestItems = _
            Array _
            ( _
                myBaseItems, _
                Types.Iterable.ToCollection(myBaseItems), _
                Types.Iterable.ToArrayList(myBaseItems), _
                Types.Iterable.ToLyst(myBaseItems) _
            )
            
        Dim myResult As Variant
        Dim myResultItems As Variant
        
        'Act:
        Dim myKeys As Variant
        For Each myKeys In myTestKeys
        
            Dim myItems As Variant
            For Each myItems In myTestItems
                
                Dim myKvp As Kvp
                Set myKvp = Kvp.Deb.Add(Array(100&, 101&, 102&), Array(1000, 2000, 3000))
                
                myKvp.InputItemsAsIterableIterable myKeys, myItems, 1
                myResult = myKvp.Keys
                myResultItems = myKvp.Items
                
                'Assert:
                Assert.SequenceEquals myExpectedKeys, myResult  ', TypeName(myKeys) & Char.Space & myPlace & CurrentProcedureName & ":Keys"
                Assert.SequenceEquals myExpectedItems, myResultItems ', TypeName(myItems) & Char.Space & myPlace & CurrentProcedureName & ":Items"
                
            Next
            
        Next
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test   raised an error: #" & Err.Number & " - " & Err.Description
     
    End Sub
'#End Region

'#Region "add"

    '@TestMethod("Kvp")
    Private Sub Test01_Add_SingleAdd_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array(0&, 1&, 2&, 3&, 4&, 5&, 6&)
        
        Dim myResult As Variant
        
        'Act:
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
        
TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Add_SingleAdd_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
        
        Dim myResult As Variant
        
        'Act:
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Add_MultiAdd_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array(0&, 1&, 2&, 3&, 4&, 5&, 6&)
        
        Dim myResult As Variant
        
        'Act:
        With myKvp
        
            .Add 0&, "Hello"
            .Add 1&, "There"
            .Add 2&, "World"
            .Add 3&, "Its"
            .Add 4&, "A"
            .Add 5&, "Nice"
            .Add 6&, "Day"
            
        End With
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_Add_MultiAdd_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
        
        Dim myResult As Variant
        
        'Act:
        With myKvp
        
            .Add 0&, "Hello"
            .Add 1&, "There"
            .Add 2&, "World"
            .Add 3&, "Its"
            .Add 4&, "A"
            .Add 5&, "Nice"
            .Add 6&, "Day"
            
        End With
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_Add_SingleAdd_Start0000_GetKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("0000", "0001", "0002", "0003", "0004", "0005", "0006")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.SetAutoKeyToAutoKeyByString
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test06_Add_Single_Add_Start0000_GetValues()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
        
        Dim myResult As Variant
        
        'Act:
        
        myKvp.SetAutoKeyToAutoKeyByString "aaaa"
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test07_Add_MultipleAdd_Start0000_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("0000", "0001", "0002", "0003", "0004", "0005", "0006")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.SetAutoKeyToAutoKeyByString "0000"
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
            
        End With
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test08_Add_MultipleAdd_Start0000_GetValues()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
        
        Dim myResult As Variant
        
        'Act:
        
        myKvp.SetAutoKeyToAutoKeyByString "0000"
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test09_Add_SingleAdd_StartOneHundred_SingleAddGetKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array(100&, 101&, 102&, 103&, 104&, 105&, 106&)
        
        Dim myResult As Variant
        
        'Act:
        
        myKvp.SetAutoKeyToAutoKeyByNumber 100&
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test10_Add_SingleAdd_StartOneHundred_SingleAddGetValues()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.SetAutoKeyToAutoKeyByNumber 100&
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test11_Add_MultipleAdd_StartOneHundred_MultiAddGetKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array(100&, 101&, 102&, 103&, 104&, 105&, 106&)
        
        Dim myResult As Variant
        
        'Act:
        
        myKvp.SetAutoKeyToAutoKeyByNumber 100&
        With myKvp
        
        .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test12_Add_MultipleAdd_StartOneHundred_MultiAddGetValues()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
        
        Dim myResult As Variant
        
        'Act:
        
        myKvp.SetAutoKeyToAutoKeyByNumber 100&
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test13_SingleAdd_Startaaaa_GetKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaab", "aaac", "aaad", "aaae", "aaaf", "aaag")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.SetAutoKeyToAutoKeyByString "aaaa"
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test14_Add_SingleAdd_Startaaaa_GetValues()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
        
        Dim myResult As Variant
        
        'Act:
        
        myKvp.SetAutoKeyToAutoKeyByString "aaaa"
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test15_Add_MultipleAdd_Startaaaa_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaab", "aaac", "aaad", "aaae", "aaaf", "aaag")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.SetAutoKeyToAutoKeyByString "aaaa"
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test16_Add_MultipleAdd_Startaaaa_GetValues()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array("Hello", "There", "World", "Its", "A", "Nice", "Day")
        
        Dim myResult As Variant
        
        'Act:
        
        myKvp.SetAutoKeyToAutoKeyByString "aaaa"
        With myKvp
        
            .Add "Hello"
            .Add "There"
            .Add "World"
            .Add "Its"
            .Add "A"
            .Add "Nice"
            .Add "Day"
            
        End With
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test17_AddArray_DefaultLongKeyZero_ArrayOfLong_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array(0&, 1&, 2&, 3&, 4&, 5&, 6&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.Add Array(100, 200, 300, 400, 500, 600, 700)
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test18_AddArray_DefaultLongKeyZero_ArrayOfLong_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array(100, 200, 300, 400, 500, 600, 700)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.Add Array(100, 200, 300, 400, 500, 600, 700)

        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test19_AddArrayDefaultStringKey0000_ArrayOfLong_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByString
        
        Dim myExpected As Variant
        
        myExpected = Array("0000", "0001", "0002", "0003", "0004", "0005", "0006")
        Dim myResult As Variant
        
        'Act:
        myKvp.Add Array(100, 200, 300, 400, 500, 600, 700)
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test20_AddArray_DefaultStringKey0000_ArrayOfLong_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByString
        
        Dim myExpected As Variant
        myExpected = Array(100, 200, 300, 400, 500, 600, 700)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.Add (Array(100, 200, 300, 400, 500, 600, 700))
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test21_AddArray_DefinedLongKeytwenty_ArrayOfLong_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByNumber 20&
        
        Dim myExpected As Variant
        myExpected = Array(20&, 21&, 22&, 23&, 24&, 25&, 26&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.Add Array(100, 200, 300, 400, 500, 600, 700)
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test22_AddArray_DefinedLongKeyTwenty_ArrayOfLong_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByNumber 20&
        Dim myExpected As Variant
        myExpected = Array(100, 200, 300, 400, 500, 600, 700)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.Add Array(100, 200, 300, 400, 500, 600, 700)

        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test23_AddArray_DefinedStringKeyaaaa_ArrayOfLong_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByString "aaaa"
        
        Dim myExpected As Variant
        
        myExpected = Array("aaaa", "aaab", "aaac", "aaad", "aaae", "aaaf", "aaag")
        Dim myResult As Variant
        
        'Act:
        myKvp.Add Array(100, 200, 300, 400, 500, 600, 700)
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test24_AddArray_DefinedStringKeyaaaa_ArrayOfLong_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByString "aaaa"
        
        Dim myExpected As Variant
        myExpected = Array(100, 200, 300, 400, 500, 600, 700)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.Add (Array(100, 200, 300, 400, 500, 600, 700))
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test25_Add_DefaultLongKeyZero_ArrayOfLong_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array(0&, 1&, 2&, 3&, 4&, 5&, 6&)
        
    Dim myResult As Variant
        
        Dim myColl As Collection
        Set myColl = New Collection
        With myColl
        
            .Add 100
            .Add 200
            .Add 300
            .Add 400
            .Add 500
            .Add 600
            .Add 700
            
        End With
        
        'Act:
        myKvp.Add myColl
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test26_Add_DefaultLongKeyZero_ArrayOfLong_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        
        Dim myExpected As Variant
        myExpected = Array(100, 200, 300, 400, 500, 600, 700)
        
        Dim myResult As Variant
        
        Dim myColl As Collection
        Set myColl = New Collection
        With myColl
        
            .Add 100
            .Add 200
            .Add 300
            .Add 400
            .Add 500
            .Add 600
            .Add 700
            
        End With
        'Act:
        myKvp.Add myColl

        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test27_Add_DefaultStringKey0000_ArrayOfLong_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByString
        
        Dim myExpected As Variant
        
        myExpected = Array("0000", "0001", "0002", "0003", "0004", "0005", "0006")
        Dim myResult As Variant
        
        Dim myColl As Collection
        Set myColl = New Collection
        With myColl
        
            .Add 100
            .Add 200
            .Add 300
            .Add 400
            .Add 500
            .Add 600
            .Add 700
            
        End With
        
        'Act:
        myKvp.Add myColl
        
        myResult = myKvp.GetKeys.ToArray
        
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test28_Add_DefaultStringKey0000_ArrayOfLong_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByString
        
        Dim myExpected As Variant
        myExpected = Array(100, 200, 300, 400, 500, 600, 700)
        
        Dim myResult As Variant
        Dim myColl As Collection
        Set myColl = New Collection
        With myColl
        
            .Add 100
            .Add 200
            .Add 300
            .Add 400
            .Add 500
            .Add 600
            .Add 700
            
        End With
        
        'Act:
        myKvp.Add myColl
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test29_Add_DefinedLongKeytwenty_ArrayOfLong_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByNumber 20&
        
        Dim myExpected As Variant
        myExpected = Array(20&, 21&, 22&, 23&, 24&, 25&, 26&)
        
        Dim myResult As Variant
        Dim myColl As Collection
        Set myColl = New Collection
        With myColl
        
            .Add 100
            .Add 200
            .Add 300
            .Add 400
            .Add 500
            .Add 600
            .Add 700
            
        End With
        
        'Act:
        myKvp.Add myColl
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test30_Add_DefinedLongKeyTwenty_ArrayOfLong_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByNumber 20&
        Dim myExpected As Variant
        myExpected = Array(100, 200, 300, 400, 500, 600, 700)
        
        Dim myResult As Variant
        Dim myColl As Collection
        Set myColl = New Collection
        With myColl
        
            .Add 100
            .Add 200
            .Add 300
            .Add 400
            .Add 500
            .Add 600
            .Add 700
            
        End With
        
        'Act:
        myKvp.Add myColl

        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test31_Add_DefinedStringKeyaaaa_ArrayOfLong_GetKeys()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByString "aaaa"
        
        Dim myExpected As Variant
        
        myExpected = Array("aaaa", "aaab", "aaac", "aaad", "aaae", "aaaf", "aaag")
        Dim myResult As Variant
        Dim myColl As Collection
        Set myColl = New Collection
        With myColl
        
            .Add 100
            .Add 200
            .Add 300
            .Add 400
            .Add 500
            .Add 600
            .Add 700
            
        End With
        
        'Act:
        myKvp.Add myColl
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test32_Add_DefinedStringKeyaaaa_ArrayOfLong_GetValues()

        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb
        myKvp.SetAutoKeyToAutoKeyByString "aaaa"
        
        Dim myExpected As Variant
        myExpected = Array(100, 200, 300, 400, 500, 600, 700)
        
        Dim myResult As Variant
        Dim myColl As Collection
        Set myColl = New Collection
        With myColl
        
            .Add 100
            .Add 200
            .Add 300
            .Add 400
            .Add 500
            .Add 600
            .Add 700
            
        End With
        
        'Act:
        myKvp.Add myColl
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test33_Add_FourPairs_GetValues()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        
        Dim myExpected As Variant
        myExpected = Array("one", "three", "five", "seven")
        Dim myResult As Variant
        'Act:
        Set myKvp = Kvp.Deb
        With myKvp
        
            .Add 1, "one"
            .Add 3, "three"
            .Add 5, "five"
            .Add 7, "seven"
            
        End With
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
        
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test34_Add_FourPairs_GetKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        
        Dim myExpected As Variant
        myExpected = Array(1, 3, 5, 7)
        Dim myResult As Variant
        'Act:
        Set myKvp = Kvp.Deb
        With myKvp
        
            .Add 1, "one"
            .Add 3, "three"
            .Add 5, "five"
            .Add 7, "seven"
            
        End With
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
        
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test35_Add_AddArrayIterableArrayIterable_GetValues()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        
        Dim myExpected As Variant
        myExpected = Array("one", "three", "five", "seven")
        Dim myResult As Variant
        'Act:
        Set myKvp = Kvp.Deb
        myKvp.Add Array(1, 3, 5, 7), Array("one", "three", "five", "seven")
        
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
        
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test36_Add_AddIterableArrayIterableArray_GetKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        
        Dim myExpected As Variant
        myExpected = Array(1, 3, 5, 7)
        Dim myResult As Variant
        'Act:
        Set myKvp = Kvp.Deb
        myKvp.Add Array(1, 3, 5, 7), Array("one", "three", "five", "seven")
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."
        
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test37_Add_ByRow_UseRow1AsKeysAndData()
        'On Error GoTo TestFail
        
        'Arrange:
        ' 2d array coordinate system is (no of columns, no of rows)
        Dim myArray(1 To 6, 1 To 5) As Long
        Dim myCol As Long
        For myCol = 1 To 5
            
            Dim myRow As Long
            For myRow = 1 To 6
            
                myArray(myRow, myCol) = (myRow + 3) * (myCol + 1)
                
            Next

        Next
        ' the above gives the following table
        '   8   12  16  20  24
        '   10  15  20  25  30
        '   12  18  24  30  36
        '   14  21  28  35  42
        '   16  24  32  40  48
        '   18  27  36  45  54
        
        Dim myKvp As Kvp: Set myKvp = Kvp.Deb
        Dim myExpected As Variant
        ' The keys are the first value of each row
        myExpected = Array(8&, 10&, 12&, 14&, 16&, 18&)
        Dim myResult As Variant
        'Act:
        myKvp.Add myArray, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsCopy)
        myResult = myKvp.GetKeys.ToArray
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."

TestExit:
        Exit Sub
        
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
        
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test38_Add_ByColumn_UseCol1AsKeysAndData()
        'On Error GoTo TestFail
        
        'Arrange:
        ' 2d array coordinate system is (no of columns, no of rows)
        Dim myArray(1 To 6, 1 To 5) As Long
        Dim myCol As Long
        For myCol = 1 To 5
            
            Dim myRow As Long
            For myRow = 1 To 6
            
                myArray(myRow, myCol) = (myRow + 3) * (myCol + 1)
                
            Next

        Next
        ' the above gives the following table
        '   8   12  16  20  24
        '   10  15  20  25  30
        '   12  18  24  30  36
        '   14  21  28  35  42
        '   16  24  32  40  48
        '   18  27  36  45  54
        
        Dim myExpected As Variant
        myExpected = Array(8&, 12&, 16&, 20&, 24&)
        Dim myKvp As Kvp
        Dim myResult As Variant
        'Act:
        
        Set myKvp = Kvp.Deb.Add(myArray, Enums.TableToLystActions.AsEnum(RankIsColumnFirstItemActionIsCopy))
        myResult = myKvp.GetKeys.ToArray
        'Assert:
        
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."

TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test39_Add_ByRows_UseRow1AsKeysOnly()
        'On Error GoTo TestFail
        
        'Arrange:
        ' 2d array coordinate system is (no of columns, no of rows)
        Dim myArray(1 To 6, 1 To 5) As Long
        Dim myCol As Long
        For myCol = 1 To 5
            
            Dim myRow As Long
            For myRow = 1 To 6
            
                myArray(myRow, myCol) = (myRow + 3) * (myCol + 1)
                
            Next

        Next
        ' the above gives the following table
        '   8   12  16  20  24
        '   10  15  20  25  30
        '   12  18  24  30  36
        '   14  21  28  35  42
        '   16  24  32  40  48
        '   18  27  36  45  54
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(8&, 10&, 12&, 14&, 16&, 18&)
        Dim MyExpectedRow1values As Variant
        MyExpectedRow1values = Array(12&, 16&, 20&, 24&)
        Dim myKvp As Kvp
        Dim myResult As Variant
        Dim MyResultRow1values As Variant
        'Act:
        
        Set myKvp = Kvp.Deb.Add(myArray, Enums.TableToLystActions.AsEnum(RankIsRowFirstItemActionIsSplit))
        
        
        myResult = myKvp.GetKeys.ToArray
        MyResultRow1values = myKvp.Item(8&).Items
        'Assert:
        Assert.Permissive.SequenceEquals myExpectedKeys, myResult  ' ' , myPlace & CurrentProcedureName & ":Keys"
        Assert.Permissive.SequenceEquals MyExpectedRow1values, MyResultRow1values ' ' , myPlace & CurrentProcedureName & ":First"
TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test40_Add_byColumn_UseCol1AsKeysOnly()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myArray(4, 4) As Long
        Dim myRow As Long
        For myRow = 0 To 4
            
            Dim myCol As Long
            For myCol = 0 To 4
            
                myArray(myRow, myCol) = (myCol + 2) * (myRow + 1)
            Next
            
        Next
        
        
        Dim myKvp As Kvp: Set myKvp = Kvp.Deb
        Dim myExpected As Variant
        myExpected = Array(2&, 3&, 4&, 5&, 6&)
        Dim myResult As Variant
        
        'Act:
        myKvp.Add myArray, Enums.TableToLystActions.AsEnum(RankIsColumnFirstItemActionIsSplit)
        myResult = myKvp.GetKeys.ToArray
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ' ' , myPlace & CurrentProcedureName & "."

TestExit:
        Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "CLone"
    Private Sub Test01_Clone_KeysAreSame()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        myKvp.Add Split("Hello,There,World", ",")
        
        Dim myExpected As Variant
        myExpected = myKvp.GetKeys.ToArray
        
        
        Dim myResult As Variant
        Dim myClone As Kvp
        
        
        'Act:
        Set myClone = myKvp.Clone
        myResult = myClone.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Clone_ValuesAreSame()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        myKvp.Add Split("Hello,There,World", ",")
        
        Dim myExpected As Variant
        myExpected = myKvp.GetItems.ToArray
        
        
        Dim myResult As Variant
        Dim myClone As Kvp
        
        
        'Act:
        Set myClone = myKvp.Clone
        myResult = myClone.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test03_Clone_OfNewKvpSucceeds()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb
        
        
        Dim myExpected As Long
        myExpected = 0
        
        
        Dim myResult As Long
        '@Ignore VariableNotUsed
        Dim myClone As Kvp
        
        
        'Act:
        On Error Resume Next
        '@Ignore AssignmentNotUsed
        Set myClone = myKvp.Clone
        myResult = Err.Number
        'On Error GoTo 0
        
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
'#End Region

'#Region "Count"

    '@TestMethod("Kvp")
    Private Sub Test01_Count_NewKvpIsZero()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                           As Kvp
        Dim myExpected                      As Long
        myExpected = 0
        
        Set myKvp = Kvp.Deb
        Dim myResult                        As Long
        
        'Act:
        myResult = myKvp.Count
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Count_PopulatedKvpIsFive()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                           As Kvp
        Dim myExpected                      As Long
        myExpected = 5
        
        Set myKvp = Kvp.Deb.Add(Array(1, 2, 3, 4, 5))
        Dim myResult                        As Long
        
        'Act:
        myResult = myKvp.Count
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Count_EmptiedKvpIsZero()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                           As Kvp
        Dim myExpected                      As Long
        myExpected = 0
        
        Set myKvp = Kvp.Deb.Add(Array(1, 2, 3, 4, 5))
        myKvp.RemoveAll
        
        Dim myResult                        As Long
        
        'Act:
        myResult = myKvp.Count
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region

'#Region "DecByKey"

    '@TestMethod("Kvp")
    Private Sub Test01_Dec_Item5DefaultOne()
    'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60, 70, 80, 90))
        
        Dim myExpected As Variant
        myExpected = 59
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Dec 5&
    
        myResult = myKvp.Item(5&)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
        
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Dec_Item5SpecifyOne()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(Array(1, 2, 3, 4, 5, 6, 7, 8, 9))
        
        Dim myExpected As Long
        myExpected = 5
        
        Dim myResult As Long
        'Act:
        
        myKvp.Dec 5&, 1
    
        myResult = myKvp.Item(5&)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Dec_Item5Specify3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(Array(1, 2, 3, 4, 5, 6, 7, 8, 9))
        
        Dim myExpected As Long
        myExpected = 3
        
        Dim myResult As Long
        'Act:
        
        myKvp.Dec 5&, 3
    
        myResult = myKvp.Item(5&)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region

'#Region "DecAll"

    '@TestMethod("Kvp")
    Private Sub Test01_DecAll_Item5DefaultOne()
    'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(0, 1, 2, 3, 4, 5, 6, 7, 8)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Dec
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_DecAll_Item5SpecifyOne()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(0, 1, 2, 3, 4, 5, 6, 7, 8)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Dec ipDecrement:=1
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_DecAll_Item5Specify3()
        'On Error GoTo TestFail
        
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(-2, -1, 0, 1, 2, 3, 4, 5, 6)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Dec ipDecrement:=3
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "Dec" '?Duplicate?

    '@TestMethod("Kvp")
    Private Sub Test01_Dec_Items147DefaultOne()
    'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        'myKvp.add  (Array(1, 2, 3, 4, 5, 6, 7, 8, 9))
        Dim myExpected As Variant
        myExpected = Array(1, 1, 3, 4, 4, 6, 7, 7, 9)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Dec Array(1&, 4&, 7&)
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Dec_Item147SpecifyOne()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(1, 1, 3, 4, 4, 6, 7, 7, 9)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Dec Array(1&, 4&, 7&), 1
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Dec_Item147Specify3()
        'On Error GoTo TestFail
        
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(1, -1, 3, 4, 2, 6, 7, 5, 9)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Dec Array(1&, 4&, 7&), 3
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region

'#Region "Queue"
    '@TestMethod("Kvp")
    Private Sub Test01_Enqueue_DefaultLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40)

        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 2&, 3&, 4&)
        
        Dim myExpectedValues As Variant
        myExpectedValues = Array(10, 20, 30, 40, 50)
        
        Dim myResultValues As Variant
        Dim myResult As Variant
        'Act:
        myKvp.EnQueue 50
        myResult = myKvp.GetKeys.ToArray
        myResultValues = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myResult, myExpectedKeys  ', CurrentProcedureName
        Assert.SequenceEquals myResultValues, myExpectedValues  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Enqueue_SpecifiedLongKey10()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10&).Add(1, 2, 3, 4)

        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(10&, 11&, 12&, 13&, 14&)
        
        Dim myExpectedValues As Variant
        myExpectedValues = Array(1, 2, 3, 4, 5)
        
        Dim myResultValues As Variant
        Dim myResult As Variant
        'Act:
        myKvp.EnQueue 5
        myResult = myKvp.GetKeys.ToArray
        myResultValues = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myResult, myExpectedKeys  ', CurrentProcedureName
        Assert.SequenceEquals myResultValues, myExpectedValues  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Enqueue_DefaultStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(1, 2, 3, 4)

        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("0000", "0001", "0002", "0003", "0004")
        
        Dim myExpectedValues As Variant
        myExpectedValues = Array(1, 2, 3, 4, 5)
        
        Dim myResultValues As Variant
        Dim myResult As Variant
        'Act:
        myKvp.EnQueue 5
        myResult = myKvp.GetKeys.ToArray
        myResultValues = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myResult, myExpectedKeys  ', CurrentProcedureName
        Assert.SequenceEquals myResultValues, myExpectedValues  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_Enqueue_SpecifiedStringKeyaaaa()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(Array(1, 2, 3, 4))

        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("aaaa", "aaab", "aaac", "aaad", "aaae")
        
        Dim myExpectedValues As Variant
        myExpectedValues = Array(1, 2, 3, 4, 5)
        
        Dim myResultValues As Variant
        Dim myResult As Variant
        'Act:
        myKvp.EnQueue 5
        myResult = myKvp.GetKeys.ToArray
        myResultValues = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myResult, myExpectedKeys  ', CurrentProcedureName
        Assert.SequenceEquals myResultValues, myExpectedValues  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test05_KeysAsKeysArrayDefaultStartIndex()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst(100, 200, 300, 400, 500, 600, 700)).Add(Array(1, 2, 3, 4))

        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(100, 200, 300, 400, 500)
        
        Dim myExpectedValues As Variant
        myExpectedValues = Array(1, 2, 3, 4, 5)
        
        Dim myResultValues As Variant
        Dim myResult As Variant
        'Act:
        myKvp.EnQueue 5
        myResult = myKvp.GetKeys.ToArray
        myResultValues = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myResult, myExpectedKeys  ', CurrentProcedureName
        Assert.SequenceEquals myResultValues, myExpectedValues  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test06_KeysAsKeysArraySpecifiedStartIndex4()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst(100, 200, 300, 400, 500, 600, 700), 4).Add(Array(1, 2, 3, 4))

        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(500, 600, 700, 100, 200)
        
        Dim myExpectedValues As Variant
        myExpectedValues = Array(1, 2, 3, 4, 5)
        
        Dim myResultValues As Variant
        Dim myResult As Variant
        'Act:
        myKvp.EnQueue 5
        myResult = myKvp.GetKeys.ToArray
        myResultValues = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myResult, myExpectedKeys  ', CurrentProcedureName
        Assert.SequenceEquals myResultValues, myExpectedValues  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test07_Dequeue_DefaultLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5)
        
        Dim myResult As Kvp
        
        Dim myExpected As String
        myExpected = "0,1"
        
        'Act:
        Set myResult = myKvp.DeQueue
        
        'Assert:
        Assert.AreEqual myExpected, myResult.GetFirst.ToString  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test08_Dequeue_SpecifiedLongKey10()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(1, 2, 3, 4, 5)
        
        Dim myResult As Kvp
        
        Dim myExpected As String
        myExpected = "10,1"
        
        'Act:
        Set myResult = myKvp.DeQueue
        
        'Assert:
        Assert.AreEqual myExpected, myResult.GetFirst.ToString  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test09_Dequeue_DefaultStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
    
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(1, 2, 3, 4, 5)
        
        Dim myResult As Kvp
        
        Dim myExpected As String
        myExpected = "0000,1"
        
        'Act:
        Set myResult = myKvp.DeQueue
        
        'Assert:
        Assert.AreEqual myExpected, myResult.GetFirst.ToString  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test10_Dequeue_SpecifiedStringKeyaaaa()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
    
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(1, 2, 3, 4, 5)
        
        Dim myResult As Kvp
        
        Dim myExpected As String
        myExpected = "aaaa,1"
        
        'Act:
        Set myResult = myKvp.DeQueue
        
        'Assert:
        Assert.AreEqual myExpected, myResult.GetFirst.ToString  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test11_ArrayKey_DefaultKeyIndex()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
    
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("Six", "Seven", "eight", "nine", "ten", "eleven", "twelve")).Add(1, 2, 3, 4, 5)
        
        Dim myResult As Kvp
        
        Dim myExpected As String
        myExpected = "Six,1"
        
        'Act:
        Set myResult = myKvp.DeQueue
        
        'Assert:
        Assert.AreEqual myExpected, myResult.GetFirst.ToString  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test12_ArrayKey_SpecifiedKeyIndex4()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
    
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("Six", "Seven", "eight", "nine", "ten", "eleven", "twelve"), 4).Add(1, 2, 3, 4, 5)
        
        Dim myResult As Kvp
        
        Dim myExpected As String
        myExpected = "ten,1"
        
        'Act:
        Set myResult = myKvp.DeQueue
        
        'Assert:
        Assert.AreEqual myExpected, myResult.GetFirst.ToString  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region

'#Region "FilterByKeys"
    '@TestMethod("Setup")
    Private Sub Test01_FilterByKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb
        myKvpA.Add 22&, "Hello World 1"
        myKvpA.Add 23&, "Hello World 2"
        myKvpA.Add 25&, "Hello World 3"
        myKvpA.Add 26&, "Hello World 4"
        myKvpA.Add 27&, "Hello World 5"
        
        ' Subset deliberately includes keys not present in host dictionary
        Dim myKeysB As Variant
        myKeysB = Array(22&, 25&, 27&, 80&)

        Dim myKeysInAandB As Variant
        myKeysInAandB = Array(22&, 25&, 27&)
        
        Dim myKeysInAOnly As Variant
        myKeysInAOnly = Array(23&, 26&)
        
        Dim myKeysInBonly As Variant
        myKeysInBonly = Array(80&)

        Dim mySubSetKvp As Kvp
        
        'Act:
        Set mySubSetKvp = myKvpA.FilterByKeys(myKeysB)
        
        Dim myAandB As Variant
        Dim myB As Variant
        Dim myA As Variant
        
        
        myAandB = mySubSetKvp.Item(0&).GetKeys.ToArray
        myA = mySubSetKvp.Item(1&).GetKeys.ToArray
        myB = mySubSetKvp.Item(2&).GetKeys.ToArray
        'Assert:
        Assert.SequenceEquals myKeysInAandB, myAandB  ', CurrentProcedureName
        Assert.SequenceEquals myKeysInBonly, myB  ', CurrentProcedureName
        Assert.SequenceEquals myKeysInAOnly, myA  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
'#End Region

'#Region "Item"
    '@TestMethod("Kvp")
    Private Sub Test01_Item_DefaultKey_Key3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50)
        
        Dim myExpected As Long
        myExpected = 40
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.Item(3&)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Item_LongKey_Key3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber().Add(10, 20, 30, 40, 50)
        
        Dim myExpected As Long
        myExpected = 40
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.Item(3&)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Item_SpecifiedLongKey_Key13()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        'Note that the key is an integer not long
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(10, 20, 30, 40, 50)
        
        Dim myExpected As Long
        myExpected = 40
        
        Dim myResult As Long
        'Act:
        ' we must retirieve by the smae type as the key
        ' which in this case is integer
        myResult = myKvp.Item(13)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_Item_DefaultStringKey_Key0003()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50)
        
        Dim myExpected As Long
        myExpected = 40
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.Item("0003")
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_Item_SpecifiedStringKey_Keyaaad()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10&, 20&, 30&, 40&, 50&)
        
        Dim myExpected As Long
        myExpected = 40
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.Item("aaad")
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test06_Item_SpecifiedKeysArray_Key4point3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst(1#, 2.1, 3.2, 4.3, 5.4)).Add(Array(10&, 20&, 30&, 40&, 50&))
        
        Dim myExpected As Long
        myExpected = 40
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.Item(4.3)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region

'#Region "GetFirst"
    '@TestMethod("Kvp")
    Private Sub Test01_GetFirst_DefaultKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50)
        Dim myExpected As Long
        myExpected = 10
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.GetFirst.Item
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_GetFirst_SpecifiedLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50)
        Dim myExpected As String
        myExpected = "0,10"
        
        Dim myResult As String
        'Act:
        
        myResult = myKvp.GetFirst.ToString
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test03_GetFirst_SpecifiedLongKey_Key10()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(10, 20, 30, 40, 50)
        Dim myExpected As String
        myExpected = "10,10"
        
        Dim myResult As String
        'Act:
        
        myResult = myKvp.GetFirst.ToString
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Get_First_DefaultStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50)
        Dim myExpected As String
        myExpected = "0000,10"
        
        Dim myResult As String
        'Act:
        
        myResult = myKvp.GetFirst.ToString
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_GetFirst_SpecifiedStringKey_Keyaaaa()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50)
        Dim myExpected As String
        myExpected = "aaaa,10"
        
        Dim myResult As String
        'Act:
        
        myResult = myKvp.GetFirst.ToString
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test05_GetFirst_KeysArray_DefaultStartIndex()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight")).Add(10, 20, 30, 40, 50)
        Dim myExpected As String
        myExpected = "One,10"
        
        Dim myResult As String
        'Act:
        
        myResult = myKvp.GetFirst.ToString
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test06_GetFirst_KeysArray_SpecifyStartIndex4()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight"), 4).Add(10, 20, 30, 40, 50)
        Dim myExpected As String
        myExpected = "Five,10"
        
        Dim myResult As String
        'Act:
        
        myResult = myKvp.GetFirst.ToString
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region

'#Region "GetIndexOfKey"
    '@TestMethod("Kvp")
    Private Sub Test01_GetIndexOfKey_DefaultKey_IndexOf3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50)
        Dim myExpected As Long
        myExpected = 3
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.GetIndexOfKey(3&)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_GetIndexOfKey_SpecifiedDefaultLongKey_IndexOf3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50)
        Dim myExpected As Long
        myExpected = 3
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.GetIndexOfKey(3&)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_GetIndexOfKey_SpecifiedLongKey10_IndexOf3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(10, 20, 30, 40, 50)
        Dim myExpected As Long
        myExpected = 3
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.GetIndexOfKey(13)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_DefaultStringKey_IndexOf0003()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50)
        Dim myExpected As Long
        myExpected = 3
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.GetIndexOfKey("0003")
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test05_GetIndexOfKey_SpecifiedStringKeyaaaa_IndexOfaaad()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50)
        Dim myExpected As Long
        myExpected = 3
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.GetIndexOfKey("aaad")
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test06_GetIndexOfKey_SpecifiedKeysArray_IndexOf4point4()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst(1.1, 2.2, 3.3, 4.4, 5.5, 6.6, 7.7)).Add(10, 20, 30, 40, 50)
        Dim myExpected As Long
        myExpected = 3
        
        Dim myResult As Long
        
        'Act:
        myResult = myKvp.GetIndexOfKey(4.4)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
'#End Region

'#Region "GetIndexOfValue"
    '@TestMethod("Kvp")
    Private Sub Test01_GetIndexOfValue_DefaultKey_IndexOf40()
        'On Error GoTo TestFail
        
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 40)
        Dim myExpected As Variant
        myExpected = Array(3&, 5&)
        
        Dim myResult As Variant
        'Act:
        
    myResult = myKvp.GetIndexOfValue(40).Items
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_GetIndexOfValue_SpecifiedDefaultLongKey_IndexOf40()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 40)
        Dim myExpected As Variant
        myExpected = Array(3&, 5&)
        
        Dim myResult As Variant
        'Act:
        
        myResult = myKvp.GetIndexOfValue(40).GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_GetIndexOfValue_SpecifiedLongKey10_IndexOf40()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(10, 20, 30, 40, 50, 40)
        Dim myExpected As Variant
        myExpected = Array(3&, 5&)
        
        Dim myResult As Variant
        'Act:
        
        myResult = myKvp.GetIndexOfValue(40).GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_GetIndexOfValue_DefaultStringKey_IndexOf40()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 40)
        Dim myExpected As Variant
        myExpected = Array(3&, 5&)
        
        Dim myResult As Variant
        'Act:
        
        myResult = myKvp.GetIndexOfValue(40).GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test05_GetIndexOfValue_SpecifiedStringKeyaaaa_IndexOf40()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 40)
        Dim myExpected As Variant
        myExpected = Array(3&, 5&)
        
        Dim myResult As Variant
        'Act:
        
        myResult = myKvp.GetIndexOfValue(40).GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test06_GetIndexOfValue_SpecifiedKeysArray_IndexOf40()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst(1.1, 2.2, 3.3, 4.4, 5.5, 6.6, 7.7)).Add(10, 20, 30, 40, 50, 40)
        Dim myExpected As Variant
        myExpected = Array(3&, 5&)
        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvp.GetIndexOfValue(40).GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
'#End Region

'#Region "GetKeys" 'empty

'#End Region

'#Region "GetKeysWithValue"
    '@TestMethod("Kvp")
    Private Sub Test01_GetKeysWithValue_GetTwoValues()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(Array(1&, 2&, 3&, 4&, 5&, 6&), Array(10, 20, 30, 40, 50, 40))
        Dim myExpected As Variant
        myExpected = Array(3&, 5&)
        
        Dim myResult As Variant
        'Act:
        
        myResult = myKvp.GetKeysWithValue(40).Keys
        
        'Assert:
        Debug.Print TypeName(myExpected(0)), TypeName(myResult(0))
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
'#End Region

'#Region "HoldsKey"



    '@TestMethod("Kvp")
    Private Sub Test01_HoldsKey_NumberKeyIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsKey(12)
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test02_HoldsKey_NumberKeyIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsKey(300)
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test03_HoldsKey_StringKeyIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsKey("aaac")
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_HoldsKey_StringKeyIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(Array(100, 200, 300, 400))
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsKey("300")
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test05_HoldsKey_ArrayKeyIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four"), 0).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsKey("Three")
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test06_HoldsKey_ArrayKeyIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four"), 0).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsKey("300")
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
'#End Region

'#Region "HoldsItem"

    '@TestMethod("Kvp")
    Private Sub Test01_HoldsItem_NumberKeyValueIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsValue(300)
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test02_HoldsItem_NumberKeyValueIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsValue(12)
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test03_HoldsItem_StringKeyValueIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsValue(300)
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_HoldsItem_StringKeyValueIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsValue("aaac")
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test05_HoldsItem_ArrayKeyValueIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four"), 0).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsValue(300)
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test06_HoldsItem_ArrayKeyValueIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four"), 0).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HoldsValue("Three")
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "Inc"
'@TestMethod("Kvp")
    Private Sub Test01_Inc_Item5DefaultOne()
    'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Long
        myExpected = 7
        
        Dim myResult As Long
        'Act:
        
        myKvp.Inc 5&
    
        myResult = myKvp.Item(5&)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Inc_Item5SpecifyOne()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Long
        myExpected = 7
        
        Dim myResult As Long
        'Act:
        
        myKvp.Inc 5&, 1
    
        myResult = myKvp.Item(5&)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Inc_Item5Specify3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Long
        myExpected = 9
        
        Dim myResult As Long
        'Act:
        
        myKvp.Inc 5&, 3
    
        myResult = myKvp.Item(5&)
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
'#End Region

'#Region "IncAll"

    '@TestMethod("Kvp")
    Private Sub Test01_IncAll_Item5DefaultOne()
    'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(2, 3, 4, 5, 6, 7, 8, 9, 10)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Inc
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_IncAll_Item5SpecifyOne()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(2, 3, 4, 5, 6, 7, 8, 9, 10)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Inc ipIncrement:=1
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_IncAll_Item5Specify3()
        'On Error GoTo TestFail
        
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(4, 5, 6, 7, 8, 9, 10, 11, 12)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Inc ipIncrement:=3
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "IncByKeys"

    '@TestMethod("Kvp")
    Private Sub Test01_IncByKeys_Items147DefaultOne()
    'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        'myKvp.add  (Array(1, 2, 3, 4, 5, 6, 7, 8, 9))
        Dim myExpected As Variant
        myExpected = Array(1, 3, 3, 4, 6, 6, 7, 9, 9)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Inc Array(1&, 4&, 7&)
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_IncByKeys_Item147SpecifyOne()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(1, 3, 3, 4, 6, 6, 7, 9, 9)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Inc Array(1&, 4&, 7&), 1
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_IncByKeys_Item147Specify3()
        'On Error GoTo TestFail
        
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5, 6, 7, 8, 9)
        
        Dim myExpected As Variant
        myExpected = Array(1, 5, 3, 4, 8, 6, 7, 11, 9)
        
        Dim myResult As Variant
        'Act:
        
        myKvp.Inc Array(1&, 4&, 7&), 3
    
        myResult = myKvp.GetItems.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName
        
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "InsertAfterKey"

    '@TestMethod("Kvp")
    Private Sub Test01_InsertAfterKey_LastItemLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10&).Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array(10&, 11&, 13&, 12&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertAt 2, "Its a nice day"
        myResult = myKvp.Keys
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test02_InsertAfterKey_IncByKeysLastItemStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaab", "aaad", "aaac")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertAt 2, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_InsertAfterKey_FirstItemLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10&).Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array(13&, 10&, 11&, 12&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertAt 0, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_InsertAfterKey_FirstItemStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array("aaad", "aaaa", "aaab", "aaac")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertAt 0, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_InsertAfterKey_MidItemLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10&).Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array(10&, 13&, 11&, 12&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertAt 1, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test06_InsertAfterKey_MidItemStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaad", "aaab", "aaac")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertAt 1, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test07_InsertAfterKey_MidItemUserLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10&).Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array(10&, 13&, 11&, 12&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertAt 1, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test08_InsertAfterKey_MidItemUserStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaad", "aaab", "aaac")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertAt 1, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "InsertBeforeKey"

    '@TestMethod("Kvp")
    Private Sub Test01_InsertBeforeKey_LastItemLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10&).Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array(10&, 11&, 13&, 12&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertBeforeKey 12&, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test02_InsertBeforeKey_LastItemStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaab", "aaad", "aaac")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertBeforeKey "aaac", "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_InsertBeforeKey_FirstItemLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10&).Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array(13&, 10&, 11&, 12&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertBeforeKey 10&, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_InsertBeforeKey_FirstItemStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array("aaad", "aaaa", "aaab", "aaac")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertBeforeKey "aaaa", "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_InsertBeforeKey_MidItemLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10&).Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array(10&, 13&, 11&, 12&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertBeforeKey 11&, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test06_InsertBeforeKey_MidItemStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaad", "aaab", "aaac")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertBeforeKey "aaab", "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test07_InsertBeforeKey_MidItemUserLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10&).Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array(10&, 13&, 11&, 12&)
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertBeforeKey 11&, "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test08_InsertBeforeKey_MidItemUserStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add("Hello", "There", "World")
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaad", "aaab", "aaac")
        
        Dim myResult As Variant
        
        'Act:
        myKvp.InsertBeforeKey "aaab", "Its a nice day"
        myResult = myKvp.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "InsertFirst"

    '@TestMethod("Kvp")
    Private Sub Test01_InsertFirst_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(6&, 0&, 1&, 2&, 3&, 4&, 5&)
        
        Dim myResult As Variant
        'Act:
        myKvp.InsertAt 0, 70
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_InsertFirst_DefaultNumberKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(6&, 0&, 1&, 2&, 3&, 4&, 5&)
        
        Dim myResult As Variant
        'Act:
        myKvp.InsertAt 0, 70
        
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_InsertFirst_SpecifiedNumberKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(106&, 100&, 101&, 102&, 103&, 104&, 105&)
        
        Dim myResult As Variant
        'Act:
        myKvp.InsertAt 0, 70
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_InsertFirst_SpecifiedNumberKeySpeciedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(106&, 100&, 101&, 102&, 103&, 104&, 105&)
        
        Dim myResult As Variant
        'Act:
        myKvp.InsertAt 0, 25&
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_InsertFirst_SpecifiedNumberKeySpeciedKeyTwoInserts()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(106&, 70&, 100&, 101&, 102&, 103&, 104&, 105&)
        
        Dim myResult As Variant
        'Act:
      '  myKvp.Prepend 70, 25&
      '  myKvp.Prepend 80
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test06_InsertFirst_SpecifiedStringKeyDefaultKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array("0006", "0000", "0001", "0002", "0003", "0004", "0005")
        
        Dim myResult As Variant
        'Act:
     '   myKvp.InsertFirst 70
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test06_InsertFirst_SpecifiedStringKeySpecifiedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array("aaag", "aaaa", "aaab", "aaac", "aaad", "aaae", "aaaf")
        
        Dim myResult As Variant
        'Act:
      '  myKvp.Prepend 70
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "InsertLast"

    '@TestMethod("Kvp")
    Private Sub Test01_InsertLast_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(0&, 1&, 2&, 3&, 4&, 5&, 6&)
        
        Dim myResult As Variant
        'Act:
       ' myKvp.Append (70)
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_InsertLast_DefaultNumberKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(0&, 1&, 2&, 3&, 4&, 5&, 6&)
        
        Dim myResult As Variant
        'Act:
       ' myKvp.Append (70)
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_InsertLast_SpecifiedNumberKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(100&, 101&, 102&, 103&, 104&, 105&, 106&)
        
        Dim myResult As Variant
        'Act:
       ' myKvp.Append (70)
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_InsertLast_SpecifiedNumberKeySpeciedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(100&, 101&, 102&, 103&, 104&, 105&, 25&)
        
        Dim myResult As Variant
        'Act:
       ' myKvp.Append 70, 25&
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_InsertLast_SpecifiedNumberKeySpeciedKeyTwoInserts()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array(100&, 101&, 102&, 103&, 104&, 105&, 25&, 26&)
        
        Dim myResult As Variant
        'Act:
     '   myKvp.Append 70, 25&
     '   myKvp.Append 80
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test06_InsertLast_SpecifiedStringKeyDefaultKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array("0000", "0001", "0002", "0003", "0004", "0005", "0006")
        
        Dim myResult As Variant
        'Act:
        ' myKvp.Append 70
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test07_InsertLast_SpecifiedStringKeySpecifiedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaab", "aaac", "aaad", "aaae", "aaaf", "aaag")
        
        Dim myResult As Variant
        'Act:
        ' myKvp.Append 70
        myResult = myKvp.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub




'#End Region

'#Region "IntMinMax"

    '@TestMethod("Kvp")
    Private Sub Test01_IntMinMax_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(20, 50, 10, 30, 40, 50, 50, 10)
        
        Dim myExpectedMinKeys As Variant
        myExpectedMinKeys = Array(2&, 7&)
        
        Dim myExpectedMaxKeys As Variant
        myExpectedMaxKeys = Array(1&, 5&, 6&)
        
        Dim myResult As Kvp
        'Act:
        Set myResult = myKvp.NumMaxMin
        
        'Assert:
        Assert.SequenceEquals myExpectedMinKeys, myResult.Item(0&).GetKeys.ToArray  ', CurrentProcedureName
        Assert.SequenceEquals myExpectedMaxKeys, myResult.Item(1&).GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "Invert"

    '@TestMethod("Setup")
    Private Sub Test01_Invert_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(Array("Hello World 1", "Hello World 2", "Hello World 3", "Hello World 4", "Hello World 5"), Array(22&, 23&, 25&, 26&, 27&))
        

        Dim myExpected As Variant
        myExpected = Array(27&, 26&, 25&, 23&, 22&)

        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvp.Reverse.GetKeys.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub




'#End Region

'#Region "HoldsItems"



    '@TestMethod("Kvp")
    Private Sub Test01_LacksItems_IsTrue()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                           As Kvp
        Dim myExpected                      As Boolean
        myExpected = True
        
        Set myKvp = Kvp.Deb
        Dim myResult                        As Boolean
        
        'Act:
        myResult = myKvp.HasNoItems
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_LacksItems_PopulatedKvpIsFalse()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Dim myExpected As Boolean
        myExpected = False
        
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5)
        Dim myResult As Boolean
        
        'Act:
        myResult = Kvp.HasNoItems
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_LacksItems_EmptiedKvpIsTrue()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                           As Kvp
        Dim myExpected                      As Boolean
        myExpected = True
        
        Set myKvp = Kvp.Deb.Add(Array(1, 2, 3, 4, 5))
        myKvp.RemoveAll
        
        Dim myResult                        As Boolean
        
        'Act:
        myResult = myKvp.HasNoItems
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    
    '@TestMethod("Kvp")
    Private Sub Test01_HoldsItems_NewKvpIsFalse()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                           As Kvp
        Dim myExpected                      As Boolean
        myExpected = False
        
        Set myKvp = Kvp.Deb
        Dim myResult                        As Boolean
        
        'Act:
        myResult = myKvp.HasItems
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_HoldsItems_PopulatedKvpIsTrue()
        'On Error GoTo TestFail
        
        'Arrange:
        
        Dim myExpected As Boolean: myExpected = True
        
        Dim myKvp As Kvp:
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5)
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.HasItems
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_HoldsItems_EmptiedKvpIsFalse()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp  As Kvp
        Dim myExpected  As Boolean
        myExpected = False
        
        Set myKvp = Kvp.Deb.Add(1, 2, 3, 4, 5)
        myKvp.RemoveAll
        
        Dim myResult   As Boolean
        
        'Act:
        myResult = myKvp.HasItems
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub
'#End Region

'#Region "ItemAt"
    '@TestMethod("Kvp")
    Private Sub Test01_ItemAt_DefaultKey_IndexAt3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.Add(10&, 20&, 30&, 40&, 50&)
        
        Dim myExpected As Long
        myExpected = 40
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.ItemAt(3)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_ItemAt_LongKey_IndexAt3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber().Add(10&, 20&, 30&, 40&, 50&)
        
        Dim myExpected As Long
        myExpected = 3
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.GetKeys.ToArray(3)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_ItemAt_SpecifiedLongKey_IndexAt3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(10&, 20&, 30&, 40&, 50&)
        
        Dim myExpected As Long
        myExpected = 13
        
        Dim myResult As Long
        'Act:
        
        myResult = myKvp.GetKeys.ToArray(3)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_ItemAt_DefaultStringKey_IndexAt3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10&, 20&, 30&, 40&, 50&)
        
        Dim myExpected As String
        myExpected = "0003"
        
        Dim myResult As String
        'Act:
        
        myResult = myKvp.GetKeys.ToArray(3)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_ItemAt_SpecifiedStringKey_IndexAt3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10&, 20&, 30&, 40&, 50&)
        
        Dim myExpected As String
        myExpected = "aaad"
        
        Dim myResult As String
        'Act:
        
        myResult = myKvp.GetKeys.ToArray(3)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test06_ItemAt_SpecifiedKeysArray_IndexAt3()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst(1#, 2.1, 3.2, 4.3, 5.4)).Add(10&, 20&, 30&, 40&, 50&)
        
        Dim myExpected As Double
        myExpected = 4.3
        
        Dim myResult As Double
        'Act:
        
        myResult = myKvp.GetKeys.ToArray(3)
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "KeysAllAandOnlyB"

    '@TestMethod("Kvp")
    Private Sub Test01_KeysAllAandOnlyB_DefaultKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0, 1, 2, 3, 4, 5))
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0, 1, 2, 3, 7, 8))
        
        Dim myExpected As Variant
        myExpected = Array(0, 1, 2, 3, 4, 5, 7, 8)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysAllAandOnlyB(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_KeysAllAandOnlyB_DefaultNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 4&, 5&))
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 7&, 8&))
        
        Dim myExpected As Variant
        myExpected = Array(0&, 1&, 2&, 3&, 4&, 5&, 7&, 8&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysAllAandOnlyB(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_KeysAllAandOnlyB_SpecifiedNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(100&, 101&, 102&, 103&, 107&, 108&))
        
        Dim myExpected As Variant
        myExpected = Array(100&, 101&, 102&, 103&, 104&, 105&, 107&, 108&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysAllAandOnlyB(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "KeysInAOnly"

    '@TestMethod("Kvp")
    Private Sub Test01_KeysInAOnly_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 7&, 8&))
        
        Dim myExpected As Variant
        myExpected = Array(4&, 5&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_KeysInAOnly_DefaultNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 4&, 5&))
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 7&, 8&))
        
        Dim myExpected As Variant
        myExpected = Array(4&, 5&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_KeysInAOnly_SpecifiedNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(100&, 101&, 102&, 103&, 107&, 108&))
        
        Dim myExpected As Variant
        myExpected = Array(104&, 105&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_KeysInAOnly_SpecifiedStringKeysDefaultKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array("0000", "0001", "0002", "0003", "0007", "0008"))
        
        Dim myExpected As Variant
        myExpected = Array("0004", "0005")
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_KeysInAOnly_SpecifiedStringKeysSpecifiedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array("aaaa", "aaab", "aaac", "aaad", "aaag", "aaah"))
        
        Dim myExpected As Variant
        myExpected = Array("aaae", "aaaf")
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "KeysInBOnly"

    '@TestMethod("Kvp")
    Private Sub Test01_KeysInBOnly_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 7&, 8&))
        
        Dim myExpected As Variant
        myExpected = Array(7&, 8&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInBOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_KeysInBOnly_DefaultNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 4&, 5&))
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 7&, 8&))
        
        Dim myExpected As Variant
        myExpected = Array(7&, 8&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInBOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_KeysInBOnly_SpecifiedNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(100&, 101&, 102&, 103&, 107&, 108&))
        
        Dim myExpected As Variant
        myExpected = Array(107&, 108&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInBOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_KeysInBOnly_SpecifiedStringKeysDefaultKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array("0000", "0001", "0002", "0003", "0007", "0008"))
        
        Dim myExpected As Variant
        myExpected = Array("0007", "0008")
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInBOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_KeysInBOnly_SpecifiedStringKeysSpecifiedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array("aaaa", "aaab", "aaac", "aaad", "aaag", "aaah"))
        
        Dim myExpected As Variant
        myExpected = Array("aaag", "aaah")
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInBOnly(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "KeysInBothAandB"

    '@TestMethod("Kvp")
    Private Sub Test01_KeysInBothAandB_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 7&, 8&))
        
        Dim myExpected As Variant
        myExpected = Array(0&, 1&, 2&, 3&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAandB(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_KeysInBothAandB_DefaultNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 4&, 5&))
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 7&, 8&))
        
        Dim myExpected As Variant
        myExpected = Array(0&, 1&, 2&, 3&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAandB(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_KeysInBothAandB_SpecifiedNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(Array(10, 20, 30, 40, 50, 60))
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(100&, 101&, 102&, 103&, 107&, 108&))
        
        Dim myExpected As Variant
        myExpected = Array(100&, 101&, 102&, 103&)
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAandB(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_KeysInBothAandB_SpecifiedStringKeysDefaultKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array("0000", "0001", "0002", "0003", "0007", "0008"))
        
        Dim myExpected As Variant
        myExpected = Array("0000", "0001", "0002", "0003")
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAandB(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_KeysInBothAandB_SpecifiedStringKeysSpecifiedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array("aaaa", "aaab", "aaac", "aaad", "aaag", "aaah"))
        
        Dim myExpected As Variant
        myExpected = Array("aaaa", "aaab", "aaac", "aaad")
        
        Dim myResult As Kvp
        
        'Act:
        Set myResult = myKvpA.KeysInAandB(myKvpB)
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "KeysNotInBothAandB"

    '@TestMethod("Kvp")
    Private Sub Test01_KeysNotInBothAandB_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 7&, 8&))
        
        Dim myExpected As Variant
        myExpected = Array(4&, 5&, 7&, 8&)
        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvpA.KeysNotInBothAandB(myKvpB).GetKeys.Sort.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_KeysNotInBothAandB_DefaultNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 4&, 5&))
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(0&, 1&, 2&, 3&, 7&, 8&))
        
        Dim myExpected As Variant
        myExpected = Array(4&, 5&, 7&, 8&)
        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvpA.KeysNotInBothAandB(myKvpB).GetKeys.Sort.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_KeysNotInBothAandB_SpecifiedNumberKeysLongKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array(100&, 101&, 102&, 103&, 107&, 108&))
        
        Dim myExpected As Variant
        myExpected = Array(104&, 105&, 107&, 108&)
        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvpA.KeysNotInBothAandB(myKvpB).GetKeys.Sort.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_KeysNotInBothAandB_SpecifiedStringKeysDefaultKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array("0000", "0001", "0002", "0003", "0007", "0008"))
        
        Dim myExpected As Variant
        myExpected = Array("0004", "0005", "0007", "0008")
        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvpA.KeysNotInBothAandB(myKvpB).GetKeys.Sort.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_KeysNotInBothAandB_SpecifiedStringKeysSpecifiedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvpA As Kvp
        Set myKvpA = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myKvpB As Kvp
        Set myKvpB = Kvp.Deb.Add(Array(10, 20, 30, 40, 50, 60), Array("aaaa", "aaab", "aaac", "aaad", "aaag", "aaah"))
        
        Dim myExpected As Variant
        myExpected = Array("aaae", "aaaf", "aaag", "aaah")
        
        Dim myResult As Variant
        
        'Act:
        myResult = myKvpA.KeysNotInBothAandB(myKvpB).GetKeys.Sort.ToArray
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "LacksKey"

    '@TestMethod("Kvp")
    Private Sub Test01_LacksKey_NumberKeyIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksKey(12)
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test02_LacksKey_NumberKeyIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksKey(300)
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test03_LacksKey_StringKeyIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksKey("aaac")
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_LacksKey_StringKeyIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksKey("300")
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test05_LacksKey_ArrayKeyIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four"), 0).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksKey("Three")
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test06_LacksKey_ArrayKeyIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four"), 0).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksKey("300")
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "LacksItem"

    '@TestMethod("Kvp")
    Private Sub Test01_LacksItem_umberKeyValueIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksValue(300)
        
        'Assert:
        Assert.IsFalse myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test02_LacksItem_NumberKeyValueIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(10).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksValue(12)
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test03_LacksItem_StringKeyValueIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksValue(300)
        
        'Assert:
        Assert.AreEqual False, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_LacksItem_StringKeyValueIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksValue("aaac")
        
        'Assert:
        Assert.IsTrue myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test05_LacksItem_ArrayKeyValueIsPresent()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four"), 0).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksValue(300)
        
        'Assert:
        Assert.AreEqual False, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test06_LacksItem_ArrayKeyValueIsMissing()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByIndex(Types.Iterable.ToLyst("One", "Two", "Three", "Four"), 0).Add(100, 200, 300, 400)
        
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.LacksValue("Three")
        
        'Assert:
        Assert.AreEqual True, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "MirrorByItem"

    '@TestMethod("Kvp")
    Private Sub Test01_MirrorByItem_DefaultKey_DefaultIntegerByIndex()
        'On Error GoTo TestFail
        
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 40)
        Dim myExpected As Variant
        myExpected = Array(10, 20, 30, 40, 50)
        
        Dim myResult As Kvp
        'Act:
        
        Set myResult = myKvp.MirrorByValue
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_MirrorByItem_DefaultKey_DefaulStringByIndex()
        'On Error GoTo TestFail
        
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(Array("Ten", "Twenty", "Thirty", "Fourty", "Fifty", "Fourty"))
        Dim myExpected As Variant
        myExpected = Array("Ten", "Twenty", "Thirty", "Fourty", "Fifty")
        
        Dim myResult As Kvp
        'Act:
        
        Set myResult = myKvp.MirrorByValue
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub




'#End Region

'#Region "MirrorByFirstItem"

    '@TestMethod("Kvp")
    Private Sub Test01_MirrorByFirstItem_DefaultKey_DefaultIntegerByIndex()
        'On Error GoTo TestFail
        
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 40, 30, 40, 50, 40)
        Dim myExpected As Variant
        myExpected = Array(10, 20, 40, 30, 50)
        
        Dim myResult As Kvp
        'Act:
        
        Set myResult = myKvp.MirrorFirstValues
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.Item(0&).GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test01_MirrorByFirstItem_DefaultKey_DefaultStringByIndex()
        'On Error GoTo TestFail
        
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add("Ten", "Twenty", "Fourty", "Thirty", "Fourty", "Fifty", "Fourty")
        Dim myExpected As Variant
        myExpected = Array("Ten", "Twenty", "Fourty", "Thirty", "Fifty")
        
        Dim myResult As Kvp
        'Act:
        
        Set myResult = myKvp.MirrorFirstValues
        
        'Assert:
        Assert.SequenceEquals myExpected, myResult.Item(0&).GetKeys.ToArray  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "NextPair"
    '@TestMethod("Kvp")
    Private Sub Test01_NextPair_DefaultNoSpecifiedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb(2&, 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.NextPair(1&)
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test02_NextPair_DefaultNumberIntegerKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb(2&, 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.NextPair(1&)
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_NextPair_SpecifiedNumberIntegerKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(CDbl(25.2)).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb(CDbl(27.2), 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.NextPair(CDbl(26.2))
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_NextPair_DefaultStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb("0002", 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.NextPair("0001")
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_NextPair_SpecifiedStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb("aaac", 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.NextPair("aaab")
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "PrevPair"

    '@TestMethod("Kvp")
    Private Sub Test01_PrevPair_DefaultNoSpecifiedKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb(0&, 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PrevPair(1&)
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test02_PrevPair_DefaultNumberIntegerKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb(0&, 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PrevPair(1&)
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_PrevPair_SpecifiedNumberIntegerKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(CDbl(25.2)).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb(CDbl(25.2), 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PrevPair(CDbl(26.2))
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_PrevPair_DefaultStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb("0000", 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PrevPair("0001")
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_PrevPair_SpecifiedStringKey()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As KVPair
        Set myExpected = KVPair.Deb("aaaa", 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PrevPair("aaab")
        
        'Assert:
        Assert.AreEqual myExpected.ToString, myResult.ToString  ', CurrentProcedureName
TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region

'#Region "Pull"

    '@TestMethod("Kvp")
    Private Sub Test01_Pull_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 3&, 4&, 5&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(2&, 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.Pull(2&)
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Pull_NumberKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 3&, 4&, 5&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(2&, 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.Pull(2&)
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Pull_NumberKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(100&, 101&, 103&, 104&, 105&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(102&, 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.Pull(102&)
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_Pull_StringKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("0000").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("0000", "0001", "0003", "0004", "0005")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("0002", 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.Pull("0002")
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_Pull_StringKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("aaaa", "aaab", "aaad", "aaae", "aaaf")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("aaac", 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.Pull("aaac")
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "PullAt"

    '@TestMethod("Kvp")
    Private Sub Test01_PullAt_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 3&, 4&, 5&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(2&, 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullAt(2)
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_PullAt_NumberKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 3&, 4&, 5&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(2&, 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullAt(2)
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_PullAt_NumberKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(100&, 101&, 103&, 104&, 105&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(102&, 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullAt(2)
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_PullAt_StringKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("0000").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("0000", "0001", "0003", "0004", "0005")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("0002", 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullAt(2)
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_PullAt_StringKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("aaaa", "aaab", "aaad", "aaae", "aaaf")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("aaac", 30)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullAt(2)
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "PullFirst"

    '@TestMethod("Kvp")
    Private Sub Test01_PullFirst_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(1&, 2&, 3&, 4&, 5&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(0&, 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullFirst
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_PullFirst_NumberKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(1&, 2&, 3&, 4&, 5&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(0&, 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullFirst
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_PullFirst_NumberKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(101&, 102&, 103&, 104&, 105&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(100&, 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullFirst
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_PullFirst_StringKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("0000").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("0001", "0002", "0003", "0004", "0005")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("0000", 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullFirst
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_PullFirst_StringKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("aaab", "aaac", "aaad", "aaae", "aaaf")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("aaaa", 10)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullFirst
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "PullLast"

    '@TestMethod("Kvp")
    Private Sub Test01_PullLast_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 2&, 3&, 4&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(5&, 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_PullLast_NumberKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 2&, 3&, 4&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(5&, 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_PullLast_NumberKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(100&, 101&, 102&, 103&, 104&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(105&, 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_PullLast_StringKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("0000").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("0000", "0001", "0002", "0003", "0004")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("0005", 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_PullLast_StringKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("aaaa", "aaab", "aaac", "aaad", "aaae")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("aaaf", 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "RemoveAll"

    '@TestMethod("Kvp")
    Private Sub Test01_RemoveAll_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Long
        myExpected = 0
        
        
        Dim myResult As Long
        'Act:
        myResult = myKvp.RemoveAll.Count
        
        
        'Assert:
        
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_RemoveAll_NumberKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Long
        myExpected = 0
        
        
        Dim myResult As Long
        'Act:
        myResult = myKvp.RemoveAll.Count
        
        
        'Assert:
        
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test03_RemoveAll_NumberKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Long
        myExpected = 0
        
        
        Dim myResult As Long
        'Act:
        myResult = myKvp.RemoveAll.Count
        
        
        'Assert:
        
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test04_RemoveAll_StringKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Long
        myExpected = 0
        
        
        Dim myResult As Long
        'Act:
        myResult = myKvp.RemoveAll.Count
        
        
        'Assert:
        
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_RemoveAll_StringKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Long
        myExpected = 0
        
        
        Dim myResult As Long
        'Act:
        myResult = myKvp.RemoveAll.Count
        
        
        'Assert:
        
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region

'#Region "Remove"

    '@TestMethod("Kvp")
    Private Sub Test01_Remove_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 2&, 3&, 4&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(5&, 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_Remove_NumberKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 2&, 3&, 4&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(5&, 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_Remove_NumberKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(100&, 101&, 102&, 103&, 104&)
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb(105&, 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_Remove_StringKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("0000").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("0000", "0001", "0002", "0003", "0004")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("0005", 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_Remove_StringKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("aaaa", "aaab", "aaac", "aaad", "aaae")
        
        Dim myExpectedPair As KVPair
        Set myExpectedPair = KVPair.Deb("aaaf", 60)
        
        Dim myResult As KVPair
        'Act:
        Set myResult = myKvp.PullLast
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myKvp.GetKeys.ToArray  ', CurrentProcedureName
        Assert.AreEqual myExpectedPair.ToString, myResult.ToString  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


'#End Region

'#Region "RemoveAt"

    '@TestMethod("Kvp")
    Private Sub Test01_RemoveAt_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 3&, 4&, 5&)
        
        
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveAt(2).GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_RemoveAt_NumberKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 3&, 4&, 5&)
        
        
        Dim myResult As Variant
        'Act:
        
        myResult = myKvp.RemoveAt(2).GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_RemoveAt_NumberKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(100&, 101&, 103&, 104&, 105&)
        
        
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveAt(2).GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
    
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_RemoveAt_StringKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("0000").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("0000", "0001", "0003", "0004", "0005")
        
            
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveAt(2).GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_RemoveAt_StringKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("aaaa", "aaab", "aaad", "aaae", "aaaf")
        
            
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveAt(2).GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "RemoveFirst"

    '@TestMethod("Kvp")
    Private Sub Test01_RemoveFirst_DefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(1&, 2&, 3&, 4&, 5&)
        
        
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveFirst.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_RemoveFirst_NumberKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(1&, 2&, 3&, 4&, 5&)
        
        
        Dim myResult As Variant
        'Act:
        
        myResult = myKvp.RemoveFirst.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_RemoveFirst_NumberKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(101&, 102&, 103&, 104&, 105&)
        
        
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveFirst.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
    
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_RemoveFirst_StringKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("0000").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("0001", "0002", "0003", "0004", "0005")
        
            
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveFirst.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_RemoveFirst_StringKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("aaab", "aaac", "aaad", "aaae", "aaaf")
        
            
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveFirst.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub




'#End Region

'#Region "RemoveLast"

    '@TestMethod("Kvp")
    Private Sub Test01_RemoveLast_efaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 2&, 3&, 4&)
        
        
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveLast.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_RemoveLast_NumberKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(0&, 1&, 2&, 3&, 4&)
        
        
        Dim myResult As Variant
        'Act:
        
        myResult = myKvp.RemoveLast.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test03_RemoveLast_NumberKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByNumber(100&).Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array(100&, 101&, 102&, 103&, 104&)
        
        
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveLast.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
    
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub


    '@TestMethod("Kvp")
    Private Sub Test04_RemoveLast_StringKeysDefaultKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("0000").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("0000", "0001", "0002", "0003", "0004")
        
            
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveLast.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test05_RemoveLast_StringKeysSpecifiedKeys()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.SetAutoKeyToAutoKeyByString("aaaa").Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpectedKeys As Variant
        myExpectedKeys = Array("aaaa", "aaab", "aaac", "aaad", "aaae")
        
            
        Dim myResult As Variant
        'Act:
        myResult = myKvp.RemoveLast.GetKeys.ToArray
        
        
        'Assert:
        Assert.SequenceEquals myExpectedKeys, myResult  ', CurrentProcedureName
        Exit Sub
        
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub




'#End Region

'#Region "SumKeys" 'empty
'@TestMethod("Kvp")
Private Sub Test01_SumKeys_IsObject()
    'On Error GoTo TestFail
    
    'Arrange:
    Dim myKvp                              As Kvp
    
    'Act:
    Set myKvp = Kvp.Deb
    
    'Assert:
    'Assert.IsTrue IsObject(myKvp)

TestExit:
    Exit Sub
TestFail:
    Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub



'#End Region

'#Region "SumItems" 'empty

    '@TestMethod("Kvp")
    Private Sub Test01_SumItems_IsObject()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp                              As Kvp
        
        'Act:
        Set myKvp = Kvp.Deb
        
        'Assert:
       ' Assert.IsTrue IsObject(myKvp)

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub



'#End Region

'#Region "UniqueItems"

    '@TestMethod("Kvp")
    Private Sub Test01_UniqueItems_DefaultKeysTrue()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Boolean
        myExpected = True
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.IsUnique
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_UniqueItems_DefaultKeysFalse()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60, 10)
        
        Dim myExpected As Boolean
        myExpected = False
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.IsUnique
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test01_NotUniqueItems_DefaultKeysTrue()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60)
        
        Dim myExpected As Boolean
        myExpected = False
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.IsNotUnique
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

    '@TestMethod("Kvp")
    Private Sub Test02_NotUniqueItems_DefaultKeysFalse()
        'On Error GoTo TestFail
        
        'Arrange:
        Dim myKvp As Kvp
        Set myKvp = Kvp.Deb.Add(10, 20, 30, 40, 50, 60, 10)
        
        Dim myExpected As Boolean
        myExpected = True
        
        Dim myResult As Boolean
        
        'Act:
        myResult = myKvp.IsNotUnique
        
        'Assert:
        Assert.AreEqual myExpected, myResult  ', CurrentProcedureName

TestExit:
        Exit Sub
TestFail:
        Debug.Print "Test raised an error: #" & Err.Number & " - " & Err.Description
    End Sub

'#End Region
