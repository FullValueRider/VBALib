VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SetH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SetH is a multipurpose Collection/ArrayList replacement.  It uses a cHashD object to store the items.  Items are access using 1 based indexing. For a Keyed collection, use one of the Kvp classes"
'@IgnoreModule
'@PredeclaredId
'@Exposed
'@Folder("Set")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

'@ModuleDescription("SetH is a multipurpose Collection/ArrayList replacement.  It uses a cHashD object to store the items.  Items are access using 1 based indexing. For a Keyed collection, use one of the Kvp classes")
' SeqC uses the VBA Collection class as a host for the items
' Other versions of SeqC are available (SeqA, SeqC, SeqAL, SetH, SeqT, SeqL)

Option Explicit

''Implements ISeq
'Private Const MY_TYPENAME                       As String = "SetH"
'Private Const FIRST_INDEX                       As Long = 1
'
'Private Type State
'    Host                                        As cHashC
'    Tmp                                         As SetH
'End Type
'
'Private s                                       As State


' Deb is a Factory method of the PredeclaredId which returns a new instance of SetH
'=================================================================================
' The input for Deb is a ParamArray.
' The ParamArray is interpreted according to the folowing rules
' If the ParamArray contains no Items
' - A empty SetHuence (count = 0)
' If the ParamArray has 2 or more items
' - A SetH of size equal to the ParamArray, items as generated by for each over the ParamArray
' If the ParamArray has a single Item
' - Number: SetH with specified capacity with each Item being the Value 'Empty'
' - String: A SetH whose capacity is the length of the string with one character per Item
' - Array: SetH of items as provided by For Each over the array.  Multidimension arrays are flattened but the position in the original array can be calculated by the user if required.
' - Dictionaries: a seq of key/value arrays as generated from for each on the Keys array.
' - Collection: The collection is copied to the Host collection.  There is no refererence link to the original collection (See Bind Property)
' - SetH: Items as generated by for each on the SetH
' - mscorlib Arraylist: Items as generated by for each on the arrylist
' - mscorlib Stack: Items as generated by for each on the stack
' - mscorlib Queue: Items as generated by for eac on the queue

''@DefaultMember
'Public Function Deb(ParamArray ipParamarray() As Variant) As SetH
'
'    With New SetH
'        Set s.Tmp = .ConstructInstance
'    End With
'
'    ' the assignment to s.tmp is a kludge that allows the various
'    ' Populate methods to operate on the new instance
'    ' which greatly simplifies the populate methods at some processing cost.
'    If ArrayOp.LacksItems(vba.cvar(ipParamArray)) Then
'        Set Deb = s.Tmp
'        Set s.Tmp = Nothing
'        Exit Function
'    End If
'
'    ' Resolve the paramarray if it only has a single item
'    Dim myParam As Variant
'    myParam = vba.cvar(ipParamArray)
'    ArrayOp.SPlat myParam
'
'    ' myParam may now be any type so we process in accordance with the rules stated above
'    Dim myGroupId As e_Group
'    myGroupId = GroupInfo.Id(myParam)
'
'    Select Case myGroupId
'
'        Case e_Group.m_Empty:               s.Tmp.Add myParam
'        Case e_Group.m_Admin:               s.Tmp.Add myParam
'        Case e_Group.m_number:              PopulateUsingInitialSize VBA.CLng(myParam)
'        Case e_Group.m_string:              PopulateFromString VBA.CStr(myParam)
'        Case e_Group.m_array:               PopulateUsingForEach myParam
'        Case e_Group.m_List:                PopulateUsingForEach myParam
'        Case e_Group.m_Dictionary:          PopulateFromDictionary myParam
'        Case e_Group.m_Other:               s.Tmp.Add myParam
'
'    End Select
'
'    Set Deb = s.Tmp
'    Set s.Tmp = Nothing
'
'End Function
'
'
'Friend Function ConstructInstance() As SetH
'    Set s.Host = New cHashC
'    Set ConstructInstance = Me
'End Function
'
'
'Private Sub PopulateUsingInitialSize(ByVal ipInitialSize As Long)
'
'    GuardInsufficientCapacity ipInitialSize, "SetH.Deb"
'
'    Dim myCount As Long
'    For myCount = p.firstindex To ipInitialSize
'       s.Tmp.Add Empty
'    Next
'
'End Sub
'
'Private Sub PopulateFromString(ByRef ipString As String)
'
'    If VBA.Len(ipString) = 0 Then
'        Exit Sub
'    End If
'
'    Dim myIndex As Long
'    For myIndex = 1 To VBA.Len(ipString)
'        s.Tmp.Add VBA.Mid$(ipString, myIndex, 1)
'    Next
'
'End Sub
'
'Private Sub PopulateUsingForEach(ByVal ipForEach As Variant)
'
'   If VBA.IsArray(ipForEach) Then
'        If ArrayOp.LacksItems(ipForEach) Then
'            Exit Sub
'        End If
'    ElseIf ipForEach.Count = 0 Then
'        Exit Sub
'    End If
'
'    Dim myItem As Variant
'    For Each myItem In ipForEach
'      s.Tmp.Add myItem
'    Next
'
'End Sub
'
'' A SetH can be constructed using Dictionary.Keys or Dictionary.Items
'' Consequently, if the single item in the paramarray is a Dictionary
'' it is assumed that the sequence will be Key/Item pairs as single pair entities
'' initially we will do this as an array
'Private Sub PopulateFromDictionary(ByVal ipD As Variant)
'
'    If ipD.Count = 0 Then
'        Exit Sub
'    End If
'
'    Dim myKey As Variant
'    For Each myKey In ipD.Keys
'         s.Tmp.Add Array(myKey, ipD.Item(myKey))
'    Next
'
'End Sub
'
'
''Public Property Get Host() As Collection
''
''    If VBA.IsMissing(ipRecipient) Then
''        Set Host = s.Host
''    Else
''        Set ipRecipient = s.Host
''    End If
''
''End Property
''
''Public Property Set Host(ByRef ipRecipient As Variant)
''    Set s.Host = ipCollection
''End Property
'
'Public Function Bind(ByRef ipHash As cHashC) As SetH
'    Set s.Host = ipHash
'    Set Bind = Me
'End Function
'
'
'Public Function Clone() As SetH
'   Set Clone = SetH.Deb(s.Host)
'End Function
'
'Public Function Items() As Variant
'    Items = s.Host.Keys
'End Function
'''@Enumerator
''Public Function NewEnum() As IUnknown
''   'Set NewEnum = s.Host.[_NewEnum]
''End Function
'
'
'Public Function Count(Optional ByRef ipItem As Variant) As Long
'
'    If VBA.IsMissing(ipItem) Then
'        Count = s.Host.Count
'        Exit Function
'    End If
'
'    Dim myCount As Long
'    myCount = 0
'
'    Dim myIndex As Long
'    For myIndex = 1 To s.Host.Count
'        If ComparerHelpers.Equals(s.Host.KeyByIndex(myIndex), ipItem) Then
'            myCount = myCount + 1
'        End If
'    Next
'
'    Count = IIf(myCount = 0, -1, myCount)
'
'End Function
'
'
''@Description("Returns a subset of the SetH")
'Public Function Slice _
'( _
'    Optional ByRef ipStart As Long = 1, _
'    Optional ByRef ipRun As Long = -1, _
'    Optional ByRef ipStep As Long = 1 _
') As SetH
'
'    If s.Host.Count = 0 Then
'        Set Slice = SetH.Deb
'        Exit Function
'    End If
'
'    GuardIndexOutOfBounds ipStart, FIRST_INDEX, s.Host.Count, "SetH.Sice"
'
'    Dim myRun As Long
'    myRun = ipRun
'    If myRun = -1 Then
'        myRun = s.Host.Count - ipStart + 1
'    Else
'        GuardIndexOutOfBounds ipStart, FIRST_INDEX, myRun + ipStart, "SetH.Sice"
'    End If
'
'    Dim myS As SetH
'    Set myS = SetH.Deb
'
'    Dim mySeqIndex As Long
'    'Debug.Print ipStart + myRun - 1
'    For mySeqIndex = ipStart To ipStart + myRun - 1 Step ipStep
'        '@Ignore FunctionReturnValueDiscarded
'        myS.Add s.Host.KeyByIndex(mySeqIndex)
'    Next
'
'    Set Slice = myS
'
'End Function
'
'
''@Description("Returns a shallow copy of the SetH as a VBA collection")
'Public Function ToCollection() As Collection
'
'    Dim myC As Collection
'    Set myC = New Collection
'
'    Dim myItem As Variant
'    For Each myItem In s.Host
'        myC.Add myItem
'    Next
'
'    Set ToCollection = myC
'
'End Function
'
'
''@Description("Returns a shallow copy of the SetH as an Array.  The Lbound of the array is 1")
'Public Function ToArray() As Variant
'
'    #If twinbasic Then
'
'        ToArray = s.Host.Keys
'
'    #Else
'
'        Dim myArray As Variant
'        ReDim myArray(1 To s.Host.Count)
'
'        Dim myIndex As Long
'        For myIndex = 1 To s.Host.Count
'            If VBA.IsObject(s.Host.KeyByIndex(myIndex)) Then
'                Set myArray(myIndex) = s.Host.KeyByIndex(myIndex)
'            Else
'                myArray(myIndex) = s.Host.KeyByIndex(myIndex)
'            End If
'        Next
'
'        ToArray = myArray
'
'    #End If
'
'End Function
'
'
''@Description("Returns the LBound of the SetH or -1 if the SetH is empty.  Defined as 1 for SetH")
'Public Function FirstIndex() As Long
'
'    If s.Host.Count < FIRST_INDEX Then
'        FirstIndex = -1
'    Else
'        FirstIndex = p.firstindex
'    End If
'
'End Function
'
'
''@Description("Returns the UBound of the SetH or -1 if the SetH is empty.  Equivalent to Count.")
'Public Function LastIndex() As Long
'
'    If s.Host.Count < FIRST_INDEX Then
'        LastIndex = -1
'    Else
'        LastIndex = s.Host.Count
'    End If
'
'End Function
'
'
''@Description("Returns the index of the second item in the SetH or -1  if the SetH has less than 2 Items")
'Public Function FBOIndex() As Long
'
'    If s.Host.Count < FIRST_INDEX + 1 Then
'       FBOIndex = -1
'    Else
'        FBOIndex = p.firstindex + 1
'    End If
'
'End Function
'
''Description("Returns the index of the second to last item in the SetH, or -1 if the SetH has lees than 2 items")
'Public Function LBOIndex() As Long
'
'    If s.Host.Count < FIRST_INDEX + 1 Then
'        LBOIndex = -1
'    Else
'        LBOIndex = s.Host.Count - 1
'    End If
'
'End Function
'
'Public Property Get myTypename() As String
'   myTypename = MY_TYPENAME
'End Property
'
''@Description("Add a single item.  Returns the index at which added.  For compatibility with ArrayList")
'Public Function Add(ByVal ipItem As Variant) As Long
'
'    s.Host.Add ipItem
'    Add = s.Host.Count
'
'End Function
'
''@Description("Adds one or more items . Return the Me instance")
'Public Function AddItems(ParamArray ipParamarray() As Variant) As SetH
'
'    If ArrayOp.LacksItems(vba.cvar(ipParamArray)) Then
'        Set AddItems = Me
'        Exit Function
'    End If
'
'    Dim myItem As Variant
'    For Each myItem In ipParamarray
'        s.Host.Add myItem
'    Next
'
'    Set AddItems = Me
'
'End Function
'
'
''@Description("Add an array, list type, dictionary type, or characters of a string")
'Public Function AddRange(ByVal ipRange As Variant) As SetH
'
'    Dim myGroupId As e_Group
'    myGroupId = GroupInfo.Id(ipRange)
'    GuardInvalidRangeObject myGroupId, ipRange, "SetH.AddRange"
'
'    Select Case myGroupId
'        Case e_Group.m_string:              AddRangeFromString VBA.CStr(ipRange)
'        Case e_Group.m_array:               AddRangeUsingForEach ipRange
'        Case e_Group.m_List:                AddRangeUsingForEach ipRange
'        Case e_Group.m_Dictionary:          AddRangeFromDictionary ipRange
'    End Select
'
'    Set AddRange = Me
'End Function
'
'
'Private Sub AddRangeFromString(ByRef ipString As String)
'
'    Dim myLen As Long
'    myLen = VBA.Len(ipString)
'
'    If myLen = 0 Then
'        Exit Sub
'    End If
'
'    Dim myIndex As Long
'    For myIndex = 1 To myLen
'        s.Host.Add VBA.Mid$(ipString, myIndex, 1)
'    Next
'
'End Sub
'
'Private Sub AddRangeUsingForEach(ByVal ipForEach As Variant)
'
'    If ArrayOp.IsArray(ipForEach) Then
'        If ArrayOp.LacksItems(ipForEach) Then
'            Exit Sub
'        End If
'    ElseIf ipForEach.Count < 1 Then
'        Exit Sub
'    End If
'
'    Dim myItem As Variant
'    For Each myItem In ipForEach
'      s.Host.Add myItem
'    Next
'
'End Sub
'
'Private Sub AddRangeFromDictionary(ByVal ipD As Variant)
'
'    If ipD.Count < 1 Then
'        Exit Sub
'    End If
'
'    Dim myKey As Variant
'    For Each myKey In ipD.Keys
'      s.Host.Add Array(myKey, ipD.Item(myKey))
'    Next
'
'End Sub
'
'
'Public Property Get Item(ByRef ipIndex As Long) As Variant
'
'    GuardIndexOutOfBounds ipIndex, FIRST_INDEX, s.Host.Count, "SetH.Get_Item"
'
'    Dim myItem As Variant
'    myItem = Array(s.Host.KeyByIndex(ipIndex))
'
'     If VBA.IsObject(myItem(0)) Then
'         Set Item = myItem(0)
'     Else
'         Item = myItem(0)
'     End If
'
'End Property
'
'Public Property Let Item(ByRef ipIndex As Long, ByVal ipItem As Variant)
'
'    GuardIndexOutOfBounds ipIndex, FIRST_INDEX, s.Host.Count, "SetH.Get_Item"
'
'    s.Host.Add ipItem
'    s.Host.Remove ipIndex
'
'End Property
'
'
'Public Property Set Item(ByRef ipIndex As Long, ByVal ipItem As Variant)
'
'    GuardIndexOutOfBounds ipIndex, FIRST_INDEX, s.Host.Count, "SetH.Get_Item"
'
'   s.Host.Add ipItem
'   s.Host.Remove ipIndex
'
'End Property
'
''@Description("Inserts one or more items starting at ipIndex. Returns the index of the inserted item")
'Public Function Insert(ByRef ipIndex As Long, ByRef ipItem As Variant) As Long
'
'    If ipIndex < FIRST_INDEX Or ipIndex > s.Host.Count Then
'        Insert = -1
'        Exit Function
'    End If
'
'    s.Host.Add ipItem
'
'    Insert = ipIndex
'
'End Function
'
''@Description("Inserts one or more items starting at ipIndex")
'Public Function InsertItems(ByRef ipIndex As Long, ParamArray ipParamarray() As Variant) As SetH
'
'    If ipIndex < FIRST_INDEX Or ipIndex > s.Host.Count Then
'        Set InsertItems = Me
'        Exit Function
'    End If
'
'    If ArrayOp.LacksItems(vba.cvar(ipParamArray)) Then
'        Set InsertItems = Me
'        Exit Function
'   End If
'
'   Set InsertItems = InsertRange(ipIndex, vba.cvar(ipParamArray))
'
'End Function
'
'
''@Description("For strings inserts individual characters starting at ipIndex.  For Containers adds each item in the Container")
'Public Function InsertRange(ByRef ipIndex As Long, ByRef ipRange As Variant) As SetH
'
'    If ipIndex < FIRST_INDEX Or ipIndex > s.Host.Count Then
'        Set InsertRange = Me
'        Exit Function
'    End If
'
'    Dim myGroupId As e_Group
'    myGroupId = GroupInfo.Id(ipRange)
'    GuardInvalidRangeObject myGroupId, ipRange, "SetH.InsertRange"
'
'    Select Case GroupInfo.Id(ipRange)
'
'        Case e_Group.m_string:          InsertRangeFromString ipIndex, VBA.CStr(ipRange)
'        Case e_Group.m_array:           InsertRangeUsingForeach ipIndex, ipRange
'        Case e_Group.m_List:            InsertRangeUsingForeach ipIndex, ipRange
'        Case e_Group.m_Dictionary:      InsertRangeFromDictionary ipIndex, ipRange
'
'    End Select
'
'End Function
'
'
'Private Sub InsertRangeFromString(ByRef ipIndex As Long, ByRef ipString As String)
'
'    Dim myLen As Long
'    myLen = VBA.Len(ipString)
'
'    If myLen = 0 Then
'        Exit Sub
'    End If
'
'    Dim mySeqIndex As Long
'    mySeqIndex = ipIndex
'
'    Dim myStrIndex As Long
'    For myStrIndex = 1 To myLen
'        s.Host.Add VBA.Mid$(ipString, myStrIndex, 1)
'        mySeqIndex = mySeqIndex + 1
'    Next
'
'End Sub
'
'
'Private Sub InsertRangeUsingForeach(ByRef ipIndex As Long, ByRef ipForEach As Variant)
'
'    Dim myLen As Long
'    If ArrayOp.IsArray(ipForEach) Then
'        myLen = ArrayOp.Count(ipForEach)
'    Else
'        myLen = ipForEach.Count
'    End If
'
'    If myLen < 1 Then
'        Exit Sub
'    End If
'
'    Dim mySeqIndex As Long
'    mySeqIndex = ipIndex
'
'    Dim myItem As Variant
'    For Each myItem In ipForEach
'        s.Host.Add myItem
'        mySeqIndex = mySeqIndex + 1
'    Next
'
'End Sub
'
'
'Private Sub InsertRangeFromDictionary(ByRef ipIndex As Long, ByRef ipD As Variant)
'
'    If ipD.Count = 0 Then
'        Exit Sub
'    End If
'
'    Dim mySeqIndex As Long
'    mySeqIndex = ipIndex
'
'    Dim myKey As Variant
'    For Each myKey In ipD.Keys
'        s.Host.Add Array(myKey, ipD.Item(myKey))
'        mySeqIndex = mySeqIndex + 1
'    Next
'
'End Sub
'
'
''@Description("Removes the first encountered item from the SetH")
'Public Function Remove(ParamArray ipParamarray() As Variant) As SetH
'
'    If s.Host.Count < 1 Then
'        Set Remove = Me
'        Exit Function
'    End If
'
'    If ArrayOp.LacksItems(vba.cvar(ipParamArray)) Then
'        Set Remove = Me
'        Exit Function
'    End If
'
'    Dim myItem As Variant
'    For Each myItem In ipParamarray
'
'        Dim myIndex As Long
'        myIndex = IndexOf(myItem)
'
'        If myIndex > 0 Then
'            s.Host.Remove myIndex
'        End If
'
'    Next
'
'    Set Remove = Me
'
'End Function
'
'
''@Description("Removes the items at the Indexes.  Indexes are sorted before use.")
'Public Function RemoveAt(ParamArray ipParamarray() As Variant) As SetH
'
'    If s.Host.Count < 1 Then
'        Set RemoveAt = Me
'        Exit Function
'    End If
'
'    If ArrayOp.LacksItems(vba.cvar(ipParamArray)) Then
'        Set RemoveAt = Me
'        Exit Function
'    End If
'
'    Dim mySorted As Variant
'    mySorted = ipParamarray
'    Sorters.ShakerSortArray mySorted
'
'    Dim myIndex As Long
'    For myIndex = UBound(mySorted) To LBound(mySorted) Step -1
'        s.Host.Remove mySorted(myIndex)
'    Next
'
'    Set RemoveAt = Me
'
'End Function
'
'
''@Description("Remove items in the specified range of Indexes")
'Public Function RemoveRange(ByRef ipSeq As SetH) As SetH
'
'    If s.Host.Count < 1 Then
'        Set RemoveRange = Me
'        Exit Function
'    End If
'
'    If ipSeq.Count < 1 Then
'        Set RemoveRange = Me
'        Exit Function
'    End If
'
'    Dim myItem As Variant
'
'    For Each myItem In ipSeq
'        s.Host.Remove IndexOf(myItem)
'    Next
'
'    Set RemoveRange = Me
'
'End Function
'
'
''@Description("Remove items in the specified range of Indexes. Indexes are sorted in reverse order before use")
'Public Function RemoveIndexesRange(ByRef ipSeq As SetH) As SetH
'
'    If s.Host.Count < 1 Then
'        Set RemoveIndexesRange = Me
'        Exit Function
'    End If
'
'    If ipSeq.Count < 1 Then
'        Set RemoveIndexesRange = Me
'        Exit Function
'    End If
'
'    ipSeq.Sort.Reverse
'
'    Dim myItem As Variant
'    For Each myItem In ipSeq
'        s.Host.Remove myItem
'    Next
'
'    Set RemoveIndexesRange = Me
'
'End Function
'
''@Description("If the ParamArray is empty, removes all items otherwise removes all ocurrences of  each item in the paramarray")
'Public Function RemoveAll(ParamArray ipParamarray() As Variant) As SetH
'
'    If ArrayOp.LacksItems(vba.cvar(ipParamArray)) Then
'
'        #If twinbasic Then
'            s.Host.Clear
'        #Else
'            RemoveAllItems
'        #End If
'
'    Else
'
'        Dim myItem As Variant
'        For Each myItem In ipParamarray
'            RemoveAllSingleItem myItem
'        Next
'
'    End If
'
'    Set RemoveAll = Me
'
'End Function
'
'
'Private Sub RemoveAllItems()
'
'    Dim myIndex As Long
'    For myIndex = s.Host.Count To 1 Step -1
'        s.Host.Remove myIndex
'    Next
'
'End Sub
'
'Private Sub RemoveAllSingleItem(ByRef ipItem As Variant)
'
'    Dim myIndex As Long
'    For myIndex = s.Host.Count To 1 Step -1
'        If ComparerHelpers.Equals(s.Host.KeyByIndex(myIndex), ipItem) Then
'            s.Host.Remove myIndex
'        End If
'    Next
'
'End Sub
'
'
''@("Removes all Items from the SetH")
'Public Function Clear() As SetH
'    Set Clear = RemoveAll
'End Function
'
'
''@("Removes all Items by replacing the Host collection with a new collection")
'Public Function Reset() As SetH
'    Set s.Host = New cHashC
'    Set Reset = Me
'End Function
'
'
'Private Function RemoveAtUsingArrayOfIndexes(ByRef ipRange As Variant) As SetH
''
'    If ArrayOp.LacksItems(VBA.CVar(ipRange)) Then
'
'         Set RemoveAtUsingArrayOfIndexes = Me
'        Exit Function
'
'    End If
'
'    Sorters.ShakerSortArray VBA.CVar(ipRange)
'
'    Dim myIndex As Long
'    For myIndex = UBound(ipRange, 1) To LBound(ipRange, 1) Step -1
'
'       If ipRange(myIndex) > 0 And ipRange(myIndex) <= s.Host.Count Then
'           s.Host.Remove ipRange(myIndex)
'       End If
'
'    Next
'
'    Set RemoveAtUsingArrayOfIndexes = Me
'
'End Function
'
'
''@Description("Adds Count of Item to the SetH")
'Public Function Fill(ByRef ipItem As Variant, Optional ByVal ipCount As Long = 1) As SetH
'
'    If ipCount < 1 Then
'        Set Fill = Me
'        Exit Function
'    End If
'
'    Dim myCount As Long
'    For myCount = 1 To ipCount
'        s.Host.Add ipItem
'    Next
'
'    Set Fill = Me
'
'End Function
'
'
''@Description("Returns the first N items of the SetH, Default 1")
'Public Function Head(Optional ByVal ipSplitAfter As Long = 1) As SetH
'
'    If s.Host.Count <= 1 Then
'        Set Head = SetH.Deb
'        Exit Function
'    End If
'
'    If ipSplitAfter < 1 Then
'        Set Head = SetH.Deb
'        Exit Function
'    End If
'
'    If ipSplitAfter >= s.Host.Count Then
'        Set Head = Me.Clone
'        Exit Function
'    End If
'
'    Set Head = Me.Slice(1, ipSplitAfter)
'
'End Function
'
'
''@Description("Returns the Items after SplitAfter, Default is Index 2")
'Public Function Tail(Optional ByVal ipSplitAfter As Long = 1) As SetH
'
'    If s.Host.Count = 0 Then
'        Set Tail = SetH.Deb
'        Exit Function
'    End If
'
'    If ipSplitAfter < 1 Then
'        Set Tail = Me.Clone
'        Exit Function
'    End If
'
'    If ipSplitAfter >= s.Host.Count Then
'        Set Tail = SetH.Deb
'        Exit Function
'    End If
'
'    Set Tail = Me.Slice(ipSplitAfter + 1)
'
'End Function
'
''@Description("True if the Item is found in the SetH")
'Public Function HoldsItem(ByRef ipItem As Variant) As Boolean
'
'    Dim myItem As Variant
'    For Each myItem In s.Host
'        If ComparerHelpers.Equals(myItem, ipItem) Then
'            HoldsItem = True
'            Exit Function
'        End If
'    Next
'
'    HoldsItem = False
'
'End Function
'
'
''@Description("True if the Item is not found in the SetH")
'Public Function LacksItem(ByRef ipItem As Variant) As Boolean
'    LacksItem = Not Me.HoldsItem(ipItem)
'End Function
'
'
''@Description("ArrayList terminology for HoldsItem")
'Public Function Contains(ByRef ipItem As Variant) As Boolean
'    Contains = Me.HoldsItem(ipItem)
'End Function
'
'
''@Description"VBALib.KvpH and cHashd terminology for HoldsItem")
'Public Function Exists(ByRef ipItem As Variant) As Boolean
'    'The new exists method for Collection is for Keys not for Items
'    Exists = HoldsItem(ipItem)
'End Function
'
'
''@Description("Alternative to Count > 0")
'Public Function HoldsItems() As Boolean
'    HoldsItems = s.Host.Count > 0
'End Function
'
''@Description("Alternative to Count=0")
'Public Function LacksItems() As Boolean
'    LacksItems = s.Host.Count < 1
'End Function
'
''@Description("Mutates the SetH in place by applying the function in IMapper to each item")
'Public Function MappedIt(ByRef ipMapper As IMapper) As SetH
'    'Mapped it works by reference but we allow a function return to enable fluid use of Mappedit
'
'    Dim myIndex As Long
'    For myIndex = FirstIndex To LastIndex
'
'        Dim myItem As Variant
'        myItem = ipMapper.ExecMapper(s.Host.KeyByIndex(myIndex))
'
'        If VBA.IsObject(myItem(0)) Then
'            Set Me.Item(myIndex) = myItem(0)
'        Else
'            Me.Item(myIndex) = myItem(0)
'        End If
'
'    Next
'
'    Set MappedIt = Me
'
'End Function
'
'
''@Description("Returns a new SetH where each item in the Host has been mutated by the function provided by IMapper")
'Public Function MapIt(ByRef ipMapper As IMapper) As SetH
'
'    Dim myS As SetH
'    Set myS = SetH.Deb
'
'    Dim myItem As Variant
'    For Each myItem In s.Host
'        myS.Add ipMapper.ExecMapper(myItem)(0)
'    Next
'
'    Set MapIt = myS
'
'End Function
'
'
''@Description("Returns a new SetH comprised of item where IComparer is true")
'Public Function FilterIt(ByVal ipComparer As IComparer) As SetH
'
'    Dim myS As SetH
'    Set myS = SetH.Deb
'
'    Dim myItem As Variant
'    For Each myItem In s.Host
'
'        If ipComparer.ExecCmp(myItem) Then
'            myS.Add myItem
'        End If
'
'    Next
'
'    Set FilterIt = myS
'
'End Function
'
'
''@Description("Returns a count where the function provided by IComparer is true")
'Public Function CountIt(ByVal ipComparer As IComparer) As Long
'
'   Dim myCount As Long
'   myCount = 0
'
'   Dim myItem As Variant
'   For Each myItem In s.Host
'
'       If ipComparer.ExecCmp(myItem) Then
'           myCount = myCount + 1
'       End If
'
'   Next
'
'   CountIt = myCount
'
'End Function
'
'
''@Description("Returns a single value based on the function applied by the IReducer").
'Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant
'
'   Dim myItem As Variant
'   myItem = Array(ipReducer.ExecReduction(Me))
'
'   If VBA.IsObject(myItem(0)) Then
'      Set ReduceIt = myItem(0)
'   Else
'       ReduceIt = myItem(0)
'   End If
'
'End Function
'
'
''@DescriptioN("Returns the item at FirstIndex")
'Public Property Get First() As Variant
'
'   GuardInvalidIndex 1, s.Host.Count, "SetH.Get_First"
'
'   If VBA.IsObject(s.Host.KeyByIndex(First_Index)) Then
'       Set First = s.Host.KeyByIndex(First_Index)
'   Else
'       First = s.Host.KeyByIndex(First_Index)
'   End If
'
'End Property
'
'
''@DescriptioN("Assigns to the Item at FirstIndex")
'Public Property Let First(ByVal ipItem As Variant)
'
'   GuardInvalidIndex 1, s.Host.Count, "SetH.Let_First"
'
'   Me.Item(First_Index) = ipItem
'
'End Property
'
''@Description("Assigns to the Item at FirstIndex")
'Public Property Set First(ByVal ipItem As Variant)
'
'   GuardInvalidIndex 1, s.Host.Count, "SetH.Set_Last"
'
'   Set Me.Item(FirstIndex) = ipItem
'
'End Property
'
''@Description("Returns the Item at LastIndex")
'Public Property Get Last() As Variant
'
'   GuardInvalidIndex s.Host.Count, s.Host.Count, "SetH.Set_Last"
'
'   Dim myLastIndex As Long
'   myLastIndex = s.Host.Count
'
'    If VBA.IsObject(s.Host.KeyByIndex(myLastIndex)) Then
'       Set Last = s.Host.KeyByIndex(myLastIndex)
'   Else
'       Last = s.Host.KeyByIndex(myLastIndex)
'   End If
'End Property
''@Description("Assigns to the Item at LastIndex")
'Public Property Let Last(ByVal ipItem As Variant)
'
'   GuardInvalidIndex s.Host.Count, s.Host.Count, "SetH.Set_First"
'
'   Me.Item(s.Host.Count) = ipItem
'
'End Property
'
'
''@Description("Assigns to the item at LastIndex")
'Public Property Set Last(ByVal ipItem As Variant)
'
'   GuardInvalidIndex s.Host.Count, s.Host.Count, "SetH.Set_First"
'
'   Set Me.Item(s.Host.Count) = ipItem
'
'End Property
'
'
''@Description("Returns the item at FirstIndex")
'Public Property Get FBO() As Variant
'
'   GuardInvalidIndex 2, s.Host.Count, "SetH.Get_FBO"
'
'   Dim FBOIndex As Long
'   FBOIndex = p.firstindex + 1
'
'   If VBA.IsObject(s.Host.KeyByIndex(FBOIndex)) Then
'       Set FBO = s.Host.KeyByIndex(FBOIndex)
'   Else
'       FBO = s.Host.KeyByIndex(FBOIndex)
'   End If
'
'End Property
'
''@Description("Assigns to the Item at FirstIndex")
'Public Property Let FBO(ByVal ipItem As Variant)
'
'   GuardInvalidIndex 2, s.Host.Count, "SetH.Let_FBO"
'
'   Set Me.Item(First_Index + 1) = ipItem
'
'End Property
'
''@Description("Assigns to the Item at FirstIndex")
'Public Property Set FBO(ByVal ipItem As Variant)
'
'   GuardInvalidIndex 2, s.Host.Count, "SetH.Set_FBO"
'
'   Set Me.Item(FirstIndex + 1) = ipItem
'
'End Property
'
''@Description("Returns the Item at LastIndex")
'Public Property Get LBO() As Variant
'
'   GuardInvalidIndex 2, s.Host.Count, "SetH.Let_FBO"
'
'   Dim myLBOIndex As Long
'   myLBOIndex = s.Host.Count - 1
'
'    If VBA.IsObject(s.Host.KeyByIndex(myLBOIndex)) Then
'       Set LBO = s.Host.KeyByIndex(myLBOIndex)
'   Else
'       LBO = s.Host.KeyByIndex(myLBOIndex)
'   End If
'
'End Property
'
'
''@Description("Assigns to the Item at LastIndex")
'Public Property Let LBO(ByVal ipItem As Variant)
'
'   GuardInvalidIndex 2, s.Host.Count, "SetH.Let_FBO"
'
'   Me.Item(s.Host.Count - 1) = ipItem
'
'End Property
'
'
''@Description("Assigns to the item at LastIndex")
'Public Property Set LBO(ByVal ipItem As Variant)
'
'    GuardInvalidIndex 2, s.Host.Count, "SetH.Let_FBO"
'
'   Set Me.Item(s.Host.Count - 1) = ipItem
'
'End Property
'
'
''@Description("Returns the index of the first Item in the specified range or -1 if not found")
'Public Function IndexOf _
'( _
'    ByVal ipItem As Variant, _
'    Optional ByVal ipStart As Long = 1, _
'    Optional ByVal ipRun As Long = -1 _
') As Long
'
'    If s.Host.Count = 0 Then
'        IndexOf = -1
'        Exit Function
'    End If
'
'    GuardIndexOutOfBounds ipStart, FIRST_INDEX, s.Host.Count, "SetH.IndexOf"
'
'    Dim myRun As Long
'    myRun = ipRun
'    If myRun = -1 Then
'        myRun = s.Host.Count - ipStart + 1
'    Else
'        GuardIndexOutOfBounds myRun, FIRST_INDEX, s.Host.Count, "SetH.IndexOf"
'    End If
'
'    Dim myIndex As Long
'    For myIndex = ipStart To ipStart + myRun - 1
'        If ComparerHelpers.Equals(s.Host.KeyByIndex(myIndex), ipItem) Then
'            IndexOf = myIndex
'            Exit Function
'        End If
'    Next
'
'    IndexOf = -1
'
' End Function
'
'
' '@Description("Returns the index of the last item inthe specified range or -1 if not found")
'Public Function LastIndexOf _
'( _
'    ByVal ipItem As Variant, _
'    Optional ByVal ipStart As Long = 1, _
'    Optional ByVal ipRun As Long = -1 _
') As Long
'
'    If s.Host.Count < 1 Then
'        LastIndexOf = -1
'        Exit Function
'    End If
'
'    GuardIndexOutOfBounds ipStart, FIRST_INDEX, s.Host.Count, "SetH.LastIndexOf"
'
'    Dim myRun As Long
'    myRun = ipRun
'    If myRun = -1 Then
'        myRun = s.Host.Count - ipStart + 1
'    Else
'        GuardIndexOutOfBounds myRun, FIRST_INDEX, s.Host.Count, "SetH.LastIndexOf"
'    End If
'
'    Dim myIndex As Long
'    For myIndex = ipStart + myRun - 1 To ipStart Step -1
'        If s.Host.KeyByIndex(myIndex) = ipItem Then
'            LastIndexOf = myIndex
'            Exit Function
'        End If
'    Next
'
'    LastIndexOf = -1
'
' End Function
'
'
' Public Property Get TypeName() As String
'    TypeName = MY_TYPENAME
'End Property
'
'' Stack and queue usage
'' Queue
'' Enqueue - add at end
'' Dequeue - removes from position 1
'
'' Stack
'' Push = add at end
'' Pop remove from end
''
''
''
'Public Function Push(ParamArray ipParamarray() As Variant) As SetH
'     Set Push = AddRange(vba.cvar(ipParamArray))
'End Function
'
'Public Function PushRange(ByRef ipRange As Variant) As SetH
'    Set PushRange = AddRange(ipRange)
'End Function
'
'Public Function Pop() As Variant
'
'    Dim myLastIndex As Long
'    myLastIndex = s.Host.Count
'
'    Dim myItem As Variant
'    myItem = Array(s.Host.KeyByIndex(myLastIndex))
'
'    If VBA.IsObject(myItem(0)) Then
'        Set Pop = myItem(0)
'    Else
'        Pop = myItem(0)
'    End If
'
'    s.Host.Remove myLastIndex
'
'End Function
'
'Public Function PopRange(Optional ByVal ipCount As Long = 1) As SetH
'
'    If ipCount < 0 Then
'        Set PopRange = SetH.Deb
'        Exit Function
'    End If
'
'    If ipCount > s.Host.Count Then
'        Set PopRange = Me.Reversed
'        Me.RemoveAll
'        Exit Function
'    End If
'
'    Dim myS As SetH
'    Set myS = SetH.Deb
'
'    Dim myIndex As Long
'    For myIndex = s.Host.Count To s.Host.Count - ipCount + 1 Step -1
'        myS.Add s.Host.KeyByIndex(myIndex)
'        s.Host.Remove myIndex
'    Next
'
'    Set PopRange = myS
'
'End Function
'
'
'Public Function enQueue(ParamArray ipParamarray() As Variant) As SetH
'   Set enQueue = AddRange(vba.cvar(ipParamArray))
'End Function
'
'
'Public Function EnqueueRange(ByRef ipRange As Variant) As SetH
'   Set EnqueueRange = AddRange(ipRange)
'End Function
'
'
'Public Function Dequeue() As Variant
'
'    Dim myItem As Variant
'    myItem = Array(s.Host.KeyByIndex(First_Index))
'
'    If VBA.IsObject(myItem(0)) Then
'        Set Dequeue = myItem(0)
'    Else
'        Dequeue = myItem(0)
'    End If
'
'    s.Host.Remove FIRST_INDEX
'
'End Function
'
'
'Public Function DequeueRange(Optional ByVal ipCount As Long = 1) As SetH
'
'    If ipCount < 0 Then
'        Set DequeueRange = SetH.Deb
'        Exit Function
'    End If
'
'    If ipCount > s.Host.Count Then
'        Set DequeueRange = Me.Clone
'        Me.RemoveAll
'        Exit Function
'    End If
'
'    Dim myS As SetH
'    Set myS = SetH.Deb
'
'    Dim myIndex As Long
'    For myIndex = 1 To ipCount
'        myS.Add s.Host.KeyByIndex(First_Index)
'        s.Host.Remove FIRST_INDEX
'    Next
'
'    Set DequeueRange = myS
'
'End Function
'
''
'' Public Function RotL(Optional ByVal ipCount As Long = 1) As SetH Implements ISeq.RotL
''     Dim myCount As Long
''     For myCount = 1 To ipCount
''         AddItems Dequeue
''     Next
''     Return Me
'' End Sub
''
'' Public Function RotR(Optional ByVal ipcount As Long = 1) As SetH Implements ISeq.RotR
''     Dim myCount As Long
''     For myCount = 1 To ipCount
''         Insert 1, Pop
''     Next
''     Return Me
'' End Sub
''
'
''
' Public Function Sort() As SetH
'     Sorters.ShakerSortByItem Me
'     Set Sort = Me
' End Function
''
'Public Function Sorted() As SetH
'    Dim myS As SetH
'    Set myS = Me.Clone
'    Sorters.ShakerSortByItem myS
'    Set Sorted = myS
'End Function
'
'
'
'
''@Description("Swaps the Items at the specified indexes")
'Public Sub Swap(ByRef ipLHSIndex As Long, ByRef ipRHSIndex As Long)
'
'    Dim myTemp As Variant
'
'    If VBA.IsObject(Item(ipLHSIndex)) Then
'        Set myTemp = Item(ipLHSIndex)
'    Else
'        myTemp = Item(ipLHSIndex)
'    End If
'
'    If VBA.IsObject(Item(ipRHSIndex)) Then
'        Set Item(ipLHSIndex) = Item(ipRHSIndex)
'    Else
'        Item(ipLHSIndex) = Item(ipRHSIndex)
'    End If
'
'    If VBA.IsObject(myTemp) Then
'        Set Item(ipRHSIndex) = myTemp
'    Else
'        Item(ipRHSIndex) = myTemp
'    End If
'
'End Sub
'
'
'Public Function Reverse() As SetH
'
'    Dim myFirst As Long
'    myFirst = 1
'
'    Dim myLast As Long
'    myLast = s.Host.Count
'
'    Dim mySwap As Variant
'
'    Do While myFirst < myLast
'
'        If VBA.IsObject(s.Host.KeyByIndex(myFirst)) Then
'            Set mySwap = s.Host.KeyByIndex(myFirst)
'        Else
'            mySwap = s.Host.KeyByIndex(myFirst)
'        End If
'
'        If VBA.IsObject(s.Host.KeyByIndex(myLast)) Then
'            Set Me.Item(myFirst) = s.Host.KeyByIndex(myLast)
'        Else
'            Me.Item(myFirst) = s.Host.KeyByIndex(myLast)
'        End If
'
'        If VBA.IsObject(mySwap) Then
'            Set Me.Item(myLast) = mySwap
'        Else
'            Me.Item(myLast) = mySwap
'        End If
'
'         myFirst = myFirst + 1
'         myLast = myLast - 1
'
'     Loop
'
'
'    Set Reverse = Me
'
' End Function
'
' Public Function Reversed() As SetH
'
'     Dim myS As SetH
'     Set myS = SetH.Deb
'
'     Dim myIndex As Long
'     For myIndex = s.Host.Count To 1 Step -1
'        myS.Add s.Host.KeyByIndex(myIndex)
'    Next
'
'    Set Reversed = myS
'
' End Function
'
' '@Description("Returns the unique subset of the host SetH")
'Public Function Unique() As SetH
'
'    Dim myS As SetH
'    Set myS = SetH.Deb
'
'    Select Case s.Host.Count
'
'        Case 0:             Set Unique = myS
'
'        Case 1:             Set Unique = Me.Clone
'
'        Case Else
'
'            Dim myItem As Variant
'            For Each myItem In s.Host
'                If myS.LacksItem(myItem) Then
'                    myS.Add myItem
'                End If
'            Next
'
'            Set Unique = myS
'
'    End Select
'
'End Function
'
'
'Public Function SetOf(ByRef ipSet As e_SetoF, ByRef ipRangeItem As SetH) As SetH
'
'    Select Case ipSet
'
'        Case e_SetoF.m_Common:           Set SetOf = SetOfCommon(ipRangeItem)
'        Case e_SetoF.m_HostOnly:         Set SetOf = SetOfHostOnly(ipRangeItem)
'        Case e_SetoF.m_ParamOnly:        Set SetOf = SetOfParamOnly(ipRangeItem)
'        Case e_SetoF.m_NotCommon:        Set SetOf = SetOfNotCommon(ipRangeItem)
'        Case e_SetoF.m_Unique:           Set SetOf = SetofUnique(ipRangeItem)
'        Case Else
'            Err.Raise 17 + vbObjectError, _
'                "Seq.SetOf", _
'                Fmt.Text("Invalid e_set member. Expecting {0} to {1}. Got {2}", e_SetoF.m_First, e_SetoF.m_Last, ipSet)
'
'    End Select
'
' End Function
'
'
''Private Function ResolveSetOfRangeItem(ByRef ipRangeItem As Variant) As Variant
''
''    Dim myRangeItem As Variant
''
''    Select Case GroupInfo.Id(ipRangeItem)
''        Case e_Group.m_string:              myRangeItem = Array(SetH.Deb(ipRangeItem))
''        Case e_Group.m_array:               myRangeItem = Array(ipRangeItem)
''        Case e_Group.m_List:                myRangeItem = Array(ipRangeItem)
''        Case e_Group.m_Dictionary:          myRangeItem = Array(SetH.Deb(ipRangeItem))
''    End Select
''
''    ResolveSetOfRangeItem = myRangeItem
''
''End Function
'
'Private Function SetOfCommon(ByRef ipRangeItem As SetH) As SetH
'
'    Dim myS As SetH
'    Set myS = SetH.Deb
'
'    Dim myItem As Variant
'    For Each myItem In ipRangeItem
'        If myS.LacksItem(myItem) Then
'            If Me.HoldsItem(myItem) Then
'                myS.Add myItem
'            End If
'        End If
'    Next
'
'    Set SetOfCommon = myS
'
'End Function
'
'
'
'Private Function SetOfHostOnly(ByRef ipRangeItem As SetH) As SetH
'
'    Dim myS As SetH
'    Set myS = Me.Dedup
'
'    Dim myItem As Variant
'    For Each myItem In ipRangeItem
'
'        If myS.HoldsItem(myItem) Then
'            myS.RemoveAll myItem
'        End If
'
'    Next
'
'    Set SetOfHostOnly = myS
'
'
'End Function
'
'Private Function SetOfParamOnly(ByRef ipRangeItem As SetH) As SetH
'
'    Dim myS As SetH
'    Set myS = SetH.Deb
'
'    Dim myItem As Variant
'    For Each myItem In ipRangeItem
'
'        If Me.LacksItem(myItem) Then
'            myS.Add myItem
'        End If
'
'    Next
'
'    Set SetOfParamOnly = myS
'
'
'End Function
'
'Private Function SetOfNotCommon(ByRef ipRangeItem As SetH) As SetH
'
'    Dim myS As SetH
'    Set myS = Me.Dedup
'
'    Dim myItem As Variant
'    For Each myItem In ipRangeItem
'
'        If myS.HoldsItem(myItem) Then
'            myS.RemoveAll myItem
'        Else
'            myS.Add myItem
'        End If
'
'    Next
'
'    Set SetOfNotCommon = myS
'
'End Function
'
'Private Function SetofUnique(ByRef ipRangeItem As SetH) As SetH
'
'    Dim myS As SetH
'    Set myS = Me.Dedup
'
'    Dim myItem As Variant
'    For Each myItem In ipRangeItem
'        If myS.LacksItem(myItem) Then
'            myS.Add myItem
'        End If
'    Next
'
'    Set SetofUnique = myS
'
'End Function
'
'
'
' Public Function Freq() As KvpH
'
'     Dim myKvp As KvpH
'     Set myKvp = KvpH.Deb
'
'    Dim myItem As Variant
'    For Each myItem In s.Host
'         If myKvp.Exists(myItem) Then
'             myKvp.Item(myItem) = myKvp.Item(myItem) + 1
'         Else
'             myKvp.Add myItem, 1
'         End If
'
'     Next
'
'     Set Freq = myKvp
'
' End Function



