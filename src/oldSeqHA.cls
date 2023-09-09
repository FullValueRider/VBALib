VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "oldSeqHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SeqHSeqH is a multipurpose Collection/ArrayList replacement.  It uses a double linked list with two sets of links to manage hash clashes and ordr of addition/insertion to store the items.  Items are access using a single 1 based index. For a Keyed collection, use one of the Kvp classes"
'@IgnoreModule
''@PredeclaredId
''@Exposed
''@Folder("Seq)
''@ModuleDescription("SeqHASeqHA is a multipurpose Collection/ArrayList replacement.  It uses a double linked list with two sets of links to manage hash clashes and ordr of addition/insertion to store the items.  Items are access using a single 1 based index. For a Keyed collection, use one of the Kvp classes")
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'' Based on cHashD from https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs)&s=f049d882b4fc4d27f4bc2f9bc383a53e
'' cHashD, Olaf Schmidt in August 2016 (enhanced about mixed, variable KeyTypes and Remove-functions in 2020)
'' Reformatted and Rubberducked by Steven Laycock July 2023
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'' Code line limit should be 120 characters.
'' Comment line limit should be 80 characters
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'' SeqHA is a list class and consequently will store items with the same value.
'' This means that hash clashes are very likely to happen
'' This class uses an array to store the unique hashes and a linked list to
'' manage the hash clashes.
'' the hash table stores the node containing the first hash for a particular value
'' subsequent hash clashes are linked after the first node for the hash
'' nodes are linked in Hash/Item sequence.
'' The order of addition of values is also preserved by adding a second pair
'' of node references to the node class
'
''
'' To allow insertion a second set of references is required to track order of addition/insertion
'' It should be noted that the Add/Insert methods will use either appendnode or insert node depending on the hash
'
'
'Option Explicit
'
'' Hash-Multiplikator ' as derived from olaf schmidts cHashD
'Private Const HMul                          As Long = 3727
'
'Private Const MY_TYPENAME                   As String = "SeqHAA"
'Private Const FIRST_INDEX                   As Long = 1
'Private Const MIN_SEQ_SIZE                  As Long = 512
'
''Private Type Host
''    Count                                   As Long
''End Type
''
'
'Private Type State
'    ' W is used in hash calculations
'    W()                                     As Integer
'    ' typedef used to access the data area of W()
'    sAW                                     As SAFEARRAY1D
'    ' a feelgood typedef to allow consistent terminology between seq classes where the count of nodes in not provided by the Host object
'   ' Host                                    As Host
'    ' A value with all bits set to 1 which represents the number of hashslots available
'    ' it is used to AND intermediate values in the the hash calculation
'    ' to ensure that the hash is always within the range of available HashSlots
'    ' i.e. it is always a power of 2 minus 1
'    ' and consequenlty the number of hashslots can only be doubled or halved
'    HashSlotMask                            As Long
'
'    ' An array where nodes are stored using their hash as the index.
'    ' Each HashSlot holds a reference to a single node, which is the first node with a particular hash
'    HashSlots()                             As SeqHNode
'
'    ' first node added to the seq
'    HashRoot                                As SeqHNode
'    ' last node appended to the seq (note not added)
'    HashLast                                As SeqHNode
'
'    ' Node containing order=1
'    OrderRoot                               As SeqHNode
'    ' Node contain the last order
'    OrderLast                               As SeqHNode
'
''    ' flag used to prevent updating of the node order field
''    ' when redistrubing hashes due to a resize.
''    Resizing                                As Boolean
'
'End Type
'
'Private s                                   As State
'
'Private Type Properties
'    Size                                    As Long ' The dimmed size of s.HashSlots
'End Type
'
'
''@Ignore MoveFieldCloserToUsage
'Private p                                   As Properties
'
'' Deb is a Factory method of the PredeclaredId which returns a new instance of SeqHA
''=================================================================================
'' The input for Deb is a ParamArray.
'' The ParamArray is interpreted according to the folowing rules
'' If the ParamArray contains no Items
'' - A empty SeqHAuence (count = 0)
'' If the ParamArray has 2 or more items
'' - An Seq of size equal to the ParamArray, items as generated by for each over the ParamArray
'' If the ParamArray has a single Item
'' - Number: SeqHA with specified capacity with each Item being the Value 'Empty'
'' - String: A SeqHA whose capacity is the length of the string with one character per Item
'' - Array: SeqHA of items as provided by For Each over the array.  Multidimension arrays are flattened but the position in the original array can be calculated by the user if required. See also Bind
'' - Dictionaries: a SeqHA of key/value arrays as generated from for each on the Keys array.
'' - Collection: The collection is copied to the Host collection.  There is no refererence link to the original collection (See Host Property)
'' - SeqHA: Items as generated by for each on the SeqHA
'' - mscorlib Arraylist: Items as generated by for each on the arrylist
'' - mscorlib Stack: Items as generated by for each on the stack
'' - mscorlib Queue: Items as generated by for eac on the queue
'
''The printby methods are used for debugging.  do not use complex keys
'Public Sub PrintByHash()
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'    Do
'        Debug.Print myNode.Hash, myNode.Order, myNode.Item
'        Set myNode = myNode.HashSucc
'    Loop Until myNode Is Nothing
'
'End Sub
'
'Public Sub PrintByOrder()
'
'    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
'    Do
'        Debug.Print myNode.Hash, myNode.Order, myNode.Item
'        Set myNode = myNode.OrderSucc
'    Loop Until myNode Is Nothing
'
'End Sub
'
''@DefaultMember
'Public Function Deb(ParamArray ipParamarray() As Variant) As oldSeqHA
'    With New oldSeqHA
'        Set Deb = .ConstructInstance(VBA.CVar(ipParamarray))
'    End With
'End Function
'
'
'Friend Function ConstructInstance(ByRef ipFPA As Variant) As oldSeqHA
'
'    s.sAW.cDims = 1
'    s.sAW.cbElements = 2
'    s.sAW.cLocks = 1
'    s.sAW.fFeatures = &H11                       'FADF_AUTO=&H1 || FADF_FIXEDSIZE=&H10
'    BindArray s.W, VarPtr(s.sAW)
'    SetSize 16384                                'Hangover from cHashD: sets the number of hashslots (we use 1 based index, cHAshD used 0 based index)
'
'     ' Check to see if we return a default 'empty' seq
'    If ArrayOp.LacksItems(ipFPA) Then
'        Set ConstructInstance = Me
'        Exit Function
'    End If
'
'    ' ipParamArray contains one or more items depending on the input type.
'    ' if the input contains only a single item we want to use that item
'    ' else we want to treat the forwrded paramarray as a single array item
'
'    Dim myItem As Variant
'    If ArrayOp.HoldsOneItem(ipFPA) Then
'        If VBA.IsObject(ipFPA(0)) Then
'            Set myItem = ipFPA(0)
'        Else
'            myItem = ipFPA(0)
'        End If
'    Else
'        myItem = ipFPA
'    End If
'
'    ' Resolve the items in the forwarded paramarray using the rules for
'    ' instantiation as described above.
'    Select Case True
'
'        Case VBA.IsArray(myItem):                   Set ConstructInstance = AddRange(ArrayOp.Linearise(myItem))
'        Case GroupInfo.IsContainer(myItem):         Set ConstructInstance = AddRange(myItem)
'        Case GroupInfo.IsString(myItem):            Set ConstructInstance = AddRange(myItem)
'        Case GroupInfo.IsNumber(myItem):            Set ConstructInstance = ConstructInstanceByCount(VBA.CLng(myItem))
'        Case Else:                                  Set ConstructInstance = AddItems(myItem)
'    End Select
'
'End Function
'
'
'Private Function ConstructInstanceByCount(Optional ByRef ipNewSize As Long = MIN_SEQ_SIZE) As oldSeqHA
'
'    If ipNewSize < MIN_SEQ_SIZE Then
'        Err.Raise 17 + vbObjectError, _
'            "VBALib.SeqHA", _
'            Fmt.Dbg("Seq should have a size of at least 512 items. Got {0}", ipNewSize)
'    End If
'
'    SetSize ipNewSize
'
'    Set ConstructInstanceByCount = Me
'
'End Function
'
''@Description("set the size of hash array to be the 2^n value just greater than ipNewSize")
''@Ignore AssignedByValParameter
'Public Sub SetSize(Optional ByVal ipNewSize As Long)
'
'    If ipNewSize < 0 Then
'        ipNewSize = MIN_SEQ_SIZE
'    End If
'
'    ' calculate a long whose value is the first 2^n-1 that is greater than ipNewSize
'    ' this provides a bitfield of 1's used to set the maximum for a hash value
'    s.HashSlotMask = (2 ^ Int((VBA.Log(ipNewSize) / VBA.Log(2)) + 1)) - 1
'    If s.HashSlotMask < ipNewSize Then
'        s.HashSlotMask = (s.HashSlotMask + 1) * 2 - 1
'    End If
'
'    ' Redim without preserve should mean that we have no entries in s.HashSlots
'    ReDim s.HashSlots(FIRST_INDEX To s.HashSlotMask + 1)
'    p.Size = s.HashSlotMask
'    If s.HashRoot Is Nothing Then
'        Exit Sub
'    End If
'
'    ' update the hashes in the nodes
'    ' then populate the new s.HashSlots array using the new hashes
'    ' this does not change the order of addition
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'    Do
'        myNode.Hash = GetHash(myNode.Item)
'        Set myNode = myNode.HashSucc
'    Loop Until myNode Is Nothing
'
'    ' repopulate s.HashSlots according to the new hashes
'    ' the order of nodes is unchanged by the reallocation of hashslots
'    ' the rootnode should be unchanged by the rehashing
'    ' s.HashLast node will have been updated by the rehashing
'    Set s.HashSlots(s.HashRoot.Hash) = s.HashRoot
'    Set s.HashLast = s.HashRoot
'    Set myNode = s.HashRoot
'
'    ' linking nodes in accordance with the hash and item
'    ' is done by the append and inset methods so
'    ' here we only need to focus on the hash value
'    Do Until myNode.HashSucc Is Nothing
'
'        Set myNode = myNode.HashSucc
'
'        If s.HashSlots(myNode.Hash) Is Nothing Then
'            Set s.HashSlots(myNode.Hash) = myNode
'            AppendByHash myNode
'        Else
'            InsertbyHash myNode
'        End If
'
'    Loop
'
'End Sub
'
'' There is no bind method for SeqHA as SeqHA is not based on a 'container' objects
'
'
''@Description("Makes a shallow copy of the seq")
'Public Function Clone() As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'
'    If s.HashRoot Is Nothing Then
'        Set Clone = myS
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
'    Do
'
'        myS.Add myNode.Item
'        Set myNode = myNode.OrderSucc
'
'    Loop Until myNode Is Nothing
'
'    Set Clone = myS
'
'End Function
'
''@Description("Returns the size of the Seq or the number of specified Items.  Returns -1 rather than 0 if there are no items in the Seq but count item will return 0 if not found")
'
'Public Function Count(Optional ByRef ipItem As Variant) As Long
'
'    If s.HashRoot Is Nothing Then
'        Count = -1
'        Exit Function
'    End If
'
'    If VBA.IsMissing(ipItem) Then
'        Count = s.OrderLast.Order
'        Exit Function
'    End If
'
'    Dim myCount As Long: myCount = 0
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'    Do
'
'        If Comparers.EQ(ipItem, myNode.Item) Then
'            myCount = myCount + 1
'        End If
'
'        Set myNode = myNode.HashSucc
'
'    Loop Until myNode Is Nothing
'
'    Count = myCount
'
'End Function
'
'
'Public Function Items() As Variant
'
'    If s.HashRoot Is Nothing Then
'        Items = Array()
'        Exit Function
'    End If
'
'    Items = Me.ToArray
'
'End Function
'
'
''@Description("Returns a subset of the SeqHA")
'Public Function Slice _
'( _
'Optional ByRef ipStart As Long = 1, _
'Optional ByRef ipRun As Long = -1, _
'Optional ByRef ipStep As Long = 1 _
') As oldSeqHA
'
'    If s.HashRoot Is Nothing Then
'        Set Slice = oldSeqHA.Deb
'        Exit Function
'    End If
'
'    GuardIndexOutOfBounds ipStart, FIRST_INDEX, s.OrderLast.Order, "SeqHA.Slice"
'
'    Dim myRun As Long
'    myRun = ipRun
'    If myRun < FIRST_INDEX Then
'        myRun = s.OrderLast.Order - ipStart + 1
'    Else
'        GuardIndexOutOfBounds myRun + ipStart - 1, FIRST_INDEX, s.OrderLast.Order, "SeqHA.Slice"
'    End If
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'
'    Dim mySeqIndex As Long
'    For mySeqIndex = ipStart To ipStart + myRun - 1 Step ipStep
'        myS.Add Me.Item(mySeqIndex)
'    Next
'
'    Set Slice = myS
'
'End Function
''@Description("Returns a shallow copy of the SeqHA as an Array.  The Lbound of the array is 1")
'Public Function ToArray() As Variant
'
'    If s.HashRoot Is Nothing Then
'        ToArray = Array()
'        Exit Function
'    End If
'
'    Dim myA As Variant: ReDim myA(s.OrderRoot.Order To s.OrderLast.Order)
'
'    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
'    Do
'
'        If VBA.IsObject(myNode.Item) Then
'            Set myA(myNode.Order) = myNode.Item
'        Else
'            myA(myNode.Order) = myNode.Item
'        End If
'
'        Set myNode = myNode.OrderSucc
'
'    Loop Until myNode Is Nothing
'
'    ToArray = myA
'
'End Function
'
'
''@Description("Returns a shallow copy of the SeqHA as a VBA collection")
'Public Function ToCollection() As Collection
'
'    Dim myC As Collection: Set myC = New Collection
'
'    If s.HashRoot Is Nothing Then
'        Set ToCollection = myC
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
'    Do
'        myC.Add myNode.Item
'        Set myNode = myNode.OrderSucc
'    Loop Until myNode Is Nothing
'
'    Set ToCollection = myC
'
'End Function
'
'
''@Description("Returns the Host of the first Index")
'Public Function FirstIndex() As Long
'
'    If s.HashRoot Is Nothing Then
'        FirstIndex = -1
'    Else
'        FirstIndex = p.FirstIndex
'    End If
'
'End Function
'
'
''@Description("Returns the UBound of the SeqHA or -1 if the SeqHA is empty.  Equivalent to Count.")
'Public Function LastIndex() As Long
'
'    If s.HashRoot Is Nothing Then
'        LastIndex = -1
'    Else
'        LastIndex = s.OrderLast.Order
'    End If
'
'End Function
'
''@Description("Returns the index of the second item in the SeqHA or -1  if the SeqHA has less than 2 Items")
'Public Function FBOIndex() As Long
'
'    If s.OrderRoot Is Nothing Then
'        FBOIndex = -1
'    ElseIf s.OrderRoot.OrderSucc Is Nothing Then
'        FBOIndex = -1
'    Else
'        FBOIndex = p.FirstIndex + 1
'    End If
'
'End Function
'
'
''Description("Returns the index of the second to last item in the SeqHA, or -1 if the SeqHA has lees than 2 items")
'Public Function LBOIndex() As Long
'
'    If s.OrderRoot Is Nothing Then
'        LBOIndex = -1
'    ElseIf s.OrderLast.OrderPred Is Nothing Then
'        LBOIndex = -1
'    Else
'        LBOIndex = s.OrderLast.OrderPred.Order
'    End If
'
'End Function
'
'
'
'
''@Description("Returns the item at FirstIndex")
'Public Property Get First() As Variant
'
'    If s.OrderRoot Is Nothing Then
'        First = Null
'        Exit Property
'    End If
'
'    If VBA.IsObject(s.OrderRoot.Item) Then
'        Set First = s.OrderRoot.Item
'    Else
'        First = s.OrderRoot.Item
'    End If
'
'End Property
'
'
''@Description("Assigns to the Item at FirstIndex")
'Public Property Let First(ByVal ipItem As Variant)
'
'    ' ToDo: review if the add action is better as an error
'    If s.OrderRoot Is Nothing Then
'        Me.Add ipItem
'        Exit Property
'    End If
'
'    s.OrderRoot.Item = ipItem
'
'End Property
'
'
''@Description("Assigns to the Item at FirstIndex")
'Public Property Set First(ByVal ipItem As Variant)
'
'    If s.OrderRoot Is Nothing Then
'        Me.Add ipItem
'        Exit Property
'    End If
'
'    Set s.OrderRoot.Item = ipItem
'
'End Property
'
'
''@Description("Returns the Item at LastIndex")
'Public Property Get Last() As Variant
'
'    If s.OrderLast Is Nothing Then
'        Last = Null
'        Exit Property
'    End If
'
'    If VBA.IsObject(s.OrderLast.Item) Then
'        Set Last = s.OrderLast.Item
'    Else
'        Last = s.OrderLast.Item
'    End If
'
'End Property
'
'
''@Description("Assigns to the Item at LastIndex")
'Public Property Let Last(ByVal ipItem As Variant)
'
'    If s.OrderLast Is Nothing Then
'        Me.Add ipItem
'        Exit Property
'    End If
'
'    s.OrderLast.Item = ipItem
'
'End Property
'
'
''@Description("Assigns to the item at LastIndex")
'Public Property Set Last(ByVal ipItem As Variant)
'
'    If s.OrderLast Is Nothing Then
'        Me.Add ipItem
'        Exit Property
'    End If
'
'    Set s.OrderLast.Item = ipItem
'
'End Property
'
''@Description("Returns the item at First but one Index")
'Public Property Get FBO() As Variant
'
'    If s.OrderRoot Is Nothing Then
'        FBO = Null
'        Exit Property
'    ElseIf s.OrderRoot.OrderSucc Is Nothing Then
'        FBO = Null
'        Exit Property
'    End If
'
'    If VBA.IsObject(s.OrderRoot.OrderSucc.Item) Then
'        Set FBO = s.OrderRoot.OrderSucc.Item
'    Else
'        FBO = s.OrderRoot.OrderSucc.Item
'    End If
'
'End Property
'
'
''@Description("Assigns to the Item at First but one Index")
'Public Property Let FBO(ByVal ipItem As Variant)
'
'    If s.OrderRoot Is Nothing Then
'        Err.Raise 17 + vbObjectError, _
'            "SeqHA.Let_FBO", _
'            Fmt.Text("Index {0} not avaiable", FIRST_INDEX + 1)
'    ElseIf s.OrderLast.Order = p.FirstIndex Then
'        Me.Add ipItem
'    Else
'        s.OrderRoot.OrderSucc.Item = ipItem
'    End If
'
'End Property
'
'
''@Description("Assigns to the Item at First but one Index")
'Public Property Set FBO(ByVal ipItem As Variant)
'
'    If s.HashRoot Is Nothing Then
'        Err.Raise 17 + vbObjectError, _
'            "SeqHA.Let_FBO", _
'            Fmt.Text("Index {0} not avaiable", FirstIndex + 1)
'    ElseIf s.OrderLast.Order = p.FirstIndex Then
'        Me.Add ipItem
'    Else
'        Set s.OrderRoot.OrderSucc.Item = ipItem
'    End If
'
'End Property
'
'
''@Description("Returns the Item at Lastbut one Index")
'Public Property Get LBO() As Variant
'
'    If s.OrderLast Is Nothing Then
'        LBO = Null
'        Exit Property
'    ElseIf s.OrderLast.OrderPred Is Nothing Then
'        LBO = Null
'        Exit Property
'    End If
'
'    If VBA.IsObject(s.OrderLast.OrderPred.Item) Then
'        Set LBO = s.OrderLast.OrderPred.Item
'    Else
'        LBO = s.OrderLast.OrderPred.Item
'    End If
'
'End Property
'
'
''@Description("Assigns to the Item at Last but one Index")
'Public Property Let LBO(ByVal ipItem As Variant)
'
'    If s.HashRoot Is Nothing Then
'        Err.Raise 17 + vbObjectError, _
'            "SeqHA.Let_FBO", _
'            Fmt.Text("Index {0} not avaiable", s.OrderLast.Order - 1)
'    ElseIf s.OrderLast.OrderPred Is Nothing Then
'        Me.InsertAt ipItem, FIRST_INDEX
'    Else
'        s.OrderLast.OrderPred.Item = ipItem
'    End If
'
'End Property
'
'
''@Description("Assigns to the item at Last but one Index")
'Public Property Set LBO(ByVal ipItem As Variant)
'
'    If s.HashRoot Is Nothing Then
'        Err.Raise 17 + vbObjectError, _
'            "SeqHA.Let_FBO", _
'            Fmt.Text("Index {0} not avaiable", s.OrderLast.Order - 1)
'    ElseIf s.OrderLast.OrderPred Is Nothing Then
'        Me.InsertAt ipItem, FIRST_INDEX
'    Else
'        s.OrderLast.OrderPred.Item = ipItem
'    End If
'
'End Property
'
'
'Public Function IndexOf _
'( _
'    ByVal ipItem As Variant, _
'    Optional ByVal ipStart As Long = 1, _
'    Optional ByVal ipRun As Long = 0 _
') As Long
'
'    IndexOf = -1
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    If ipStart < FIRST_INDEX Or ipStart > s.OrderLast.Order Then
'        Exit Function
'    End If
'
'    ' exit  if the hash slot is not occupied
'    Dim myHash As Long: myHash = GetHash(ipItem)
'    If s.HashSlots(myHash) Is Nothing Then
'        Exit Function
'    End If
'
'    ' we now know that ipItem exists in the seq
'    ' but does it exist in the range defined by ipstart,iprun
'
'    ' ensure that ipstart+iprun only goes to the end of the seq
'    Dim myEnd As Long
'    If ipRun = 0 Then
'        myEnd = s.OrderLast.Order
'    Else
'        myEnd = ipStart + ipRun - 1
'        If myEnd > s.OrderLast.Order Then
'            myEnd = s.OrderLast.Order
'        End If
'    End If
'
'    ' check nodes between ipStart and ipEnd for th e first match to item
'    Dim myNode As SeqHNode: Set myNode = GetNodeAt(ipStart)
'
'    'The first found node may be the first in a sequence so we need to
'    ' search for the first node whose order is in the range ipStart to ipEnd
'    Do
'
'        If Comparers.EQ(myNode.Item, ipItem) Then
'            IndexOf = myNode.Order
'        End If
'
'        If myNode.OrderSucc Is Nothing Then
'            Exit Function
'        End If
'
'        ' If myNode.Hash <> myNode.OrderSucc.Hash Then
'        '     Exit Function
'        ' End If
'
'        Set myNode = myNode.OrderSucc
'
'    Loop Until myNode.Order > myEnd
'
'End Function
'
'Public Function LastIndexOf _
'( _
'ByVal ipItem As Variant, _
'Optional ByVal ipStart As Long = 1, _
'Optional ByVal ipRun As Long = 0 _
') As Long
'
'    LastIndexOf = -1
'
'    If s.OrderRoot Is Nothing Then
'        Exit Function
'    End If
'
'    If ipStart < FIRST_INDEX Or ipStart > s.OrderLast.Order Then
'        Exit Function
'    End If
'
'    ' ensure that ipstart+iprun only goes to the end of the seq
'    Dim myEnd As Long
'    If ipRun = 0 Then
'        myEnd = s.OrderLast.Order
'    Else
'        myEnd = ipStart + ipRun - 1
'        If myEnd > s.OrderLast.Order Then
'            myEnd = s.OrderLast.Order
'        End If
'    End If
'
'    'cannot use Iteritems unless we convert managing nodes into managing a list object
'    ' which is probably not worth the effort as it will duplicate Seq classes
'
'    Dim myNode As SeqHNode: Set myNode = GetNodeAt(myEnd)
'
'    Dim myCount As Long
'    For myCount = ipStart To myEnd
'
'        If Comparers.EQ(myNode.Item, ipItem) Then
'            LastIndexOf = myNode.Order
'            Exit Function
'        End If
'
'        Set myNode = myNode.OrderPred
'
'    Next
'
'    LastIndexOf = -1
'
'End Function
'
''@Description("Add a single item.  Returns the index at which the item was added for compatibility with the mscorlib ArrayList.")
'Public Function Add(ByRef ipItem As Variant) As Long
'
'    ' The Add actions depends on the root and the hash slot status
'    ' If s.HashRoot is nothing: add the first node to the list
'    ' elseIf the hash slot is nothing: append the new node to the node list
'    ' else the hash slot is occupied: insert the new node into the node list
'
'    ' for an add we always appendbyorder
'
'    Dim myHash As Long: myHash = GetHash(ipItem)
'    Dim myNewNode As SeqHNode: Set myNewNode = SeqHNode.Deb(myHash, ipItem)
'
'    If s.HashRoot Is Nothing Then
'
'        SetupFirstNode myNewNode
'
'    ElseIf s.HashSlots(myHash) Is Nothing Then
'
'        ' if the hash slot is nothing then we just append
'        Set s.HashSlots(myHash) = myNewNode
'        AppendByHash myNewNode
'        AppendByOrder myNewNode
'
'
'    Else
'
'        'Insert ipNewNode in the node list at the end of the hash/item sequence
'        ' append the new node to the order seq
'        InsertbyHash myNewNode
'        AppendByOrder myNewNode
'
'    End If
'
'    Add = myNewNode.Order
'
'End Function
'
'
''@Description("Add Items in the paramarray to the  )
'Public Function AddItems(ParamArray ipParamarray() As Variant) As oldSeqHA
'
'    If ArrayOp.LacksItems(VBA.CVar(ipParamarray)) Then
'        Set AddItems = Me
'        Exit Function
'    End If
'
'    Dim myParamarray As Variant: myParamarray = ipParamarray
'    myParamarray = ArrayOp.Splat(myParamarray)
'
'    AddRange myParamarray
'
'    Set AddItems = Me
'
'End Function
'
''@Description("Add an array,  list type, dictionary type, or characters of a string")
'Public Function AddRange(ByVal ipRange As Variant) As oldSeqHA
'
'    Set AddRange = Me
'
'     If GroupInfo.IsNotContainer(ipRange) And GroupInfo.IsNotString(ipRange) Then
'        Err.Raise 17 + vbObjectError, _
'            "VBALib.SeqHA.AddRange", _
'            Fmt.Text("Expecting a container type.  Got {0}", VBA.TypeName(ipRange))
'    End If
'
'    Dim myItems As IterItems: Set myItems = IterItems(ipRange)
'
'    If myItems.LacksItems Then
'        Exit Function
'    End If
'
'    Do
'        If GroupInfo.IsDictionary(ipRange) Then
'            Dim myK As Variant: myK = Array(myItems.CurKey(0))
'            Dim myI As Variant: myI = Array(myItems.CurItem(0))
'            Me.Add Array(myK(0), myI(0))
'        Else
'            Me.Add myItems.CurItem(0)
'        End If
'
'    Loop While myItems.MoveNext
'
'End Function
'
'
'Public Property Get Item(ByVal ipIndex As Long) As Variant
'
'    Item = Null
'
'    If ipIndex < FIRST_INDEX Or ipIndex > s.OrderLast.Order Then
'        Exit Property
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = GetNodeAt(ipIndex)
'
'   If myNode Is Nothing Then
'        Exit Property
'    ElseIf VBA.IsObject(myNode.Item) Then
'        Set Item = myNode.Item
'    Else
'        Item = myNode.Item
'    End If
'
'End Property
'
'
'Public Property Let Item(ByVal ipIndex As Long, ByVal ipItem As Variant)
'
'    GuardIndexOutOfBounds ipIndex, s.OrderRoot.Order, s.OrderLast.Order, "SeqHA.Let_Item"
'
'    Dim myNode As SeqHNode: Set myNode = GetNodeAt(ipIndex)
'    myNode.Item = ipItem
'
'End Property
'
'
'Public Property Set Item(ByVal ipIndex As Long, ByVal ipItem As Variant)
'
'    GuardIndexOutOfBounds ipIndex, FIRST_INDEX, s.OrderLast.Order, "SeqHA.Set_Item"
'
'    Dim myNode As SeqHNode: Set myNode = GetNodeAt(ipIndex)
'    Set myNode.Item = ipItem
'
'End Property
'
'
'Public Function InsertAt(ByVal ipIndex As Long, ByVal ipItem As Variant) As Long
'
'    If ipIndex < FIRST_INDEX Or ipIndex > s.OrderLast.Order Then
'        InsertAt = -1
'        Exit Function
'    End If
'
'    ' create the new node
'    Dim myHash As Long: myHash = GetHash(ipItem)
'    Dim myNewNode As SeqHNode: Set myNewNode = SeqHNode.Deb(myHash, ipItem, ipOrder:=ipIndex)
'
'    ' there are two strands to update, byHash and ByOrder
'    ' deal with byHash first
'    If s.HashSlots(myHash) Is Nothing Then
'        AppendByHash myNewNode
'    Else
'        InsertbyHash myNewNode
'    End If
'
'    ' byorder
'    If ipIndex = p.FirstIndex Then
'        PrependByOrder myNewNode
'    Else
'    'if ipIndex = s.OrderLast.Order Then
'    'End If
'    '     AppendByOrder myNewNode
'    ' Else
'        InsertByOrder myNewNode
'    End If
'
'    's.OrderLast.Order = s.OrderLast.Order + 1
'    InsertAt = myNewNode.Order
'
'End Function
'
'Public Function InsertAtItems(ByVal ipIndex As Long, ParamArray ipParamarray() As Variant) As oldSeqHA
'
'    Dim myParamarray As Variant: myParamarray = ipParamarray
'    myParamarray = ArrayOp.Splat(myParamarray)
'
'    Set InsertAtItems = InsertAtRange(ipIndex, myParamarray)
'
'End Function
'
'Public Function InsertAtRange(ByVal ipIndex As Long, ByRef ipRange As Variant) As oldSeqHA
'
'    If ipIndex < s.OrderRoot.Order Or ipIndex > s.OrderLast.Order Then
'        Set InsertAtRange = Me
'        Exit Function
'    End If
'
'    GuardInvalidRangeItem ipRange, "SeqHA.InsertAtRange"
'
'    ' iterate in reverse so we can just keep inserting at ipIndex
'    Dim myItems As IterItems: Set myItems = IterItems(ipRange).MoveToEndIndex
'
'    If myItems.LacksItems Then
'        Set InsertAtRange = Me
'        Exit Function
'    End If
'
'    Do
'
'        If GroupInfo.IsDictionary(ipRange) Then
'            InsertAt ipIndex, Array(myItems.CurKey(0), myItems.CurItem(0))
'        Else
'            InsertAt ipIndex, myItems.CurItem(0)
'        End If
'
'    Loop While myItems.MovePrev
'
'    Set InsertAtRange = Me
'
'End Function
'
'Public Function Remove(ByRef ipItem As Variant) As Long
'
'    If s.HashRoot Is Nothing Then
'        Remove = -1
'        Exit Function
'    End If
'
'    Dim myHash As Long: myHash = GetHash(ipItem)
'    Dim myNode As SeqHNode: Set myNode = s.HashSlots(myHash)
'
'    If myNode Is Nothing Then
'        Exit Function
'    End If
'
'    Remove = myNode.Order
'    RemoveNode myNode
'
'End Function
'
'
''@Description "Removes the first encountered item from the Seq for each item in the ParamArray"
'Public Function RemoveItems(ParamArray ipParamarray() As Variant) As oldSeqHA
'
'    Dim myParamarray As Variant: myParamarray = ipParamarray
'    myParamarray = ArrayOp.Splat(myParamarray)
'
'    Set RemoveItems = RemoveRange(myParamarray)
'
'    Set RemoveItems = Me
'
'End Function
'
'Public Function RemoveRange(ByRef ipRange As Variant) As oldSeqHA
'
'    Set RemoveRange = Me
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    Dim myItems As IterItems: Set myItems = IterItems(ipRange)
'
'    If myItems.LacksItems Then
'        Exit Function
'    End If
'
'    Do
'        RemoveNode GetNodeOf(myItems.CurItem(0))
'    Loop While myItems.MoveNext
'
'    Set RemoveRange = Me
'
'End Function
'
'Public Function RemoveAll(ParamArray ipParamarray() As Variant) As oldSeqHA
'
'    Set RemoveAll = Me
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    If ArrayOp.LacksItems(VBA.CVar(ipParamarray)) Then
'        RemoveAllItems
'        Exit Function
'    End If
'
'    Dim myParamarray As Variant
'    '@Ignore AssignmentNotUsed
'    myParamarray = ipParamarray
'    myParamarray = ArrayOp.Splat(ipParamarray)
'
'    Dim myParam As Variant
'    For Each myParam In myParamarray
'        RemoveAllSingleItem myParam
'    Next
'
'    Set RemoveAll = Me
'
'End Function
'
'
'Public Function Clear() As oldSeqHA
'    RemoveAll
'    Set Clear = Me
'End Function
'
'
'Public Function Reset() As oldSeqHA
'    Me.RemoveAll
'    Set Reset = Me
'End Function
'
'
''Public Function RemoveAfter(ByRef ipIndex As Long, Optional ByVal ipRun As Long = 1) As SeqHA
''
''    Set RemoveAfter = Me
''
''    If s.HashRoot Is Nothing Then
''        Exit Function
''    End If
''
''    If Me.LacksKey(ipKey) Then
''        Exit Function
''    End If
''
''    If ipRun < 1 Then
''        Err.Raise 17 + vbObjectError, _
''        "VALib.SeqHA.RemoveAfter", _
''        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
''    End If
''
''    Dim myNode As SeqHNode: Set myNode = s.HashRoot
''    Do Until ComparerHelpers.Equals(ipKey, myNode.Key)
''        Set myNode = myNode.Succ
''        If myNode Is Nothing Then
''            Exit Function
''        End If
''
''    Loop
''
''    If myNode.Succ Is Nothing Then
''        Exit Function
''    End If
''
''    '@Ignore AssignmentNotUsed
''    Set myNode = myNode.Succ
''
''    Dim myCount As Long
''    For myCount = 1 To ipRun
''
''        RemoveNode myNode.Succ
''
''        If myNode.Succ Is Nothing Then
''            Exit Function
''        End If
''
''    Next
''
''End Function
''
''
''Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As SeqHA
''
''    Set RemoveBefore = Me
''
''    If VBA.IsNull(ipKey) Then
''        Err.Raise 17 + vbObjectError, _
''        "VBALib.SeqHA.RemoveBefore", _
''        "Null is not a valid Key"
''    End If
''
''    If Me.LacksKey(ipKey) Then
''        Exit Function
''    End If
''
''    If ipRun < 1 Then
''        Err.Raise 17 + vbObjectError, _
''        "VALib.SeqHA.RemoveBefore", _
''        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
''    End If
''
''    Dim myNode As SeqHNode: Set myNode = s.HashRoot
''    Do Until ComparerHelpers.Equals(ipKey, myNode.Key)
''        Set myNode = myNode.Succ
''        If myNode Is Nothing Then
''            Exit Function
''        End If
''
''    Loop
''
''    If myNode.Pred Is Nothing Then
''        Exit Function
''    End If
''
''    '@Ignore AssignmentNotUsed
''    Set myNode = myNode.Pred
''
''    Dim myCount As Long
''    For myCount = 1 To ipRun
''
''        RemoveNode myNode.Pred
''
''        If myNode.Pred Is Nothing Then
''            Exit Function
''        End If
''
''    Next
''
''End Function
'
'
''@Description(Removes the node at the specified Index
'Public Function RemoveAt(ByRef ipIndex As Long) As Long
'
'    RemoveAt = -1
'
'    If s.OrderRoot Is Nothing Then
'        Exit Function
'    End If
'
'    If ipIndex < s.OrderRoot.Order Or ipIndex > s.OrderLast.Order Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = GetNodeAt(ipIndex)
'
'    RemoveNode myNode
'    RemoveAt = ipIndex
'
'End Function
'
''ToDo: rename RemoveIndexes to removeIndexes
''@Description("RemoveAt for a list of indexes.  Indexes are sorted and removed in reverse order")
'Public Function RemoveIndexes(ParamArray ipParamarray() As Variant) As oldSeqHA
'
'    Dim myParamarray As Variant: myParamarray = ipParamarray
'    myParamarray = ArrayOp.Splat(myParamarray)
'
'    Set RemoveIndexes = RemoveIndexesRange(myParamarray)
'
'    Set RemoveIndexes = Me
'
'End Function
'
''@Description("Removes the Items at the indexes in the range list")
'Public Function RemoveIndexesRange(ByRef ipRange As Variant) As oldSeqHA
'
'    Set RemoveIndexesRange = Me
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    Dim myS As SeqA: Set myS = SeqA(ipRange).Sorted
'
'    If myS.LacksItems Then
'        Exit Function
'    End If
'
'    ' we iterate backwards so we don't so we don't compromise indexes yet to be removed
'    'Dim myIndex As Long
'    Dim myItems As IterItems: Set myItems = IterItems(myS).MoveToEndIndex
'    Do
'        Dim myNode As SeqHNode: Set myNode = GetNodeAt(myItems.CurItem(0))
'        RemoveNode myNode
'    Loop While myItems.MovePrev
'
'    Set RemoveIndexesRange = Me
'
'End Function
'
'
'Public Function Fill(ByRef ipItem As Variant, Optional ByVal ipCount As Long = 1) As oldSeqHA
'
'    Set Fill = Me
'
'    If ipCount < 1 Then
'        Exit Function
'    End If
'
'    Dim myCount As Long
'    For myCount = 1 To ipCount
'        Me.Add ipItem
'    Next
'
'End Function
'
'
'
''@Description("Returns the first N items of the SeqHA, Default 1")
'Public Function Head(Optional ByVal ipSplitAfter As Long = 1) As oldSeqHA
'
'    Set Head = oldSeqHA.Deb
'
'    If s.OrderLast.Order <= p.FirstIndex Then
'        Exit Function
'    End If
'
'    If ipSplitAfter < FIRST_INDEX Then
'        Exit Function
'    End If
'
'    If ipSplitAfter >= s.OrderLast.Order Then
'        Set Head = Me.Clone
'        Exit Function
'    End If
'
'    Set Head = Me.Slice(FIRST_INDEX, ipSplitAfter)
'
'End Function
'
'
''@Description("Returns the Items after SplitAfter, Default is Index 2")
'Public Function Tail(Optional ByVal ipSplitAfter As Long = p.FirstIndex) As oldSeqHA
'
'    Set Tail = oldSeqHA.Deb
'
'    If s.OrderLast.Order < FIRST_INDEX Then
'        Exit Function
'    End If
'
'    If ipSplitAfter >= s.OrderLast.Order Then
'        Exit Function
'    End If
'
'    If ipSplitAfter < FIRST_INDEX Then
'        Set Tail = Me.Clone
'        Exit Function
'    End If
'
'    Set Tail = Me.Slice(ipSplitAfter + 1)
'
'End Function
'
'Public Function HoldsItem(ByRef ipItem As Variant) As Boolean
'
'    HoldsItem = False
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'    Do
'
'        If Comparers.EQ(myNode.Item, ipItem) Then
'            HoldsItem = True
'            Exit Function
'        End If
'
'        Set myNode = myNode.HashSucc
'
'    Loop Until myNode Is Nothing
'
'End Function
'
'
'Public Function Contains(ByRef ipItem As Variant) As Boolean
'    Contains = Me.HoldsItem(ipItem)
'End Function
'
'Public Function LacksItem(ByRef ipItem As Variant) As Boolean
'    LacksItem = Not HoldsItem(ipItem)
'End Function
'
'
'Public Function HoldsItems() As Boolean
'    HoldsItems = Not s.HashRoot Is Nothing
'End Function
'
'
'Public Function LacksItems() As Boolean
'    LacksItems = s.HashRoot Is Nothing
'End Function
'
'
'
''@Description("Mutates the KvpA in place by applying the function in IMapper to each item")
'Public Function MappedIt(ByVal ipMapper As IMapper) As oldSeqHA
'
'    Set MappedIt = Me
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set SeqHNode = s.OrderRoot
'    Do
'
'        Dim myItem As Variant: myItem = ipMapper.ExecMapper(myNode.Item)
'
'        If VBA.IsObject(myItem(0)) Then
'            Set myNode.Item = myItem(0)
'        Else
'            myNode.Item = myItem(0)
'        End If
'
'        Set myNode = myNode.OrderSucc
'
'    Loop Until myNode Is Nothing
'
'    Set MappedIt = Me
'
'End Function
'
'
''@Description("Returns a KvpA.Deb by applying the function in IMapper to each Item")
'Public Function MapIt(ByVal ipMapper As IMapper) As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'    Set MapIt = myS
'
'    If s.HashRoot Is Nothing Then
'        Set MapIt = Me
'        Exit Function
'    End If
'
'
'    Dim myNode As SeqHNode: Set SeqHNode = s.OrderRoot
'    Do
'        myS.Add ipMapper.ExecMapper(myNode.Item)(0)
'        Set myNode = myNode.OrderSucc
'    Loop Until myNode Is Nothing
'
'    Set MapIt = myS
'
'End Function
'
'
''@Description("Returns a subset of the seq composed of items where IComparer was true")
'Public Function FilterIt(ByVal ipComparer As IComparer) As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'    Set FilterIt = myS
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set SeqHNode = s.OrderRoot
'    Do
'
'        If ipComparer.ExecCmp(myNode.Item) Then
'            myS.Add myNode.Item
'        End If
'
'        Set myNode = myNode.OrderSucc
'
'    Loop Until myNode Is Nothing
'
'    Set FilterIt = myS
'
'End Function
'
''@Description("Returns the count of where ICOmparer is true. Does not generate an interim filtered seq"
'Public Function CountIt(ByVal ipComparer As IComparer) As Long
'
'    If s.HashRoot Is Nothing Then
'        CountIt = -1
'        Exit Function
'    End If
'
'    Dim myCount As Long: myCount = 0
'
'    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
'    Do
'
'        If ipComparer.ExecCmp(myNode.Item) Then
'            myCount = myCount + 1
'        End If
'
'        Set myNode = myNode.OrderSucc
'
'    Loop Until myNode Is Nothing
'
'    CountIt = myCount
'
'End Function
'
''@Description("Returns a single Host based on the Reducer").
'Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant
'
'    If s.HashRoot Is Nothing Then
'        ReduceIt = Null
'        Exit Function
'    End If
'
'    Dim myItem As Variant: myItem = Array(ipReducer.ExecReduction(Me.ToArray))
'
'    If VBA.IsObject(myItem(0)) Then
'        Set ReduceIt = myItem(0)
'    Else
'        ReduceIt = myItem(0)
'    End If
'
'End Function
'
'Public Property Get TypeName() As String
'    TypeName = MY_TYPENAME
'End Property
'
'Public Function Push(ParamArray ipParamarray() As Variant) As oldSeqHA
'
'    Dim myParamarray As Variant: myParamarray = ipParamarray
'    myParamarray = ArrayOp.Splat(myParamarray)
'
'    Set Push = AddRange(myParamarray)
'
'End Function
'
'
'Public Function PushRange(ByRef ipRange As Variant) As oldSeqHA
'    Set PushRange = AddRange(ipRange)
'End Function
'
'
''@Description("Returns the last item of the seq.  Removes the last item of the seq")
'Public Function Pop() As Variant
'
'    If s.HashRoot Is Nothing Then
'        Pop = Null
'        Exit Function
'    End If
'
'    If VBA.IsObject(s.OrderLast.Item) Then
'        Set Pop = s.OrderLast.Item
'    Else
'        Pop = s.OrderLast.Item
'    End If
'
'    RemoveNode s.OrderLast
'
'End Function
'
''ToDo change output to array for all seq to avoid large memort committment for small sequences
'Public Function PopRange(Optional ByVal ipCount As Long = 1) As SeqA
'
'    Set PopRange = SeqA.Deb
'
'    If ipCount < 1 Then
'        Exit Function
'    End If
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    If ipCount > s.OrderLast.Order Then
'        Set PopRange = SeqA(ToArray).Reverse
'        Me.RemoveAll
'        Exit Function
'    End If
'
'    Dim myS As SeqA: Set myS = SeqA.Deb
'
'    Dim myCount As Long
'    For myCount = 1 To ipCount
'        myS.Add Me.Pop
'    Next
'
'    Set PopRange = myS
'
'End Function
'
'Public Function enQueue(ParamArray ipParamarray() As Variant) As oldSeqHA
'
'    Dim myParamarray As Variant: myParamarray = ipParamarray
'    myParamarray = ArrayOp.Splat(myParamarray)
'
'    Set enQueue = AddRange(myParamarray)
'
'End Function
'
'Public Function EnqueueRange(ByRef ipRange As Variant) As oldSeqHA
'    Set EnqueueRange = AddRange(ipRange)
'End Function
'
'Public Function Dequeue() As Variant
'
'    If s.HashRoot Is Nothing Then
'        Dequeue = Null
'        Exit Function
'    End If
'
'    If VBA.IsObject(s.OrderRoot.Item) Then
'        Set Dequeue = s.OrderRoot.Item
'    Else
'        Dequeue = s.OrderRoot.Item
'    End If
'
'    RemoveNode s.OrderRoot
'
'End Function
''@Description("SeqA is deliberate")
'Public Function DequeueRange(Optional ByVal ipCount As Long = 1) As SeqA
'
'    Set DequeueRange = SeqA.Deb
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    If ipCount < 1 Then
'        Exit Function
'    End If
'
'    If ipCount > s.OrderLast.Order Then
'        Set DequeueRange = SeqA(ToArray)
'        Me.RemoveAll
'        Exit Function
'    End If
'
'    Dim myS As SeqA: Set myS = SeqA.Deb
'
'    Dim myIndex As Long
'    For myIndex = 1 To ipCount
'        myS.Add Me.Dequeue
'    Next
'
'    Set DequeueRange = myS
'
'End Function
'
'Public Function Sort() As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = Me.Clone
'
'    If s.OrderLast.Order < FIRST_INDEX Then
'        Set Sort = myS
'        Exit Function
'    End If
'
'    Sorters.ShakerSortByItem myS
'    Set Sort = myS
'
'End Function
'
'
'Public Function Sorted() As oldSeqHA
'
'    If s.OrderLast.Order < FIRST_INDEX Then
'        Set Sorted = Me
'        Exit Function
'    End If
'
'    Sorters.ShakerSortByItem Me
'    Set Sorted = Me
'
'End Function
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
''@Description("Returns a new seq in reverse order of host seq")
'Public Function Reverse() As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'    Set Reverse = myS
'
'    If s.OrderRoot Is Nothing Then
'        Exit Function
'    End If
'
'    Dim myIndex As Long
'    For myIndex = s.OrderLast.Order To FIRST_INDEX Step -1
'        myS.Add GetNodeAt(myIndex).Item
'    Next
'
'    Set Reverse = myS
'
'End Function
'
''@Description("Reverses the seq in place")
'Public Function Reversed() As oldSeqHA
'
'    If s.OrderLast.Order < FIRST_INDEX + 1 Then
'        Set Reversed = Me
'        Exit Function
'    End If
'
'    Dim myFirstIndex As Long: myFirstIndex = p.FirstIndex
'    Dim myLastIndex As Long: myLastIndex = s.OrderLast.Order
'
'    Do
'
'        Me.Swap myFirstIndex, myLastIndex
'        myFirstIndex = myFirstIndex + 1
'        myLastIndex = myLastIndex - 1
'
'    Loop Until myFirstIndex >= myLastIndex
'
'    Set Reversed = Me
'
'End Function
'
'Public Function IsUnique() As Boolean
'
'    If LacksItems Then
'        IsUnique = False
'        Exit Function
'    End If
'
'    Dim myS As cHashC: Set myS = New cHashC
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'
'    Do
'
'        If myS.Exists(myNode.Item) Then
'            IsUnique = False
'            Exit Function
'        Else
'            myS.Add myNode.Item
'        End If
'
'        Set myNode = myNode.HashSucc
'
'    Loop Until myNode Is Nothing
'
'    IsUnique = True
'
'End Function
'
'
'Public Function IsNotUnique() As Boolean
'    IsNotUnique = Not IsUnique
'End Function
'
''@Description("Returns the a new seq with the unique subset of the host Seq")
'Public Function Dedup() As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'
'    If s.HashRoot Is Nothing Then
'        Set Dedup = myS
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
'
'    Select Case s.OrderLast.Order
'
'        Case Is < FIRST_INDEX:              Set Dedup = myS
'
'        Case FIRST_INDEX:                   Set Dedup = Me.Clone
'
'        Case Else
'
'            Do
'
'                If myS.LacksItem(myNode.Item) Then
'                    myS.Add myNode.Item
'                End If
'
'                Set myNode = myNode.OrderSucc
'
'            Loop Until myNode Is Nothing
'
'            Set Dedup = myS
'
'    End Select
'
'End Function
'
''@Description("Removes duplicate items from the seq")
'Public Function Deduped() As oldSeqHA
'
'    If s.HashRoot Is Nothing Then
'        Set Deduped = Me
'        Exit Function
'    End If
'
'
'    Dim myH As cHashC: Set myH = New cHashC
'
'    Dim myNode As SeqHNode: Set myNode = s.OrderLast
'    Do
'
'        If myH.Exists(myNode.Item) Then
'
'            Dim myTmpNode As SeqHNode: Set myTmpNode = myNode
'            Set myNode = myNode.OrderPred
'            RemoveNode myTmpNode
'
'
'        Else
'
'            myH.Add myNode.Item
'            Set myNode = myNode.Order.Pred
'        End If
'
'    Loop Until myNode Is Nothing
'
'    Set Deduped = Me
'
'End Function
'
'
'Public Function Freq() As KvpA
'
'    Dim myK As KvpA: Set myK = KvpA.Deb
'    Set Freq = myK
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'    Do
'
'        If myK.HoldsItem(myNode.Item) Then
'            myK.Item(myNode.Item) = myK.Item(myNode.Item) + 1
'        Else
'            myK.Add myNode.Item, 1
'        End If
'
'        Set myNode = myNode.HashSucc
'
'    Loop Until myNode Is Nothing
'
'    Set Freq = myK
'
'End Function
'
'Public Function SetOf(ByRef ipSet As e_SetoF, ByRef ipRangeItem As oldSeqHA) As oldSeqHA
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
'            "VBALib.SeqHA.SetOf", _
'            Fmt.Text("Invalid e_set member. Expecting {0} to {1}. Got {2}", e_SetoF.m_First, e_SetoF.m_Last, ipSet)
'
'    End Select
'
'End Function
'
'
'
'
'
'
'
'
''Public Property Get ItemAt(ByVal ipIndex As Long) As Variant
''
''    If ipIndex < FIRST_INDEX Or ipIndex > s.OrderLast.order Then
''        ItemAt = Null
''        Exit Property
''    End If
''
''    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
''
''    Do Until myNode.Order = ipIndex
''        Set myNode = myNode.OrderSucc
''    Loop
''
''    If VBA.IsObject(myNode.Item) Then
''        Set ItemAt = myNode.Item
''    Else
''        ItemAt = myNode.Item
''    End If
''
''End Property
''
''
''Public Property Let ItemAt(ByVal ipIndex As Long, ByVal ipItem As Variant)
''
''    If ipIndex < FIRST_INDEX Or ipIndex > s.OrderLast.order Then
''        Err.Raise 17 + vbObjectError, _
''        "VBALib.SeqHA.Let_ItemByIndex", _
''        Fmt.Text("Bounds are {0} to {1}. Got index of {2}", FIRST_INDEX, s.OrderLast.order, ipIndex)
''    End If
''
''    If VBA.IsNull(ipItem) Then
''        Err.Raise 17 + vbObjectError, _
''        "Let_ItemByIndex", _
''        "Cannot store a value of null"
''    End If
''
''    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
''
''    Do Until myNode.Order = ipIndex
''        Set myNode = myNode.OrderSucc
''    Loop
''
''    myNode.Item = ipItem
''
''End Property
''
''
''Public Property Set ItemAt(ByVal ipIndex As Long, ByVal ipItem As Variant)
''
''    If ipIndex < FIRST_INDEX Or ipIndex > s.OrderLast.order Then
''        Err.Raise 17 + vbObjectError, _
''        "VBALib.SeqHA.Let_ItemByIndex", _
''        Fmt.Text("Bounds are {0} to {1}. Got index of {2}", FIRST_INDEX, s.OrderLast.order, ipIndex)
''    End If
''
''    If VBA.IsNull(ipItem) Then
''        Err.Raise 17 + vbObjectError, _
''        "Let_ItemByIndex", _
''        "Cannot store a value of null"
''    End If
''
''    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
''
''    Dim myIndex As Long: myIndex = 1
''    Do Until myIndex = ipIndex
''        Set myNode = myNode.OrderSucc
''        myIndex = myIndex + 1
''    Loop
''
''    Set myNode.Item = ipItem
''
''End Property
'
'
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'' Private methods
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'Private Function SetOfCommon(ByRef ipSeq As oldSeqHA) As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'    Set SetOfCommon = myS
'
'    If Me.LacksItems Or ipSeq.LacksItems Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'
'    Do
'        If ipSeq.HoldsItem(myNode.Item) Then
'            myS.Add myNode.Item
'        End If
'
'        Set myNode = myNode.HashSucc
'
'    Loop Until myNode Is Nothing
'
'    Set SetOfCommon = myS
'
'End Function
'
'
'Private Function SetOfHostOnly(ByRef ipSeq As oldSeqHA) As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'    Set SetOfHostOnly = myS
'
'    If s.HashRoot Is Nothing Then
'        Exit Function
'    End If
'
'    Set myS = Me.Dedup
'
'    If ipSeq.LacksItems Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'    Do
'
'        If ipSeq.HoldsItem(myNode.Item) Then
'            myS.RemoveAll myNode.Item
'        End If
'
'        Set myNode = myNode.HashSucc
'
'    Loop Until myNode Is Nothing
'
'    Set SetOfHostOnly = myS
'
'End Function
'
'
'Private Function SetOfParamOnly(ByRef ipSeq As oldSeqHA) As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'    Set SetOfParamOnly = myS
'
'    If ipSeq.LacksItems Then
'        Exit Function
'    End If
'
'    Set myS = ipSeq.Dedup
'
'    If Me.LacksItems Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'    Do
'
'        If myS.HoldsItem(myNode.Item) Then
'            myS.RemoveAll myNode.Item
'        End If
'
'        Set myNode = myNode.HashSucc
'
'    Loop Until myNode Is Nothing
'
'    Set SetOfParamOnly = myS
'
'End Function
'
'
'Private Function SetOfNotCommon(ByRef ipSeq As oldSeqHA) As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'    Set SetOfNotCommon = myS
'
'    If LacksItems And ipSeq.LacksItems Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
'    Do
'
'        If ipSeq.LacksItem(myNode.Item) Then
'            If myS.LacksItem(myNode.Item) Then
'                myS.Add myNode.Item
'            End If
'        End If
'
'        Set myNode = myNode.OrderSucc
'
'    Loop Until myNode Is Nothing
'
'    Dim myItems As IterItems: Set myItems = IterItems(ipSeq)
'    Do
'
'        If Me.LacksItem(myItems.CurItem(0)) Then
'            If myS.LacksItem(myItems.CurItem(0)) Then
'                myS.Add myItems.CurItem(0)
'            End If
'        End If
'
'    Loop While myItems.MoveNext
'
'    Set SetOfNotCommon = myS
'
'End Function
'
'
'Private Function SetofUnique(ByRef ipSeq As oldSeqHA) As oldSeqHA
'
'    Dim myS As oldSeqHA: Set myS = oldSeqHA.Deb
'    Set SetofUnique = myS
'
'    If Me.LacksItems And ipSeq.LacksItems Then
'        Exit Function
'    End If
'
'    Set myS = Me.Dedup
'
'    Dim myItems As IterItems: Set myItems = IterItems(ipSeq)
'    Do
'
'        If myS.LacksItem(myItems.CurItem(0)) Then
'            myS.Add myItems.CurItem(0)
'        End If
'
'    Loop While myItems.MoveNext
'
'    Set SetofUnique = myS
'
'End Function
'
'
'
'Private Sub SetupFirstNode(ByRef ipNode As SeqHNode)
'
'        Set s.HashRoot = ipNode
'        Set s.HashLast = s.HashRoot
'
'        ' nodes for managining sequence by order field
'        Set s.OrderRoot = s.HashRoot
'        Set s.OrderLast = s.HashRoot
'
'        s.OrderRoot.UpdateOrder
'
'        ' populate hash slot
'        Set s.HashSlots(s.HashRoot.Hash) = s.HashRoot
'
'End Sub
'
'Private Sub AppendByHash(ByRef ipNode As SeqHNode)
'
'    ' The append node method adds a node s.HashLast
'    ' link nodes following the hash
'    ' If s.HashSlots(ipNode.Hash) Is Nothing Then
'    '     Set s.HashSlots(ipNode.Hash) = ipNode
'    ' End If
'
'
'    Set ipNode.HashPred = s.HashLast
'    Set ipNode.HashSucc = Nothing
'    Set s.HashLast.HashSucc = ipNode
'    Set s.HashLast = ipNode
'
'End Sub
'
'Private Sub AppendByOrder(ByRef ipNode As SeqHNode)
'
'    ' no need to look at hashslots as this is handled by  'byHash' methods
'    Set ipNode.OrderPred = s.OrderLast
'    Set ipNode.OrderSucc = Nothing
'    Set s.OrderLast.OrderSucc = ipNode
'    Set s.OrderLast = ipNode
'    s.OrderLast.OrderPred.UpdateOrder
'
'End Sub
'Private Sub InsertbyHash(ByVal ipNewNode As SeqHNode)
'
'    ' we know that the new node is not the first node.
'    ' we know that the new node exists in hashslots
'    ' we insert the newnode after the last node that matches Hash/Item
'
'    ' get the first node for the newnode hash
'    Dim myCurNode As SeqHNode: Set myCurNode = s.HashSlots(ipNewNode.Hash)
'
'    'The new node is inserted at the end of the current hash/item combination
'    ' we know that curnode and new node have the same hash
'    ' we don't know if they have the same items
'
'    Do While myCurNode.Hash = ipNewNode.Hash
'        ' check item is the same item as ipnewnode
'        If Comparers.EQ(myCurNode.Item, ipNewNode.Item) Then
'            ' check the next node is not nothing
'            If myCurNode.HashSucc Is Nothing Then
'                ' if ipnode.succ is nothing then by default
'                ' myNode is s.last so we can just append the new node
'                AppendByHash ipNewNode
'                Exit Sub
'            Else
'                Set myCurNode = myCurNode.HashSucc
'            End If
'        End If
'    Loop
'
'    ' at this point we have matching hashes but different items
'    ' curnode is the node after the last item/hash match so we
'    ' insert newnode before curnode
'
'    Set ipNewNode.HashPred = myCurNode.HashPred 'myPredNode
'    Set ipNewNode.HashSucc = myCurNode 'myPred.Succ
'
'    Set myCurNode.HashPred = ipNewNode
'    Set ipNewNode.HashPred.HashSucc = ipNewNode
'
'End Sub
'
'
'
'Private Sub InsertByOrder(ByVal ipNewNode As SeqHNode)
'
'    ' we know that the new node is not the first or last node.
'    ' we know that the new node exists in hashslots
'    ' we insert the newnode before the node with the same order
'
'    Dim myCurNode As SeqHNode: Set myCurNode = GetNodeAt(ipNewNode.Order)
'
'    Set ipNewNode.OrderPred = myCurNode.OrderPred
'    Set ipNewNode.OrderSucc = myCurNode
'
'    Set myCurNode.OrderPred.OrderSucc = ipNewNode
'    Set myCurNode.OrderPred = ipNewNode
'
'    ipNewNode.UpdateOrder
'
'    ' Dim myCurNode As SeqHNode: Set myCurNode = GetNodeAt(ipNewNode.Order)
'
'    ' Set ipNewNode.OrderPred = myCurNode.OrderPred
'    ' Set myCurNode.OrderPred.OrderSucc = ipNewNode
'
'    ' Set ipNewNode.OrderSucc = myCurNode
'    ' Set myCurNode.OrderPred = ipNewNode
'
'    ' ipNewNode.UpdateOrder
'
'End Sub
'Private Sub PrependByOrder(ByRef ipNode As SeqHNode)
'
'    Set ipNode.OrderPred = Nothing
'    Set ipNode.OrderSucc = s.OrderRoot
'    Set s.OrderRoot.OrderPred = ipNode
'    Set s.OrderRoot = ipNode
'    s.OrderRoot.UpdateOrder
'
'End Sub
'
''@Description("Returns the node where .Order = ipIndex")
'Private Function GetNodeAt(ByRef ipIndex As Long) As Variant
'
'    Set GetNodeAt = Nothing
'
'    If s.OrderRoot Is Nothing Then
'        Exit Function
'    End If
'
'    If ipIndex < s.OrderRoot.Order Or ipIndex > s.OrderLast.Order Then
'        Exit Function
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.OrderRoot
'    Do
'        If myNode.Order = ipIndex Then
'            Set GetNodeAt = myNode
'            Exit Function
'        End If
'
'        Set myNode = myNode.OrderSucc
'
'    Loop Until myNode Is Nothing
'
'End Function
'
''@Description("Returns the first node containing ipItem")
'Private Function GetNodeOf(ByRef ipItem As Variant) As Variant
'
'    Set GetNodeOf = Nothing
'
'    If s.OrderRoot Is Nothing Then
'        Exit Function
'    End If
'
'
'    Dim myNode As SeqHNode: Set myNode = s.HashSlots(GetHash(ipItem))
'    If myNode Is Nothing Then
'        Exit Function
'    End If
'
'    Do
'        If Comparers.EQ(myNode.Item, ipItem) Then
'            Set GetNodeOf = myNode
'            Exit Function
'        End If
'
'        If myNode.Hash <> myNode.OrderSucc.Hash Then
'            Exit Function
'        End If
'
'        Set myNode = myNode.OrderSucc
'
'    Loop Until myNode Is Nothing
'
'End Function
'
'
'Private Function GetHash(ByRef ipItem As Variant) As Long
'
'    Dim D As Double
'    Dim L As LongPtr
'    Dim F As Single
'    Dim VT As VbVarType
'    Dim myHash As Long
'
'    Dim C As Variant
'    C = VBA.CDec(C)
'
'    myHash = s.HashSlotMask                              'init the HashValue (all bits to 1)
'    Dim myItem As Variant
'    If GroupInfo.IsContainer(ipItem) Or GroupInfo.IsItemObject(ipItem) Or GroupInfo.IsAdmin(ipItem) Then
'        myItem = GetItemAsString(ipItem)
'    Else
'        myItem = ipItem
'    End If
'
'    VT = VarType(myItem)
'    Select Case VT
'        Case vbString
'
'            Dim myIndex As Long
'            For myIndex = 1 To VBA.Len(myItem)
'
'                Dim myChar As Long
'                myChar = VBA.AscW(VBA.Mid$(myItem, myIndex, 1))
'                myHash = (myHash + myChar) * HMul And s.HashSlotMask
'
'            Next
'
'
''        Case vbObject
''
''            L = ObjPtr(myItem)
''            s.saW.cElements1D = 2
''            s.saW.pvData = VarPtr(L)
''
''            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 2 16Bit integers in the array)
''            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
''
'        Case vbCurrency
'
'            C = myItem
'            s.sAW.cElements1D = 4
'            s.sAW.pvData = VarPtr(C)
'
'            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 4 16Bit integers in the array)
'            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
'            myHash = (myHash + s.W(2)) * HMul And s.HashSlotMask
'            myHash = (myHash + s.W(3)) * HMul And s.HashSlotMask
'
'        Case vbLong, vbInteger, vbByte
'
'            L = myItem
'            s.sAW.cElements1D = 2
'            s.sAW.pvData = VarPtr(L)
'
'            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 2 16Bit integers in the array)
'            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
'
'        Case vbDouble
'
'            D = myItem: s.sAW.cElements1D = 4
'            s.sAW.pvData = VarPtr(D)
'
'            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 4 16Bit integers in the array)
'            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
'            myHash = (myHash + s.W(2)) * HMul And s.HashSlotMask
'            myHash = (myHash + s.W(3)) * HMul And s.HashSlotMask
'
'        Case vbDate
'
'            D = ipItem
'            s.sAW.cElements1D = 4
'            s.sAW.pvData = VarPtr(D)
'
'            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 4 16Bit integers in the array)
'            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
'            myHash = (myHash + s.W(2)) * HMul And s.HashSlotMask
'            myHash = (myHash + s.W(3)) * HMul And s.HashSlotMask
'
'        Case vbSingle
'
'            F = myItem
'            s.sAW.cElements1D = 2
'            s.sAW.pvData = VarPtr(F)
'            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 2 16Bit integers in the array)
'            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
'
'    End Select
'    ' the hashslot array starts at 1 so we must add 1 below for the case where the hash is zero
'    GetHash = myHash + 1
'
'End Function
'
'Private Function GetItemAsString(ByVal ipItem As Variant) As String
'
'    ' containers and arrays must be sorted if we are to have comparable hashes
'    Dim myResult As Variant
'    Select Case True
'        Case GroupInfo.IsArray(ipItem):
'            myResult = ArrayOp.Linearise(ipItem)
'            Sorters.ShakerSortArray myResult
'
'        Case GroupInfo.IsList(ipItem):
'            Sorters.ShakerSortByItem ipItem
'            Set myResult = ipItem
'
'        Case GroupInfo.IsItemByToArray(ipItem)
'            myResult = ipItem.ToArray
'            Sorters.ShakerSortArray myResult
'
'        Case GroupInfo.IsDictionary(ipItem)
'            myResult = ipItem.SPairs
'
'        Case GroupInfo.IsItemObject(ipItem)
'            myResult = Stringifier.SetNoMarkup.StringifyItemObject(ipItem)
'    End Select
'
'    GetItemAsString = Fmt.NoMarkup.Text("{0}", myResult)
'
'End Function
'
''@Ignore ProcedureCanBeWrittenAsFunction
'Private Sub RemoveNode(ByRef ipNode As SeqHNode)
'
'    ' when removing a node we must check that the node is not the node registered in hashslots
'    ' if it is we must replace the node pointed to by s.HashSlots(hash) with
'    ' s.HashSlots.succ or nothing
'    ' wemust also remember to update roo/last nodes as appropriate
'
'    ' check for the case of a single node
'    If s.OrderLast.Order = p.FirstIndex Then
'        Set s.HashSlots(ipNode.Hash) = Nothing
'        Set ipNode = Nothing
'        Set s.HashRoot = Nothing
'        Set s.HashLast = Nothing
'        Set s.OrderRoot = Nothing
'        Set s.OrderLast = Nothing
'        Exit Sub
'    End If
'
'    ' we know that item will have a hash for an occupied slot because the absense of its
'    ' item from the seq was previously tested
'    ' remove node from hashslot if succ is nothing or if the hash of the next node is different
'    If s.HashSlots(ipNode.Hash) Is ipNode Then
'        If ipNode.HashSucc Is Nothing Then
'            Set s.HashSlots(ipNode.Hash) = Nothing
'        ElseIf ipNode.Hash = ipNode.HashSucc.Hash Then
'            Set s.HashSlots(ipNode.Hash) = ipNode.HashSucc
'        Else
'            Set s.HashSlots(ipNode.Hash) = Nothing
'        End If
'    End If
'
'
'    '@Ignore VariableNotUsed
'    Dim myTmp As SeqHNode
'    ' Disconnect the node from hash seq
'    ' node deletion only occurs after order seq is updates
'    If ipNode.HashPred Is Nothing Then
'
'
'        Set s.HashRoot = ipNode.HashSucc
'        Set s.HashRoot.HashPred = Nothing
'
'    ElseIf ipNode.HashSucc Is Nothing Then
'
'        Set s.HashLast = ipNode.HashPred
'        Set s.HashLast.HashSucc = Nothing
'
'    Else
'
'        Set ipNode.HashSucc.HashPred = ipNode.HashPred
'        Set ipNode.HashPred.HashSucc = ipNode.HashSucc
'
'    End If
'
'    ' disconnect the node from order seq
'    ' is this the first order node
'    If ipNode.OrderPred Is Nothing Then
'        Set myTmp = ipNode
'        Set s.OrderRoot = ipNode.OrderSucc
'        Set s.OrderRoot.OrderPred = Nothing
'        s.OrderRoot.UpdateOrder
'        Set myTmp = Nothing
'        Exit Sub
'    End If
'
'    ' is ipnode the last order node
'    If ipNode.OrderSucc Is Nothing Then
'        ' if ipnode.ordersucc is nothing this means that ipnode **is** ipNode.last
'        Set myTmp = ipNode
'        Set s.OrderLast = ipNode.OrderPred
'        Set s.OrderLast.OrderSucc = Nothing
'        s.OrderLast.OrderPred.UpdateOrder
'        Set myTmp = Nothing
'        Exit Sub
'    End If
'
'    Set ipNode.OrderSucc.OrderPred = ipNode.OrderPred
'    Set ipNode.OrderPred.OrderSucc = ipNode.OrderSucc
'    ipNode.OrderPred.UpdateOrder
'    Set ipNode = Nothing
'
'End Sub
'
'Private Sub RemoveAllItems()
'
'    Dim myNode As SeqHNode:  Set myNode = s.OrderLast
'
'    Do
'
'        Dim myTmpNode As SeqHNode: Set myTmpNode = myNode
'        Set myNode = myNode.OrderPred
'        Set myTmpNode = Nothing
'
'    Loop Until myTmpNode Is Nothing
'
'    Set s.HashLast = Nothing
'    Set s.HashRoot = Nothing
'    Set s.OrderRoot = Nothing
'    Set s.OrderLast = Nothing
'   ' s.OrderLast.order = 0
'
'End Sub
'
'
'Private Sub RemoveAllSingleItem(ByRef ipItem As Variant)
'
'    ' all nodes with the same hash /item are linked
'    ' Where there is a hash clash the node at s.hashslots(myHash) may not hold ipItem
'    ' in which case we have to search forward to find the item
'    ' we also have to be mindful that whilst the hash of an item may be valid (the relevant hashslot is occupied)
'    ' the item itself may not be in the seq
'    Dim myHash As Long: myHash = GetHash(ipItem)
'    Dim myNode As SeqHNode: Set myNode = s.HashSlots(myHash)
'
'    'check that that hashslot is occupied
'    If myNode Is Nothing Then
'        Exit Sub
'    End If
'
'    ' search for the first matching item with the current hash
'    ' also realising that becasue we have a valid hash doesn't mean that the item is in the sequence
'    Do Until Comparers.EQ(myNode.Item, ipItem)
'        Set myNode = myNode.HashSucc
'        ' check for the case of a valid hash but no item in seq
'        ' i.e. the node hash has changed without finding a matching item
'        If myNode.Hash <> myHash Then
'            Exit Sub
'        End If
'    Loop
'
'    ' we have found a node with matching hash/item
'    ' but the next noede might be nothing
'    Do
'
'        If myNode.HashSucc Is Nothing Then
'            RemoveNode myNode
'            Exit Do
'        End If
'
'        Dim myTmp As SeqHNode: Set myTmp = myNode
'        Set myNode = myNode.HashSucc
'        RemoveNode myTmp
'
'    Loop While Comparers.EQ(myNode.Item, ipItem)
'
'End Sub

