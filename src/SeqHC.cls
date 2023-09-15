VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SeqHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "SeqHc is a Collection/ArrayList replacement. Items are stored in an Array/Quadruple linkied list using a hash value. Order of addition is preserved.  Indexing is 1 based. Keys are not supported (See Kvp classes)"
'@PredeclaredId
'@Exposed
'@Folder("Seq)
'@ModuleDescription("SeqHc is a Collection/ArrayList replacement. Items are stored in an Array/Quadruple linkied list using a hash value. Order of addition is preserved.  Indexing is 1 based. Keys are not supported (See Kvp classes)")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' uses hashing algorith from cHashD from https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs)&s=f049d882b4fc4d27f4bc2f9bc383a53e
' cHashD, Olaf Schmidt in August 2016 (enhanced about mixed, variable KeyTypes and Remove-functions in 2020)
'
' SeqHc is a list class and consequently will store multiple items with the same
' value.
' SeqHc uses a a linked list accssed by a HashTable and Collection
' hashing to allow for fast lookup of items or items by index.
' Lack of variability in the input data will significantly degrade the
' lookup performance of items.
'
' Items are stored in the nodes of a linked list and references to the nodes are
' added to a HashTable array and an order collection.  Hash clashes are managed by
' linking nodes in the HashList to get sequences of Hash/Item pairs.
'
' Nodes containing the first instance of a hash value
' are stored in an array indexed by the hash value.
' Nodes are also stored in a collection in order of addition/insertion
' Hash clashes are managed by inserting items into the linked list at the
' location defined by the hash. Hash clashes are inserted into the list following
' the last ocurrence of the Hash/Item combination
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Items are accessed using 1 based indexing
' Enumeration via NewEnum is not implemented (See Iteritems for a richer enumration experience)
' Keys are not suported (see Kvp classes)
' Preserves order of addition.
' Insertion at a defined position is supported
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
Option Explicit
'Aug 2023 - Hashing algorith moved to a separate class - Hasher with method OSHash
' constants common to all seq classes
Private Const MY_TYPENAME                       As String = "SeqHC"
Private Const FIRST_INDEX                       As Long = 1
Private Const WHOLE_SEQ                         As Long = -1
Private Const TO_LAST_INDEX                     As Long = 0
Private Const DEFAULT_STEP                      As Long = 1
Private Const LACKS_ITEMS                       As Long = -1
Private Const INDEX_OUT_OF_BOUNDS               As Long = -1
Private Const NOT_FOUND                         As Long = -1

' constants specific to class
'@Ignore ConstantNotUsed
Private Const MIN_SEQ_SIZE                  As Long = 128

 
Private Type State
    ' An array where nodes are stored using their hash as the index.
    ' Each HashSlot holds a reference to a single node, which is the
    ' first node with a particular hash value
    ' hash clashes are inserted into the list after this first node
    ' in hash/key precedence
    HashSlots()                             As SeqHNode
    
    ' first node added to the seq
    HashRoot                                As SeqHNode
    ' last node appended to the seq (note not added)
    HashLast                                As SeqHNode
    
    Hasher                                  As Hasher ' hash generator via OSHash
    
    ' a collection which stores nodes in the order in which
    ' they are added/inserted
    Order                                   As SeqC

End Type

Private s                                   As State


Private Type Properties
    ' the defined first index
    FirstIndex                              As Long
    ' The dimmed size of s.HashSlots
    Size                                    As Long
End Type

Private p                                   As Properties

' Deb is a Factory method of the PredeclaredId which returns a new instance of SeqHc
'=================================================================================
' The input for Deb is a ParamArray.
' The ParamArray is interpreted according to the folowing rules
' If the ParamArray contains no Items
' - A empty SeqHcuence (count = 0)
' If the ParamArray has 2 or more items
' - An Seq of size equal to the ParamArray, items as generated by for each over the ParamArray
' If the ParamArray has a single Item
' - Number: Set the Hashrange to the nearest biggest power of 2
' - String: A Seq whose capacity is the length of the string with one character per Item
' - Array: Seq of items as provided by For Each over the array.  Multidimension arrays are flattened but the position in the original array can be calculated by the user if required. See also Bind
' - Dictionaries: a Seq of key/value arrays as generated from for each on the Keys array.
' - Collection: The collection is copied to the Host collection.  There is no refererence link to the original collection (See Host Property)
' - SeqHc: Items as generated by for each on the SeqHc
' - mscorlib Arraylist: Items as generated by for each on the arrylist
' - mscorlib Stack: Items as generated by for each on the stack
' - mscorlib Queue: Items as generated by for eac on the queue

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'Setup and reinitialisation methods
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'@DefaultMember
Public Function Deb(ParamArray ipParamarray() As Variant) As SeqHC
Attribute Deb.VB_UserMemId = 0
    With New SeqHC
        Set Deb = .ConstructInstance(VBA.CVar(ipParamarray))
    End With
End Function


Friend Function ConstructInstance(ByRef ipFpa As Variant) As SeqHC

    p.FirstIndex = FIRST_INDEX
    Set s.Order = SeqC.Deb
    
    ' count is obtained from s.order.count
    ' Check to see if we return a default 'empty' seq
    If ArrayOp.LacksItems(ipFpa) Then
        Set s.Hasher = Hasher.Deb
        ReDim s.HashSlots(0 To s.Hasher.MaxHash)
        Set ConstructInstance = Me
        Exit Function
    End If
    
    ' ipParamArray contains one or more items depending on the input type.
    ' if the input contains only a single item we want to use that item
    ' else we want to treat the forwrded paramarray as a single array item
    Dim myItem As Variant
    If ArrayOp.HoldsOneItem(ipFpa) Then
        If VBA.IsObject(ipFpa(0)) Then
            Set myItem = ipFpa(0)
        Else
            myItem = ipFpa(0)
        End If
    Else
        myItem = ipFpa
    End If
    
    If GroupInfo.IsNumber(myItem) Then
        Set s.Hasher = Hasher(myItem)
        ReDim s.HashSlots(0 To s.Hasher.MaxHash)
        Set ConstructInstance = Me
        Exit Function
    Else
        Set s.Hasher = Hasher.Deb
        ReDim s.HashSlots(0 To s.Hasher.MaxHash)
    End If
   
    ' Resolve the items in the forwarded paramarray using the rules
    ' as described above.
    Select Case True
    
        Case VBA.IsArray(myItem):                   Set ConstructInstance = AddRange(ArrayOp.Linearise(myItem))
        Case GroupInfo.IsContainer(myItem):         Set ConstructInstance = AddRange(myItem)
        Case GroupInfo.IsString(myItem):            Set ConstructInstance = AddRange(myItem)
        Case Else:                                  Set ConstructInstance = AddItems(myItem)
    End Select
    
End Function


'@Description("Clears all entries and resets the size of the hastable to accomodate the new size")
'@Ignore AssignedByValParameter
Public Sub Reinit(Optional ByRef ipNewSize As Variant)
Attribute Reinit.VB_Description = "Clears all entries and resets the size of the hastable to accomodate the new size"

    Set s.Order = SeqC.Deb
    
    If s.Hasher Is Nothing Then
        If VBA.IsMissing(ipNewSize) Then
            If p.Size > 0 Then
                Set s.Hasher = Hasher(p.Size)
            Else
                Set s.Hasher = Hasher.Deb  ' uses DEFAULT_HASH_TABLE_SIZE
            End If
        End If
    Else
        Set s.Hasher = Hasher(ipNewSize)
    End If
   
    ' Redim without preserve should mean that we have no entries in s.HashSlots
    ReDim s.HashSlots(0 To s.Hasher.MaxHash)
    p.Size = s.Hasher.Size
   
    If Me.LacksItems Then
        Exit Sub
    End If
    
    ' reinit is clearing the seq so we need to
    ' delete all existing nodes
    Dim myNode As SeqHNode: Set myNode = s.HashLast
    Do Until myNode.HashPred Is Nothing
    
        Set myNode = myNode.HashPred
        Set myNode.HashSucc = Nothing
     Loop
     
     Set s.HashRoot = Nothing
    
End Sub

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Core methods - essential for making the class work and status of the object
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'@Description("Add a single item.  Returns the index at which the item was added for compatibility with the mscorlib ArrayList.")
Public Function Add(ByRef ipItem As Variant) As Long
Attribute Add.VB_Description = "Add a single item.  Returns the index at which the item was added for compatibility with the mscorlib ArrayList."
    
    Dim myHash As Long: myHash = s.Hasher.OSHash(ipItem)
    Dim myNewNode As SeqHNode: Set myNewNode = SeqHNode.Deb(myHash, ipItem)
    
    If Me.LacksItems Then
        
        SetupFirstNode myNewNode
        
    ElseIf s.HashSlots(myHash) Is Nothing Then
    
        ' if the hash slot is nothing then we just append
        Set s.HashSlots(myHash) = myNewNode
        AppendByHash myNewNode
        AppendByOrder myNewNode
        
    Else
        
        'The hashslot is occupied so we must insert in the hash list
        ' but append by order
        InsertbyHash myNewNode
        AppendByOrder myNewNode
        
    End If
    
    Add = s.Order.Count
    
End Function

'@Description("Returns the size of the Seq or the number of specified Items.  Returns -1 rather than 0 if there are no items in the Seq but count item will return 0 if not found")
Public Function Count(Optional ByRef ipItem As Variant) As Long
Attribute Count.VB_Description = "Returns the size of the Seq or the number of specified Items.  Returns -1 rather than 0 if there are no items in the Seq but count item will return 0 if not found"
    
    Count = LACKS_ITEMS
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If VBA.IsMissing(ipItem) Then
        Count = s.Order.Count
        Exit Function
    End If
        
    ' we know that items with the same hash are sequential
    ' in the Hash list so we can jump straight to the
    ' first node with the hash
    Dim myHash As Long: myHash = s.Hasher.OSHash(ipItem)
    Dim myNode As SeqHNode: Set myNode = s.HashSlots(myHash)
    
    ' The hash/Item pair may not match so find the
    ' first node with a matching item
    Do Until Comparers.EQ(myNode.Item, ipItem)
    
        If myNode.HashSucc Is Nothing Then
            Exit Function
        End If
        
        If myNode.Hash <> myNode.HashSucc.Hash Then
            Exit Function
        End If
        
        Set myNode = myNode.HashSucc
        
    Loop
    
    ' we have found the first matching hash/item pair
    Dim myCount As Long: myCount = 0
    Do While Comparers.EQ(myNode.Item, ipItem)
        
        myCount = myCount + 1
        If myNode.HashSucc Is Nothing Then
            Exit Do
        End If
        
        Set myNode = myNode.HashSucc
    
    Loop
    
    Count = myCount
    
End Function


'@Description("Inserts a single item at the specified index")
Public Function InsertAt(ByVal ipIndex As Long, ByVal ipItem As Variant) As Long
Attribute InsertAt.VB_Description = "Inserts a single item at the specified index"

    InsertAt = INDEX_OUT_OF_BOUNDS
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "InsertAt", Guard.REPORT_BACK) Then
        Exit Function
    End If
    
    ' create the new node
    Dim myHash As Long: myHash = s.Hasher.OSHash(ipItem)
    Dim myNewNode As SeqHNode: Set myNewNode = SeqHNode.Deb(myHash, ipItem)
        
    ' there are two strands to update, byHash and ByOrder
    ' deal with byHash first
    If s.HashSlots(myHash) Is Nothing Then
        AppendByHash myNewNode
    Else
        InsertbyHash myNewNode
    End If
    
    Select Case ipIndex
        Case Me.FirstIndex:  PrependByOrder myNewNode
        Case Me.Lastindex:   InsertByOrder ipIndex, myNewNode
        Case Else:           InsertByOrder ipIndex, myNewNode
    End Select

    InsertAt = ipIndex
    
End Function


Public Property Get Item(ByVal ipIndex As Long) As Variant

    Item = Null
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "Get Item", Guard.REPORT_BACK) Then
        Exit Property
    End If
    
    Dim myItem As Variant: myItem = Array(s.Order.Item(ipIndex).Item)
    
    If VBA.IsObject(myItem(0)) Then
        Set Item = myItem(0)
    Else
        Item = myItem(0)
    End If
    
End Property


Public Property Let Item(ByVal ipIndex As Long, ByVal ipItem As Variant)

    Guard.IndexOutOfBounds ipIndex, Me, "SeqHc.Let_Item"

    s.Order.Item(ipIndex).Item = ipItem
    
End Property


Public Property Set Item(ByVal ipIndex As Long, ByVal ipItem As Variant)
 
    Guard.IndexOutOfBounds ipIndex, Me, "SeqHc.Set_Item"

    Set s.Order.Item(ipIndex).Item = ipItem
    
End Property



'@description("Removes a single item.  Returns the index of the item)
Public Function Remove(ByRef ipItem As Variant) As Long

    Remove = NOT_FOUND

    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As SeqHNode: Set myNode = GetNodeOf(ipItem)
    
    If NotFound(myNode) Then
        Exit Function
    End If
    
    Remove = myNode.Order
    RemoveNode myNode
    
End Function


'@Description("Removes the item at the specified index")
Public Function RemoveAt(ByRef ipIndex As Long) As Long
Attribute RemoveAt.VB_Description = "Removes the item at the specified index"

    RemoveAt = NOT_FOUND
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "RemoveAt", Guard.REPORT_BACK) Then
        Exit Function
    End If
    
    Dim myNode As SeqHNode: Set myNode = s.Order.Item(ipIndex)
    
    RemoveNode myNode
    RemoveAt = ipIndex
    
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Sugar methods' Add to the flexibility of the seq class using the above methods
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'@Description("Makes a shallow copy of the seq")
Public Function Clone() As SeqHC
Attribute Clone.VB_Description = "Makes a shallow copy of the seq"

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set Clone = myS

    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myIndex As Long
     For myIndex = Me.FirstIndex To Me.Lastindex
        myS.Add s.Order.Item(myIndex).Item
    Next
    
    Set Clone = myS
    
End Function


'@Description("Returns the Items in the seq as an Array() of variants")
Public Function Items() As Variant
Attribute Items.VB_Description = "Returns the Items in the seq as an Array() of variants"

    Items = Array()

    If Me.LacksItems Then
        Exit Function
    End If
    
    Items = ToArray

End Function


'@Description("Returns a subset of the SeqHc")
Public Function Slice _
( _
Optional ByRef ipStart As Long = FIRST_INDEX, _
Optional ByRef ipRun As Long = WHOLE_SEQ, _
Optional ByRef ipStep As Long = DEFAULT_STEP _
) As SeqHC
Attribute Slice.VB_Description = "Returns a subset of the SeqHc"

    ' ToDo: Should all Seq return a SeqA for SLice?
    Set Slice = SeqHC.Deb

    If Me.LacksItems Then
        Exit Function
    End If

    Guard.IndexOutOfBounds ipStart, Me, "Slice"

    Dim myEnd As Long
    If ipRun < Me.FirstIndex Then
        myEnd = Me.Lastindex
    Else
        myEnd = ipStart + ipRun - 1
        If myEnd > Me.Lastindex Then
            myEnd = Me.Lastindex
        End If
    End If

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Dim mySeqIndex As Long
    For mySeqIndex = ipStart To myEnd Step ipStep
        myS.Add s.Order.Item(mySeqIndex).Item
    Next

    Set Slice = myS

End Function
'@Description("Returns a shallow copy of the SeqHc as an Array.  The Lbound of the array is 1")
Public Function ToArray() As Variant
Attribute ToArray.VB_Description = "Returns a shallow copy of the SeqHc as an Array.  The Lbound of the array is 1"

    ToArray = Array()

    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myS As SeqA: Set myS = SeqA(Me.Lastindex)
    Dim myIndex As Long
    For myIndex = s.Order.FirstIndex To s.Order.Lastindex
        myS.Add s.Order.Item(myIndex).Item
    Next

    ToArray = myS.ToArray

End Function


'@Description("Returns a shallow copy of the SeqHc as a VBA collection")
Public Function ToCollection() As Collection
Attribute ToCollection.VB_Description = "Returns a shallow copy of the SeqHc as a VBA collection"

    Dim myC As Collection: Set myC = New Collection
    Set ToCollection = myC
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myIndex As Long
    For myIndex = s.Order.FirstIndex To s.Order.Lastindex
        myC.Add s.Order.Item(myIndex).Item
    Next
    
    Set ToCollection = myC

End Function


'@Description("Returns the predefined firstindex")
Public Function FirstIndex() As Long
Attribute FirstIndex.VB_Description = "Returns the predefined firstindex"
    FirstIndex = p.FirstIndex
End Function


'@Description("Returns the LastIndex of the Seq or -1 if the Seq is empty.  Equivalent to Count.")
Public Function Lastindex() As Long
Attribute Lastindex.VB_Description = "Returns the LastIndex of the Seq or -1 if the Seq is empty.  Equivalent to Count."

    If Me.Count < Me.FirstIndex Then
        Lastindex = LACKS_ITEMS
    Else
        Lastindex = Me.Count
    End If

End Function

'@Description("Returns the predefined First but One index")
Public Function FBOIndex() As Long
Attribute FBOIndex.VB_Description = "Returns the predefined First but One index"
    FBOIndex = Me.FirstIndex + 1
End Function


'Description("Returns the index of the second to last item in the SeqHc, or -1 if the SeqHc has lees than 2 items")
Public Function LBOIndex() As Long

    If Me.Lastindex < Me.FBOIndex Then
        LBOIndex = LACKS_ITEMS
    Else
        LBOIndex = Me.Lastindex - 1
    End If
   
End Function

'@Description("Returns the item at FirstIndex")
Public Property Get First() As Variant
Attribute First.VB_Description = "Returns the item at FirstIndex"

    First = Null
    
    If Me.LacksItems Then
        Exit Property
    End If

    If VBA.IsObject(s.Order.Item(Me.FirstIndex).Item) Then
        Set First = s.Order.Item(Me.FirstIndex).Item
    Else
        First = s.Order.Item(Me.FirstIndex).Item
    End If

End Property


'@Description("Assigns to the Item at FirstIndex")
Public Property Let First(ByVal ipItem As Variant)
Attribute First.VB_Description = "Assigns to the Item at FirstIndex"

    ' ToDo: review if the add action is better as an error
    If Me.LacksItems Then
        Me.Add ipItem
        Exit Property
    End If

    s.Order.Item(Me.FirstIndex).Item = ipItem

End Property


'@Description("Assigns to the Item at FirstIndex")
Public Property Set First(ByVal ipItem As Variant)
Attribute First.VB_Description = "Assigns to the Item at FirstIndex"

    If Me.LacksItems Then
        Me.Add ipItem
        Exit Property
    End If

    Set s.Order.Item(Me.FirstIndex).Item = ipItem

End Property


'@Description("Returns the Item at LastIndex")
Public Property Get Last() As Variant
Attribute Last.VB_Description = "Returns the Item at LastIndex"

    Last = Null
    
    If Me.LacksItems Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Order.Item(Me.Lastindex).Item) Then
        Set Last = s.Order.Item(Me.Lastindex).Item
    Else
        Last = s.Order.Item(Me.Lastindex).Item
    End If

End Property


'@Description("Assigns to the Item at LastIndex")
Public Property Let Last(ByVal ipItem As Variant)
Attribute Last.VB_Description = "Assigns to the Item at LastIndex"

    If Me.LacksItems Then
        Me.Add ipItem
        Exit Property
    End If

    s.Order.Item(Me.Lastindex).Item = ipItem

End Property


'@Description("Assigns to the item at LastIndex")
Public Property Set Last(ByVal ipItem As Variant)
Attribute Last.VB_Description = "Assigns to the item at LastIndex"

    If Me.LacksItems Then
        Me.Add ipItem
        Exit Property
    End If

    Set s.Order.Item(Me.Lastindex).Item = ipItem
   
End Property

'@Description("Returns the item at First but one Index")
Public Property Get FBO() As Variant
Attribute FBO.VB_Description = "Returns the item at First but one Index"

    FBO = Null
 
    If Guard.IndexNotFound(FBOIndex, Me, "Get FBO", Guard.REPORT_BACK) Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Order.Item(Me.FBOIndex).Item) Then
        Set FBO = s.Order.Item(Me.FBOIndex).Item
    Else
        FBO = s.Order.Item(Me.FBOIndex).Item
    End If

End Property


'@Description("Assigns to the Item at First but one Index")
Public Property Let FBO(ByVal ipItem As Variant)
Attribute FBO.VB_Description = "Assigns to the Item at First but one Index"

    Guard.IndexNotFound FBOIndex, Me, "Let FBO"

    s.Order.Item(Me.FBOIndex).Item = ipItem

End Property


'@Description("Assigns to the Item at First but one Index")
Public Property Set FBO(ByVal ipItem As Variant)
Attribute FBO.VB_Description = "Assigns to the Item at First but one Index"

    Guard.IndexNotFound Me.FBOIndex, Me, "Set LBO"

    Set s.Order.Item(Me.FBOIndex).Item = ipItem
    

End Property


'@Description("Returns the Item at Lastbut one Index")
Public Property Get LBO() As Variant
Attribute LBO.VB_Description = "Returns the Item at Lastbut one Index"

    LBO = Null
    
    If Guard.IndexNotFound(Me.LBOIndex, Me, "Get LBO", Guard.REPORT_BACK) Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Order.Item(Me.LBOIndex).Item) Then
        Set LBO = s.Order.Item(Me.LBOIndex).Item
    Else
        LBO = s.Order.Item(Me.LBOIndex).Item
    End If

End Property


'@Description("Assigns to the Item at Last but one Index")
Public Property Let LBO(ByVal ipItem As Variant)
Attribute LBO.VB_Description = "Assigns to the Item at Last but one Index"

     Guard.IndexNotFound Me.LBOIndex, Me, "Let LBO"

     s.Order.Item(Me.LBOIndex).Item = ipItem

End Property


'@Description("Assigns to the item at Last but one Index")
Public Property Set LBO(ByVal ipItem As Variant)
Attribute LBO.VB_Description = "Assigns to the item at Last but one Index"

    Guard.IndexNotFound Me.LBOIndex, Me, "Set LBO"

    Set s.Order.Item(Me.LBOIndex).Item = ipItem
   
End Property

'@Description("Returns the index of the first Item found or -1.  Search range may be specified")
Public Function IndexOf _
( _
    ByVal ipItem As Variant, _
    Optional ByVal ipStart As Long = FIRST_INDEX, _
    Optional ByVal ipRun As Long = TO_LAST_INDEX _
) As Long
Attribute IndexOf.VB_Description = "Returns the index of the first Item found or -1.  Search range may be specified"
    
    IndexOf = NOT_FOUND
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.IndexOutOfBounds(ipStart, Me, "IndexOf", Guard.REPORT_BACK) Then
        Exit Function
    End If
    
    ' this quick exit is an advantage of the hashing approach
    If Me.LacksItem(ipItem) Then
        Exit Function
    End If
    
    ' we now know that ipItem exists in the seq
    ' but does it exist in the range defined by ipstart,iprun

    ' ensure that ipstart+iprun only goes to the end of the seq
    Dim myEnd As Long
    If ipRun = TO_LAST_INDEX Then
        myEnd = Me.Lastindex
    ElseIf ipRun > 0 Then
        myEnd = ipStart + ipRun - 1
        If myEnd > Me.Lastindex Then
            myEnd = Me.Lastindex
        End If
    Else
        Exit Function
    End If
    
    'The first found node may be the first in a sequence so we need to
    ' search for the first node whose order is in the range ipStart to ipEnd
    Dim myIndex As Long
    For myIndex = ipStart To myEnd
        
        If Comparers.EQ(s.Order.Item(myIndex).Item, ipItem) Then
            IndexOf = myIndex
            Exit Function
        End If
            
    Next
    
End Function

Public Function LastIndexOf _
( _
ByVal ipItem As Variant, _
Optional ByVal ipStart As Long = FIRST_INDEX, _
Optional ByVal ipRun As Long = TO_LAST_INDEX _
) As Long
    
    LastIndexOf = NOT_FOUND
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.IndexOutOfBounds(ipStart, Me, "LastIndexOf", Guard.REPORT_BACK) Then
        Exit Function
    End If
    
    ' this quick exit is an advantage of the hashing approach
    If Me.LacksItem(ipItem) Then
        Exit Function
    End If
    
    ' ensure that ipstart+iprun only goes to the end of the seq
    Dim myEnd As Long
    If ipRun = TO_LAST_INDEX Then
        myEnd = Me.Lastindex
    ElseIf ipRun > 0 Then
        myEnd = ipStart + ipRun - 1
        If myEnd > Me.Lastindex Then
            myEnd = Me.Lastindex
        End If
    Else
        Exit Function
    End If
        
    Dim myIndex As Long
    For myIndex = myEnd To ipStart Step -1
        
        If Comparers.EQ(s.Order.Item(myIndex).Item, ipItem) Then
            LastIndexOf = myIndex
            Exit Function
        End If
    
    Next
        
    LastIndexOf = NOT_FOUND
    
End Function


'@Description("Add Items in the paramarray to the  )
Public Function AddItems(ParamArray ipParamarray() As Variant) As SeqHC

    If ArrayOp.LacksItems(VBA.CVar(ipParamarray)) Then
        Set AddItems = Me
        Exit Function
    End If

    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(VBA.CVar(ipParamarray))

    AddRange myParamarray

    Set AddItems = Me

End Function

'@Description("Add an array,  list type, dictionary type, or characters of a string")
Public Function AddRange(ByVal ipRange As Variant) As SeqHC
Attribute AddRange.VB_Description = "Add an array,  list type, dictionary type, or characters of a string"

    Set AddRange = Me
    
    Guard.InvalidRangeItem ipRange, MY_TYPENAME, "AddRange", Guard.REPORT_BACK
    
    Dim myItems As IterItems: Set myItems = IterItems(ipRange)
    
    If myItems.LacksItems Then
        Exit Function
    End If
    
    Do
        If GroupInfo.IsDictionary(ipRange) Then
            Me.Add Array(myItems.CurKey(0), myItems.CurItem(0))
        Else
            Me.Add myItems.CurItem(0)
        End If
        
    Loop While myItems.MoveNext
    
    Set AddRange = Me
    
End Function


'@Description("Inserts one or more items starting at ipIndex")
Public Function InsertAtItems(ByVal ipIndex As Long, ParamArray ipParamarray() As Variant) As SeqHC
Attribute InsertAtItems.VB_Description = "Inserts one or more items starting at ipIndex"
    
    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(VBA.CVar(ipParamarray))
    
    Set InsertAtItems = InsertAtRange(ipIndex, myParamarray)
    
End Function


'@Description("Inserts items in a cotainer starting at ipIndex")
Public Function InsertAtRange(ByVal ipIndex As Long, ByRef ipRange As Variant) As SeqHC
Attribute InsertAtRange.VB_Description = "Inserts items in a cotainer starting at ipIndex"

    Set InsertAtRange = Me

    If Guard.IndexOutOfBounds(ipIndex, Me, "InsertAtRange", Guard.REPORT_BACK) Then
        Exit Function
    End If
    
    Guard.InvalidRangeItem ipRange, "SeqHc", "InsertAtRange"
    
    ' insert in reverse order so we don't have to increment ipIndex
    Dim myItems As IterItems: Set myItems = IterItems(ipRange).MoveToEndIndex
    
    If myItems.LacksItems Then
        Exit Function
    End If
    
    Do
        
        If GroupInfo.IsDictionary(ipRange) Then
            InsertAt ipIndex, Array(myItems.CurKey(0), myItems.CurItem(0))
        Else
            InsertAt ipIndex, myItems.CurItem(0)
        End If
   
    Loop While myItems.MovePrev
    
    Set InsertAtRange = Me
    
End Function


'@Description "Removes the first encountered item from the Seq for each item in the ParamArray"
Public Function RemoveItems(ParamArray ipParamarray() As Variant) As SeqHC
Attribute RemoveItems.VB_Description = "Removes the first encountered item from the Seq for each item in the ParamArray"

    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(VBA.CVar(ipParamarray))
    
    Set RemoveItems = RemoveRange(myParamarray)
    
    Set RemoveItems = Me
    
End Function

Public Function RemoveRange(ByRef ipRange As Variant) As SeqHC

    Set RemoveRange = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myItems As IterItems: Set myItems = IterItems(ipRange)
    
    If myItems.LacksItems Then
        Exit Function
    End If
    
    Do
        RemoveNode GetNodeOf(myItems.CurItem(0))
    Loop While myItems.MoveNext
    
    Set RemoveRange = Me
    
End Function

Public Function RemoveAll(ParamArray ipParamarray() As Variant) As SeqHC

    Set RemoveAll = Me
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(VBA.CVar(ipParamarray))

    If ArrayOp.LacksItems(myParamarray) Then
        ' For SeqHc removeall is effectively the same as reset
        RemoveAllItems
        Exit Function
    End If
    
    Dim myParam As Variant
    For Each myParam In myParamarray
        RemoveAllSingleItem myParam
    Next
    
    Set RemoveAll = Me

End Function


Public Function Clear() As SeqHC
    Me.RemoveAll
    Set Clear = Me
End Function


Public Function Reset() As SeqHC
    
    Set Reset = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As SeqHNode: Set myNode = s.HashLast
    Do Until myNode.HashPred Is Nothing
        Set myNode = myNode.HashPred
        RemoveNode myNode.HashSucc
    Loop
    
    ReDim s.HashSlots(0 To s.Hasher.MaxHash)
    
    Set s.Order = SeqC.Deb

    Set s.HashRoot = Nothing
    Set s.HashLast = Nothing
    
End Function


'Public Function RemoveAfter(ByRef ipIndex As Long, Optional ByVal ipRun As Long = 1) As SeqHc
'
'    Set RemoveAfter = Me
'
'    If Me.lacksItems Then
'        Exit Function
'    End If
'
'    If Me.LacksKey(ipKey) Then
'        Exit Function
'    End If
'
'    If ipRun < 1 Then
'        Err.Raise 17 + vbObjectError, _
'        "VALib.SeqHc.RemoveAfter", _
'        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'    Do Until ComparerHelpers.Equals(ipKey, myNode.Key)
'        Set myNode = myNode.Succ
'        If myNode Is Nothing Then
'            Exit Function
'        End If
'
'    Loop
'
'    If myNode.Succ Is Nothing Then
'        Exit Function
'    End If
'
'    '@Ignore AssignmentNotUsed
'    Set myNode = myNode.Succ
'
'    Dim myCount As Long
'    For myCount = 1 To ipRun
'
'        RemoveNode myNode.Succ
'
'        If myNode.Succ Is Nothing Then
'            Exit Function
'        End If
'
'    Next
'
'End Function
'
'
'Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As SeqHc
'
'    Set RemoveBefore = Me
'
'    If VBA.IsNull(ipKey) Then
'        Err.Raise 17 + vbObjectError, _
'        "VBALib.SeqHc.RemoveBefore", _
'        "Null is not a valid Key"
'    End If
'
'    If Me.LacksKey(ipKey) Then
'        Exit Function
'    End If
'
'    If ipRun < 1 Then
'        Err.Raise 17 + vbObjectError, _
'        "VALib.SeqHc.RemoveBefore", _
'        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
'    End If
'
'    Dim myNode As SeqHNode: Set myNode = s.HashRoot
'    Do Until ComparerHelpers.Equals(ipKey, myNode.Key)
'        Set myNode = myNode.Succ
'        If myNode Is Nothing Then
'            Exit Function
'        End If
'
'    Loop
'
'    If myNode.Pred Is Nothing Then
'        Exit Function
'    End If
'
'    '@Ignore AssignmentNotUsed
'    Set myNode = myNode.Pred
'
'    Dim myCount As Long
'    For myCount = 1 To ipRun
'
'        RemoveNode myNode.Pred
'
'        If myNode.Pred Is Nothing Then
'            Exit Function
'        End If
'
'    Next
'
'End Function




'ToDo: rename RemoveIndexes to removeIndexes
'@Description("RemoveAt for a list of indexes.  Indexes are sorted and removed in reverse order")
Public Function RemoveIndexes(ParamArray ipParamarray() As Variant) As SeqHC
Attribute RemoveIndexes.VB_Description = "RemoveAt for a list of indexes.  Indexes are sorted and removed in reverse order"

    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(VBA.CVar(ipParamarray))
    
    Set RemoveIndexes = RemoveIndexesRange(myParamarray)
    
    Set RemoveIndexes = Me
    
End Function

'@Description("Removes the Items at the indexes in the range list")
Public Function RemoveIndexesRange(ByRef ipRange As Variant) As SeqHC
Attribute RemoveIndexesRange.VB_Description = "Removes the Items at the indexes in the range list"

    Set RemoveIndexesRange = Me
    
    If Me.LacksItems Then
        Exit Function
    End If

    ' We cannot use iteritems on ipRange directly because
    ' we need to ensure that the indexes are removed in reverse
    ' order to avoid indexes being disconnected from thier original positions
    Dim myS As SeqA
    ' SeqA will combine keys and items to a an array
    ' if ipRange is a dictionary
    ' so we choose to select the Items as the indexes
    If GroupInfo.IsDictionary(ipRange) Then
        Set myS = SeqA(ipRange.Items)
    Else
        Set myS = SeqA(ipRange)
    End If
    
    If myS.LacksItems Then
        Exit Function
    End If
    
    myS.Sorted
    
    ' we iterate backwards so we don't so we don't compromise indexes yet to be removed
    'Dim myIndex As Long
    Dim myItems As IterItems: Set myItems = IterItems(myS).MoveToEndIndex
    Do
        RemoveNode GetNodeAt(myItems.CurItem(0))
    Loop While myItems.MovePrev
    
    Set RemoveIndexesRange = Me
    
End Function


Public Function Fill(ByRef ipItem As Variant, Optional ByVal ipCount As Long = 1) As SeqHC

    Set Fill = Me
    
    Dim myCount As Long
    For myCount = 1 To ipCount
        Me.Add ipItem
    Next

End Function



'@Description("Returns the first N items of the SeqHc, Default 1")
Public Function Head(Optional ByVal ipSplitAfter As Long = FIRST_INDEX) As SeqHC
Attribute Head.VB_Description = "Returns the first N items of the SeqHc, Default 1"

    Set Head = SeqHC.Deb
    
    If Me.LacksItems Then
        Exit Function
    End If

    If ipSplitAfter < Me.FirstIndex Then
        Exit Function
    End If

    If ipSplitAfter >= Me.Lastindex Then
        Set Head = Me.Clone
        Exit Function
    End If

    Set Head = Me.Slice(Me.FirstIndex, ipSplitAfter)

End Function


'@Description("Returns the Items after SplitAfter, Default is Index 2")
Public Function Tail(Optional ByVal ipSplitAfter As Long = FIRST_INDEX) As SeqHC
Attribute Tail.VB_Description = "Returns the Items after SplitAfter, Default is Index 2"

    Set Tail = SeqHC.Deb
    
    If Me.LacksItems Then
        Exit Function
    End If

    If ipSplitAfter >= Me.Lastindex Then
        Exit Function
    End If
    
    If ipSplitAfter < Me.FirstIndex Then
        Set Tail = Me.Clone
        Exit Function
    End If

    Set Tail = Me.Slice(ipSplitAfter + 1)

End Function

Public Function HoldsItem(ByRef ipItem As Variant) As Boolean

    HoldsItem = False
    
    If Me.LacksItems Then
        Exit Function
    End If
        
    Dim myNode As SeqHNode: Set myNode = GetNodeOf(ipItem)
    
    If NotFound(myNode) Then
        Exit Function
    End If
    
    Dim myHash As Long: myHash = myNode.Hash
    ' found hass but item may not match
    Do While myHash = myNode.Hash
    
        If Comparers.EQ(myNode.Item, ipItem) Then
            HoldsItem = True
            Exit Function
        End If
        
        If myNode.HashSucc Is Nothing Then
            Exit Function
        End If
        
        Set myNode = myNode.HashSucc
        
    Loop
    
End Function


Public Function Contains(ByRef ipItem As Variant) As Boolean
    Contains = Me.HoldsItem(ipItem)
End Function

Public Function LacksItem(ByRef ipItem As Variant) As Boolean
    LacksItem = Not HoldsItem(ipItem)
End Function


Public Function HoldsItems() As Boolean
    HoldsItems = Not s.HashRoot Is Nothing
End Function


Public Function LacksItems() As Boolean
    LacksItems = Not Me.HoldsItems
End Function

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' An attempt as  providing some 'lambda' functionality - see IMapper and mp/cmp/rd classes
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@Description("Mutates the Seq in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As SeqHC
Attribute MappedIt.VB_Description = "Mutates the Seq in place by applying the function in IMapper to each item"

    Set MappedIt = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As SeqHNode
    For Each myNode In s.Order
    
        Dim myA() As Variant: myA = ipMapper.ExecMapper(myNode.Item)
        
        If VBA.IsObject(myA(0)) Then
            Set myNode.Item = myA(0)
        Else
            myNode.Item = myA(0)
        End If
        
    Next
   
    Set MappedIt = Me

End Function


'@Description("Returns a new Seq by applying the function in IMapper to each Item")
Public Function MapIt(ByVal ipMapper As IMapper) As SeqHC
Attribute MapIt.VB_Description = "Returns a new Seq by applying the function in IMapper to each Item"

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set MapIt = myS
    
    If Me.LacksItems Then
        Set MapIt = Me
        Exit Function
    End If

    Dim myIndex As Long
    For myIndex = s.Order.FirstIndex To s.Order.Lastindex
        myS.Add ipMapper.ExecMapper(s.Order.Item(myIndex).Item)(0)
    Next
          
    Set MapIt = myS

End Function


'@Description("Returns a new Seq where each item returns true from the IComparer")
Public Function FilterIt(ByVal ipComparer As IComparer) As SeqHC
Attribute FilterIt.VB_Description = "Returns a new Seq where each item returns true from the IComparer"

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set FilterIt = myS
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myIndex As Long
    For myIndex = s.Order.FirstIndex To s.Order.Lastindex
    
        If ipComparer.ExecCmp(s.Order.Item(myIndex).Item) Then
            myS.Add s.Order.Item(myIndex).Item
        End If
        
        'Set myNode = myNode.OrderSucc
        
    Next
         
    Set FilterIt = myS

End Function


'@Description("Returns the count the number of items where IComparer returns true")
Public Function CountIt(ByVal ipComparer As IComparer) As Long
Attribute CountIt.VB_Description = "Returns the count the number of items where IComparer returns true"

    CountIt = -1
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myCount As Long: myCount = 0
    
    Dim myIndex As Long
    For myIndex = s.Order.FirstIndex To s.Order.Lastindex
    
        If ipComparer.ExecCmp(s.Order.Item(myIndex).Item) Then
            myCount = myCount + 1
        End If
    
        'Set myNode = myNode.OrderSucc
        
    Next
    
    CountIt = myCount

End Function

'@Description("Returns the single value computed by the IReducer").
Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant

    ReduceIt = Null

    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myItem As Variant: myItem = Array(ipReducer.ExecReduction(Me.ToArray))
    
    If VBA.IsObject(myItem(0)) Then
        Set ReduceIt = myItem(0)
    Else
        ReduceIt = myItem(0)
    End If
    
End Function

'@Description("Adds items to the seq at the end of the seq")
Public Function Push(ParamArray ipParamarray() As Variant) As SeqHC
Attribute Push.VB_Description = "Adds items to the seq at the end of the seq"

    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(VBA.CVar(ipParamarray))
    
    Set Push = AddRange(myParamarray)
    
End Function

'@Description("Add the Items in ipRange to the end of the Seq")
Public Function PushRange(ByRef ipRange As Variant) As SeqHC
Attribute PushRange.VB_Description = "Add the Items in ipRange to the end of the Seq"
    Set PushRange = AddRange(ipRange)
End Function


'@Description("Removes the last item from the seq and returns it")
Public Function Pop() As Variant
Attribute Pop.VB_Description = "Removes the last item from the seq and returns it"

    Pop = Null

    If Me.LacksItems Then
        Exit Function
    End If
   
    If VBA.IsObject(s.Order.Item(Me.Lastindex).Item) Then
        Set Pop = s.Order.Item(Me.Lastindex).Item
    Else
        Pop = s.Order.Item(Me.Lastindex).Item
    End If

    RemoveNode s.Order.Item(Me.Lastindex)
        
End Function

'@Description("Removes ipCount Items from the  end of the Seq. Returns a SeqA of the items in order of removal"
Public Function PopRange(Optional ByVal ipCount As Long = 1) As SeqA

    Set PopRange = SeqA.Deb
    
    If ipCount < Me.FirstIndex Then
        Exit Function
    End If
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If ipCount > Me.Lastindex Then
        'Bug: SeqA(Me).reverse instantiates a seqc
        Set PopRange = SeqA(Me.Reverse)
        Me.RemoveAll
        Exit Function
    End If
    
    Dim myS As SeqA: Set myS = SeqA.Deb
    
    Dim myCount As Long
    For myCount = 1 To ipCount
        myS.Add Me.Pop
    Next
    
    Set PopRange = myS
    
End Function

'@Description("Adds items to the seq at the end of the seq")
Public Function enQueue(ParamArray ipParamarray() As Variant) As SeqHC
Attribute enQueue.VB_Description = "Adds items to the seq at the end of the seq"

    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(VBA.CVar(ipParamarray))
    
    Set enQueue = AddRange(myParamarray)
    
End Function

Public Function EnqueueRange(ByRef ipRange As Variant) As SeqHC
    Set EnqueueRange = AddRange(ipRange)
End Function

Public Function Dequeue() As Variant

    Dequeue = Null

    If Me.LacksItems Then
        Exit Function
    End If
    
    If VBA.IsObject(s.Order.Item(Me.FirstIndex).Item) Then
        Set Dequeue = s.Order.Item(Me.FirstIndex).Item
    Else
        Dequeue = s.Order.Item(Me.FirstIndex).Item
    End If

    RemoveNode s.Order.Item(Me.FirstIndex)

End Function
'@Description("SeqA is deliberate")
Public Function DequeueRange(Optional ByRef ipCount As Long = 1) As SeqA
Attribute DequeueRange.VB_Description = "SeqA is deliberate"

    Set DequeueRange = SeqA.Deb
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If ipCount < 1 Then
        Exit Function
    End If
    
    If ipCount > Me.Lastindex Then
        ipCount = Me.Lastindex
    End If

    Dim myS As SeqA: Set myS = SeqA.Deb
    
    Dim myCount As Long
    For myCount = 1 To ipCount
        myS.Add Me.Dequeue
    Next
    
    Set DequeueRange = myS
    
End Function


'@Description("Returns a new sequence sorted in 'Ascending' order"
Public Function Sort() As SeqHC

    Dim myS As SeqHC: Set myS = SeqHC.Deb
        
    If Me.LacksItems Then
        Set Sort = myS
        Exit Function
    End If
        
    '@Ignore AssignmentNotUsed
    Set myS = Me.Clone
    Sorters.ShakerSortByItem myS
    Set Sort = myS
    
End Function


'@Description("Sorts the seq in place in 'Ascending' order ")
Public Function Sorted() As SeqHC
Attribute Sorted.VB_Description = "Sorts the seq in place in 'Ascending' order "
            
    Set Sorted = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
        
    Sorters.ShakerSortByItem Me
    Set Sorted = Me
    
End Function


'@Description("Returns a new seq with Items in reverse order")
Public Sub Swap(ByRef ipLHSIndex As Long, ByRef ipRHSIndex As Long)
Attribute Swap.VB_Description = "Returns a new seq with Items in reverse order"

    Dim myTemp As Variant
    
    If VBA.IsObject(Item(ipLHSIndex)) Then
        Set myTemp = Item(ipLHSIndex)
    Else
        myTemp = Item(ipLHSIndex)
    End If
    
    If VBA.IsObject(Item(ipRHSIndex)) Then
        Set Item(ipLHSIndex) = Item(ipRHSIndex)
    Else
        Item(ipLHSIndex) = Item(ipRHSIndex)
    End If
    
    If VBA.IsObject(myTemp) Then
        Set Item(ipRHSIndex) = myTemp
    Else
        Item(ipRHSIndex) = myTemp
    End If

End Sub

'@Description("Returns a new seq in reverse order of host seq")
Public Function Reverse() As SeqHC
Attribute Reverse.VB_Description = "Returns a new seq in reverse order of host seq"

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set Reverse = myS
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myIndex As Long
    For myIndex = Me.Lastindex To Me.FirstIndex Step -1
        myS.Add s.Order.Item(myIndex).Item
    Next
         
    Set Reverse = myS

End Function

'@Description("Reverses the order of items in place")
Public Function Reversed() As SeqHC
Attribute Reversed.VB_Description = "Reverses the order of items in place"
    
    If Me.LacksItems Then
        Set Reversed = Me
        Exit Function
    End If
     
    Dim myFirstIndex As Long: myFirstIndex = Me.FirstIndex
    Dim myLastIndex As Long: myLastIndex = Me.Lastindex
    
    Do
    
        Me.Swap myFirstIndex, myLastIndex
        myFirstIndex = myFirstIndex + 1
        myLastIndex = myLastIndex - 1
    
    Loop Until myFirstIndex >= myLastIndex

    Set Reversed = Me

End Function


'@Description("Returns true if there are no duplicate items")
Public Function IsUnique() As Boolean
Attribute IsUnique.VB_Description = "Returns true if there are no duplicate items"

    If LacksItems Then
        IsUnique = False
        Exit Function
    End If
    
    Dim myS As cHashC: Set myS = New cHashC
    
    Dim myNode As SeqHNode: Set myNode = s.HashRoot
    
    Do
    
        If myS.Exists(myNode.Item) Then
            IsUnique = False
            Exit Function
        Else
            myS.Add myNode.Item
        End If
        
        Set myNode = myNode.HashSucc
        
    Loop Until myNode Is Nothing

    IsUnique = True

End Function


Public Function IsNotUnique() As Boolean
    IsNotUnique = Not IsUnique
End Function

'@Description("Returns the a new seq with the unique subset of the host Seq")
Public Function Dedup() As SeqHC
Attribute Dedup.VB_Description = "Returns the a new seq with the unique subset of the host Seq"

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set Dedup = myS

    If Me.LacksItems Then
        Exit Function
    End If

    Dim myNode As SeqHNode: Set myNode = s.HashRoot
    Do
        If myS.LacksItem(myNode.Item) Then
            myS.Add myNode.Item
        End If
        
        Set myNode = myNode.HashSucc
        
    Loop Until myNode Is Nothing

    Set Dedup = myS
            
End Function

'@Description("Removes duplicate items from the seq")
Public Function Deduped() As SeqHC
Attribute Deduped.VB_Description = "Removes duplicate items from the seq"

    Set Deduped = Me
    
    If LacksItems Then
        Exit Function
    End If
    
    Dim myS As SeqHC: Set myS = SeqHC.Deb
    
    Dim myNode As SeqHNode
    For Each myNode In s.Order
        
        If myS.HoldsItem(myNode.Item) Then
        
            Dim myTmpNode As SeqHNode: Set myTmpNode = myNode
            Set myNode = myNode.OrderSucc
            RemoveNode myTmpNode
            
        Else
        
            myS.Add myNode.Item
            Set myNode = myNode.OrderSucc
        End If
        
    Next
    Set Deduped = Me
    
End Function


Public Function Freq() As KvpA

    Dim myK As KvpA: Set myK = KvpA.Deb
    Set Freq = myK
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As SeqHNode: Set myNode = s.HashRoot
    Do
    
        If myK.HoldsItem(myNode.Item) Then
            myK.Item(myNode.Item) = myK.Item(myNode.Item) + 1
        Else
            myK.Add myNode.Item, 1
        End If
         
        Set myNode = myNode.HashSucc
         
    Loop Until myNode Is Nothing

    Set Freq = myK
    
End Function

Public Function SetOf(ByRef ipSet As e_SetoF, ByRef ipRangeItem As SeqHC) As SeqHC

    Select Case ipSet

        Case e_SetoF.m_Common:           Set SetOf = SetOfCommon(ipRangeItem)
        Case e_SetoF.m_HostOnly:         Set SetOf = SetOfHostOnly(ipRangeItem)
        Case e_SetoF.m_ParamOnly:        Set SetOf = SetOfParamOnly(ipRangeItem)
        Case e_SetoF.m_NotCommon:        Set SetOf = SetOfNotCommon(ipRangeItem)
        Case e_SetoF.m_Unique:           Set SetOf = SetofUnique(ipRangeItem)
        Case Else
            Err.Raise 17 + vbObjectError, _
            "VBALib.SeqHc.SetOf", _
            Fmt.Text("Invalid e_set member. Expecting {0} to {1}. Got {2}", e_SetoF.m_First, e_SetoF.m_Last, ipSet)

    End Select

End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Private methods
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Function SetOfCommon(ByRef ipSeq As SeqHC) As SeqHC

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set SetOfCommon = myS
    
    If Me.LacksItems Or ipSeq.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As SeqHNode: Set myNode = s.HashRoot
    
    Do
        If ipSeq.HoldsItem(myNode.Item) Then
            myS.Add myNode.Item
        End If
        
        Set myNode = myNode.HashSucc
        
    Loop Until myNode Is Nothing

    Set SetOfCommon = myS

End Function


Private Function SetOfHostOnly(ByRef ipSeq As SeqHC) As SeqHC

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set SetOfHostOnly = myS
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Set myS = Me.Dedup

    If ipSeq.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As SeqHNode: Set myNode = s.HashRoot
    Do
    
        If ipSeq.HoldsItem(myNode.Item) Then
            myS.RemoveAll myNode.Item
        End If
        
        Set myNode = myNode.HashSucc
        
    Loop Until myNode Is Nothing
    
    Set SetOfHostOnly = myS
    
End Function


Private Function SetOfParamOnly(ByRef ipSeq As SeqHC) As SeqHC

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set SetOfParamOnly = myS
    
    If ipSeq.LacksItems Then
        Exit Function
    End If
    
    Set myS = ipSeq.Dedup
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As SeqHNode: Set myNode = s.HashRoot
    Do
    
        If myS.HoldsItem(myNode.Item) Then
            myS.RemoveAll myNode.Item
        End If
        
        Set myNode = myNode.HashSucc
        
    Loop Until myNode Is Nothing
    
    Set SetOfParamOnly = myS
    
End Function


Private Function SetOfNotCommon(ByRef ipSeq As SeqHC) As SeqHC

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set SetOfNotCommon = myS
    
    If Me.LacksItems And ipSeq.LacksItems Then
        Exit Function
    End If
    
    Dim myIndex As Long
    For myIndex = s.Order.FirstIndex To s.Order.Lastindex
    
        If ipSeq.LacksItem(s.Order.Item(myIndex).Item) Then
            If myS.LacksItem(s.Order.Item(myIndex).Item) Then
                myS.Add s.Order.Item(myIndex).Item
            End If
        End If
                
    Next
    
    Dim myItems As IterItems: Set myItems = IterItems(ipSeq)
    Do
        Dim myA() As Variant: myA = Array(myItems.CurItem(0))
        If Me.LacksItem(myA(0)) Then
            If myS.LacksItem(myA(0)) Then
                myS.Add myA(0)
            End If
        End If
        
    Loop While myItems.MoveNext

    Set SetOfNotCommon = myS

End Function


Private Function SetofUnique(ByRef ipSeq As SeqHC) As SeqHC

    Dim myS As SeqHC: Set myS = SeqHC.Deb
    Set SetofUnique = myS
    
    If Me.LacksItems And ipSeq.LacksItems Then
        Exit Function
    End If
    
    Set myS = Me.Dedup
    
    Dim myItems As IterItems: Set myItems = IterItems(ipSeq)
    Do
    
        If myS.LacksItem(myItems.CurItem(0)) Then
            myS.Add myItems.CurItem(0)
        End If
        
    Loop While myItems.MoveNext
    
    Set SetofUnique = myS

End Function



Private Sub SetupFirstNode(ByRef ipNode As SeqHNode)

        Set s.HashRoot = ipNode
        Set s.HashLast = s.HashRoot
        
        ' populate hash slot
        Set s.HashSlots(s.HashRoot.Hash) = s.HashRoot
        
        ' add to s.order
        ipNode.Order = 1
        s.Order.Add ipNode
        
End Sub

Private Sub AppendByHash(ByRef ipNode As SeqHNode)

    ' The append node method adds a node s.HashLast
    ' at the end of the hashlist
    Set ipNode.HashPred = s.HashLast
    Set ipNode.HashSucc = Nothing
    Set s.HashLast.HashSucc = ipNode
    Set s.HashLast = ipNode
    
End Sub

Private Sub AppendByOrder(ByRef ipNode As SeqHNode)
    
  s.Order.Add ipNode
  ipNode.Order = s.Order.Count
    
End Sub

Private Sub InsertbyHash(ByVal ipNewNode As SeqHNode)

    ' we know that the new node is not the first node.
    ' we know that the new node exists in hashslots
    ' we insert the newnode after the last node that matches Hash/Item
    
    ' get the first node for the newnode hash
    Dim myCurNode As SeqHNode: Set myCurNode = s.HashSlots(ipNewNode.Hash)

    'The new node is inserted at the end of the current hash/item combination
    ' we know that curnode and new node have the same hash
    ' we don't know if they have the same items

    Do While myCurNode.Hash = ipNewNode.Hash
        ' check item is the same item as ipnewnode
        If Comparers.NEQ(myCurNode.Item, ipNewNode.Item) Then
            Exit Do
        End If
        
        ' check the next node is not nothing
        If myCurNode.HashSucc Is Nothing Then
            ' if ipnode.succ is nothing then by default
            ' myNode is s.last so we can just append the new node
            AppendByHash ipNewNode
            Exit Sub
        Else
            Set myCurNode = myCurNode.HashSucc
        End If
    Loop
    
    ' at this point we have matching hashes but different items
    ' curnode is the node after the last item/hash match so we
    ' insert newnode before curnode
    ' are we trying to insert at the start of the hashlist
    If myCurNode.HashPred Is Nothing Then
        Set ipNewNode.HashSucc = myCurNode 'myPred.Succ
        Set myCurNode.HashPred = ipNewNode
    
    End If
    
    Set ipNewNode.HashPred = myCurNode.HashPred 'myPredNode
    Set ipNewNode.HashSucc = myCurNode 'myPred.Succ
    
    Set myCurNode.HashPred = ipNewNode
    Set ipNewNode.HashPred.HashSucc = ipNewNode
    
End Sub


Private Sub InsertByOrder(ByRef ipIndex As Long, ByRef ipNewNode As SeqHNode)

    ' The new node is not the first or last node.
    s.Order.InsertAt ipIndex, ipNewNode
    
    UpdateOrder ipIndex
   
End Sub


Private Sub PrependByOrder(ByRef ipNode As SeqHNode)

    ipNode.Order = Me.FirstIndex
    s.Order.InsertAt Me.FirstIndex, ipNode
    
    UpdateOrder 1 'ipNode.Order + 1
    
End Sub


Private Sub UpdateOrder(ByRef ipIndex As Long)

    Dim myIndex As Long
    For myIndex = ipIndex To Me.Lastindex
        With s.Order.Item(myIndex)
            .Order = myIndex
        End With
    Next
    
End Sub

'@Description("Returns the node where .Order = ipIndex")
Private Function GetNodeAt(ByRef ipIndex As Long) As Variant
Attribute GetNodeAt.VB_Description = "Returns the node where .Order = ipIndex"

    Set GetNodeAt = Nothing

    If Guard.IndexOutOfBounds(ipIndex, Me, "GetNodeAt", Guard.REPORT_BACK) Then
        Exit Function
    End If
    
    Set GetNodeAt = s.Order.Item(ipIndex)

End Function

'@Description("Returns the first node containing ipItem")
Private Function GetNodeOf(ByRef ipItem As Variant) As Variant
Attribute GetNodeOf.VB_Description = "Returns the first node containing ipItem"

    Set GetNodeOf = Nothing
    
    Dim myNode As SeqHNode: Set myNode = s.HashSlots(s.Hasher.OSHash(ipItem))
    
    If NotFound(myNode) Then
        Exit Function
    End If
    
    ' the item of the hash value may not match
    Do Until Comparers.EQ(myNode.Item, ipItem)
    
        If myNode.HashSucc Is Nothing Then
            Exit Function
        End If
        
        If myNode.Hash <> myNode.HashSucc.Hash Then
            Exit Function
        End If
        
        Set myNode = myNode.HashSucc
        
    Loop
    
    ' we now have a matching hash and item
    Set GetNodeOf = myNode

End Function


'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub RemoveNode(ByRef ipNode As SeqHNode)

    Dim myTmp As SeqHNode: Set myTmp = ipNode
    RemoveOrderNode myTmp
    RemoveHashTableNode myTmp
    RemoveHashNode myTmp

End Sub

Private Sub RemoveOrderNode(ByRef ipNode As SeqHNode)
    ' when removing a node we must check that the node is not the node registered in hashslots
    ' if it is we must replace the node pointed to by s.HashSlots(hash) with
    ' s.HashSlots.succ or nothing
    ' wemust also remember to update root/last nodes as appropriate
    
    ' remove the node from s.order
    Dim myIndex As Long: myIndex = ipNode.Order
    s.Order.RemoveAt myIndex
    UpdateOrder myIndex
    
End Sub

Private Sub RemoveHashTableNode(ByRef ipNode As SeqHNode)

    If Not ipNode Is s.HashSlots(ipNode.Hash) Then
        Exit Sub
    End If
    
    ' the node is in a hashslot so
    ' replace the node with hashsucc if hashsucc has the same hash
    ' else replace with nothing
    
    If ipNode.HashSucc Is Nothing Then
        Set s.HashSlots(ipNode.Hash) = Nothing
    ElseIf ipNode.Hash = ipNode.HashSucc.Hash Then
        Set s.HashSlots(ipNode.Hash) = ipNode.HashSucc
    Else
        Set s.HashSlots(ipNode.Hash) = Nothing
    End If
    
End Sub

'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub RemoveHashNode(ByRef ipNode As SeqHNode)
    ' we can now focus on just the hashlist
    If ipNode Is s.HashRoot Then
        RemoveFirstHashNode
    ElseIf ipNode Is s.HashLast Then
        RemoveLastHashNode
    Else
        RemoveLinkedHashNode ipNode
    End If
    
End Sub


Private Sub RemoveFirstHashNode()

    If s.HashRoot.HashSucc Is Nothing Then
        Set s.HashRoot = Nothing
    Else
        '@Ignore VariableNotUsed
        Dim myTmp As SeqHNode: Set myTmp = s.HashRoot
        Set s.HashRoot = s.HashRoot.HashSucc
        Set s.HashRoot.HashPred = Nothing
        Set myTmp = Nothing
    End If
    
End Sub

Private Sub RemoveLastHashNode()

    If s.HashLast.HashPred Is Nothing Then
        Set s.HashLast = Nothing
    Else
        '@Ignore VariableNotUsed
        Dim myTmp As SeqHNode: Set myTmp = s.HashLast
        Set s.HashLast = s.HashLast.HashPred
        Set s.HashLast.HashSucc = Nothing
        Set myTmp = Nothing
    End If
End Sub

'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub RemoveLinkedHashNode(ByRef ipNode As SeqHNode)

    Set ipNode.HashSucc.HashPred = ipNode.HashPred
    Set ipNode.HashPred.HashSucc = ipNode.HashSucc
    Set ipNode = Nothing
    
End Sub


Private Sub RemoveAllItems()

    Dim myNode As SeqHNode:  Set myNode = s.Order.Item(Me.Lastindex)
    
    Do
    
        '@Ignore VariableNotUsed
        Dim myTmpNode As SeqHNode: Set myTmpNode = myNode
        Set myNode = myNode.OrderPred
        Set myTmpNode = Nothing
        
    Loop Until myNode Is Nothing
    
    Set s.HashLast = Nothing
    Set s.HashRoot = Nothing
    
End Sub


Private Sub RemoveAllSingleItem(ByRef ipItem As Variant)

    ' all nodes with the same hash /item are linked
    ' Where there is a hash clash the node at s.hashslots(myHash) may not hold ipItem
    ' in which case we have to search forward to find the item
    ' we also have to be mindful that whilst the hash of an item may be valid (the relevant hashslot is occupied)
    ' the item itself may not be in the seq
    Dim myHash As Long: myHash = s.Hasher.OSHash(ipItem)
    Dim myNode As SeqHNode: Set myNode = s.HashSlots(myHash)
    
    'check that that hashslot is occupied
    If myNode Is Nothing Then
        Exit Sub
    End If
    
    ' search for the first matching item with the current hash
    ' also realising that becasue we have a valid hash doesn't mean that the item is in the sequence
    Do Until Comparers.EQ(myNode.Item, ipItem)
        Set myNode = myNode.HashSucc
        ' check for the case of a valid hash but no item in seq
        ' i.e. the node hash has changed without finding a matching item
        If myNode.Hash <> myHash Then
            Exit Sub
        End If
    Loop

    ' we have found a node with matching hash/item
    ' but the next noede might be nothing
    Do
    
        If myNode.HashSucc Is Nothing Then
            RemoveNode myNode
            Exit Do
        End If
        
        Dim myTmp As SeqHNode: Set myTmp = myNode
        Set myNode = myNode.HashSucc
        RemoveNode myTmp
        
    Loop While Comparers.EQ(myNode.Item, ipItem)
    
End Sub

'The following methods are used for debugging.  do not use complex keys
Public Sub Analyse()

    Dim myCount As Long: myCount = 0
    Dim myIndex As Long
    For myIndex = LBound(s.HashSlots) To UBound(s.HashSlots)
        If Not s.HashSlots(myIndex) Is Nothing Then
            myCount = myCount + 1
        End If
    Next
    
    Fmt.Dbg "Hashslots used = {0}/{1}/{2}", myCount, Me.Count, s.Hasher.MaxHash
End Sub

Private Function NotFound(ByRef ipItem As Variant) As Boolean

    If VBA.IsObject(ipItem) Then
        NotFound = ipItem Is Nothing
    ElseIf GroupInfo.IsNumber(ipItem) Then
        NotFound = ipItem = -1
    Else
        NotFound -VBA.IsEmpty(ipItem)
    End If
        
End Function
