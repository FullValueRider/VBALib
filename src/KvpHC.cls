VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@IgnoreModule
''@PredeclaredId
''@Exposed
''@Folder("Seq)
''@ModuleDescription("SeqHc is a Collection/ArrayList replacement. Items are stored in an Array/Quadruple linkied list using a hash value. Order of addition is preserved.  Indexing is 1 based. Keys are not supported (See Kvp classes)")
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'' Code line limit should be 120 characters.
'' Comment line limit should be 80 characters
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'' uses hashing algorith from cHashD from https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs)&s=f049d882b4fc4d27f4bc2f9bc383a53e
'' cHashD, Olaf Schmidt in August 2016 (enhanced about mixed, variable KeyTypes and Remove-functions in 2020)
''
'' SeqHc is a list class and consequently will store multiple items with the same
'' value.
'' SeqHc uses a a linked list accssed by a HashTable and Collection
'' hashing to allow for fast lookup of items or items by index.
'' Lack of variability in the input data will significantly degrade the
'' lookup performance of items.
''
'' Items are stored in the nodes of a linked list and references to the nodes are
'' added to a HashTable array and an order collection.  Hash clashes are managed by
'' linking nodes in the HashList to get sequences of Hash/Item pairs.
''
'' Nodes containing the first instance of a hash value
'' are stored in an array indexed by the hash value.
'' Nodes are also stored in a collection in order of addition/insertion
'' Hash clashes are managed by inserting items into the linked list at the
'' location defined by the hash. Hash clashes are inserted into the list following
'' the last ocurrence of the Hash/Item combination
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'' Items are accessed using 1 based indexing
'' Enumeration via NewEnum is not implemented (See Iteritems for a richer enumration experience)
'' Keys are not suported (see Kvp classes)
'' Preserves order of addition.
'' Insertion at a defined position is supported
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'Option Explicit
''Aug 2023 - Hashing algorith moved to a separate class - Hasher with method OSHash
'' constants common to all seq classes
'Private Const MY_TYPENAME                       As String = "SeqHC"
'Private Const FIRST_INDEX                       As Long = 1
'Private Const WHOLE_SEQ                         As Long = -1
'Private Const TO_LAST_INDEX                     As Long = 0
'Private Const DEFAULT_STEP                      As Long = 1
'Private Const LACKS_ITEMS                       As Long = -1
'Private Const INDEX_OUT_OF_BOUNDS               As Long = -1
'Private Const NOT_FOUND                         As Long = -1
'
'' constants specific to class
''@Ignore ConstantNotUsed
'Private Const MIN_SEQ_SIZE                  As Long = 128
'
'
'Private Type State
'    ' An array where nodes are stored using their hash as the index.
'    ' Each HashSlot holds a reference to a single node, which is the
'    ' first node with a particular hash value
'    ' hash clashes are inserted into the list after this first node
'    ' in hash/key precedence
'    KeysHash()                             As SeqHNode
'    ItemsHash()
'    ' first node added to the seq
'    HashRoot                                As SeqHNode
'    ' last node appended to the seq (note not added)
'    HashLast                                As SeqHNode
'
'    Hasher                                  As Hasher ' hash generator via OSHash
'
'    ' a collection which stores nodes in the order in which
'    ' they are added/inserted
'    Order                                   As SeqC
'
'End Type
'
'Private s                                   As State
'
'
'Private Type Properties
'    ' the defined first index
'    FirstIndex                              As Long
'    ' The dimmed size of s.HashSlots
'    Size                                    As Long
'End Type
'
'Private p                                   As Properties
'
'' Deb is a Factory method of the PredeclaredId which returns a new instance of SeqHc
''=================================================================================
'' The input for Deb is a ParamArray.
'' The ParamArray is interpreted according to the folowing rules
'' If the ParamArray contains no Items
'' - A empty SeqHcuence (count = 0)
'' If the ParamArray has 2 or more items
'' - An Seq of size equal to the ParamArray, items as generated by for each over the ParamArray
'' If the ParamArray has a single Item
'' - Number: Set the Hashrange to the nearest biggest power of 2
'' - String: A Seq whose capacity is the length of the string with one character per Item
'' - Array: Seq of items as provided by For Each over the array.  Multidimension arrays are flattened but the position in the original array can be calculated by the user if required. See also Bind
'' - Dictionaries: a Seq of key/value arrays as generated from for each on the Keys array.
'' - Collection: The collection is copied to the Host collection.  There is no refererence link to the original collection (See Host Property)
'' - SeqHc: Items as generated by for each on the SeqHc
'' - mscorlib Arraylist: Items as generated by for each on the arrylist
'' - mscorlib Stack: Items as generated by for each on the stack
'' - mscorlib Queue: Items as generated by for eac on the queue
'
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
''Setup and reinitialisation methods
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
''@DefaultMember
'Public Function Deb(ParamArray ipParamarray() As Variant) As SeqHC
'    With New SeqHC
'        Set Deb = .ConstructInstance(VBA.CVar(ipParamarray))
'    End With
'End Function
'
'
'Friend Function ConstructInstance(ByRef ipFPA As Variant) As SeqHC
'
'    p.FirstIndex = FIRST_INDEX
'    Set s.Order = SeqC.Deb
'
'    ' count is obtained from s.order.count
'    ' Check to see if we return a default 'empty' seq
'    If ArrayOp.LacksItems(ipFPA) Then
'        Set s.Hasher = Hasher.Deb
'        ReDim s.HashSlots(0 To s.Hasher.MaxHash)
'        Set ConstructInstance = Me
'        Exit Function
'    End If
'
'    ' ipParamArray contains one or more items depending on the input type.
'    ' if the input contains only a single item we want to use that item
'    ' else we want to treat the forwrded paramarray as a single array item
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
'    If GroupInfo.IsNumber(myItem) Then
'        Set s.Hasher = Hasher(myItem)
'        ReDim s.HashSlots(0 To s.Hasher.MaxHash)
'        Set ConstructInstance = Me
'        Exit Function
'    Else
'        Set s.Hasher = Hasher.Deb
'        ReDim s.HashSlots(0 To s.Hasher.MaxHash)
'    End If
'
'    ' Resolve the items in the forwarded paramarray using the rules
'    ' as described above.
'    Select Case True
'
'        Case VBA.IsArray(myItem):                   Set ConstructInstance = AddRange(ArrayOp.Linearise(myItem))
'        Case GroupInfo.IsContainer(myItem):         Set ConstructInstance = AddRange(myItem)
'        Case GroupInfo.IsString(myItem):            Set ConstructInstance = AddRange(myItem)
'        Case Else:                                  Set ConstructInstance = AddItems(myItem)
'    End Select
'
'End Function
'
'
''@Description("Clears all entries and resets the size of the hastable to accomodate the new size")
''@Ignore AssignedByValParameter
'Public Sub Reinit(Optional ByRef ipNewSize As Variant)
'
'    Set s.Order = SeqC.Deb
'
'    If s.Hasher Is Nothing Then
'        If VBA.IsMissing(ipNewSize) Then
'            If p.Size > 0 Then
'                Set s.Hasher = Hasher(p.Size)
'            Else
'                Set s.Hasher = Hasher.Deb  ' uses DEFAULT_HASH_TABLE_SIZE
'            End If
'        End If
'    Else
'        Set s.Hasher = Hasher(ipNewSize)
'    End If
'
'    ' Redim without preserve should mean that we have no entries in s.HashSlots
'    ReDim s.HashSlots(0 To s.Hasher.MaxHash)
'    p.Size = s.Hasher.Size
'
'    If Me.LacksItems Then
'        Exit Sub
'    End If
'
'    ' reinit is clearing the seq so we need to
'    ' delete all existing nodes
'    Dim myNode As SeqHNode: Set myNode = s.HashLast
'    Do Until myNode.HashPred Is Nothing
'
'        Set myNode = myNode.HashPred
'        Set myNode.HashSucc = Nothing
'     Loop
'
'     Set s.HashRoot = Nothing
'
'End Sub

