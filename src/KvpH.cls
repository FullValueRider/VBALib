VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Dictionary using an array of hashed keys and a linked list to preserve order of addition."
'@PredeclaredId
'@Exposed
'@Folder("Kvp")
'@ModuleDescription("Dictionary using an array of hashed keys and a linked list to preserve order of addition.")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Based on cHashD from https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs)&s=f049d882b4fc4d27f4bc2f9bc383a53e
' cHashD, Olaf Schmidt in August 2016 (enhanced about mixed, variable KeyTypes and Remove-functions in 2020)
' Reformatted and Rubberducked by Steven Laycock July 2023
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' KvpH is a Dictionary class that allows the use of duplicate keys.
' This means that hash clashes are very likely to happen
' This class uses an array to store the unique key hashes and a linked list to
' manage the hash clashes.
' the hash table stores the node containing the first hash for a particular value
' subsequent hash clashes are linked after the first node for the hash
' nodes are linked in Hash/Item sequence.
' The order of addition of values is also preserved by adding a second pair
' of node references to the node class
' KvpH allows arrays and containerobjects to be used as keys.  This use will be
' slow because these items will be stringified for any comparison purposes
' the stringified data is not chaced so will be recalculated for every comparison
' see stringifier class

Option Explicit

' Hash-Multiplikator
Private Const HMul                          As Long = 3727

Private Const MY_TYPENAME                   As String = "KvpH"
Private Const FIRST_INDEX                   As Long = 1
Private Const MIN_SEQ_SIZE                  As Long = 512
Private Type State
    ' W is used in hash calculations
    W()                                     As Integer
    ' typedef used to access the data area of W()
    saW                                     As SAFEARRAY1D
          
   ' Count                                   As Long
   ' HashSlotMask                            As Long 'represents the number of bits which can represent the positive integers that encapsulate the Nodes array size
    
    ' A mask value with all bits set to 1 which represents the number of hashslots available
    ' it is used to AND intermediate values in the the hash calculation
    ' to ensure that the hash is always within the range of available HashSlots
    ' i.e. it is always a power of 2 minus 1
    ' and consequenlty the number of hashslots can only be doubled or halved
    HashSlotMask                            As Long
    ' An array where nodes are stored using their hash as the index.
    ' Each HashSlot holds a reference to a single node, which is the first node with a particular hash
    HashSlots()                             As KvpHNode
    ' first node added to the seq
    HashRoot                                As KvpHNode
    ' last node appended to the seq (note not added)
    HashLast                                As KvpHNode
    ' Node containing order=1
    OrderRoot                               As KvpHNode
    ' Node contain the last order
    OrderLast                               As KvpHNode
    
End Type


Private s                                   As State

Private Type Properties
    EnsureUniqueKeys                        As Boolean
    Size                                    As Long ' The dimmed size of s.HashSlots
End Type


Private p                                   As Properties

'@DefaultMember
Public Function Deb(Optional ByRef ipCapacity As Long = MIN_SEQ_SIZE) As KvpH
Attribute Deb.VB_UserMemId = 0
    With New KvpH
        Set Deb = .ConstructInstance(ipCapacity)
    End With
End Function


Friend Function ConstructInstance(ByRef ipCapacity As Long) As KvpH

    s.saW.cDims = 1
    s.saW.cbElements = 2
    s.saW.cLocks = 1
    s.saW.fFeatures = &H11                       'FADF_AUTO=&H1 || FADF_FIXEDSIZE=&H10
    BindArray s.W, VarPtr(s.saW)
    SetSize ipCapacity
    p.EnsureUniqueKeys = True
    Set ConstructInstance = Me
    
End Function


'The printby methods are used for debugging.  do not use complex keys
Public Sub PrintByHash()

    Dim myNode As KvpHNode: Set myNode = s.HashRoot
    Do
        Debug.Print myNode.Hash, myNode.Order, myNode.Key, myNode.Item
        Set myNode = myNode.HashSucc
    Loop Until myNode Is Nothing

End Sub

Public Sub PrintByOrder()

    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
        Debug.Print myNode.Hash, myNode.Order, myNode.Key, myNode.Item
        Set myNode = myNode.OrderSucc
    Loop Until myNode Is Nothing

End Sub
'@Description("set the size of hash array to be the 2^n value just greater than ipNewSize")
'@Ignore AssignedByValParameter
Public Sub SetSize(Optional ByVal ipNewSize As Long)
Attribute SetSize.VB_Description = "set the size of hash array to be the 2^n value just greater than ipNewSize"

    If ipNewSize < 0 Then
        ipNewSize = MIN_SEQ_SIZE
    End If

    'calculate a long whose value is the first 2^n-1 that is greater than ipNewSize
    ' this provides a bitfield of 1's used to set the maximum value for a hash value
    s.HashSlotMask = (2 ^ Int((VBA.Log(ipNewSize) / VBA.Log(2)) + 1)) - 1
    If s.HashSlotMask < ipNewSize Then
        s.HashSlotMask = s.HashSlotMask * 2 + 1
    End If
    'Debug.Print Hex$(ipNewSize), Hex$(s.hashslotmask)
    
    ' Redim without preserve should mean that we have no entries in s.HashSlots
    ReDim s.HashSlots(FIRST_INDEX To s.HashSlotMask + 1)
    p.Size = s.HashSlotMask
    If s.HashRoot Is Nothing Then
        Exit Sub
    End If
    
    ' update the hashes in the nodes
    ' then populate the new s.modes using the new hashes
    ' this does not change the order of addition,
    
    Dim myNode As KvpHNode: Set myNode = s.HashRoot
    Do
        myNode.Hash = GetHash(myNode.Key)
        Set myNode = myNode.HashSucc
    Loop Until myNode Is Nothing
    
    ' repopulate s.HashSlots according to the new hashes
    ' the order of nodes is unchanged by the reallocation of hashslots
    ' the rootnode should be unchanged by the rehashing
    ' s.HashLast node will have been updated by the rehashing
    
    Set s.HashSlots(s.HashRoot.Hash) = s.HashRoot
    Set s.HashLast = s.HashRoot
    Set myNode = s.HashRoot
    
    ' linking nodes in accordance with the hash and item
    ' is done by the append and inset methods so
    ' here we only need to focus on the hash value
    Do Until myNode.HashSucc Is Nothing
    
        Set myNode = myNode.HashSucc
        
        If s.HashSlots(myNode.Hash) Is Nothing Then
            Set s.HashSlots(myNode.Hash) = myNode
            AppendByHash myNode
        Else
            InsertbyHash myNode
        End If
        
    Loop
    
End Sub


Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpH

    Set Add = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Add", _
            "Null is not a valid Key"
    End If
    
    If p.EnsureUniqueKeys Then
        If Me.HoldsKey(ipKey) Then
            Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Add", _
            Fmt.Text("Duplicate key: {0}", ipKey)
        End If
    End If

    Dim myHash As Long: myHash = GetHash(ipKey)
    Dim myNewNode As KvpHNode: Set myNewNode = KvpHNode(myHash, ipKey, ipItem)
    
    If s.OrderRoot Is Nothing Then
    
       SetupFirstNode myNewNode
     
    'we are not the first key of this hash
    ElseIf s.HashSlots(myHash) Is Nothing Then
'        ' append node to s.last
'        AppendByHash KvpHNode(myHash, ipKey, ipItem)
'
'        Set Add = Me
'        Exit Function
'    ElseIf s.HashSlots(myHash) Is Nothing Then
    
        ' if thed hash slot is nothing then we just append
        Set s.HashSlots(myHash) = myNewNode
        AppendByHash myNewNode
        AppendByOrder myNewNode
        
        
    Else
        
        'Insert ipNewNode in the node list at the end of the hash/item sequence
        ' append the new node to the order seq
        InsertbyHash myNewNode
        AppendByOrder myNewNode
        
    End If
   
    Set Add = Me
    
End Function


Public Function AddPairs(ByRef ipKeys As Variant, ByRef ipItems As Variant) As KvpH
   
    Dim myKeys As IterItems: Set myKeys = IterItems(ipKeys)
    Dim myItems As IterItems: Set myItems = IterItems(ipItems)
    
    If myKeys.LacksItems Or myItems.LacksItems Then
        Set AddPairs = Me
        Exit Function
    End If
     
    Do
    
        If VBA.IsNull(myKeys.CurItem(0)) Then
            Err.Raise 17 + vbObjectError, _
                "VBALib.KvpH.AddPairs", _
                Fmt.Text("At Index '{0}': Null is not a valid Key", myKeys.CurKey(0))
        End If
    
        If p.EnsureUniqueKeys Then
    
            If Me.HoldsKey(myKeys.CurItem(0)) Then
                ' As a convenience ignore duplicate keys unless a warning is specifically requested
                Err.Raise 17 + vbObjectError, _
                    "VBALib.KvpH.AddPairs", _
                    Fmt.Text("Duplicate key: Got {0}", myKeys.CurItem(0))
            End If
        End If
        
        Me.Add myKeys.CurItem(0), myItems.CurItem(0)
          
    Loop While myKeys.MoveNext And myItems.MoveNext
   
    Set AddPairs = Me
   
End Function


Public Property Get Item(ByRef ipKey As Variant) As Variant

    Item = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Property
    End If
                   
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_Item", _
            "Null is not a valid Key"
    End If
    
'    If p.EnsureUniqueKeys Then
'
'        If Me.HoldsKey(ipKey) Then
'            ' As a convenience ignore duplicate keys unless a warning is specifically requested
'            Err.Raise 17 + vbObjectError, _
'                "VBALib.KvpH.AddPairs", _
'                Fmt.Text("Duplicate key: Got {0}", ipKey)
'        End If
'    End If
    
    Dim myNode As KvpHNode: Set myNode = GetNodeOf(ipKey)
    
    If myNode Is Nothing Then
        Exit Function
    ElseIf VBA.IsObject(myNode.Item) Then
        Set Item = myNode.Item
    Else
        Item = myNode.Item
    End If
    
End Property


Public Property Let Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_Item", _
            "Null is not a valid Key"
    End If

    Dim myNode As KvpHNode: Set myNode = GetNodeOf(ipKey)
    
    If myNode Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_Item", _
            Fmt.Text("No such Key: '{0}'", ipKey)
    End If
    
    myNode.Item = ipItem
    
End Property


Public Property Set Item(ByRef ipKey As Variant, ByVal ipItem As Variant)
 
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Set_Item", _
            "Null is not a valid Key"
    End If
    
    Dim myNode As KvpHNode: Set myNode = GetNodeOf(ipKey)
    
    If myNode Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Set_Item", _
            Fmt.Text("Key {0} not found", ipKey)
    End If
    
    Set myNode.Item = ipItem
    
End Property


Public Property Get ItemAt(ByVal ipIndex As Long) As Variant

    If ipIndex < s.OrderRoot.Order Or ipIndex > s.OrderLast.Order Then
        ItemAt = Null
        Exit Property
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do Until myNode.Order = ipIndex
        Set myNode = myNode.OrderSucc
    Loop
    
    If VBA.IsObject(myNode.Item) Then
        Set ItemAt = myNode.Item
    Else
        ItemAt = myNode.Item
    End If
    
End Property


Public Property Let ItemAt(ByVal ipIndex As Long, ByVal ipItem As Variant)

    If ipIndex < s.OrderRoot.Order Or ipIndex > s.OrderLast.Order Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_ItemAt", _
            Fmt.Text("Bounds are {0} to {1}. Got index of {2}", s.OrderRoot.Order, s.OrderLast.Order, ipIndex)
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    
    Do Until myNode.Order = ipIndex
        Set myNode = myNode.OrderSucc
    Loop
    
    myNode.Item = ipItem
   
End Property


Public Property Set ItemAt(ByVal ipIndex As Long, ByVal ipItem As Variant)

    If ipIndex < s.OrderRoot.Order Or ipIndex > s.OrderLast.Order Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_ItemAt", _
            Fmt.Text("Bounds are {0} to {1}. Got index of {2}", s.OrderRoot.Order, s.OrderLast.Order, ipIndex)
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do Until myNode.Order = ipIndex
        Set myNode = myNode.OrderSucc
    Loop
    
    Set myNode.Item = ipItem
    
End Property

'@Description("Removes the first found key/item pair for each key in the paramarray")
Public Function Remove(ParamArray ipParamarray() As Variant) As KvpH
Attribute Remove.VB_Description = "Removes the first found key/item pair for each key in the paramarray"

    Set Remove = Me
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If
    
    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(VBA.CVar(ipParamarray))
    
    If Not VBA.IsArray(myParamarray) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Remove", _
            Fmt.Text("Expecting an array.  Got {0}", VBA.TypeName(myParamarray))
    End If
    
    Dim myKey As Variant
    For Each myKey In myParamarray
          
        If VBA.IsNull(myKey) Then
            Err.Raise 17 + vbObjectError, _
                "VBALib.KvpH.Remove", _
                "Null is not a valid Key"
        End If
    
        Dim myNode As KvpHNode: Set myNode = GetNodeOf(myKey)
        If Not myNode Is Nothing Then
            RemoveNode myNode
        End If
    Next
    
    Set Remove = Me
    
End Function

'@Description("Removes the first found key/item pair for each key in the range parameter")
Public Function RemoveRange(ByRef ipRange As Variant) As KvpH
Attribute RemoveRange.VB_Description = "Removes the first found key/item pair for each key in the range parameter"

    Set RemoveRange = Me
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If
    
    Dim myKeys As IterItems: Set myKeys = IterItems(ipRange)
    
    If myKeys.LacksItems Then
        Exit Function
    End If
    
    Do
    
        If VBA.IsNull(myKeys.CurItem(0)) Then
            Err.Raise 17 + vbObjectError, _
                "VBALib.KvpH.Remove", _
                "Null is not a valid Key"
        End If
        
        RemoveNode GetNodeOf(myKeys.CurItem(0))
        
    Loop While myKeys.MoveNext
    
    Set RemoveRange = Me
    
End Function

Public Function RemoveAll(ParamArray ipParamarray() As Variant) As KvpH

    Set RemoveAll = Me
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If
    
    If ArrayOp.LacksItems(VBA.CVar(ipParamarray)) Then
        RemoveAllItems
        Exit Function
    End If
    
    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(VBA.CVar(ipParamarray))
  
    
    Dim myKey As Variant
    For Each myKey In myParamarray
        RemoveAllSingleItem myKey
    Next
    
    Set RemoveAll = Me

End Function

Public Function Clear() As KvpH
    Me.RemoveAll
    Set Clear = Me
End Function


Public Function Reset() As KvpH
    Me.RemoveAll
    Set Reset = Me
End Function

Public Function RemoveAt(ByVal ipIndex As Long) As Long

    RemoveAt = -1
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If
    
    If ipIndex < s.OrderLast.Order Or ipIndex > s.OrderLast.Order Then
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = GetNodeAt(ipIndex)
    
    RemoveNode myNode
    RemoveAt = ipIndex
        
End Function

Public Function RemoveIndexes(ParamArray ipParamarray() As Variant) As KvpH

    Dim myParamarray As Variant: myParamarray = ipParamarray
    myParamarray = ArrayOp.Splat(myParamarray)
    
    Set RemoveIndexes = RemoveIndexesRange(myParamarray)
    
    Set RemoveIndexes = Me
    
End Function

'@Description("Removes the Items at the indexes in the range list")
Public Function RemoveIndexesRange(ByRef ipRange As Variant) As KvpH
Attribute RemoveIndexesRange.VB_Description = "Removes the Items at the indexes in the range list"

    Set RemoveIndexesRange = Me
    
    If s.HashRoot Is Nothing Then
        Exit Function
    End If
    
    Dim myS As SeqA: Set myS = SeqA(ipRange).Sorted
    
    If myS.LacksItems Then
        Exit Function
    End If
    
    ' we iterate backwards so we don't so we don't compromise indexes yet to be removed
    'Dim myIndex As Long
    Dim myIndexes As IterItems: Set myIndexes = IterItems(myS).MoveToEndIndex
    Do
        Dim myNode As KvpHNode: Set myNode = GetNodeAt(myIndexes.CurItem(0))
        RemoveNode myNode
    Loop While myIndexes.MovePrev
    
    Set RemoveIndexesRange = Me
    
End Function

'@Description("Remove ipRun key/value pairs after the last of  Hash/Key pairs
Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpH

    Set RemoveAfter = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.RemoveAfter", _
            "Null is not a valid Key"
    End If
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
            "VALib.KvpH.RemoveAfter", _
            Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
    End If
    
    Dim myNode As KvpHNode: Set myNode = GetNodeOf(ipKey)
    
    If myNode Is Nothing Then
        Exit Function
    End If
    
   ' Move to the end of Hash/Key pairs.  We only need to look at the keys at this stage
    Do Until myNode.HashSucc Is Nothing
    
        If Not ComparerHelpers.Equals(myNode.HashSucc.Key, ipKey) Then
            Exit Do
        End If
        Set myNode = myNode.HashSucc
        
    Loop
    
    ' we are now at the first node with a different hash/keypair
    Dim myCount As Long
    For myCount = 1 To ipRun
    
        If myNode.HashSucc Is Nothing Then
            Exit Function
        End If
        
        RemoveNode myNode.HashSucc
        
    Next
    
End Function

'@Description("Remove key/item pairs before the specified Key
Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpH

    Set RemoveBefore = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.RemoveBefore", _
            "Null is not a valid Key"
    End If
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
            "VALib.KvpH.RemoveAfter", _
            Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
    End If
    
    Dim myNode As KvpHNode: Set myNode = GetNodeOf(ipKey)
    
    If myNode Is Nothing Then
        Exit Function
    End If
    
    If myNode.HashPred Is Nothing Then
        Exit Function
    End If
    
    ' we are now at the first node if ipKey so we need to move
    ' to the pred
    '@Ignore AssignmentNotUsed
    'Set myNode = myNode.HashPred
    
    Dim myCount As Long
    For myCount = 1 To ipRun
    
        If myNode.HashPred Is Nothing Then
            Exit Function
        End If
        
        RemoveNode myNode.HashPred
        
    Next
    
End Function

Public Function InsertAt(ByRef ipIndex As Long, ByVal ipKey As Variant, ByVal ipItem As Variant) As Long

    If ipIndex < s.OrderRoot.Order Or ipIndex > s.OrderLast.Order Then
        InsertAt = -1
        Exit Function
    End If
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.InsertAt", _
            " Null is not a valid Key"
    End If
    
    If p.EnsureUniqueKeys Then
        If Me.HoldsKey(ipKey) Then
            Err.Raise 17 + vbObjectError, _
                "VBALib.KvpH.InsertAt", _
                Fmt.Text(" Duplicate Key. Got {0}", ipKey)
        End If
    End If

    ' create the new node
    Dim myHash As Long: myHash = GetHash(ipItem)
    Dim myNewNode As KvpHNode: Set myNewNode = KvpHNode.Deb(myHash, ipItem, ipOrder:=ipIndex)
    
    ' there are two strands to update, byHash and ByOrder
    ' deal with byHash first
    If s.HashSlots(myHash) Is Nothing Then
        AppendByHash myNewNode
    Else
        InsertbyHash myNewNode
    End If
    
    ' byorder
    If ipIndex = s.OrderRoot.Order Then
        PrependByOrder myNewNode
    Else
    'if ipIndex = s.OrderLast.Order Then
    'End If
    '     AppendByOrder myNewNode
    ' Else
        InsertByOrder myNewNode
    End If
        
    's.OrderLast.Order = s.OrderLast.Order + 1
    InsertAt = myNewNode.Order
    
End Function
Public Function Clone() As KvpH

    Dim myK As KvpH: Set myK = KvpH.Deb
    
    If s.HashRoot Is Nothing Then
        Set Clone = myK
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        myK.Add myNode.Key, myNode.Item
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing
    
    Set Clone = myK
    
End Function


Public Function Exists(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Exists", _
            "Null is not a valid Key"
    End If

    Exists = False
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If
    
    Dim myH As Long: myH = GetHash(ipKey)
    If s.HashSlots(myH) Is Nothing Then
        Exit Function
    End If
    
    ' found a matching hash, now check for matching key
    Dim myNode As KvpHNode: Set myNode = s.HashSlots(myH)
    Do Until myNode.Hash <> myH
        If ComparerHelpers.Equals(myNode.Key, ipKey) Then
            Exists = True
            Exit Function
        End If
        
        Set myNode = myNode.OrderSucc
        If myNode Is Nothing Then
            Exit Function
        End If
    
    Loop
    
End Function


Public Function HoldsKey(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.HoldsKey", _
            "Null is not a valid Key"
    End If

    If s.OrderRoot Is Nothing Then
        HoldsKey = False
        Exit Function
    End If
    
    HoldsKey = Me.Exists(ipKey)
    
End Function


Public Function ContainsKey(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.ContainsKey", _
            "Null is not a valid Key"
    End If

    If s.OrderRoot Is Nothing Then
        ContainsKey = False
        Exit Function
    End If
    
    ContainsKey = Me.Exists(ipKey)
    
End Function


Public Function LacksKey(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.LacksKey", _
            "Null is not a valid Key"
    End If

    If s.OrderRoot Is Nothing Then
        LacksKey = True
        Exit Function
    End If
    
    LacksKey = Not Exists(ipKey)
    
End Function


Public Function ExistsItem(ByRef ipItem As Variant) As Boolean

    If s.OrderRoot Is Nothing Then
        ExistsItem = False
        Exit Function
    End If

    ExistsItem = Me.HoldsItem(ipItem)
    
End Function


Public Function HoldsItem(ByRef ipItem As Variant) As Boolean

    If s.OrderRoot Is Nothing Then
        HoldsItem = False
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.HashRoot
    Do
        If ComparerHelpers.Equals(myNode.Item, ipItem) Then
            HoldsItem = True
            Exit Function
        End If
        Set myNode = myNode.HashSucc
    Loop Until myNode Is Nothing
    
End Function


Public Function Contains(ByRef ipItem As Variant) As Boolean
    
    If s.OrderRoot Is Nothing Then
        Contains = False
        Exit Function
    End If
    
    Contains = Me.HoldsItem(ipItem)
    
End Function


Public Function LacksItem(ByRef ipItem As Variant) As Boolean

    If s.OrderRoot Is Nothing Then
        LacksItem = True
        Exit Function
    End If

    LacksItem = Not HoldsItem(ipItem)
End Function


Public Function HoldsItems() As Boolean
    HoldsItems = Not s.OrderRoot Is Nothing
End Function


Public Function LacksItems() As Boolean
    LacksItems = s.OrderRoot Is Nothing
End Function


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property


'@Description("Mutates the KvpA in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As KvpH
Attribute MappedIt.VB_Description = "Mutates the KvpA in place by applying the function in IMapper to each item"

    If s.OrderRoot Is Nothing Then
        Set MappedIt = Me
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        Dim myItem As Variant
        myItem = ipMapper.ExecMapper(myNode.Item)
        
        If VBA.IsObject(myItem(0)) Then
            Set myNode.Item = myItem(0)
        Else
            myNode.Item = myItem(0)
        End If
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing
   
    Set MappedIt = Me

End Function


'@Description("Returns a KvpA.Deb by applying the function in IMapper to each Item")
Public Function MapIt(ByVal ipMapper As IMapper) As KvpH
Attribute MapIt.VB_Description = "Returns a KvpA.Deb by applying the function in IMapper to each Item"

    If s.OrderRoot Is Nothing Then
        Set MapIt = Me
        Exit Function
    End If

    Dim myK As KvpH: Set myK = KvpH.Deb
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        Dim myItem As Variant
        myItem = ipMapper.ExecMapper(myNode.Item)
        
        myK.Add myNode.Key, myItem(0)
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing
      
    Set MapIt = myK

End Function


'@Description("Returns a KvpA.Deb where the ICOmparer has been applied to the Items of the original Kvp")
Public Function FilterIt(ByVal ipComparer As IComparer) As KvpH
Attribute FilterIt.VB_Description = "Returns a KvpA.Deb where the ICOmparer has been applied to the Items of the original Kvp"

    Dim myK As KvpH: Set myK = KvpH.Deb
    
    If s.OrderRoot Is Nothing Then
        Set FilterIt = myK
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        If ipComparer.ExecCmp(myNode.Item) Then
            myK.Add myNode.Key, myNode.Item
        End If
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing
      
    Set FilterIt = myK

End Function


Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant

    If s.OrderRoot Is Nothing Then
        ReduceIt = Null
        Exit Function
    End If
    
    Dim myItem As Variant
    myItem = Array(ipReducer.ExecReduction(Me.Items))
    
    If VBA.IsObject(myItem(0)) Then
        Set ReduceIt = myItem(0)
    Else
        ReduceIt = myItem(0)
    End If
    
End Function


'@Description("Returns the size of the Kvp or the number of specified Items.  Returns -1 rather than 0 if there are no items in the kvp but count item will return 0 if not found")
Public Function Count(Optional ByRef ipItem As Variant) As Long
Attribute Count.VB_Description = "Returns the size of the Kvp or the number of specified Items.  Returns -1 rather than 0 if there are no items in the kvp but count item will return 0 if not found"
    
    If s.HashRoot Is Nothing Then
        Count = -1
        Exit Function
    End If
    
    If VBA.IsMissing(ipItem) Then
        Count = s.OrderLast.Order
        Exit Function
    End If
    
    Dim myCount As Long: myCount = 0
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        If ComparerHelpers.Equals(ipItem, myNode.Item) Then
            myCount = myCount + 1
        End If
        
        Set myNode = myNode.OrderSucc
    
    Loop Until myNode Is Nothing
    
    Count = myCount
    
End Function


Public Function CountIt(ByVal ipComparer As IComparer) As Long

    If s.OrderRoot Is Nothing Then
        CountIt = -1
        Exit Function
    End If

    Dim myCount As Long: myCount = 0
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        If ipComparer.ExecCmp(myNode.Item) Then
            myCount = myCount + 1
        End If
    
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing

    CountIt = myCount

End Function


Public Function KeyByIndex(ByVal ipIndex As Long) As Variant

    KeyByIndex = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If
    
    If ipIndex < s.OrderRoot.Order Or ipIndex > s.OrderLast.Order Then
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do While myNode.Order < ipIndex
        Set myNode = myNode.OrderSucc
    Loop
    
    If VBA.IsObject(myNode.Key) Then
        Set KeyByIndex = myNode.Key
    Else
        KeyByIndex = myNode.Key
    End If
    
End Function


Public Function KeyOf(ByVal ipItem As Variant) As Variant

    KeyOf = Null
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        If ComparerHelpers.Equals(myNode.Item, ipItem) Then
            If VBA.IsObject(myNode.Key) Then
                Set KeyOf = myNode.Key
            Else
                KeyOf = myNode.Key
            End If
            
            Exit Function
            
        End If
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing
    
End Function


Public Function HasUniqueKeys() As Boolean
    HasUniqueKeys = p.EnsureUniqueKeys
End Function


Public Function ForbidSameKeys() As KvpH
    p.EnsureUniqueKeys = True
    Set ForbidSameKeys = Me
End Function


Public Function AllowSameKeys() As KvpH
    p.EnsureUniqueKeys = False
    Set AllowSameKeys = Me
End Function


Public Function Items() As Variant

    If s.HashRoot Is Nothing Then
        Items = Array()
        Exit Function
    End If

    Dim myA As Variant: ReDim myA(s.OrderRoot.Order To s.OrderLast.Order)

    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
        
        If VBA.IsObject(myNode.Item) Then
            Set myA(myNode.Order) = myNode.Item
        Else
            myA(myNode.Order) = myNode.Item
        End If

        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing

    Items = myA
    
End Function

Public Function ItemsAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    
    If s.HashRoot Is Nothing Then
        Set ItemsAsColl = myC
        Exit Function
    End If

    

    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
        
        myC.Add myNode.Item
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing

    Set ItemsAsColl = myC
    
End Function

Public Function ItemsAsSeq() As SeqH
    
    Dim myS As SeqH: Set myS = SeqH.Deb(p.Size)

    If s.OrderRoot Is Nothing Then
        Set ItemsAsSeq = myS
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        myS.Add myNode.Item
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing
    
    Set ItemsAsSeq = myS
    
End Function


Public Function Keys() As Variant
    
    If s.HashRoot Is Nothing Then
        Keys = Array()
        Exit Function
    End If

    Dim myA As Variant: ReDim myA(s.OrderRoot.Order To s.OrderLast.Order)

    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
        
        If VBA.IsObject(myNode.Item) Then
            Set myA(myNode.Order) = myNode.Key
        Else
            myA(myNode.Order) = myNode.Key
        End If

        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing

    Keys = myA
    
End Function

Public Function KeysAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    
    If s.HashRoot Is Nothing Then
        Set KeysAsColl = myC
        Exit Function
    End If

    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
        
        myC.Add myNode.Key
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing

    Set KeysAsColl = myC
    
End Function

'@Description("Returns a seq of keys.")
Public Function KeysAsSeq() As SeqH
Attribute KeysAsSeq.VB_Description = "Returns a seq of keys."

    Dim myS As SeqH: Set myS = SeqH.Deb(p.Size)
    
    If s.OrderRoot Is Nothing Then
        Set KeysAsSeq = myS
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.HashRoot
    Do
    
        myS.Add myNode.Key

        Set myNode = myNode.HashSucc
        
    Loop Until myNode Is Nothing
    
    Set KeysAsSeq = myS
    
End Function


Public Function Pairs() As Variant

    If s.OrderRoot Is Nothing Then
        Pairs = Array()
        Exit Function
    End If
    
    Dim myA As Variant: ReDim myA(s.OrderRoot.Order To s.OrderLast.Order)
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        myA(myNode.Order) = Array(myNode.Key, myNode.Item)
        
        Set myNode = myNode.OrderSucc

    Loop Until myNode Is Nothing
    
    Pairs = myA
    
End Function

Public Function PairsAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    
    If s.HashRoot Is Nothing Then
        Set PairsAsColl = myC
        Exit Function
    End If

    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
        
        myC.Add Array(myNode.Key, myNode.Item)
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing

    Set PairsAsColl = myC
    
End Function

Public Function PairsAsSeq() As SeqH

    Dim myS As SeqH: Set myS = Me.KeysAsSeq

    If myS.LacksItems Then
        Set PairsAsSeq = myS
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
    
        myS.Add Array(myNode.Key, myNode.Item)
        
        Set myNode = myNode.OrderSucc
    
    Loop Until myNode Is Nothing
    
    Set PairsAsSeq = myS

End Function


Public Property Get FirstIndex() As Long
    
    If s.HashRoot Is Nothing Then
        FirstIndex = -1
    Else
        FirstIndex = FIRST_INDEX
    End If
End Property


Public Property Get LastIndex() As Long

    If s.HashRoot Is Nothing Then
        LastIndex = -1
    Else
        LastIndex = s.OrderLast.Order
    End If
End Property

'@Description("Returns the index of the second item in the kvp or -1  if the kvp has less than 2 Items")
Public Function FBOIndex() As Long
Attribute FBOIndex.VB_Description = "Returns the index of the second item in the kvp or -1  if the kvp has less than 2 Items"

    If s.OrderRoot Is Nothing Then
        FBOIndex = -1
    ElseIf s.OrderRoot.OrderSucc Is Nothing Then
        FBOIndex = -1
    Else
        FBOIndex = FIRST_INDEX + 1
    End If
    
End Function


'Description("Returns the index of the second to last item in the SeqH, or -1 if the SeqH has lees than 2 items")
Public Function LBOIndex() As Long

    If s.OrderRoot Is Nothing Then
        LBOIndex = -1
    ElseIf s.OrderLast.OrderPred Is Nothing Then
        LBOIndex = -1
    Else
        LBOIndex = s.OrderLast.OrderPred.Order
    End If
   
End Function

Public Property Get FirstKey() As Variant

    FirstKey = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.OrderRoot.Key) Then
        Set FirstKey = s.OrderRoot.Key
    Else
        FirstKey = s.OrderRoot.Key
    End If
    
End Property


Public Function LastKey() As Variant

    LastKey = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If
    
    If VBA.IsObject(s.OrderLast.Key) Then
        Set LastKey = s.OrderLast.Key
    Else
        LastKey = s.OrderLast.Key
    End If
    
End Function


Public Property Get First() As Variant

    First = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.OrderRoot.Item) Then
        Set First = s.OrderRoot.Item
    Else
        First = s.OrderRoot.Item
    End If
    
End Property


Public Property Set First(ByVal ipItem As Variant)
    
    If s.OrderRoot Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_FirstItem", _
            "Kvp has no Keys"
    End If
    
    Set s.OrderRoot.Item = ipItem

End Property


Public Property Let First(ByVal ipItem As Variant)

    If s.OrderRoot Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_FirstItem", _
            "Kvp has no Keys"
    End If
    
    s.OrderRoot.Item = ipItem

End Property


Public Property Get Last() As Variant

    Last = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.OrderLast.Item) Then
        Set Last = s.OrderLast.Item
    Else
        Last = s.OrderLast.Item
    End If
    
End Property


Public Property Set Last(ByVal ipItem As Variant)

    If s.OrderRoot Is Nothing Then
        Err.Raise 17, _
            "VBALib.KvpL.Set LastItem", _
            "VBALib.Kvp has no Keys"
    End If

    Set s.OrderLast.Item = ipItem
    
End Property


Public Property Let Last(ByVal ipItem As Variant)

    If s.OrderRoot Is Nothing Then
        Err.Raise 17, _
            "VBALib.KvpL.Let LastItem", _
            "VBALib.Kvp has no Keys"
    End If
    
    s.OrderLast.Item = ipItem
    
End Property

Public Property Get FBOKey() As Variant

    FBOKey = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Property
    ElseIf s.OrderRoot.OrderSucc Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.OrderRoot.Key) Then
        Set FBOKey = s.OrderRoot.Key
    Else
        FBOKey = s.OrderRoot.Key
    End If
    
End Property


Public Property Get LBOKey() As Variant

    LBOKey = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Property
    ElseIf s.OrderRoot.OrderPred Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.OrderLast.OrderPred.Key) Then
        Set LastKey = s.OrderLast.OrderPred.Key
    Else
        LastKey = s.OrderLast.OrderPred.Key
    End If
    
End Property


Public Property Get FBO() As Variant

    FBO = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Property
    ElseIf s.OrderLast.OrderPred Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.OrderLast.OrderPred.Item) Then
        Set First = s.OrderLast.OrderPred.Item
    Else
        First = s.OrderLast.OrderPred.Item
    End If
    
End Property


Public Property Set FBO(ByVal ipItem As Variant)
    
    If s.OrderRoot Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_FBO", _
            "Kvp has no Keys"
    ElseIf s.OrderRoot.OrderSucc Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_FBO", _
            "Kvp has no FBO Key"
    Else
        Set s.OrderRoot.OrderSucc.Item = ipItem
    End If

End Property


Public Property Let FBO(ByVal ipItem As Variant)

    If s.OrderRoot Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_FBO", _
            "Kvp has no Keys"
    ElseIf s.OrderRoot.OrderSucc Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_FBO", _
            "Kvp has FBO Key"
    Else
        s.OrderRoot.OrderSucc.Item = ipItem
    End If

End Property


Public Property Get LBO() As Variant

    FBO = Null
    
    If s.OrderRoot Is Nothing Then
        Exit Property
    ElseIf s.OrderLast.OrderPred Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.OrderLast.Item) Then
        Set LBO = s.OrderLast.OrderPred.Item
    Else
        LBO = s.OrderLast.OrderPred.Item
    End If
    
End Property


Public Property Set LBO(ByVal ipItem As Variant)

    If s.OrderRoot Is Nothing Then
        Err.Raise 17, _
            "VBALib.KvpH.Set LBO", _
            "VBALib.Kvp has no Keys"
    ElseIf s.OrderLast.OrderPred Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_LBO", _
            "Kvp has LBO Key"
    Else
        Set s.OrderLast.OrderPred.Item = ipItem
    End If
    
End Property


Public Property Let LBO(ByVal ipItem As Variant)

   If s.OrderRoot Is Nothing Then
        Err.Raise 17, _
            "VBALib.KvpH.Set LBO", _
            "VBALib.Kvp has no Keys"
    ElseIf s.OrderLast.OrderPred Is Nothing Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_LBO", _
            "Kvp has LBO Key"
    Else
        s.OrderLast.OrderPred.Item = ipItem
    End If
    
End Property
'@Description( returns a Kvp with values as keys.  Each Value is linked to a Kvp of Key vs Index for the host value
Public Function Mirror(Optional ByVal ipMirrorType As e_MirrorType = e_MirrorType.m_ByAllValues) As KvpH

    If Me.LacksItems = 0 Then
        Set Mirror = KvpH.Deb
        Exit Function
    End If

    Select Case ipMirrorType

        Case m_ByAllValues:             Set Mirror = MirrorByAllValues
        Case m_ByFirstValue:            Set Mirror = MirrorByFirstValue
        Case Else
            Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Mirror", _
            Fmt.Text("Mirror enum not known. Got {0}", ipMirrorType)
                
    End Select

End Function


'@Description("Returns an KvpH of two KvpH.  Item(1) is an KvpH of the first found Host value vs Host Key. .Item(2) is a copy of the original KvpL with the  first Key/Value pairs removed
Private Function MirrorByAllValues() As KvpH

    Dim myK As KvpH
    Set myK = KvpH.Deb
    
    If s.OrderRoot Is Nothing Then
        Set MirrorByAllValues = KvpH.Deb
        Exit Function
    End If

    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
        If myK.LacksKey(myNode.Item) Then
            myK.Add myNode.Item, KvpL.Deb
        End If

        myK.Item(myNode.Item).Add myNode.Item, myNode.Order
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing

    Set MirrorByAllValues = myK

End Function


Private Function MirrorByFirstValue() As KvpH

    Dim myFirst As KvpH
    Set myFirst = KvpH.Deb
    
    Dim myOther As KvpH
    Set myOther = KvpH.Deb
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do

        If myFirst.LacksKey(myNode.Item) Then
            myFirst.Add myNode.Item, myNode.Key
        Else
            myOther.Add myNode.Key, myNode.Item
        End If

    Loop Until myNode Is Nothing

    Set MirrorByFirstValue = KvpH.Deb.Add(0&, myFirst).Add(1&, myOther)

End Function




Public Function IndexOfKey(ByRef ipKey As Variant) As Long

    If s.OrderRoot Is Nothing Then
        IndexOfKey = -1
        Exit Function
    End If
    
    Dim myH As Long: myH = GetHash(ipKey)
    
    If s.HashSlots(myH) Is Nothing Then
        IndexOfKey = -1
        Exit Function
    End If
    
    IndexOfKey = s.HashSlots(myH).Order
    
End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'Private functions
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
Private Function GetHash(ByRef ipKey As Variant) As Long

    Dim D As Double
    Dim L As LongPtr
    Dim F As Single
    Dim VT As VbVarType
    Dim myHash As Long
     
    Dim C As Variant
    C = VBA.CDec(C)
     
    
    myHash = s.HashSlotMask                              'init the HashValue (all bits to 1)
    Dim myKey As Variant
    If GroupInfo.IsContainer(ipKey) Or GroupInfo.IsItemObject(ipKey) Or GroupInfo.IsAdmin(ipKey) Then
        myKey = GetItemAsString(ipKey)
    Else
        myKey = ipKey
    End If


    VT = VarType(myKey)
    Select Case VT
        Case vbString
        
            Dim myIndex As Long
            For myIndex = 1 To VBA.Len(myKey)
                
                Dim myChar As Long
                myChar = VBA.AscW(VBA.Mid$(myKey, myIndex, 1))
                myHash = (myHash + myChar) * HMul And s.HashSlotMask
                
            Next
                   
            
'        Case vbObject
'
'            L = ObjPtr(Key)
'            s.saW.cElements1D = 2
'            s.saW.pvData = VarPtr(L)
'
'            myHash = (myHash + s.W(0)) * HMul And s.hashslotmask 'loop-unrolling (we have only 2 16Bit integers in the array)
'            myHash = (myHash + s.W(1)) * HMul And s.hashslotmask
         
        Case vbCurrency
        
            C = myKey
            s.saW.cElements1D = 4
            s.saW.pvData = VarPtr(C)
            
            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 4 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
            myHash = (myHash + s.W(2)) * HMul And s.HashSlotMask
            myHash = (myHash + s.W(3)) * HMul And s.HashSlotMask
         
        Case vbLong, vbInteger, vbByte
        
            L = myKey
            s.saW.cElements1D = 2
            s.saW.pvData = VarPtr(L)
            
            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 2 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
            
        Case vbDouble
        
            D = myKey: s.saW.cElements1D = 4
            s.saW.pvData = VarPtr(D)
            
            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 4 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
            myHash = (myHash + s.W(2)) * HMul And s.HashSlotMask
            myHash = (myHash + s.W(3)) * HMul And s.HashSlotMask
            
        Case vbDate
        
            D = myKey
            s.saW.cElements1D = 4
            s.saW.pvData = VarPtr(D)
            
            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 4 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
            myHash = (myHash + s.W(2)) * HMul And s.HashSlotMask
            myHash = (myHash + s.W(3)) * HMul And s.HashSlotMask

        Case vbSingle
        
            F = myKey
            s.saW.cElements1D = 2
            s.saW.pvData = VarPtr(F)
            myHash = (myHash + s.W(0)) * HMul And s.HashSlotMask 'loop-unrolling (we have only 2 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HashSlotMask
            
    End Select
    ' the hashslot array starts at 1 so we must add 1 below for the case where the hash is zero
    GetHash = myHash + 1
    
End Function


Private Function GetItemAsString(ByVal ipItem As Variant) As String

    ' containers and arrays must be sorted if we are to have comparable hashes
    Dim myResult As Variant
    Select Case True
        Case GroupInfo.IsArray(ipItem):
            myResult = ArrayOp.Linearise(ipItem)
            Sorters.ShakerSortArray myResult

        Case GroupInfo.IsList(ipItem):
            Sorters.ShakerSortByItem ipItem
            Set myResult = ipItem
            
        Case GroupInfo.IsItemByToArray(ipItem)
            myResult = ipItem.ToArray
            Sorters.ShakerSortArray myResult
            
        Case GroupInfo.IsDictionary(ipItem)
            myResult = ipItem.SPairs
            
        Case GroupInfo.IsItemObject(ipItem)
            myResult = Stringifier.SetNoMarkup.StringifyItemObject(ipItem)
    End Select
        
    GetItemAsString = Fmt.SetNoMarkup.Text("{0}", myResult)
        
End Function

Private Sub SetupFirstNode(ByRef ipNode As KvpHNode)

        Set s.HashRoot = ipNode
        Set s.HashLast = ipNode
        
        ' nodes for managining sequence by order field
        Set s.OrderRoot = ipNode
        Set s.OrderLast = ipNode
        
        s.OrderRoot.UpdateOrder
        
        ' populate hash slot
        Set s.HashSlots(s.HashRoot.Hash) = s.HashRoot
        
End Sub
Private Sub AppendByHash(ByRef ipNode As KvpHNode)

    ' The append node method adds a node s.HashLast
    ' link nodes following the hash
    ' If s.HashSlots(ipNode.Hash) Is Nothing Then
    '     Set s.HashSlots(ipNode.Hash) = ipNode
    ' End If
    
    
    Set ipNode.HashPred = s.HashLast
    Set ipNode.HashSucc = Nothing
    Set s.HashLast.HashSucc = ipNode
    Set s.HashLast = ipNode
    
End Sub

Private Sub AppendByOrder(ByRef ipNode As KvpHNode)
    
    ' no need to look at hashslots as this is handled by  'byHash' methods
    Set ipNode.OrderPred = s.OrderLast
    Set ipNode.OrderSucc = Nothing
    Set s.OrderLast.OrderSucc = ipNode
    Set s.OrderLast = ipNode
    s.OrderLast.OrderPred.UpdateOrder
    
End Sub


Private Sub InsertbyHash(ByVal ipNewNode As KvpHNode)

    ' we know that the new node is not the first node.
    ' we know that the new node exists in hashslots
    ' we insert the newnode after the last node that matches Hash/Item
    
    ' get the first node for the newnode hash
    Dim myCurNode As KvpHNode: Set myCurNode = s.HashSlots(ipNewNode.Hash)

    'The new node is inserted at the end of the current hash/item combination
    ' we know that curnode and new node have the same hash
    ' we don't know if they have the same items

    Do While myCurNode.Hash = ipNewNode.Hash
        ' check item is the same item as ipnewnode
        If ComparerHelpers.Equals(myCurNode.Key, ipNewNode.Key) Then
            ' check the next node is not nothing
            If myCurNode.HashSucc Is Nothing Then
                ' if ipnode.succ is nothing then by default
                ' myNode is s.last so we can just append the new node
                AppendByHash ipNewNode
                Exit Sub
            Else
                Set myCurNode = myCurNode.HashSucc
            End If
        End If
    Loop
    
    ' at this point we have matching hashes but different items
    ' curnode is the node after the last item/hash match so we
    ' insert newnode before curnode
    
    Set ipNewNode.HashPred = myCurNode.HashPred 'myPredNode
    Set ipNewNode.HashSucc = myCurNode 'myPred.Succ
    
    Set myCurNode.HashPred = ipNewNode
    Set ipNewNode.HashPred.HashSucc = ipNewNode
    
End Sub
Private Sub InsertByOrder(ByVal ipNewNode As KvpHNode)

    ' we know that the new node is not the first or last node.
    ' we know that the new node exists in hashslots
    ' we insert the newnode before the node with the same order
  
    Dim myCurNode As KvpHNode: Set myCurNode = GetNodeAt(ipNewNode.Order)
    
    Set ipNewNode.OrderPred = myCurNode.OrderPred
    Set ipNewNode.OrderSucc = myCurNode
    
    Set myCurNode.OrderPred.OrderSucc = ipNewNode
    Set myCurNode.OrderPred = ipNewNode
    
    ipNewNode.UpdateOrder
    
    ' Dim myCurNode As KvpHNode: Set myCurNode = GetNodeAt(ipNewNode.Order)
    
    ' Set ipNewNode.OrderPred = myCurNode.OrderPred
    ' Set myCurNode.OrderPred.OrderSucc = ipNewNode
    
    ' Set ipNewNode.OrderSucc = myCurNode
    ' Set myCurNode.OrderPred = ipNewNode
    
    ' ipNewNode.UpdateOrder
    
End Sub

Private Sub PrependByOrder(ByRef ipNode As KvpHNode)

    Set ipNode.OrderPred = Nothing
    Set ipNode.OrderSucc = s.OrderRoot
    Set s.OrderRoot.OrderPred = ipNode
    Set s.OrderRoot = ipNode
    s.OrderRoot.UpdateOrder
    
End Sub
'@Ignore ProcedureCanBeWrittenAsFunction
'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub RemoveNode(ByRef ipNode As KvpHNode)

    ' when removing a node we must check that the node is not the node registered in hashslots
    ' if it is we must replace the node pointed to by s.HashSlots(hash) with
    ' s.HashSlots.succ or nothing
    ' wemust also remember to update roo/last nodes as appropriate
    
    ' check for the case of a single node
    If s.OrderLast.Order = FIRST_INDEX Then
        Set s.HashSlots(ipNode.Hash) = Nothing
        Set ipNode = Nothing
        Set s.HashRoot = Nothing
        Set s.HashLast = Nothing
        Set s.OrderRoot = Nothing
        Set s.OrderLast = Nothing
        Exit Sub
    End If
    
    ' we know that item will have a hash for an occupied slot because the absense of its
    ' item from the seq was previously tested
    ' remove node from hashslot if succ is nothing or if the hash of the next node is different
    If s.HashSlots(ipNode.Hash) Is ipNode Then
        If ipNode.HashSucc Is Nothing Then
            Set s.HashSlots(ipNode.Hash) = Nothing
        ElseIf ipNode.Hash = ipNode.HashSucc.Hash Then
            Set s.HashSlots(ipNode.Hash) = ipNode.HashSucc
        Else
            Set s.HashSlots(ipNode.Hash) = Nothing
        End If
    End If

    
    '@Ignore VariableNotUsed
    Dim myTmp As KvpHNode
    ' Disconnect the node from hash seq
    ' node deletion only occurs after order seq is updates
    If ipNode.HashPred Is Nothing Then
        
       
        Set s.HashRoot = ipNode.HashSucc
        Set s.HashRoot.HashPred = Nothing
        
    ElseIf ipNode.HashSucc Is Nothing Then
        
        Set s.HashLast = ipNode.HashPred
        Set s.HashLast.HashSucc = Nothing
        
    Else
    
        Set ipNode.HashSucc.HashPred = ipNode.HashPred
        Set ipNode.HashPred.HashSucc = ipNode.HashSucc
      
    End If
    
    ' disconnect the node from order seq
    ' is this the first order node
    If ipNode.OrderPred Is Nothing Then
        Set myTmp = ipNode
        Set s.OrderRoot = ipNode.OrderSucc
        Set s.OrderRoot.OrderPred = Nothing
        s.OrderRoot.UpdateOrder
        Set myTmp = Nothing
        Exit Sub
    End If
        
    ' is ipnode the last order node
    If ipNode.OrderSucc Is Nothing Then
        ' if ipnode.ordersucc is nothing this means that ipnode **is** ipNode.last
        Set myTmp = ipNode
        Set s.OrderLast = ipNode.OrderPred
        Set s.OrderLast.OrderSucc = Nothing
        s.OrderLast.OrderPred.UpdateOrder
        Set myTmp = Nothing
        Exit Sub
    End If
    
    Set ipNode.OrderSucc.OrderPred = ipNode.OrderPred
    Set ipNode.OrderPred.OrderSucc = ipNode.OrderSucc
    ipNode.OrderPred.UpdateOrder
    Set ipNode = Nothing
    
End Sub

'@Description("Returns the node where .Order = ipIndex")
Private Function GetNodeAt(ByRef ipIndex As Long) As Variant
Attribute GetNodeAt.VB_Description = "Returns the node where .Order = ipIndex"

    Set GetNodeAt = Nothing
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If

    If ipIndex < s.OrderRoot.Order Or ipIndex > s.OrderLast.Order Then
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.OrderRoot
    Do
        If myNode.Order = ipIndex Then
            Set GetNodeAt = myNode
            Exit Function
        End If
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing

End Function

'@Description("Returns the first node containing ipItem")
Private Function GetNodeOf(ByRef ipKey As Variant) As Variant
Attribute GetNodeOf.VB_Description = "Returns the first node containing ipItem"

    Set GetNodeOf = Nothing
    
    If s.OrderRoot Is Nothing Then
        Exit Function
    End If

                   
    Dim myNode As KvpHNode: Set myNode = s.HashSlots(GetHash(ipKey))
    If myNode Is Nothing Then
        Exit Function
    End If
    
    Do
        If ComparerHelpers.Equals(myNode.Key, ipKey) Then
            Set GetNodeOf = myNode
            Exit Function
        End If
        
        If myNode.Hash <> myNode.OrderSucc.Hash Then
            Exit Function
        End If
        
        Set myNode = myNode.OrderSucc
        
    Loop Until myNode Is Nothing

End Function

Private Sub RemoveAllItems()

    Dim myNode As KvpHNode:  Set myNode = s.OrderLast
    
    Do
    
        Dim myTmpNode As KvpHNode: Set myTmpNode = myNode
        Set myNode = myNode.OrderPred
        Set myTmpNode = Nothing
        
    Loop Until myTmpNode Is Nothing
    
    Set s.HashLast = Nothing
    Set s.HashRoot = Nothing
    Set s.OrderRoot = Nothing
    Set s.OrderLast = Nothing
   ' s.OrderLast.order = 0
    
End Sub


Private Sub RemoveAllSingleItem(ByRef ipItem As Variant)

    ' all nodes with the same hash /item are linked
    ' Where there is a hash clash the node at s.hashslots(myHash) may not hold ipItem
    ' in which case we have to search forward to find the item
    ' we also have to be mindful that whilst the hash of an item may be valid (the relevant hashslot is occupied)
    ' the item itself may not be in the seq
    Dim myHash As Long: myHash = GetHash(ipItem)
    Dim myNode As KvpHNode: Set myNode = s.HashSlots(myHash)
    
    'check that that hashslot is occupied
    If myNode Is Nothing Then
        Exit Sub
    End If
    
    ' search for the first matching item with the current hash
    ' also realising that becasue we have a valid hash doesn't mean that the item is in the sequence
    Do Until ComparerHelpers.Equals(myNode.Item, ipItem)
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
        
        Dim myTmp As KvpHNode: Set myTmp = myNode
        Set myNode = myNode.HashSucc
        RemoveNode myTmp
        
    Loop While ComparerHelpers.Equals(myNode.Item, ipItem)
    
End Sub
