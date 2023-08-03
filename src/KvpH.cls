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
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
Option Explicit

'fixed HashSlot-amount and Hash-Multiplikator
Private Const HMul                          As Long = 3727
Private Const MY_TYPENAME                   As String = "KvpH"
Private Const FIRST_INDEX                   As Long = 1

Private Type State
    W()                                     As Integer
    saW                                     As SAFEARRAY1D
          
    Count                                   As Long
    HTUB                                    As Long 'represents the number of bits which can represent the positive integers that encapsulate the Nodes array size
    
    ' Nodes store the first node with a given Hash value
    ' if a second node has the same hash value then it is inserted
    ' in the list of nodes at the end of the nodes with the current hash/key combination
    ' the first node.
    HashSlots()                             As KvpHNode
    Root                                    As KvpHNode
    Last                                    As KvpHNode
     
End Type


Private s                                   As State

Private Type Properties
    EnsureUniqueKeys                        As Boolean
    Size                                    As Long ' The dimmed size of s.HashSlots
End Type


Private p                                   As Properties


Public Function Deb() As KvpH
    With New KvpH
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As KvpH

    s.saW.cDims = 1
    s.saW.cbElements = 2
    s.saW.cLocks = 1
    s.saW.fFeatures = &H11                       'FADF_AUTO=&H1 || FADF_FIXEDSIZE=&H10
    BindArray s.W, VarPtr(s.saW)
    SetSize 16383                                'Hangover from cHashD: at startup we set it up, to behave ideally for up to 16384 Items
    p.EnsureUniqueKeys = True
    Set ConstructInstance = Me
    
End Function


'@Description("set the size of hash array to be the 2^n value just greater than ipNewSize")
'@Ignore AssignedByValParameter
Public Sub SetSize(Optional ByVal ipNewSize As Long)
Attribute SetSize.VB_Description = "set the size of hash array to be the 2^n value just greater than ipNewSize"

    If ipNewSize < 0 Then ipNewSize = 1023

    'calculate a long whose value is the first 2^n-1 that is greater than ipNewSize
    ' this provides a bitfield of 1's used to set the maximum value for a hash value
    s.HTUB = (2 ^ Int((VBA.Log(ipNewSize) / VBA.Log(2)) + 1)) - 1
    If s.HTUB < ipNewSize Then
        s.HTUB = s.HTUB * 2 + 1
    End If
    'Debug.Print Hex$(ipNewSize), Hex$(s.HTUB)
    
    ' Redim without preserve should mean that we have no entries in s.HashSlots
    ReDim s.HashSlots(FIRST_INDEX To s.HTUB)
    p.Size = s.HTUB
    If s.Root Is Nothing Then
        Exit Sub
    End If
    
    ' update the hashes in the nodes
    ' then populate the new s.modes using the new hashes
    ' this does not change the order of addition,
    ' just the slot used in s.HashSlots to store the node
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Do
        myNode.Hash = GetHash(myNode.Key)
        Set myNode = myNode.Succ
    Loop Until myNode Is Nothing
    
    'repopulate s.HashSlots according to the new hashes
    
    Set myNode = s.Root
    Do
        If VBA.IsEmpty(s.HashSlots(myNode.Hash)) Then
        
            Set s.HashSlots(myNode.Hash) = myNode
            
        Else
            ' if the next node has the same hash then skip over it
            Do While myNode.Hash = myNode.Succ.Hash
                Set myNode = myNode.Succ
            Loop
            
        End If
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
    
End Sub


Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpH

    Set Add = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Add", _
        "Null is not a valid Key"
    End If
    
    If p.EnsureUniqueKeys Then
        If HoldsKey(ipKey) Then
            Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Add", _
            Fmt.Text("Duplicate key: {0}", ipKey)
        End If
    End If

    Dim myHash As Long: myHash = GetHash(ipKey)
    'Debug.Print ipKey, myHash
    If s.Count < FIRST_INDEX Then
    
        ' add the first hash/node
        Set s.Root = KvpHNode.Deb(Nothing, Nothing, myHash, ipKey, ipItem)
        Set s.Last = s.Root
        Set s.HashSlots(myHash) = s.Root
        s.Count = s.Count + 1
        Set Add = Me
        Exit Function
    End If
    
    'we are not the first key of this hash
    If s.HashSlots(myHash) Is Nothing Then
        ' append node to s.last
        AppendNode KvpHNode(Nothing, Nothing, myHash, ipKey, ipItem)
        
        Set Add = Me
        Exit Function
    End If
    
    If p.EnsureUniqueKeys Then
        
        ' the hash slot is already occupied
        ' so check is the node key matches inut key
        If ComparerHelpers.Equals(ipKey, s.HashSlots(myHash).Key) Then
            Err.Raise 17 + vbObject, _
            "VBAlib.KvpH.Add", _
            Fmt.Text("Duplicate Keys not allowed. Got {0}", ipKey)
        End If
        
    End If
    
    InsertNode s.HashSlots(myHash), KvpHNode(Nothing, Nothing, myHash, ipKey, ipItem)
   
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
            "Null is not a valid Key"
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
    
    If s.Root Is Nothing Then
        Exit Property
    End If
                   
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Let_Item", _
        "Null is not a valid Key"
    End If

    Dim myH As Long: myH = GetHash(ipKey)
    If VBA.IsEmpty(s.HashSlots(myH)) Then
        Item = Null
        Exit Property
    End If
    
    If VBA.IsObject(s.HashSlots(myH).Item) Then
        Set Item = s.HashSlots(myH).Item
    Else
        Item = s.HashSlots(myH).Item
    End If
    
End Property


Public Property Let Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    Dim myH As Long: myH = GetHash(ipKey)
    
    If s.HashSlots(myH) Is Nothing Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Let_Item", _
        Fmt.Text("No such Key: '{0}'", ipKey)
    End If
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Let_Item", _
        "Null is not a valid Key"
    End If
    
    s.HashSlots(myH).Item = ipItem
    
End Property


Public Property Set Item(ByRef ipKey As Variant, ByVal ipItem As Variant)
 
    Dim myH As Long: myH = GetHash(ipKey)
    
    If s.HashSlots(myH) Is Nothing Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Set_Item", _
        Fmt.Text("Key {0} not found", ipKey)
    End If
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Set_Item", _
        "Null is not a valid Key"
    End If
    
    Set s.HashSlots(myH).Item = ipItem
    
End Property


Public Function Remove(ParamArray ipParamarray() As Variant) As KvpH

    Set Remove = Me
    
    If s.Root Is Nothing Then
        Exit Function
    End If
    
    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(ipParamarray)
    
    If Not VBA.IsArray(myParamarray) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Remove", _
        Fmt.Text("Expecting an array.  Got {0}", VBA.Typename(myParamarray))
    End If
    
    Dim myItem As Variant
    For Each myItem In myParamarray
          
        If VBA.IsNull(myItem) Then
            Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Remove", _
            "Null is not a valid Key"
        End If
    
        Dim myH As Long: myH = GetHash(myItem)
        If Not s.HashSlots(myH) Is Nothing Then
            RemoveNode s.HashSlots(myH)
        End If
    Next
    
    Set Remove = Me
    
End Function


Public Function RemoveAt(ByVal ipIndex As Long) As KvpH

    Set RemoveAt = Me
    
    If ipIndex < FIRST_INDEX Or ipIndex > s.Count Then
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Do Until ipIndex = myNode.Order
        Set myNode = myNode.Succ
    Loop
    
    RemoveNode myNode
    
End Function


Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpH

    Set RemoveAfter = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.RemoveAfter", _
        "Null is not a valid Key"
    End If
    
    If s.Root Is Nothing Then
        Exit Function
    End If
    
    If Me.LacksKey(ipKey) Then
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
        "VALib.KvpH.RemoveAfter", _
        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    Do Until ComparerHelpers.Equals(ipKey, myNode.Key)
        Set myNode = myNode.Succ
        If myNode Is Nothing Then
            Exit Function
        End If
    
    Loop
    
    If myNode.Succ Is Nothing Then
        Exit Function
    End If
    
    '@Ignore AssignmentNotUsed
    Set myNode = myNode.Succ
    
    Dim myCount As Long
    For myCount = 1 To ipRun
    
        RemoveNode myNode.Succ
        
        If myNode.Succ Is Nothing Then
            Exit Function
        End If
        
    Next
    
End Function


Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpH

    Set RemoveBefore = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.RemoveBefore", _
        "Null is not a valid Key"
    End If
    
    If Me.LacksKey(ipKey) Then
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
        "VALib.KvpH.RemoveBefore", _
        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    Do Until ComparerHelpers.Equals(ipKey, myNode.Key)
        Set myNode = myNode.Succ
        If myNode Is Nothing Then
            Exit Function
        End If
    
    Loop
    
    If myNode.Pred Is Nothing Then
        Exit Function
    End If
    
    '@Ignore AssignmentNotUsed
    Set myNode = myNode.Pred
    
    Dim myCount As Long
    For myCount = 1 To ipRun
    
        RemoveNode myNode.Pred
        
        If myNode.Pred Is Nothing Then
            Exit Function
        End If
        
    Next
    
End Function


Public Function Clear() As KvpH
    
    ' this may be belt and braces but I suspect not
    ' Nodes need to be explicily set to nothing
    ' to ensure vba can reclaim the node memory
    ' within node references are set to nothing by the node terminate event.
    
    Dim myNode As KvpHNode: Set myNode = s.Last
    Do Until myNode.Pred Is Nothing
        Set myNode = myNode.Pred
        '@Ignore VariableNotUsed
        Dim myTmp As KvpHNode
        Set myTmp = myNode.Succ
        Set myTmp = Nothing
    Loop
    Set s.Root = Nothing
    Set s.Last = Nothing
    
    ' now clear the array holding node references
    SetSize UBound(s.HashSlots)
    s.Count = 0
    Set Clear = Me
    
End Function


Public Function RemoveAll() As KvpH
    Me.Clear
    Set RemoveAll = Me
End Function


Public Function Reset() As KvpH
    Me.Clear
    Set Reset = Me
End Function


'===================================================================================
' ToDo: add insert methods
'===================================================================================

Public Function Clone() As KvpH

    Dim myK As KvpH: Set myK = KvpH.Deb
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Do
    
        myK.Add myNode.Key, myNode.Item
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
    
    Set Clone = myK
    
End Function


Public Function Exists(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Exists", _
        "Null is not a valid Key"
    End If

    Dim myH As Long: myH = GetHash(ipKey)
    Exists = False
    
    If s.Root Is Nothing Then
        Exit Function
    End If
    
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
        
        Set myNode = myNode.Pred
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

    If s.Root Is Nothing Then
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

    If s.Root Is Nothing Then
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

    If s.Root Is Nothing Then
        LacksKey = True
        Exit Function
    End If
    
    LacksKey = Not Exists(ipKey)
    
End Function


Public Function ExistsItem(ByRef ipItem As Variant) As Boolean

    If s.Root Is Nothing Then
        ExistsItem = False
        Exit Function
    End If

    ExistsItem = Me.HoldsItem(ipItem)
    
End Function


Public Function HoldsItem(ByRef ipItem As Variant) As Boolean

    If s.Root Is Nothing Then
        HoldsItem = False
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    Do
        If ComparerHelpers.Equals(myNode.Item, ipItem) Then
            HoldsItem = True
            Exit Function
        End If
        Set myNode = myNode.Succ
    Loop Until myNode Is Nothing
    
End Function


Public Function Contains(ByRef ipItem As Variant) As Boolean
    
    If s.Root Is Nothing Then
        Contains = False
        Exit Function
    End If
    
    Contains = Me.HoldsItem(ipItem)
    
End Function


Public Function LacksItem(ByRef ipItem As Variant) As Boolean

    If s.Root Is Nothing Then
        LacksItem = True
        Exit Function
    End If

    LacksItem = Not HoldsItem(ipItem)
End Function


Public Function HoldsItems() As Boolean
    HoldsItems = Not s.Root Is Nothing
End Function


Public Function LacksItems() As Boolean
    LacksItems = s.Root Is Nothing
End Function


Public Property Get Typename() As String
    Typename = MY_TYPENAME
End Property


'@Description("Mutates the KvpA in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As KvpH
Attribute MappedIt.VB_Description = "Mutates the KvpA in place by applying the function in IMapper to each item"

    If s.Root Is Nothing Then
        Set MappedIt = Me
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set KvpHNode = s.Root
    Do
    
        Dim myItem As Variant
        myItem = ipMapper.ExecMapper(myNode.Item)
        
        If VBA.IsObject(myItem(0)) Then
            Set myNode.Item = myItem(0)
        Else
            myNode.Item = myItem(0)
        End If
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
   
    Set MappedIt = Me

End Function


'@Description("Returns a KvpA.Deb by applying the function in IMapper to each Item")
Public Function MapIt(ByVal ipMapper As IMapper) As KvpH
Attribute MapIt.VB_Description = "Returns a KvpA.Deb by applying the function in IMapper to each Item"

    If s.Root Is Nothing Then
        Set MapIt = Me
        Exit Function
    End If

    Dim myK As KvpH: Set myK = KvpH.Deb
    
    Dim myNode As KvpHNode: Set KvpHNode = s.Root
    Do
    
        Dim myItem As Variant
        myItem = ipMapper.ExecMapper(myNode.Item)
        
        myK.Add myNode.Key, myItem(0)
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
      
    Set MapIt = myK

End Function


'@Description("Returns a KvpA.Deb where the ICOmparer has been applied to the Items of the original Kvp")
Public Function FilterIt(ByVal ipComparer As IComparer) As KvpH
Attribute FilterIt.VB_Description = "Returns a KvpA.Deb where the ICOmparer has been applied to the Items of the original Kvp"

    Dim myK As KvpH: Set myK = KvpH.Deb
    
    If s.Root Is Nothing Then
        Set FilterIt = myK
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set KvpHNode = s.Root
    Do
    
        If ipComparer.ExecCmp(myNode.Item) Then
            myK.Add myNode.Key, myNode.Item
        End If
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
      
    Set FilterIt = myK

End Function


Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant

    If s.Root Is Nothing Then
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


Public Function Count(Optional ByRef ipItem As Variant) As Long
    
    If s.Root Is Nothing Then
        Count = -1
        Exit Function
    End If
    
    If VBA.IsMissing(ipItem) Then
        Count = s.Count
        Exit Function
    End If
    
    Dim myCount As Long: myCount = 0
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    Do
    
        If ComparerHelpers.Equals(ipItem, myNode.Item) Then
            myCount = myCount + 1
        End If
        
        Set myNode = myNode.Succ
    
    Loop Until myNode Is Nothing
    
    Count = myCount
    
End Function


Public Function CountIt(ByVal ipComparer As IComparer) As Long

    If s.Root Is Nothing Then
        CountIt = 0
        Exit Function
    End If

    Dim myCount As Long: myCount = 0
    Dim myNode As KvpHNode: Set myNode = s.Root
    Do
    
        If ipComparer.ExecCmp(myNode.Item) Then
            myCount = myCount + 1
        End If
    
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing

    CountIt = myCount

End Function


Public Function KeyByIndex(ByVal ipIndex As Long) As Variant

    KeyByIndex = Null
    
    If s.Root Is Nothing Then
        Exit Function
    End If
    
    If ipIndex < FIRST_INDEX Or ipIndex > s.Count Then
        Exit Function
    End If
    
    Dim myNode As KvpHNode: Set KvpHNode = s.Root
    Do While myNode.Order < ipIndex
        Set myNode = myNode.Succ
    Loop
    
    If VBA.IsObject(myNode.Key) Then
        Set KeyByIndex = myNode.Key
    Else
        KeyByIndex = myNode.Key
    End If
    
End Function


Public Function KeyOf(ByVal ipItem As Variant) As Variant

    KeyOf = Null
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    Do
    
        If ComparerHelpers.Equals(myNode.Item, ipItem) Then
            If VBA.IsObject(myNode.Key) Then
                Set KeyOf = myNode.Key
            Else
                KeyOf = myNode.Key
            End If
            
            Exit Function
            
        End If
        
        Set myNode = myNode.Succ
        
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

    If s.Root Is Nothing Then
        Items = Array()
        Exit Function
    End If

    Dim myA() As Variant: ReDim myA(1 To s.Count)
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Dim myIndex As Long: myIndex = 1
    Do
    
        If VBA.IsObject(myNode.Key) Then
            Set myA(myIndex) = myNode.Item
        Else
            myA(myIndex) = myNode.Item
        End If
        
        myIndex = myIndex + 1
        
        Set myNode = myNode.Succ

    Loop Until myNode Is Nothing
    
    Items = myA
    
End Function


Public Function ItemsAsSeq() As SeqH

    If s.Root Is Nothing Then
        Set ItemsAsSeq = SeqH.Deb
        Exit Function
    End If
    
    Dim myS As SeqH: Set myS = SeqH.Deb
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Do
    
        myS.Add myNode.Item
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
    
    Set ItemsAsSeq = myS
    
End Function


Public Function Keys() As Variant
    
    If s.Count < 1 Then
        Keys = Array()
        Exit Function
    End If
    
    Dim myA() As Variant: ReDim myA(1 To s.Count)
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Dim myIndex As Long
    myIndex = 1
    Do
    
        If VBA.IsObject(myNode.Key) Then
            Set myA(myIndex) = myNode.Key
        Else
            myA(myIndex) = myNode.Key
        End If
        myIndex = myIndex + 1
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
    
    Keys = myA
    
End Function


Public Function KeysAsSeq() As SeqH

    If s.Root Is Nothing Then
        Set KeysAsSeq = SeqH.Deb
        Exit Function
    End If
    
    Dim myS As SeqH: Set myS = SeqH.Deb
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    '    Dim myIndex As Long
    '    myIndex = 1
    Do
    
        myS.Add myNode.Key
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
    
    Set KeysAsSeq = myS
End Function


Public Function Pairs() As Variant()             'hand-out the Key-Values in a Pairs-Array

    If s.Root Is Nothing Then
        Pairs = Array()
        Exit Function
    End If
    
    Dim myA() As Variant: ReDim myA(FIRST_INDEX To s.Count)
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Dim myIndex As Long: myIndex = 1
    Do
    
        myA(myIndex) = Array(myNode.Key, myNode.Item)
        
        myIndex = myIndex + 1
        
        Set myNode = myNode.Succ

    Loop Until myNode Is Nothing
    
    Pairs = myA
    
End Function

Public Function SPairs() As Variant

    Dim myKeys As SeqH: Set myKeys = Me.KeysAsSeq

    If myKeys.LacksItems Then
        SPairs = Array()
        Exit Function
    End If
    
    Dim myA As Variant: ReDim myA(myKeys.FirstIndex To myKeys.LastIndex)
    
    Dim myIndex As Long
    For myIndex = myKeys.FirstIndex To myKeys.LastIndex
        myA(myIndex) = Array(myKeys.Item(myIndex), Me.Item(myKeys.Item(myIndex)))
    Next

    SPairs = myA

End Function
Public Function PairsAsSeq() As SeqH
    
    If s.Root Is Nothing Then
        Set PairsAsSeq = SeqH.Deb
        Exit Function
    End If
    
    Dim myS As SeqH: Set myS = SeqH.Deb
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    
    Do
    
        myS.Add Array(myNode.Key, myNode.Item)
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
    
    Set PairsAsSeq = myS
    
End Function


Public Property Get FirstIndex() As Long
    FirstIndex = FIRST_INDEX
End Property


Public Property Get LastIndex() As Long
    LastIndex = s.Count
End Property


Public Property Get FirstKey() As Variant

    FirstKey = Null
    
    If s.Root Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Root.Key) Then
        Set FirstKey = s.Root.Key
    Else
        FirstKey = s.Root.Key
    End If
    
End Property


Public Function LastKey() As Variant

    LastKey = Null
    
    If s.Root Is Nothing Then
        Exit Function
    End If
    
    If VBA.IsObject(s.Last.Key) Then
        Set LastKey = s.Last.Key
    Else
        LastKey = s.Last.Key
    End If
    
End Function


Public Property Get First() As Variant

    First = Null
    
    If s.Root Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Root.Item) Then
        Set First = s.Root.Item
    Else
        First = s.Root.Item
    End If
    
End Property


Public Property Set First(ByVal ipItem As Variant)
    
    If s.Root Is Nothing Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Let_FirstItem", _
        "Kvp has no Keys"
    End If
    
    Set s.Root.Item = ipItem

End Property


Public Property Let First(ByVal ipItem As Variant)

    If s.Root Is Nothing Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Let_FirstItem", _
        "Kvp has no Keys"
    End If
    
    s.Root.Item = ipItem

End Property


Public Property Get Last() As Variant

    Last = Null
    
    If s.Root Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Last.Item) Then
        Set Last = s.Last.Item
    Else
        Last = s.Last.Item
    End If
    
End Property


Public Property Set Last(ByVal ipItem As Variant)

    If s.Root Is Nothing Then
        Err.Raise 17, "VBALib.KvpL.Set LastItem", "VBALib.Kvp has no Keys"
    End If

    Set s.Last.Item = ipItem
    
End Property


Public Property Let Last(ByVal ipItem As Variant)

    If s.Root Is Nothing Then
        Err.Raise 17, "VBALib.KvpL.Let LastItem", "VBALib.Kvp has no Keys"
    End If
    
    s.Last.Item = ipItem
    
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
    
    If s.Root Is Nothing Then
        Set Mirror = KvpH.Deb
        Exit Function
    End If

    Dim myNode As KvpHNode: Set myNode = s.Root
    Do
        If myK.LacksKey(myNode.Item) Then
            myK.Add myNode.Item, KvpL.Deb
        End If

        myK.Item(myNode.Item).Add myNode.Item, myNode.Order
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing

    Set MirrorByAllValues = myK

End Function


Private Function MirrorByFirstValue() As KvpH

    Dim myFirst As KvpH
    Set myFirst = KvpH.Deb
    
    Dim myOther As KvpH
    Set myOther = KvpH.Deb
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    Do

        If myFirst.LacksKey(myNode.Item) Then
            myFirst.Add myNode.Item, myNode.Key
        Else
            myOther.Add myNode.Key, myNode.Item
        End If

    Loop Until myNode Is Nothing

    Set MirrorByFirstValue = KvpH.Deb.Add(0&, myFirst).Add(1&, myOther)

End Function


'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub InsertNode(ByRef ipNode As KvpHNode, ByVal ipNewNode As KvpHNode)

    ' ipNode is the first node at the current hash
    ' There may be other linked nodes with the same hash and hence the same key.
    ' consequently we need to search successor nodes
    ' until we find a successor that doesn't have the same key
    ' we also need to be mindful that we don't go past the end of the list of nodes
    
    Do Until ipNode.Succ Is Nothing              ' or we find Key <> .succ.key
        
        If ipNode.Key = ipNode.Succ.Key Then
            Set ipNode = ipNode.Succ
        Else
            Exit Do
        End If
        
    Loop
    
    If ipNode.Succ Is Nothing Then
        
        Set ipNode.Succ = ipNewNode
        Set ipNewNode.Pred = ipNode
        Set s.Last = ipNewNode
        s.Count = s.Count + 1
        Exit Sub
    End If

    Dim myPredNode As KvpHNode: Set myPredNode = ipNode.Pred
    
    Set ipNewNode.Pred = myPredNode
    Set ipNewNode.Succ = ipNode
    
    Set myPredNode.Succ = ipNewNode
    Set ipNode.Pred = ipNewNode
    s.Count = s.Count + 1
    
End Sub


Public Function IndexByKey(ByRef ipKey As Variant) As Long

    If s.Count < FIRST_INDEX Then
        IndexByKey = -1
        Exit Function
    End If
    
    Dim myH As Long: myH = GetHash(ipKey)
    
    If VBA.IsEmpty(s.HashSlots(myH)) Then
        IndexByKey = -1
        Exit Function
    End If
    
    IndexByKey = s.HashSlots(myH).Order
    
End Function


Public Property Get ItemByIndex(ByVal ipIndex As Long) As Variant

    If ipIndex < FIRST_INDEX Or ipIndex > s.Count Then
        ItemByIndex = Null
        Exit Property
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Dim myIndex As Long: myIndex = 1
    Do Until myIndex = ipIndex
        Set myNode = myNode.Succ
        myIndex = myIndex + 1
    Loop
    
    If VBA.IsObject(myNode.Item) Then
        Set ItemByIndex = myNode.Item
    Else
        ItemByIndex = myNode.Item
    End If
    
End Property


Public Property Let ItemByIndex(ByVal ipIndex As Long, ByVal ipItem As Variant)

    If ipIndex < FIRST_INDEX Or ipIndex > s.Count Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.KvpH.Let_ItemByIndex", _
            Fmt.Text("Bounds are {0} to {1}. Got index of {2}", FIRST_INDEX, s.Count, ipIndex)
    End If
    
    If VBA.IsNull(ipItem) Then
        Err.Raise 17 + vbObjectError, _
            "Let_ItemByIndex", _
            "Cannot store a value of null"
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Dim myIndex As Long: myIndex = 1
    Do Until myIndex = ipIndex
        Set myNode = myNode.Succ
        myIndex = myIndex + 1
    Loop
    
    myNode.Item = ipItem
   
End Property


Public Property Set ItemByIndex(ByVal ipIndex As Long, ByVal ipItem As Variant)

    If ipIndex < FIRST_INDEX Or ipIndex > s.Count Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpH.Let_ItemByIndex", _
        Fmt.Text("Bounds are {0} to {1}. Got index of {2}", FIRST_INDEX, s.Count, ipIndex)
    End If
    
    If VBA.IsNull(ipItem) Then
        Err.Raise 17 + vbObjectError, _
        "Let_ItemByIndex", _
        "Cannot store a value of null"
    End If
    
    Dim myNode As KvpHNode: Set myNode = s.Root
    
    Dim myIndex As Long: myIndex = 1
    Do Until myIndex = ipIndex
        Set myNode = myNode.Succ
        myIndex = myIndex + 1
    Loop
    
    Set myNode.Item = ipItem
    
End Property


Private Sub AppendNode(ByVal ipNode As KvpHNode)

    Set s.HashSlots(ipNode.Hash) = ipNode
    Set s.Last.Succ = ipNode
    Set ipNode.Pred = s.Last
    Set s.Last = ipNode
    ipNode.AdjustNodeOrder
    s.Count = s.Count + 1
    
End Sub


Private Function GetHash(ByRef Key As Variant) As Long

    Dim D As Double
    Dim L As LongPtr
    Dim F As Single
    Dim VT As VbVarType
    Dim myHash As Long
    'Dim s.HTUB As Long
     
    Dim C As Variant
    C = VBA.CDec(C)
     
    
    myHash = s.HTUB                              'init the HashValue (all bits to 1)


    VT = VarType(Key)
    Select Case VT
        Case vbString
        
            Dim myIndex As Long
            For myIndex = 1 To VBA.Len(Key)
                
                Dim myChar As Long
                myChar = VBA.AscW(VBA.Mid$(Key, myIndex, 1))
                myHash = (myHash + myChar) * HMul And s.HTUB
                
            Next
                   
            
        Case vbObject
        
            L = ObjPtr(Key)
            s.saW.cElements1D = 2
            s.saW.pvData = VarPtr(L)
            
            myHash = (myHash + s.W(0)) * HMul And s.HTUB 'loop-unrolling (we have only 2 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HTUB
         
        Case vbCurrency
        
            C = Key
            s.saW.cElements1D = 4
            s.saW.pvData = VarPtr(C)
            
            myHash = (myHash + s.W(0)) * HMul And s.HTUB 'loop-unrolling (we have only 4 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HTUB
            myHash = (myHash + s.W(2)) * HMul And s.HTUB
            myHash = (myHash + s.W(3)) * HMul And s.HTUB
         
        Case vbLong, vbInteger, vbByte
        
            L = Key
            s.saW.cElements1D = 2
            s.saW.pvData = VarPtr(L)
            
            myHash = (myHash + s.W(0)) * HMul And s.HTUB 'loop-unrolling (we have only 2 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HTUB
            
        Case vbDouble
        
            D = Key: s.saW.cElements1D = 4
            s.saW.pvData = VarPtr(D)
            
            myHash = (myHash + s.W(0)) * HMul And s.HTUB 'loop-unrolling (we have only 4 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HTUB
            myHash = (myHash + s.W(2)) * HMul And s.HTUB
            myHash = (myHash + s.W(3)) * HMul And s.HTUB
            
        Case vbDate
        
            D = Key
            s.saW.cElements1D = 4
            s.saW.pvData = VarPtr(D)
            
            myHash = (myHash + s.W(0)) * HMul And s.HTUB 'loop-unrolling (we have only 4 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HTUB
            myHash = (myHash + s.W(2)) * HMul And s.HTUB
            myHash = (myHash + s.W(3)) * HMul And s.HTUB

        Case vbSingle
        
            F = Key
            s.saW.cElements1D = 2
            s.saW.pvData = VarPtr(F)
            myHash = (myHash + s.W(0)) * HMul And s.HTUB 'loop-unrolling (we have only 2 16Bit integers in the array)
            myHash = (myHash + s.W(1)) * HMul And s.HTUB
            
    End Select
     
    'Debug.Print Key, myHash + 1
    GetHash = myHash + 1
    
End Function


'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub RemoveNode(ByRef ipNode As KvpHNode)

    ' when removing a node we must also remove the pointer registered in s.HashSlots
    ' consequently we must replace the node pointed to by s.HashSlots(hash) with
    ' s.HashSlots.succ or nothing
    ' It is simpler to check if ipnode is the root or last node
    
    ' check for the case of a single node
    If s.Count = 1 Then
        Set s.HashSlots(ipNode.Hash) = Nothing
        Set ipNode = Nothing
        s.Count = s.Count - 1
        Exit Sub
    End If
    
    ' there are at least two nodes in the list
    ' check for last node
    If ipNode Is s.Last Then
    
        Set s.Last = ipNode.Pred
        Set s.Last.Succ = Nothing
        ' as there are no successor node we don't need to any hash checking
        ' just remove the node from s.HashSlots
        Set s.HashSlots(ipNode.Hash) = Nothing
        Set ipNode = Nothing
        s.Count = s.Count - 1
        Exit Sub
        
    End If
    
    ' check for first node
    If ipNode Is s.Root Then
    
        Set s.Root = ipNode.Succ
        Set s.Root.Pred = Nothing
        ' there is a successor node so we need to check if the hash is the same
        ' and replace the reference in s.HashSlots if it is
        If ipNode.Hash = ipNode.Succ.Hash Then
            Set s.HashSlots(ipNode.Hash) = ipNode.Succ
        End If
        Set ipNode = Nothing
        s.Count = s.Count - 1
        Exit Sub
    End If
    
    ' the node wasn't the first or last node so it has a pred and a succ
    ' so we need to link the link the pred and succ nodes
    Set ipNode.Pred.Succ = ipNode.Succ
    Set ipNode.Succ.Pred = ipNode.Pred
    
    ' check for same hash in succ
    If ipNode.Hash = ipNode.Succ.Hash Then
        Set s.HashSlots(ipNode.Hash) = ipNode.Succ
    End If
    
    Set ipNode = Nothing
    s.Count = s.Count - 1
    
End Sub
