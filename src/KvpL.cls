VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "KvpL is a Key/Value pair dictionary where the keys and values are held in a double linked list"
'@PredeclaredId
'@Exposed
'@Folde(Kvp")
'@ModuleDescription "KvpL is a Key/Value pair dictionary where the keys and values are held in a double linked list"
Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========


Private Const ALLOW_DUPLICATE_KEYS              As Boolean = False
'@Ignore ConstantNotUsed
Private Const FORBID_DUPLICATE_KEYS             As Boolean = True

Private Const FIRST_INDEX                       As Long = 1
Private Const MY_TYPENAME                       As String = "KvpL"

Private Type State
 
    ' first node added to the seq
    Root                                        As KvpLNode
    ' last node appended to the seq (note not added)
    Last                                        As KvpLNode

End Type


Private s                                       As State


Private Type Properties
    EnsureUniqueKeys                            As Boolean          ' True to enforce unique keys
    Count                                       As Long             ' the current number of nodes
    FirstIndex                                  As Long
End Type

Private p                                       As Properties



'@DefaultMember
Public Function Deb(Optional ByVal ipUniqueKeys As Boolean = ALLOW_DUPLICATE_KEYS) As KvpL
Attribute Deb.VB_UserMemId = 0
    With New KvpL
        Set Deb = .ConstructInstance(ipUniqueKeys)
    End With
End Function


Friend Function ConstructInstance(ByVal ipUniqueKeys As Boolean) As KvpL

    p.FirstIndex = FIRST_INDEX
    Set s.Root = Nothing
    Set s.Last = Nothing
    p.Count = 0
    p.EnsureUniqueKeys = ipUniqueKeys
    Set ConstructInstance = Me
    
End Function


Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpL

    Set Add = Me
    
    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Add"
    Guard.EnsureUniqueKeys ipKey, Me, "Add"

    Dim myNewNode As KvpLNode: Set myNewNode = KvpLNode(ipKey, ipItem)
    
    If Me.LacksItems Then
       SetupFirstNode myNewNode
    Else
        AppendNode myNewNode
    End If
   
    Set Add = Me
    
End Function


Public Function AddPairs(ByRef ipKeys As Variant, ByRef ipItems As Variant) As KvpL
   
    Set AddPairs = Me
    
    Dim myKeys As IterItems: Set myKeys = IterItems(ipKeys)
    Dim myItems As IterItems: Set myItems = IterItems(ipItems)
    
    If myKeys.LacksItems Or myItems.LacksItems Then
        Exit Function
    End If
     
    Do
        Me.Add myKeys.CurItem(0), myItems.CurItem(0)
    Loop While myKeys.MoveNext And myItems.MoveNext
   
    Set AddPairs = Me
   
End Function


Public Property Get Item(ByRef ipKey As Variant) As Variant

    Item = Null
    
    If Me.LacksItems Then
        Exit Property
    End If
                   
    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Get Item"
    
    ' Kvp classes do not permit adding a key by assignment
    Dim myNode As KvpLNode: Set myNode = GetNodeByKey(ipKey)
    If Guard.KeyNotFound(myNode, ipKey, MY_TYPENAME & ".Get Item", REPORT_BACK) Then
        Exit Property
    End If
    
    If VBA.IsObject(myNode.Item) Then
        Set Item = myNode.Item
    Else
        Item = myNode.Item
    End If
    
End Property


Public Property Let Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Get Item"

    ' Kvp classes do not permit adding a key by assignment
    Dim myNode As KvpLNode: Set myNode = GetNodeByKey(ipKey)
    
    Guard.KeyNotFound myNode, ipKey, MY_TYPENAME & ".Let Item"
    
    myNode.Item = ipItem
    
End Property


Public Property Set Item(ByRef ipKey As Variant, ByVal ipItem As Variant)
 
    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Get Item"
    
    Dim myNode As KvpLNode: Set myNode = GetNodeByKey(ipKey)
    
    Guard.KeyNotFound myNode, ipKey, MY_TYPENAME & ".Let Item"
    
    Set myNode.Item = ipItem
    
End Property


Public Property Get ItemByIndex(ByVal ipIndex As Long) As Variant

    ItemByIndex = Null
    
    If Me.LacksItems Then
        Exit Property
    End If
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "Get ItemByIndex", REPORT_BACK) Then
        Exit Property
    End If
    
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(ipIndex)
    
    If VBA.IsObject(myNode.Item) Then
        Set ItemByIndex = myNode.Item
    Else
        ItemByIndex = myNode.Item
    End If
    
End Property


Public Property Let ItemByIndex(ByVal ipIndex As Long, ByVal ipItem As Variant)

    Guard.IndexOutOfBounds ipIndex, Me, "Let ItemByIndex"

    Dim myNode As KvpLNode: Set myNode = GetNodeAt(ipIndex)
    
    myNode.Item = ipItem
   
End Property


Public Property Set ItemByIndex(ByVal ipIndex As Long, ByVal ipItem As Variant)

    Guard.IndexOutOfBounds ipIndex, Me, "Let ItemByIndex"
    
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(ipIndex)
    
    Set myNode.Item = ipItem
    
End Property

Public Property Get ItemByOffset(ByVal ipoffset As Long) As Variant

    ItemByOffset = Null
    
    If Me.LacksItems Then
        Exit Property
    End If
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    If Guard.IndexOutOfBounds(myIndex, Me, "Let ItemByIndex", REPORT_BACK) Then
        Exit Property
    End If
    
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(myIndex)
    
    If VBA.IsObject(myNode.Item) Then
        Set ItemByOffset = myNode.Item
    Else
        ItemByOffset = myNode.Item
    End If
    
End Property


Public Property Let ItemByOffset(ByVal ipoffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset

    Guard.IndexOutOfBounds myIndex, Me, "Let ItemByIndex"
    
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(myIndex)
    
    myNode.Item = ipItem
   
End Property


Public Property Set ItemByOffset(ByVal ipoffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset

    Guard.IndexOutOfBounds myIndex, Me, "Set ItemByIndex"
    
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(myIndex)
    
    Set myNode.Item = ipItem
    
End Property

Public Function IndexByKey(ByRef ipKey As Variant) As Long

    IndexByKey = -1
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".Indexbykey", REPORT_BACK) Then
        Exit Function
    End If
    
    IndexByKey = GetNodeByKey(ipKey).User
    
End Function

'@Ignore WriteOnlyProperty
Public Property Let Key(ByRef ipKey As Variant, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Let Key"
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let Key"

    Dim myNode As KvpLNode: Set myNode = GetNodeByKey(ipKey)
    Guard.KeyNotFound myNode, ipKey, MY_TYPENAME & ".Let Key"
    
    myNode.Key = ipNewKey

End Property


Public Property Set Key(ByRef ipKey As Variant, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Set Key"
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set Key"

    Dim myNode As KvpLNode: Set myNode = GetNodeByKey(ipKey)
    Guard.KeyNotFound myNode, ipKey, MY_TYPENAME & ".Set Key"
    
    Set myNode.Key = ipNewKey
     
End Property

Public Property Get KeyByIndex(ByRef ipIndex As Long) As Variant

    KeyByIndex = Null
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "Let ItemByIndex", REPORT_BACK) Then
        Exit Property
    End If
                        
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(ipIndex)
    
    If VBA.IsObject(myNode.Key) Then
        Set KeyByIndex = myNode.Key
    Else
        KeyByIndex = myNode.Key
    End If
    
End Property


Public Property Let KeyByIndex(ByRef ipIndex As Long, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let KeyByIndex"
    Guard.IndexOutOfBounds ipIndex, Me, "Let ItemByIndex"
    Guard.EnsureUniqueKeys Me, ipNewKey, "Set KeyByIndex"
    
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(ipIndex)

    myNode.Key = ipNewKey
   
End Property


Public Property Set KeyByIndex(ByRef ipIndex As Long, ByVal ipNewKey As Variant)
    
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set KeyByIndex"
    Guard.IndexOutOfBounds ipIndex, Me, "Set KeyByIndex"
    Guard.EnsureUniqueKeys Me, ipNewKey, "Set KeyByIndex"
                
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(ipIndex)
    
    Set myNode.Key = ipNewKey
   
End Property


Public Property Get KeyByOffset(ByRef ipoffset As Long) As Variant

    KeyByOffset = Null
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    If Guard.IndexOutOfBounds(myIndex, Me, "Get KeyByOffset", REPORT_BACK) Then
        Exit Property
    End If
                   
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(myIndex)
    
    If VBA.IsObject(myNode.Key) Then
        Set KeyByOffset = myNode.Key
    Else
        KeyByOffset = myNode.Key
    End If
    
End Property


Public Property Let KeyByOffset(ByRef ipoffset As Long, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let KeyByOffset"
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    Guard.IndexOutOfBounds myIndex, Me, "Let KeyByOffset"
    Guard.EnsureUniqueKeys ipNewKey, Me, "Set KeyByIndex"
    
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(myIndex)
    
    myNode.Key = myNode.Key
    
End Property


Public Property Set KeyByOffset(ByRef ipoffset As Long, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set KeyByOffset"

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    Guard.IndexOutOfBounds myIndex, Me, "Set KeyByOffset"
    Guard.EnsureUniqueKeys Me, ipNewKey, "Set KeyByIndex"
    
    Dim myNode As KvpLNode: Set myNode = GetNodeAt(myIndex)
    
    Set myNode.Key = ipNewKey
   
End Property
'@Description("Removes the first found key/item pair for each key in the paramarray")
Public Function Remove(ParamArray ipParamarray() As Variant) As KvpL
Attribute Remove.VB_Description = "Removes the first found key/item pair for each key in the paramarray"

    Set Remove = Me
    
    Dim myParamarray As Variant: myParamarray = ipParamarray

    If ArrayOp.LacksItems(myParamarray) Then
        Exit Function
    End If
    
    '@Ignore AssignmentNotUsed
    myParamarray = ArrayOp.Splat(myParamarray)
    
    Set Remove = RemoveRange(myParamarray)
    
End Function

'@Description("Removes the first found key/item pair for each key in the range parameter")
Public Function RemoveRange(ByRef ipRange As Variant) As KvpL
Attribute RemoveRange.VB_Description = "Removes the first found key/item pair for each key in the range parameter"

    Set RemoveRange = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myKeys As IterItems: Set myKeys = IterItems(ipRange)
    
    If myKeys.LacksItems Then
        Exit Function
    End If
    
    Do
    
        If Guard.KeyIsAdmin(myKeys.CurItem(0), MY_TYPENAME & ".RemoveRange", REPORT_BACK) Then
            GoTo ContinueLoop
        End If
        
        RemoveNode GetNodeByKey(myKeys.CurItem(0))
        
ContinueLoop:
    Loop While myKeys.MoveNext
    
    Set RemoveRange = Me
    
End Function

'@Description("Remove ipRun key/value pairs after the last of  /Key pairs
Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpL

    Set RemoveAfter = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".RemoveRange", REPORT_BACK) Then
        Exit Function
    End If
    
    If Me.LacksKey(ipKey) Then
        Exit Function
    End If
    
    If Guard.InvalidRun(ipRun, MY_TYPENAME & ".RemoveAfter", REPORT_BACK) Then
        Exit Function
    End If
    
    Dim myNode As KvpLNode: Set myNode = GetNodeByKey(ipKey)
    
   ' we are at the first node of the  key pair
    Dim myCount As Long
    For myCount = 1 To ipRun
    
        If myNode.Succ Is Nothing Then
            Exit Function
        End If
        
        RemoveNode myNode.Succ
        
    Next
    
End Function

'@Description("Remove key/item pairs before the specified Key
Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpL

    Set RemoveBefore = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".RemoveBefore", REPORT_BACK) Then
        Exit Function
    End If
    
    If s.Root Is Nothing Then
        Exit Function
    End If
    
    If Guard.InvalidRun(ipRun, MY_TYPENAME & ".RemoveBefore", REPORT_BACK) Then
        Exit Function
    End If
    
    Dim myNode As KvpLNode: Set myNode = GetNodeByKey(ipKey)
    
    ' we are at the first node of the /key pair
    Dim myCount As Long
    For myCount = 1 To ipRun
    
        If myNode.Pred Is Nothing Then
            Exit Function
        End If
        
        RemoveNode myNode.Pred
        
    Next
    
End Function

Public Function RemoveAll(ParamArray ipParamarray() As Variant) As KvpL

    Set RemoveAll = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myParamarray As Variant: myParamarray = ipParamarray 'eliminates the need for vba.cvar()
    
    If ArrayOp.LacksItems(myParamarray) Then
        RemoveAllItems
        Exit Function
    End If
    
    '@Ignore AssignmentNotUsed
    myParamarray = ArrayOp.Splat(myParamarray)
  
    Dim myKey As Variant
    For Each myKey In myParamarray
        RemoveAllSingleItem myKey
    Next
    
    Set RemoveAll = Me

End Function

Public Function Clear() As KvpL
    
    Dim myNode As KvpLNode: Set myNode = s.Last
    
    Do Until myNode.Pred Is Nothing
        Set myNode = myNode.Pred
        RemoveNode myNode.Succ
    Loop
    
    RemoveFirstNode
    
    Set Clear = Me
    
End Function


Public Function Reset() As KvpL
    Set Reset = Clear
End Function


Public Function RemoveByIndex(ParamArray ipParamarray() As Variant) As KvpL

    Set RemoveByIndex = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myParamarray As Variant: myParamarray = ipParamarray
    
    If ArrayOp.LacksItems(myParamarray) Then
        Exit Function
    End If
    
    '@Ignore AssignmentNotUsed
    myParamarray = ArrayOp.Splat(myParamarray)
    
    Set RemoveByIndex = RemoveRangeByIndex(myParamarray)
    
End Function

'@Description("Removes the Items at the indexes in the range list")
Public Function RemoveRangeByIndex(ByRef ipRange As Variant) As KvpL
Attribute RemoveRangeByIndex.VB_Description = "Removes the Items at the indexes in the range list"

    Set RemoveRangeByIndex = Me
    
    If Me.LacksItems Then
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
    
        Dim myIndex As Long: myIndex = myIndexes.CurItem(0)
        
        If GroupInfo.IsNotNumber(myIndex) Then
            GoTo ContinueLoop:
        End If
        
        If Guard.IndexOutOfBounds(myIndex, Me, "RemoveRangeByIndex", REPORT_BACK) Then
            GoTo ContinueLoop
        End If
    
        Dim myNode As KvpLNode: Set myNode = GetNodeAt(myIndex)
        RemoveNode myNode
        
ContinueLoop:
    Loop While myIndexes.MovePrev
    
    Set RemoveRangeByIndex = Me
    
End Function



Public Function InsertAt(ByRef ipIndex As Long, ByVal ipKey As Variant, ByVal ipItem As Variant) As Long

    InsertAt = -1
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "InsertAt", Guard.REPORT_BACK) Then
        Exit Function
    End If
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".InsertAt", Guard.REPORT_BACK) Then
        Exit Function
    End If
    
    
    If Guard.EnsureUniqueKeys(ipKey, Me, "InsertAt", Guard.REPORT_BACK) Then
        Exit Function
    End If
    

    ' create the new node
    Dim myNewNode As KvpLNode: Set myNewNode = KvpLNode.Deb(ipKey, ipItem)
    
    InsertAt = InsertNode(ipIndex, myNewNode)
    
End Function
Public Function Clone() As KvpL

    Dim myK As KvpL: Set myK = KvpL.Deb
    
    Dim myNode As KvpLNode: Set myNode = s.Root
    
    Do Until myNode Is Nothing
        myK.Add myNode.Key, myNode.Item
        Set myNode = myNode.Succ
    Loop
    
    Set Clone = myK
    
End Function


Public Function Exists(ByRef ipKey As Variant) As Boolean

    Exists = False
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".Exists", REPORT_BACK) Then
        Exit Function
    End If
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    
    ' found a matching , now check for matching key
    Dim myNode As KvpLNode: Set myNode = GetNodeByKey(ipKey)
    
    Exists = Not (myNode Is Nothing)
    
End Function


Public Function HoldsKey(ByRef ipKey As Variant) As Boolean
    HoldsKey = Me.Exists(ipKey)
End Function


Public Function ContainsKey(ByRef ipKey As Variant) As Boolean
    ContainsKey = Me.Exists(ipKey)
End Function

Public Function LacksKey(ByRef ipKey As Variant) As Boolean
    LacksKey = Not Me.Exists(ipKey)
End Function

'Public Function ContainsKey(ByRef ipKey As Variant) As Boolen
'    ContainsKey = Me.Exists(ipKey)
'End Function


Public Function ExistsItem(ByRef ipItem As Variant) As Boolean

    ExistsItem = False

    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As KvpLNode: Set myNode = GetNodeByItem(ipItem)
    
    ExistsItem = Not (myNode Is Nothing)
    
End Function


Public Function HoldsItem(ByRef ipItem As Variant) As Boolean
    HoldsItem = Me.ExistsItem(ipItem)
    
End Function


Public Function Contains(ByRef ipItem As Variant) As Boolean
    Contains = Me.ExistsItem(ipItem)
End Function


Public Function LacksItem(ByRef ipItem As Variant) As Boolean
    LacksItem = Not ExistsItem(ipItem)
End Function


Public Function HoldsItems() As Boolean
    HoldsItems = Not s.Root Is Nothing
End Function


Public Function LacksItems() As Boolean
    LacksItems = s.Root Is Nothing
End Function


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property


Public Function HasUniqueKeys() As Boolean
    HasUniqueKeys = p.EnsureUniqueKeys
End Function

'@Description("Mutates the KvpA in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As KvpL
Attribute MappedIt.VB_Description = "Mutates the KvpA in place by applying the function in IMapper to each item"

    If Me.LacksItems Then
        Set MappedIt = Me
        Exit Function
    End If
    
    Dim myNode As KvpLNode: Set myNode = s.Root
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
Public Function MapIt(ByVal ipMapper As IMapper) As KvpL
Attribute MapIt.VB_Description = "Returns a KvpA.Deb by applying the function in IMapper to each Item"

    Dim myK As KvpL: Set myK = KvpL.Deb
    Set MapIt = myK
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
    
        Dim myItem As Variant
        myItem = ipMapper.ExecMapper(myNode.Item)
        
        myK.Add myNode.Key, myItem(0)
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
      
    Set MapIt = myK

End Function


'@Description("Returns a KvpA.Deb where the ICOmparer has been applied to the Items of the original Kvp")
Public Function FilterIt(ByVal ipComparer As IComparer) As KvpL
Attribute FilterIt.VB_Description = "Returns a KvpA.Deb where the ICOmparer has been applied to the Items of the original Kvp"

    Dim myK As KvpL: Set myK = KvpL.Deb
    Set FilterIt = myK
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
    
        If ipComparer.ExecCmp(myNode.Item) Then
            myK.Add myNode.Key, myNode.Item
        End If
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
      
    Set FilterIt = myK

End Function


Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant

    ReduceIt = Null
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myA() As Variant: myA = ipReducer.ExecReduction(Me.Items)
    
    If VBA.IsObject(myA(0)) Then
        Set ReduceIt = myA(0)
    Else
        ReduceIt = myA(0)
    End If
    
End Function


'@Description("Returns the size of the Kvp or the number of specified Items.  Returns -1 rather than 0 if there are no items in the kvp but count item will return 0 if not found")
Public Function Count(Optional ByRef ipItem As Variant) As Long
Attribute Count.VB_Description = "Returns the size of the Kvp or the number of specified Items.  Returns -1 rather than 0 if there are no items in the kvp but count item will return 0 if not found"
    
    If s.Root Is Nothing Then
        Count = -1
        Exit Function
    End If
    
    If VBA.IsMissing(ipItem) Then
        Count = p.Count
        Exit Function
    End If
    
    Dim myCount As Long: myCount = 0
    
    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
    
        If Comparers.EQ(ipItem, myNode.Item) Then
            myCount = myCount + 1
        End If
        
        Set myNode = myNode.Succ
    
    Loop Until myNode Is Nothing
    
    Count = myCount
    
End Function


Public Function CountIt(ByVal ipComparer As IComparer) As Long

    CountIt = -1
    
    If s.Root Is Nothing Then
        Exit Function
    End If

    Dim myCount As Long: myCount = 0
    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
    
        If ipComparer.ExecCmp(myNode.Item) Then
            myCount = myCount + 1
        End If
    
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing

    CountIt = myCount

End Function


Public Function KeyByItem(ByVal ipItem As Variant) As Variant

    Dim myNode As KvpLNode: Set myNode = GetNodeByItem(ipItem)
     
    If VBA.IsObject(myNode.Item) Then
        Set KeyByItem = myNode.Key
    Else
        KeyByItem = myNode.Key
    End If
    
End Function

'@Description("Returns a string representation of the Items")
Public Function Join(Optional ByRef ipSeparator As String = Chars.twComma) As String
    Join = Strs.Inner(Fmt.ItemSeparator(ipSeparator).Text("{0}", Me.Items))
End Function

Public Function Items() As Variant

    Items = Array()
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myA As Variant: ReDim myA(Me.FirstIndex To Me.LastIndex)
    Dim myIndex As Long: myIndex = 1
    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
        
        If VBA.IsObject(myNode.Item) Then
            Set myA(myIndex) = myNode.Item
        Else
            myA(myIndex) = myNode.Item
        End If

        Set myNode = myNode.Succ
        myIndex = myIndex + 1
        
    Loop Until myNode Is Nothing

    Items = myA
    
End Function

Public Function ItemsAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    Set ItemsAsColl = myC
    
    If Me.LacksItems Then
        Exit Function
    End If


    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
        
        myC.Add myNode.Item
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing

    Set ItemsAsColl = myC
    
End Function

Public Function ItemsAsSeq() As SeqL
    
    Dim myS As SeqL: Set myS = SeqL.Deb
    Set ItemsAsSeq = myS
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
    
        myS.Add myNode.Item
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
    
    Set ItemsAsSeq = myS
    
End Function


Public Function Keys() As Variant
    
    Keys = Array()
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myA As Variant: ReDim myA(Me.FirstIndex To Me.LastIndex)
    Dim myIndex As Long: myIndex = 1
    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
        
        If VBA.IsObject(myIndex) Then
            Set myA(myIndex) = myNode.Key
        Else
            myA(myIndex) = myNode.Key
        End If

        Set myNode = myNode.Succ
        myIndex = myIndex + 1
        
    Loop Until myNode Is Nothing

    Keys = myA
    
End Function

Public Function KeysAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    Set KeysAsColl = myC
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
        
        myC.Add myNode.Key
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing

    Set KeysAsColl = myC
    
End Function

'@Description("Returns a seq of keys.")
Public Function KeysAsSeq() As SeqL
Attribute KeysAsSeq.VB_Description = "Returns a seq of keys."

    Dim myS As SeqL: Set myS = SeqL.Deb
    
    If Me.LacksItems Then
        Set KeysAsSeq = myS
        Exit Function
    End If
    
    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
    
        myS.Add myNode.Key

        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
    
    Set KeysAsSeq = myS
    
End Function


Public Function Pairs() As Variant
    
    Pairs = Array()

    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myA As Variant: ReDim myA(Me.FirstIndex To Me.LastIndex)
    Dim myIndex As Long: myIndex = 1
    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
    
        myA(myIndex) = Array(myNode.Key, myNode.Item)
        
        Set myNode = myNode.Succ
        myIndex = myIndex = 1
        
    Loop Until myNode Is Nothing
    
    Pairs = myA
    
End Function

Public Function PairsAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    Set PairsAsColl = myC
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
        
        myC.Add Array(myNode.Key, myNode.Item)
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing

    Set PairsAsColl = myC
    
End Function

Public Function PairsAsSeq() As SeqL

    Dim myS As SeqL: Set myS = SeqL.Deb
    Set PairsAsSeq = myS
    
    If myS.LacksItems Then
        Exit Function
    End If
    
    Dim myNode As KvpLNode: Set myNode = s.Root
    Do
    
        myS.Add Array(myNode.Key, myNode.Item)
        
        Set myNode = myNode.Succ
    
    Loop Until myNode Is Nothing
    
    Set PairsAsSeq = myS

End Function



Public Property Get FirstIndex() As Long
    
    If Me.LacksItems Then
        FirstIndex = -1
    Else
        FirstIndex = p.FirstIndex
    End If
    
End Property


Public Property Get LastIndex() As Long

    If Me.LacksItems Then
        LastIndex = -1
    Else
        LastIndex = p.Count
    End If
End Property

'@Description("Returns the index of the second item in the kvp or -1  if the kvp has less than 2 Items")
Public Function FBOIndex() As Long
Attribute FBOIndex.VB_Description = "Returns the index of the second item in the kvp or -1  if the kvp has less than 2 Items"

    If Me.LacksItems Then
        FBOIndex = -1
    ElseIf s.Root.Succ Is Nothing Then
        FBOIndex = -1
    Else
        FBOIndex = p.FirstIndex + 1
    End If
    
End Function


'Description("Returns the index of the second to last item in the SeqH, or -1 if the SeqH has lees than 2 items")
Public Function LBOIndex() As Long

    If s.Root Is Nothing Then
        LBOIndex = -1
    ElseIf s.Last.Pred Is Nothing Then
        LBOIndex = -1
    Else
        LBOIndex = p.Count - 1
    End If
   
End Function

Public Property Get FirstKey() As Variant

    FirstKey = Null
    
    If Me.LacksItems Then
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
    
    If LacksItems Then
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
    
    If Me.LacksItems Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Root.Item) Then
        Set First = s.Root.Item
    Else
        First = s.Root.Item
    End If
    
End Property


Public Property Set First(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Set First"
    
    Set s.Root.Item = ipItem

End Property


Public Property Let First(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Let First"
    
    s.Root.Item = ipItem

End Property


Public Property Get Last() As Variant

    Last = Null
    
    If Me.LacksItems Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Last.Item) Then
        Set Last = s.Last.Item
    Else
        Last = s.Last.Item
    End If
    
End Property


Public Property Set Last(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Set Last"

    Set s.Last.Item = ipItem
    
End Property


Public Property Let Last(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Let Last"
    
    s.Last.Item = ipItem
    
End Property

Public Property Get FBOKey() As Variant

    FBOKey = Null
    
    If s.Root Is Nothing Then
        Exit Property
    ElseIf s.Root.Succ Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Root.Key) Then
        Set FBOKey = s.Root.Key
    Else
        FBOKey = s.Root.Key
    End If
    
End Property


Public Property Get LBOKey() As Variant

    LBOKey = Null
    
    If s.Root Is Nothing Then
        Exit Property
    ElseIf s.Last.Pred Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Last.Pred.Key) Then
        Set LastKey = s.Last.Pred.Key
    Else
        LastKey = s.Last.Pred.Key
    End If
    
End Property


Public Property Get FBO() As Variant

    FBO = Null
    
    If Me.LacksItems Then
        Exit Property
    ElseIf s.Root.Succ Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Root.Succ.Item) Then
        Set First = s.Root.Succ.Item
    Else
        FBO = s.Root.Succ.Item
    End If
    
End Property


Public Property Set FBO(ByVal ipItem As Variant)
    
    Guard.KeyNotFound s.Root, Nothing, MY_TYPENAME & ".Set FBO"
    Guard.KeyNotFound s.Root.Succ, Nothing, MY_TYPENAME & ".Set FBO"
    
    Set s.Root.Succ.Item = ipItem

End Property


Public Property Let FBO(ByVal ipItem As Variant)

    Guard.KeyNotFound s.Root, Nothing, MY_TYPENAME & ".Let FBO"
    Guard.KeyNotFound s.Root.Succ, Nothing, MY_TYPENAME & ".Let FBO"
    
    s.Root.Succ.Item = ipItem

End Property


Public Property Get LBO() As Variant

    FBO = Null
    
    If Me.LacksItems Then
        Exit Property
    ElseIf s.Last.Pred Is Nothing Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Last.Item) Then
        Set LBO = s.Last.Pred.Item
    Else
        LBO = s.Last.Pred.Item
    End If
    
End Property


Public Property Set LBO(ByVal ipItem As Variant)

    Guard.KeyNotFound s.Root, Nothing, MY_TYPENAME & ".Set LBO"
    Guard.KeyNotFound s.Root.Succ, Nothing, MY_TYPENAME & ".Set LBO"
    
    Set s.Last.Pred.Item = ipItem
    
End Property


Public Property Let LBO(ByVal ipItem As Variant)

    Guard.KeyNotFound s.Root, Nothing, MY_TYPENAME & ".Set FBO"
    Guard.KeyNotFound s.Root.Succ, Nothing, MY_TYPENAME & ".Set FBO"
    
    s.Last.Pred.Item = ipItem

End Property

Public Property Get EnsureUniqueKeys() As Boolean
    EnsureUniqueKeys = p.EnsureUniqueKeys
End Property

Public Property Let EnsureUniqueKeys(ByVal ipEnsureUniqueKeys As Boolean)
    p.EnsureUniqueKeys = ipEnsureUniqueKeys
End Property

'@Description( returns a Kvp with values as keys.  Each Value is linked to a Kvp of Key vs Index for the host value
Public Function Mirror(Optional ByVal ipMirrorType As e_MirrorType = e_MirrorType.m_ByAllValues) As KvpL

    If Me.LacksItems = 0 Then
        Set Mirror = KvpL.Deb
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



'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'Private functions
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'@Description("Returns an KvpH of two KvpH.  Item(1) is an KvpH of the first found Host value vs Host Key. .Item(2) is a copy of the original KvpL with the  first Key/Value pairs removed
Private Function MirrorByAllValues() As KvpL

    Dim myK As KvpL
    Set myK = KvpL.Deb
    
    If s.Root Is Nothing Then
        Set MirrorByAllValues = KvpL.Deb
        Exit Function
    End If

    Dim myNode As KvpLNode: Set myNode = s.Root
    Dim myIndex As Long: myIndex = 1
    Do
        If myK.LacksKey(myNode.Item) Then
            myK.Add myNode.Item, KvpLP.Deb
        End If

        myK.Item(myNode.Item).Add myNode.Item, myIndex
        myIndex = myIndex + 1
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing

    Set MirrorByAllValues = myK

End Function


Private Function MirrorByFirstValue() As KvpL

    Dim myFirst As KvpL: Set myFirst = KvpL.Deb
    Dim myOther As KvpL: Set myOther = KvpL.Deb

    Dim myPairs As Variant: Set myPairs = Me.Pairs
    
    Dim myPair As Variant
    For Each myPair In myPairs

        If myFirst.LacksKey(myPair(1)) Then
            myFirst.Add myPair(1), myPair(0)
        Else
            myOther.Add myPair(0), myPair(1)
        End If

    Next

    Set MirrorByFirstValue = KvpL.Deb.Add(0&, myFirst).Add(1&, myOther)

End Function

Private Sub SetupFirstNode(ByRef ipNode As KvpLNode)

        Set s.Root = ipNode
        Set s.Last = ipNode
        
        p.Count = 1
        
End Sub
Private Sub AppendNode(ByRef ipNode As KvpLNode)

    ' The append node method adds a node s.Last
    ' link nodes following the
    ' If s.Slots(ipNode.) Is Nothing Then
    '     Set s.Slots(ipNode.) = ipNode
    ' End If
    
    Set ipNode.Pred = s.Last
    Set s.Last.Succ = ipNode
    Set s.Last = ipNode
    
    p.Count = p.Count + 1
    
End Sub

Private Function InsertNode(ByVal ipIndex As Long, ByVal ipNewNode As KvpLNode) As Long

    ' we know that the new node is not the first or last node.
    ' we know that the new node exists in slots
    ' we insert the newnode before the node with the same
    InsertNode = ipIndex
    
    If ipIndex = p.FirstIndex Then
        PrependNode ipNewNode
        Exit Function
    End If
    
    If ipIndex = p.Count Then
        AppendNode ipNewNode
        Exit Function
    End If
  
    Dim myCurNode As KvpLNode: Set myCurNode = GetNodeAt(ipIndex)
    
    Set ipNewNode.Pred = myCurNode.Pred
    Set ipNewNode.Succ = myCurNode
    
    Set myCurNode.Pred.Succ = ipNewNode
    Set myCurNode.Pred = ipNewNode
    
    p.Count = p.Count + 1
    
End Function

Private Sub PrependNode(ByRef ipNode As KvpLNode)

    Set ipNode.Succ = s.Root
    Set s.Root.Pred = ipNode
    Set s.Root = ipNode
    
    p.Count = p.Count + 1
   
End Sub

Private Sub RemoveFirstNode()

'    '@Ignore AssignmentNotUsed, VariableNotUsed
'    Dim myTmp As KvpLNode: Set myTmp = s.Root
'    Set s.Root = s.Root.Succ
'    Set s.Root.Pred = Nothing
'    Set myTmp = Nothing
    Set s.Root = Nothing
    p.Count = p.Count - 1

End Sub

Private Sub RemoveLastNode()

    '@Ignore AssignmentNotUsed, VariableNotUsed
    Dim myTmp As KvpLNode: Set myTmp = s.Last
    Set s.Last = s.Last.Pred
    Set s.Last.Succ = Nothing
    Set myTmp = Nothing
    
    p.Count = p.Count - 1
    
End Sub

'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub RemoveNode(ByRef ipNode As KvpLNode)
 
    If ipNode Is s.Root Then
        Set s.Root = Nothing
        Exit Sub
    End If
    
    If ipNode Is s.Last Then
        RemoveLastNode
        Exit Sub
    End If
    
    '@Ignore AssignmentNotUsed, VariableNotUsed
    Dim myTmp As KvpLNode: Set myTmp = ipNode
    Set ipNode.Pred.Succ = ipNode.Succ
    Set ipNode.Succ.Pred = ipNode.Pred
    Set myTmp = Nothing
    
    p.Count = p.Count - 1
    
End Sub

Private Function GetNodeByKey(ByRef ipKey As Variant) As KvpLNode

    Set GetNodeByKey = Nothing
    
    Dim myNode As KvpLNode: Set myNode = s.Root
    Dim myIndex As Long: myIndex = 1
    Do Until myNode Is Nothing
    
        If Comparers.EQ(myNode.Key, ipKey) Then
            myNode.User = myIndex
            Set GetNodeByKey = myNode
            Exit Function
        End If
        
        Set myNode = myNode.Succ
        myIndex = myIndex + 1
        
    Loop
    
End Function

Private Function GetNodeByItem(ByRef ipItem As Variant) As KvpLNode

    Set GetNodeByItem = Nothing

    Dim myNode As KvpLNode: Set myNode = s.Root
    Dim myIndex As Long: myIndex = 1
    Do Until myNode Is Nothing
    
        If Comparers.EQ(myNode.Item, ipItem) Then
            myNode.User = myIndex
            Set GetNodeByItem = myNode
            Exit Function
        End If
        
        Set myNode = myNode.Succ
        myIndex = myIndex = 1
        
    Loop
    
End Function


'@Description("Returns the node where . = ipIndex")
Private Function GetNodeAt(ByRef ipIndex As Long) As KvpLNode
Attribute GetNodeAt.VB_Description = "Returns the node where . = ipIndex"

    Set GetNodeAt = Nothing
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "Private GetNodeAt", Guard.REPORT_BACK) Then
        Exit Function
    End If
    
    If p.Count \ 2 > ipIndex Then
        Set GetNodeAt = GetNodeFromRoot(ipIndex)
    Else
        Set GetNodeAt = GetNodeFromLast(ipIndex)
    End If

End Function

'@Description("Reurns the node at ipIndex.  The search starts at the node at s.root and follows the succ nodes")
Private Function GetNodeFromRoot(ByVal ipIndex As Long) As KvpLNode
Attribute GetNodeFromRoot.VB_Description = "Reurns the node at ipIndex.  The search starts at the node at s.root and follows the succ nodes"
            
    Dim myNode As KvpLNode: Set myNode = s.Root
    
    Dim myIndex As Long:
    Dim myMaxIndex As Long: myMaxIndex = (p.Count + 1) \ 2
    For myIndex = 1 To myMaxIndex
    
        If ipIndex = myIndex Then
            myNode.User = myIndex
            Set GetNodeFromRoot = myNode
            Exit Function
        End If
        
        Set myNode = myNode.Succ
    
    Next

End Function


'@Description("Reurns the node at ipIndex.  The search starts at the node at s.Last and follows the pred nodes")
Private Function GetNodeFromLast(ByVal ipIndex As Long) As KvpLNode
Attribute GetNodeFromLast.VB_Description = "Reurns the node at ipIndex.  The search starts at the node at s.Last and follows the pred nodes"
                   
    Dim myNode As KvpLNode: Set myNode = s.Last
    
    Dim myIndex As Long:
    Dim myStartIndex As Long: myStartIndex = (p.Count - 1) \ 2
    For myIndex = p.Count To myStartIndex Step -1
    
        If ipIndex = myIndex Then
            myNode.User = myIndex
            Set GetNodeFromLast = myNode
            Exit Function
        End If
        
        Set myNode = myNode.Pred
        
       Next

End Function

Private Sub RemoveAllItems()

    Dim myNode As KvpLNode:  Set myNode = s.Last
    
    Do
    
        Dim myTmpNode As KvpLNode: Set myTmpNode = myNode
        Set myNode = myNode.Pred
        Set myTmpNode = Nothing
        
    Loop Until myTmpNode Is Nothing
    
    Set s.Last = Nothing
    Set s.Root = Nothing
       
End Sub


Private Sub RemoveAllSingleItem(ByRef ipItem As Variant)

    ' all nodes with the same  /item are linked
    ' Where there is a  clash the node at s.slots(my) may not hold ipItem
    ' in which case we have to search forward to find the item
    ' we also have to be mindful that whilst the  of an item may be valid (the relevant slot is occupied)
    ' the item itself may not be in the seq
    'Dim my As Long: my = s.er.OS(ipItem)
    Dim myNode As KvpLNode: Set myNode = s.Root
    
    'check that that slot is occupied
    If myNode Is Nothing Then
        Exit Sub
    End If
    
    ' search for the first matching item with the current
    ' also realising that becasue we have a valid  doesn't mean that the item is in the sequence
    Do
    
        If Comparers.EQ(myNode.Item, ipItem) Then
        
            Set myNode = myNode.Pred
            RemoveNode myNode.Succ
            
        End If
        
        Set myNode = myNode.Succ
        
    Loop Until myNode Is Nothing
    
End Sub
