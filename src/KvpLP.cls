VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Dictionary using two SeqL for Keys and Items"
'@PredeclaredId
'@Exposed
'@Folder("Kvp")
'@ModuleDescription("Dictionary using two SeqL for Keys and Items")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
Option Explicit

Private Const MY_TYPENAME                       As String = "KvpLP"
Private Const FIRST_INDEX                       As Long = 1

' currently implemented using two seql, but could be slightly
' quicker if we implemented a Key/Item node and based kvpL on a node list
Private Type State
    Keys                                        As SeqL
    Items                                       As SeqL
End Type


Private s                              As State


' Checking for duplicate keys is a linear time process
' for KvpA, KvpC and KvpL.  Consequently for **ALL** Kvp classes
' EnsureUniqueKeys is set to false by default
' Setting EnsureUniqueKeys to true will allow a check for a
' duplicate key each time a key is added.
' From a performance perspective it may be better
' to disallow duplicate keys and to then subsequently
' extract unique keys using the UniqueKeys method.

Private Type Properties
    EnsureUniqueKeys                    As Boolean
    FirstIndex                          As Long
End Type

Private p                               As Properties

'@DefaultMember
Public Function Deb() As KvpLP
Attribute Deb.VB_UserMemId = 0
    With New KvpLP
        Set Deb = .ConstructInstance()
    End With
End Function


Friend Function ConstructInstance() As KvpLP
    p.FirstIndex = FIRST_INDEX
    Set s.Keys = SeqL.Deb
    Set s.Items = SeqL.Deb
    p.EnsureUniqueKeys = False
    Set ConstructInstance = Me
End Function


Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpLP
    
    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Add"
    Guard.EnsureUniqueKeys ipKey, Me, "Add"
    
    s.Keys.Add ipKey
    s.Items.Add ipItem
    
    Set Add = Me
    
End Function


Public Function AddPairs(ByRef ipKeys As Variant, ByRef ipItems As Variant) As KvpLP
   
    Dim myKeys As IterItems: Set myKeys = IterItems(ipKeys)
    Dim myItems As IterItems: Set myItems = IterItems(ipItems)
    
    If myKeys.LacksItems Or myItems.LacksItems Then
        Set AddPairs = Me
        Exit Function
    End If
    
    Do
        
        Dim myKey As Variant: myKey = Array(myKeys.CurItem(0))
        
        Guard.KeyIsAdmin myKey(0), MY_TYPENAME & ".AddPairs"
        Guard.EnsureUniqueKeys myKey(0), Me, "AddPairs"
    
        Me.Add myKey(0), myItems.CurItem(0)

    Loop While myKeys.MoveNext And myItems.MoveNext
   
    Set AddPairs = Me
   
End Function


'Three methods are provided for retrieving items
' .Item
' .ItemByIndex
' .ItemByOffset

'which correspond to the Key,Item and Offset positions used by Iteritems
'@Description("Returns the Item associated with ipKey or Null is the key is not found")
Public Property Get Item(ByRef ipKey As Variant) As Variant
Attribute Item.VB_Description = "Returns the Item associated with ipKey or Null is the key is not found"

    Item = Null
    
    If Me.LacksItems Then
        Exit Property
    End If
    
    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Get Item"
    
    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    If Guard.KeyNotFound(myIndex, ipKey, MY_TYPENAME & ".Get Item", REPORT_BACK) Then
        Exit Property
    End If

    Dim myItem As Variant: myItem = Array(s.Items.Item(myIndex))
    If VBA.IsObject(myItem(0)) Then
        Set Item = myItem(0)
    Else
        Item = myItem(0)
    End If
    
End Property


Public Property Let Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Let Item"
    
    ' Kvp classes do not permit adding a key by assignment
    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Let Item"
    
    s.Items.Item(myIndex) = ipItem
   
End Property


Public Property Set Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Set Item"
    
    ' Kvp classes do not allow adding a key by assignment
    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Set Item"
    
    Set s.Items.Item(myIndex) = ipItem
   
End Property

Public Property Get ItemByIndex(ByRef ipIndex As Long) As Variant

    ItemByIndex = Null
    
    If Me.LacksItems Then
        Exit Property
    End If
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "Get ItemByIndex", REPORT_BACK) Then
        Exit Property
    End If
                    
    If VBA.IsObject(s.Items.Item(ipIndex)) Then
        Set ItemByIndex = s.Items.Item(ipIndex)
    Else
        ItemByIndex = s.Items.Item(ipIndex)
    End If
    
End Property


Public Property Let ItemByIndex(ByRef ipIndex As Long, ByVal ipItem As Variant)

    Guard.IndexOutOfBounds ipIndex, Me, "Let ItemByIndex"
                
    s.Items.Item(ipIndex) = ipItem
   
End Property


Public Property Set ItemByIndex(ByRef ipIndex As Long, ByVal ipItem As Variant)

    Guard.IndexOutOfBounds ipIndex, Me, "Set ItemByIndex"
                
    Set s.Items.Item(ipIndex) = ipItem
   
End Property


Public Property Get ItemByOffset(ByRef ipoffset As Long) As Variant

    ItemByOffset = Null
    
    If Me.LacksItems Then
        Exit Property
    End If
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    If Guard.IndexOutOfBounds(myIndex, Me, "Let ItemByIndex", REPORT_BACK) Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Items.Item(myIndex)) Then
        Set ItemByOffset = s.Items.Item(myIndex)
    Else
        ItemByOffset = s.Items.Item(myIndex)
    End If
    
End Property


Public Property Let ItemByOffset(ByRef ipoffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    Guard.IndexOutOfBounds myIndex, Me, "Let ItemByIndex"
                
    s.Items.Item(myIndex) = ipItem
   
End Property


Public Property Set ItemByOffset(ByRef ipoffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    Guard.IndexOutOfBounds myIndex, Me, "Let ItemByIndex"
                
    Set s.Items.Item(myIndex) = ipItem
   
End Property


Public Function IndexByKey(ByRef ipKey As Variant) As Long

    IndexByKey = -1
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".Indexbykey", REPORT_BACK) Then
        Exit Function
    End If
    
    IndexByKey = s.Keys.IndexOf(ipKey)
    
End Function

' No point in Get Key because it requires the Key which we already have

'@Ignore WriteOnlyProperty
Public Property Let Key(ByRef ipKey As Variant, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Let Key"
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let Key"

    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Let Key"
    
    s.Keys.Item(myIndex) = ipNewKey
   
End Property


Public Property Set Key(ByRef ipKey As Variant, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Set Key"
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set Key"

    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Set Key"
    
   Set s.Keys.Item(myIndex) = ipNewKey
     
End Property

Public Property Get KeyByIndex(ByRef ipIndex As Long) As Variant

    KeyByIndex = Null
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "Let ItemByIndex", REPORT_BACK) Then
        Exit Property
    End If
                    
    If VBA.IsObject(s.Keys.Item(ipIndex)) Then
        Set KeyByIndex = s.Keys.Item(ipIndex)
    Else
        KeyByIndex = s.Keys.Item(ipIndex)
    End If
    
End Property


Public Property Let KeyByIndex(ByRef ipIndex As Long, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let KeyByIndex"
    Guard.IndexOutOfBounds ipIndex, Me, "Let ItemByIndex"
                
    s.Keys.Item(ipIndex) = ipNewKey
   
End Property


Public Property Set KeyByIndex(ByRef ipIndex As Long, ByVal ipNewKey As Variant)
    
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set KeyByIndex"
    Guard.IndexOutOfBounds ipIndex, Me, "Set ItemByIndex"
                
    Set s.Keys.Item(ipIndex) = ipNewKey
   
End Property


Public Property Get KeyByOffset(ByRef ipoffset As Long) As Variant

    KeyByOffset = Null
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    If Guard.IndexOutOfBounds(myIndex, Me, "Get KeyByOffset", REPORT_BACK) Then
        Exit Property
    End If
                    
    If VBA.IsObject(s.Keys.Item(myIndex)) Then
        Set KeyByOffset = s.Keys.Item(myIndex)
    Else
        KeyByOffset = s.Keys.Item(myIndex)
    End If
    
End Property


Public Property Let KeyByOffset(ByRef ipoffset As Long, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let KeyByOffset"
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    Guard.IndexOutOfBounds myIndex, Me, "Let KeyByOffset"
                
    s.Keys.Item(myIndex) = ipNewKey
   
End Property


Public Property Set KeyByOffset(ByRef ipoffset As Long, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set KeyByOffset"

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    Guard.IndexOutOfBounds myIndex, Me, "Set KeyByOffset"
                
    Set s.Keys.Item(myIndex) = ipNewKey
   
End Property

'@Description("Removes the first found key/item pair for each key in the paramarray")
Public Function Remove(ParamArray ipParamarray() As Variant) As KvpLP
Attribute Remove.VB_Description = "Removes the first found key/item pair for each key in the paramarray"

    Set Remove = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    Dim myParamarray As Variant: myParamarray = ipParamarray
    
    If ArrayOp.LacksItems(myParamarray) Then
        Exit Function
    End If
    
    '@Ignore AssignmentNotUsed
    myParamarray = ArrayOp.Splat(myParamarray)
    
    Set Remove = RemoveRange(myParamarray)
    
End Function

'@Description("Removes the first found key/item pair for each key in the range parameter")
Public Function RemoveRange(ByRef ipRange As Variant) As KvpLP
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
        
        Dim myIndex As Long: myIndex = IndexByKey(myKeys.CurItem(0))
        s.Keys.RemoveAt myIndex
        s.Items.RemoveAt myIndex
        
ContinueLoop:
    Loop While myKeys.MoveNext
    
    Set RemoveRange = Me
    
End Function

Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpLP

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
    
    Dim myRemoveIndex As Long: myRemoveIndex = s.Keys.IndexOf(ipKey) + 1
    If myRemoveIndex >= Me.LastIndex Then
        Set RemoveAfter = Me
        Exit Function
    End If
    
    Dim myCount As Long
    For myCount = 1 To ipRun
        If myRemoveIndex > Me.LastIndex Then
            Exit For
        End If
        s.Keys.RemoveAt myRemoveIndex
        s.Items.RemoveAt myRemoveIndex
    Next
    
    Set RemoveAfter = Me
    
End Function


Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpLP

    Set RemoveBefore = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.RemoveBefore", _
        "Null is not a valid Key"
    End If
    
    If Me.LacksKey(ipKey) Then
        Exit Function
    End If

    Dim myRemoveIndex As Long
    myRemoveIndex = s.Keys.IndexOf(ipKey) - 1
    If myRemoveIndex < 1 Then
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.RemoveAfter", _
        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
    End If
    
    ' when removing before we need to reduce the
    ' reduceindex by 1 after each deletion
    Dim myCount As Long
    For myCount = 1 To ipRun
        If myRemoveIndex < s.Keys.FirstIndex Then
            Exit For
        End If
        s.Keys.RemoveAt myRemoveIndex
        s.Items.RemoveAt myRemoveIndex
        myRemoveIndex = myRemoveIndex - 1
        
    Next
    
    Set RemoveBefore = Me
    
End Function


Public Function RemoveAll(ParamArray ipParamarray() As Variant) As KvpLP

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


Public Function Clear() As KvpLP
    s.Keys.RemoveAll
    s.Items.RemoveAll
    Set Clear = Me
End Function


Public Function Reset() As KvpLP
    Set s.Keys = SeqL.Deb
    Set s.Items = SeqL.Deb
    Set Reset = Me
End Function


Public Function RemoveByIndex(ParamArray ipParamarray() As Variant) As KvpLP

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
Public Function RemoveRangeByIndex(ByRef ipRange As Variant) As KvpLP
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
        Dim myIndex As Variant: myIndex = Array(myIndexes.CurItem(0))
        
        If GroupInfo.IsNotNumber(myIndex(0)) Then
            GoTo ContinueLoop:
        End If
        
        If Guard.IndexOutOfBounds(myIndex(0), Me, "RemoveRangeByIndex", REPORT_BACK) Then
            GoTo ContinueLoop
        End If
        
        s.Keys.RemoveAt myIndex(0)
        s.Items.RemoveAt myIndex(0)
        
ContinueLoop:
    Loop While myIndexes.MovePrev
    
    Set RemoveRangeByIndex = Me
    
End Function

'ToDO: add removeitems methods


Public Function Clone() As KvpLP
    Set Clone = KvpLP.Deb.AddPairs(s.Keys, s.Items)
End Function


Public Property Get EnsureUniqueKeys() As Boolean
    EnsureUniqueKeys = p.EnsureUniqueKeys
End Property

Public Property Let EnsureUniqueKeys(ByVal ipEnsureUniqueKeys As Boolean)
    p.EnsureUniqueKeys = ipEnsureUniqueKeys
End Property


Public Function Exists(ByRef ipKey As Variant) As Boolean

    Exists = False
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".Exists", REPORT_BACK) Then
        Exit Function
    End If

    If Me.LacksItems Then
        Exit Function
    End If
    
    Exists = s.Keys.HoldsItem(ipKey)
    
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


Public Function ExistsItem(ByRef ipItem As Variant) As Boolean

    ExistsItem = False
    
    If Me.LacksItems Then
        Exit Function
    End If

    ExistsItem = s.Items.HoldsItem(ipItem)
    
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
    HoldsItems = s.Keys.HoldsItems
End Function


Public Function LacksItems() As Boolean
    LacksItems = s.Keys.LacksItems
End Function


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property


Public Property Get HasUniqueKeys() As Boolean
    HasUniqueKeys = p.EnsureUniqueKeys
End Property

'@Description("Mutates the KvpL in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As KvpLP
Attribute MappedIt.VB_Description = "Mutates the KvpL in place by applying the function in IMapper to each item"

    If s.Keys.LacksItems Then
        Set MappedIt = Me
        Exit Function
    End If

    s.Items.MappedIt ipMapper
    Set MappedIt = Me

End Function


'@Description("Returns a KvpL.Deb by applying the function in IMapper to each Item")
Public Function MapIt(ByVal ipMapper As IMapper) As KvpLP
Attribute MapIt.VB_Description = "Returns a KvpL.Deb by applying the function in IMapper to each Item"

    If s.Keys.LacksItems Then
        Set MapIt = KvpLP.Deb
        Exit Function
    End If

    Dim myS As SeqL
    Set myS = s.Items.MapIt(ipMapper)
    
    Set MapIt = KvpLP.Deb.AddPairs(s.Keys, myS)

End Function


'@Description("Returns a KvpL.Deb where the ICOmparer has been applied to the Items of the original Kvp")
Public Function FilterIt(ByVal ipComparer As IComparer) As KvpLP
Attribute FilterIt.VB_Description = "Returns a KvpL.Deb where the ICOmparer has been applied to the Items of the original Kvp"

    Dim myK As KvpLP: Set myK = KvpLP.Deb
    Set FilterIt = myK
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myPairs As Variant: myPairs = Me.Pairs
    
    Dim myPair As Variant
    For Each myPair In myPairs
        If ipComparer.ExecCmp(myPair(1)) Then
            myK.Add myPair(0), myPair(1)
        End If
    Next

    Set FilterIt = myK

End Function


Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant

    ReduceIt = Null
    
    If s.Keys.LacksItems Then
        Exit Function
    End If
    
    Dim myA() As Variant: myA = ipReducer.ExecReduction(Me.Items)
    
    If VBA.IsObject(myA(0)) Then
        Set ReduceIt = myA(0)
    Else
        ReduceIt = myA(0)
    End If
    
End Function


Public Function Count(Optional ByRef ipItem As Variant) As Long
    
    If s.Keys.LacksItems Then
        Count = -1
        Exit Function
    End If
    
    If VBA.IsMissing(ipItem) Then
        Count = s.Keys.Count
        Exit Function
    End If
    
    Dim myCount As Long: myCount = 0
    Dim myIndex As Long
    For myIndex = s.Items.FirstIndex To s.Items.LastIndex
        If Comparers.EQ(ipItem, s.Items(myIndex)) Then
            myCount = myCount + 1
        End If
    Next
    
    Count = myCount
      
End Function


Public Function CountIt(ByVal ipComparer As IComparer) As Long

    If s.Keys.LacksItems Then
        CountIt = -1
        Exit Function
    End If

    Dim myCount As Long
    myCount = 0
    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex

        If ipComparer.ExecCmp(s.Items.Item(myIndex)) Then
            myCount = myCount + 1
        End If
    Next

    CountIt = myCount

End Function



Public Function KeyOf(ByVal ipItem As Variant) As Variant

    KeyOf = Null
    
    If s.Keys.LacksItems Then
        Exit Function
    End If
    
    Dim myIndex As Long: myIndex = s.Items.IndexOf(ipItem)
    
    If myIndex < 1 Then
        Exit Function
    End If
    
    If VBA.IsObject(s.Keys.Item(myIndex)) Then
        Set KeyOf = s.Keys.Item(myIndex)
    Else
        KeyOf = s.Keys.Item(myIndex)
    End If
    
End Function


Public Function Items() As Variant
    
    If s.Keys.LacksItems Then
        Items = Array()
        Exit Function
    End If
    
    Items = s.Items.ToArray
    
End Function

Public Function ItemsAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    
    If s.Items.LacksItems Then
        Set ItemsAsColl = myC
        Exit Function
    End If
    
    Dim myItems As Variant: myItems = s.Items.ToArray
    Dim myItem As Variant
    For Each myItem In myItems
        myC.Add myItem
    Next
    
    Set ItemsAsColl = myC
        
End Function
Public Function ItemsAsSeq() As SeqL

    If s.Keys.LacksItems Then
        Set ItemsAsSeq = SeqL.Deb
        Exit Function
    End If
    
    Set ItemsAsSeq = s.Items.Clone
    
End Function


Public Function Keys() As Variant
    
    If s.Keys.LacksItems Then
        Keys = Array()
        Exit Function
    End If
    
    Keys = s.Keys.ToArray
    
End Function

Public Function KeysAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    
    If s.Keys.LacksItems Then
        Set KeysAsColl = myC
        Exit Function
    End If
    
    Dim myKeys As Variant: myKeys = s.Keys.ToArray
    Dim myKey As Variant
    For Each myKey In myKeys
        myC.Add myKey
    Next
    
    Set KeysAsColl = myC
        
End Function

Public Function KeysAsSeq() As SeqL

    If s.Keys.LacksItems Then
        Set Keys = SeqL.Deb
        Exit Function
    End If

    Set KeysAsSeq = s.Keys.Clone
    
End Function


Public Function Pairs() As Variant

    If s.Keys.LacksItems Then
        Set Pairs = Array()
        Exit Function
    End If
    
    Dim myA As Variant: ReDim myA(s.Keys.FirstIndex To s.Keys.LastIndex)
    
    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex
        myA(myIndex) = Array(s.Keys.Item(myIndex), s.Items.Item(myIndex))
    Next

    Pairs = myA

End Function

Public Function PairsAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    
    If s.Keys.LacksItems Then
        Set PairsAsColl = myC
        Exit Function
    End If
    
    Dim myPairs As Variant: myPairs = Me.Pairs
    Dim myPair As Variant
    For Each myPair In myPairs
        myC.Add myPair
    Next
    
    Set PairsAsColl = myC
        
End Function

Public Function PairsAsSeq() As SeqL

    If s.Keys.LacksItems Then
        Set PairsAsSeq = SeqL.Deb
        Exit Function
    End If
    
    Dim myS As SeqL
    Set myS = SeqL.Deb
    
    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex
        myS.Add Array(s.Keys.Item(myIndex), s.Items.Item(myIndex))
    Next

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
        LastIndex = s.Keys.Count
    End If

End Property

'@Description("Returns the index of the second item in the kvp or -1  if the kvp has less than 2 Items")
Public Function FBOIndex() As Long
Attribute FBOIndex.VB_Description = "Returns the index of the second item in the kvp or -1  if the kvp has less than 2 Items"

    If Me.LacksItems Then
        FBOIndex = -1
    ElseIf s.Keys.Count = p.FirstIndex Then
        FBOIndex = -1
    Else
        FBOIndex = p.FirstIndex + 1
    End If

End Function


'Description("Returns the index of the second to last item in the SeqH, or -1 if the SeqH has lees than 2 items")
Public Function LBOIndex() As Long

    If Me.LacksItems Then
        LBOIndex = -1
    ElseIf s.Keys.Count = p.FirstIndex Then
        LBOIndex = -1
    Else
        LBOIndex = s.Keys.Count - 1
    End If

End Function

Public Property Get FirstKey() As Variant

    FirstKey = Null

    If Me.LacksItems Then
        Exit Property
    End If

    If VBA.IsObject(s.Keys.First) Then
        Set FirstKey = s.Keys.First
    Else
        FirstKey = s.Keys.First
    End If

End Property


Public Property Get LastKey() As Variant

    LastKey = Null

    If Me.LacksItems Then
        Exit Property
    End If

    If VBA.IsObject(s.Keys.Last) Then
        Set LastKey = s.Keys.Last
    Else
        LastKey = s.Keys.Last
    End If

End Property


Public Property Get First() As Variant

    First = Null

    If Me.LacksItems Then
        Exit Property
    End If

    If VBA.IsObject(s.Items.First) Then
        Set First = s.Items.First
    Else
        First = s.Items.First
    End If

End Property


Public Property Set First(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Set First"

    Set s.Items.First = ipItem

End Property


Public Property Let First(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Let First"

    s.Items.First = ipItem

End Property


Public Property Get Last() As Variant

    Last = Null

    If Me.LacksItems Then
        Exit Property
    End If

    If VBA.IsObject(s.Items.Last) Then
        Set Last = s.Items.Last
    Else
        Last = s.Items.Last
    End If

End Property


Public Property Set Last(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Set Last"

    Set s.Items.Last = ipItem

End Property


Public Property Let Last(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Let Last"

    s.Items.Last = ipItem

End Property

Public Property Get FBOKey() As Variant

    Dim myKey As Variant: myKey = Array(s.Keys.FBO)
    If VBA.IsObject(myKey(0)) Then
        Set FBOKey = myKey(0)
    Else
        FBOKey = myKey(0)
    End If

End Property

'Public Property Get FBOKey() As Variant
'    Dim myKey As Variant: myKey = Array(s.Keys.LBO)
'    If VBA.IsObject(myKey(0)) Then
'        Set LBOKey = myKey(0)
'    Else
'        LBOKey = myKey(0)
'    End If
'End Property

Public Property Get LBOKey() As Variant
    Dim myKey As Variant: myKey = Array(s.Keys.LBO)
    If VBA.IsObject(myKey(0)) Then
        Set LBOKey = myKey(0)
    Else
        LBOKey = myKey(0)
    End If
End Property


Public Property Get FBO() As Variant
    Dim myItem As Variant: myItem = Array(s.Items.FBO)
    If VBA.IsObject(myItem(0)) Then
        Set LBO = myItem(0)
    Else
        FBO = myItem(0)
    End If
End Property

Public Property Set FBO(ByVal ipItem As Variant)
    Set s.Items.FBO = ipItem
End Property


Public Property Let LBO(ByVal ipItem As Variant)
    s.Items.FBO = ipItem
End Property


Public Property Get LBO() As Variant
    Dim myItem As Variant: myItem = Array(s.Items.LBO)
    If VBA.IsObject(myItem(0)) Then
        Set LBO = myItem(0)
    Else
        LBO = myItem(0)
    End If
End Property

Public Property Set LBO(ByVal ipItem As Variant)
    Set s.Items.LBO = ipItem
End Property

'Public Property Let LBO(ByVal ipItem As Variant)
'    s.Items.LBO = ipItem
'End Property

'Public Property Get EnsureUniqueKeys() As Boolean
'    EnsureUniqueKeys = p.EnsureUniqueKeys
'End Property
'
'Public Property Let EnsureUniqueKeys(ByVal ipEnsureUniqueKeys As Boolean)
'    p.EnsureUniqueKeys = ipEnsureUniqueKeys
'End Property


'@Description( returns a Kvp with values as keys.  Each Value is linked to a Kvp of Key vs Index for the host value
Public Function Mirror(Optional ByVal ipMirrorType As e_MirrorType = e_MirrorType.m_ByAllValues) As KvpLP

    If Me.LacksItems = 0 Then
        Set Mirror = KvpLP.Deb
        Exit Function
    End If

    Select Case ipMirrorType

        Case m_ByAllValues:             Set Mirror = MirrorByAllValues
        Case m_ByFirstValue:            Set Mirror = MirrorByFirstValue
        Case Else
            Err.Raise 17 + vbObjectError, _
                "VBALib.KvpL.Mirror", _
                Fmt.Text("Mirror enum not known. Got {0}", ipMirrorType)
                
    End Select

End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'Private functions
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'@Description("Returns an KvpL of two KvpL.  Item(1) is an KvpL of the first found Host value vs Host Key. .Item(2) is a copy of the original KvpL with the  first Key/Value pairs removed
Private Function MirrorByAllValues() As KvpLP

    If Me.LacksItems = 0 Then
        Set MirrorByAllValues = KvpLP.Deb
        Exit Function
    End If
    
    Dim myK As KvpLP
    Set myK = KvpLP.Deb

    Dim myPairs As Variant: myPairs = Me.Pairs
    Dim myIndex As Long: myIndex = 1
    Dim myPair As Variant
    For Each myPair In myPairs
        If myK.LacksKey(myPair(1)) Then
            myK.Add myPair(1), KvpLP.Deb
        End If

        myK.Item(myPair(1)).Add myPair(0), myIndex
        myIndex = myIndex + 1
    Next
    
    Set MirrorByAllValues = myK

End Function


Private Function MirrorByFirstValue() As KvpLP

    Dim myFirst As KvpLP
    Set myFirst = KvpLP.Deb
    
    Dim myOther As KvpLP
    Set myOther = KvpLP.Deb

    Dim myPairs As Variant: Set myPairs = Me.Pairs
    
    Dim myPair As Variant
    For Each myPair In myPairs

        If myFirst.LacksKey(myPair(1)) Then
            myFirst.Add myPair(1), myPair(0)
        Else
            myOther.Add myPair(0), myPair(1)
        End If

    Next

    Set MirrorByFirstValue = KvpLP.Deb.Add(0&, myFirst).Add(1&, myOther)

End Function


Private Sub RemoveAllItems()
    s.Keys.RemoveAll
    s.Items.RemoveAll
End Sub


Private Sub RemoveAllSingleItem(ByRef ipItem As Variant)

    Do
        Dim myIndex As Long: myIndex = IndexByKey(ipItem)
        If myIndex = -1 Then
            Exit Do
        End If
        s.Keys.RemoveAt myIndex
        s.Items.RemoveAt myIndex
    Loop
    
End Sub
