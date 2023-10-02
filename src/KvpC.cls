VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Dictionary using a SeqC for Keys and Items"
'@PredeclaredId
'@Exposed
'@Folder("Kvp")
'@ModuleDescription("Dictionary using a SeqC for Keys and Items")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
Option Explicit

Private Const MY_TYPENAME                       As String = "KvpC"
Private Const FIRST_INDEX                       As Long = 1
'Private Const OUT_OF_BOUNDS                     As Long = -1
Private Const REPORT_BACK                       As Boolean = True

Private Type State
    Keys                                 As SeqC
    Items                                As SeqC
End Type


Private s                               As State

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
Public Function Deb(Optional ByVal ipInitialSize As Long) As KvpC
Attribute Deb.VB_UserMemId = 0
    With New KvpC
        Set Deb = .ConstructInstance(ipInitialSize)
    End With
End Function


'@Ignore ParameterNotUsed
Friend Function ConstructInstance(Optional ByRef ipInitialSize As Long = 0) As KvpC
    p.FirstIndex = FIRST_INDEX
    Set s.Keys = SeqC.Deb.Fill(Empty, ipInitialSize)
    Set s.Items = SeqC.Deb.Fill(Empty, ipInitialSize)
    p.EnsureUniqueKeys = False
    Set ConstructInstance = Me
End Function


Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpC
    
    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".AddPairs"
    Guard.EnsureUniqueKeys ipKey, Me, "AddPairs"
    
    s.Keys.Add ipKey
    s.Items.Add ipItem
    
    Set Add = Me
    
End Function


Public Function AddPairs(ByRef ipKeys As Variant, ByRef ipItems As Variant) As KvpC
   
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
        
    
        s.Keys.Add myKey(0)
        s.Items.Add myItems.CurItem(0)
        
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
                    
    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Get Item"
    
    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    If Guard.KeyNotFound(myIndex, ipKey, MY_TYPENAME & ".Get Item", REPORT_BACK) Then
        Exit Property
    End If
    
    Dim myArray As Variant:  myArray = Array(s.Items.Item(myIndex))
    If VBA.IsObject(myArray(0)) Then
        Set Item = myArray(0)
    Else
        Item = myArray(0)
    End If
    
End Property


Public Property Let Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Let Item"
    
    ' Kvp classes do not allow adding a key by assignment
    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Let Item"
    
    s.Items.Item(myIndex) = ipItem
   
End Property


Public Property Set Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Let Item"
    
    ' Kvp classes do not allow adding a key by assignment
    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Set Item"
    
    Set s.Items.Item(myIndex) = ipItem
   
End Property

'@Description("Returns the Item at index or null if the index is not valid")
Public Property Get ItemByIndex(ByRef ipIndex As Long) As Variant
Attribute ItemByIndex.VB_Description = "Returns the Item at index or null if the index is not valid"

    ItemByIndex = Null
    
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
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    If Guard.IndexOutOfBounds(myIndex, Me, "Let ItemByIndex", REPORT_BACK) Then
        Exit Property
    End If
                    
    ItemByOffset = s.Items.Item(myIndex)
    
End Property


Public Property Let ItemByOffset(ByRef ipoffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    Guard.IndexOutOfBounds myIndex, Me, "Let ItemByIndex"
    
    s.Items.Item(myIndex) = ipItem
   
End Property


Public Property Set ItemByOffset(ByRef ipoffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    Guard.IndexOutOfBounds myIndex, Me, "Set ItemByIndex"
                
    Set s.Items.Item(myIndex) = ipItem
   
End Property

Public Function IndexByKey(ByRef ipKey As Variant) As Long
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
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "Get KeyByIndex", REPORT_BACK) Then
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

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let KeyByOffset"
    Guard.IndexOutOfBounds myIndex, Me, "Let KeyByOffset"
                
    s.Keys.Item(myIndex) = ipNewKey
   
End Property


Public Property Set KeyByOffset(ByRef ipoffset As Long, ByVal ipNewKey As Variant)

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set KeyByOffset"
    Guard.IndexOutOfBounds myIndex, Me, "Set KeyByOffset"
                
    Set s.Keys.Item(myIndex) = ipNewKey
   
End Property

''@Description(Removes the specified Key/Item pair.  Returns the Index of the Key)
'Public Function Remove(ByRef ipKey As Variant) As Long
'
'    Remove = -1
'
'    If Me.LacksItems Then
'        Exit Function
'    End If
'
'    If Guard.KeyIsAdmin(ipKey,MY_TYPENAME & ".Remove", REPORT_BACK) Then
'        Exit Function
'    End If
'
'    Dim myIndex As Long: myIndex = Me.IndexByKey(ipKey)
''    If Guard.KeyNotFound(myIndex, ipkey, myTypename, "Remove", REPORT_BACK) Then
''        Exit Function
''    End If
'
'    s.Keys.RemoveAt myIndex
'    s.Items.RemoveAt myIndex
'
'    Remove = myIndex
'
'End Function


'@Description("Removes Key/Value pairs that exist in the KvpC")
Public Function Remove(ParamArray ipParamarray() As Variant) As KvpC
Attribute Remove.VB_Description = "Removes Key/Value pairs that exist in the KvpC"
 
    Set Remove = Me
    
    If s.Keys.LacksItems Then
        Exit Function
    End If
    
    If ArrayOp.LacksItems(VBA.CVar(ipParamarray)) Then
        Exit Function
    End If
    
    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(ipParamarray)
    
    Dim myKeys As IterItems: Set myKeys = IterItems(myParamarray)
    Do

        If Guard.KeyIsAdmin(myKeys.CurItem(0), MY_TYPENAME & ".Remove", REPORT_BACK) Then
            GoTo ContinueLoop:
        End If
            
        Dim myIndex As Long: myIndex = s.Keys.IndexOf(myKeys.CurItem(0))
        If Guard.KeyNotFound(myIndex, myKeys.CurItem(0), MY_TYPENAME & ".RemoveKeys", REPORT_BACK) Then
            GoTo ContinueLoop
        End If
        
        s.Keys.RemoveAt myIndex
        s.Items.RemoveAt myIndex
        
ContinueLoop:
    Loop While myKeys.MoveNext
    
    Set Remove = Me
    
End Function


'Public Function RemoveByIndex(ByRef ipIndex As Long) As KvpC
'
'    Set RemoveByIndex = Me
'
'    If Me.LacksItems Then
'        Exit Function
'    End If
'
'    If Guard.IndexOutOfBounds(ipIndex, myTypename, "RemoveByIndex", REPORT_BACK) Then
'        Exit Function
'    End If
'
'    s.Keys.RemoveAt ipIndex
'    s.Items.RemoveAt ipIndex
'
'End Function


Public Function RemoveByIndex(ParamArray ipParamarray() As Variant) As KvpC

    Set RemoveByIndex = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If ArrayOp.LacksItems(VBA.CVar(ipParamarray)) Then
        Exit Function
    End If
    
    Dim myParamarray As Variant: myParamarray = ipParamarray
    myParamarray = ArrayOp.Splat(myParamarray)
    
    Sorters.ShakerSortArrayByIndex myParamarray
    
    Dim myIndexes As IterItems: Set myIndexes = IterItems(myParamarray)
    Do
        Dim myIndex As Long
        If GroupInfo.IsNotNumber(myIndexes.CurItem(0)) Then
            GoTo ContinueLoop
        Else
            myIndex = myIndexes.CurItem(0)
        End If
        
        If Guard.IndexOutOfBounds(myIndex, Me, "RemoveByIndexes", REPORT_BACK) Then
            GoTo ContinueLoop
        End If
        
        s.Keys.RemoveAt myIndex
        s.Items.RemoveAt myIndex
    
ContinueLoop:
    Loop While myIndexes.MoveNext
    
    
End Function

Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpC

    Set RemoveAfter = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".RemoveAfter", REPORT_BACK) Then
        Exit Function
    End If
    
    
    If Me.LacksKey(ipKey) Then
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.RemoveAfter", _
        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
    End If
    
    Dim myRemoveIndex As Long: myRemoveIndex = s.Keys.IndexOf(ipKey) + 1
    If myRemoveIndex < 1 Or myRemoveIndex > s.Keys.Count Then
        Exit Function
    End If

    Dim myCount As Long
    For myCount = 1 To ipRun
        If myRemoveIndex >= s.Keys.LastIndex Then
            Exit For
        End If
        s.Keys.RemoveAt myRemoveIndex
        s.Items.RemoveAt myRemoveIndex
        
    Next
    
    Set RemoveAfter = Me
    
End Function


Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpC

    Set RemoveBefore = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.RemoveBefore", _
        "Null is not a valid Key"
    End If
    
    If Me.LacksKey(ipKey) Then
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.RemoveBefore", _
        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
    End If

    Dim myRemoveIndex As Long
    myRemoveIndex = s.Keys.IndexOf(ipKey) - 1
    If myRemoveIndex < 1 Or myRemoveIndex > s.Keys.Count Then
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.RemoveAfter", _
        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
    End If
    
    Dim myCount As Long
    For myCount = 1 To ipRun
        If myRemoveIndex < s.Keys.FirstIndex Then
            Exit For
        End If
        s.Keys.RemoveAt myRemoveIndex
        s.Items.RemoveAt myRemoveIndex
        ' unlike removeafter we have to keep reducing the removeindex
        myRemoveIndex = myRemoveIndex - 1
        
    Next
    
    Set RemoveBefore = Me
    
End Function


Public Function Clear() As KvpC
    s.Keys.RemoveAll
    s.Items.RemoveAll
    Set Clear = Me
End Function


Public Function RemoveAll() As KvpC
    s.Keys.RemoveAll
    s.Items.RemoveAll
    Set RemoveAll = Me
End Function


Public Function Reset() As KvpC
    Set s.Keys = SeqC.Deb
    Set s.Items = SeqC.Deb
    Set Reset = Me
End Function


'===================================================================================
' Add insert methods
'===================================================================================

Public Function Clone() As KvpC
    Set Clone = KvpC.Deb.AddPairs(s.Keys, s.Items)
End Function

Public Property Get EnsureUniqueKeys() As Boolean
    EnsureUniqueKeys = p.EnsureUniqueKeys
End Property

Public Property Let EnsureUniqueKeys(ByVal ipEnsureUniqueKeys As Boolean)
    p.EnsureUniqueKeys = ipEnsureUniqueKeys
End Property

Public Function Exists(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.Exists", _
        "Null is not a valid Key"
    End If

    If s.Keys.LacksItems Then
        Exists = False
        Exit Function
    End If
    
    Exists = s.Keys.HoldsItem(ipKey)
    
End Function


Public Function HoldsKey(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.HoldsKey", _
        "Null is not a valid Key"
    End If

    If s.Keys.LacksItems Then
        HoldsKey = False
        Exit Function
    End If
    
    HoldsKey = s.Keys.HoldsItem(ipKey)
    
End Function


Public Function ContainsKey(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.ContainsKey", _
        "Null is not a valid Key"
    End If

    If s.Keys.LacksItems Then
        ContainsKey = False
        Exit Function
    End If

    ContainsKey = s.Keys.HoldsItem(ipKey)
    
End Function


Public Function LacksKey(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.LacksKey", _
        "Null is not a valid Key"
    End If

    If s.Keys.LacksItems Then
        LacksKey = True
        Exit Function
    End If
    
    LacksKey = s.Keys.LacksItem(ipKey)
    
End Function


Public Function ExistsItem(ByRef ipItem As Variant) As Boolean

    If s.Keys.LacksItems Then
        ExistsItem = False
        Exit Function
    End If
    
    ExistsItem = s.Items.HoldsItem(ipItem)
    
End Function


Public Function HoldsItem(ByRef ipItem As Variant) As Boolean

    If s.Keys.LacksItems Then
        HoldsItem = False
        Exit Function
    End If
    
    HoldsItem = s.Items.HoldsItem(ipItem)
    
End Function


Public Function Contains(ByRef ipItem As Variant) As Boolean
        
    If s.Keys.LacksItems Then
        Contains = False
        Exit Function
    End If
    
    Contains = s.Items.HoldsItem(ipItem)
    
End Function


Public Function LacksItem(ByRef ipItem As Variant) As Boolean
    
    If s.Keys.LacksItems Then
        LacksItem = True
        Exit Function
    End If
    
    LacksItem = s.Items.LacksItem(ipItem)
    
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


'@Description("Mutates the Kvpc in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As KvpC
Attribute MappedIt.VB_Description = "Mutates the Kvpc in place by applying the function in IMapper to each item"

    If s.Keys.LacksItems Then
        Set MappedIt = Me
        Exit Function
    End If

    s.Items.MappedIt ipMapper
    Set MappedIt = Me

End Function


'@Description("Returns a Kvpc.Deb by applying the function in IMapper to each Item")
Public Function MapIt(ByVal ipMapper As IMapper) As KvpC
Attribute MapIt.VB_Description = "Returns a Kvpc.Deb by applying the function in IMapper to each Item"

    If s.Keys.LacksItems Then
        Set MapIt = Me
        Exit Function
    End If

    Dim myS As SeqC
    Set myS = s.Items.MapIt(ipMapper)
    
    Set MapIt = KvpC.Deb.AddPairs(s.Keys, myS)

End Function


'@Description("Returns a Kvpc.Deb where the ICOmparer has been applied to the Items of the original Kvp")
Public Function FilterIt(ByVal ipComparer As IComparer) As KvpC
Attribute FilterIt.VB_Description = "Returns a Kvpc.Deb where the ICOmparer has been applied to the Items of the original Kvp"

    Dim myK As KvpC: Set myK = KvpC.Deb
    
    If s.Keys.LacksItems Then
        Set FilterIt = myK
        Exit Function
    End If

    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex
        If ipComparer.ExecCmp(s.Items.Item(myIndex)) Then
            myK.Add s.Keys.Item(myIndex), s.Items.Item(myIndex)
        End If
    Next

    Set FilterIt = myK

End Function


Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant

    If s.Keys.LacksItems Then
        ReduceIt = Null
        Exit Function
    End If
    
    Dim myA As Variant: myA = ipReducer.ExecReduction(s.Items)
    
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
    
    Dim myItem As Variant
    For Each myItem In s.Items
        If Comparers.EQ(ipItem, myItem) Then
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


'Public Function KeyByIndex(ByVal ipIndex As Long) As Variant
'
'    KeyByIndex = Null
'
'    If s.Keys.LacksItems Then
'        Exit Function
'    End If
'
'    If ipIndex < s.Keys.FirstIndex Or ipIndex > s.Keys.LastIndex Then
'        Exit Function
'    End If
'
'    Dim myItem As Variant: myItem = Array(s.Keys.Item(ipIndex))
'
'    If VBA.IsObject(myItem(0)) Then
'        Set KeyByIndex = myItem(0)
'    Else
'        KeyByIndex = myItem(0)
'    End If
'
'End Function


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


'Public Function HasUniqueKeys() As Boolean
'    HasUniqueKeys = p.EnsureUniqueKeys
'End Function
'
'
'Public Function ForbidSameKeys() As KvpC
'    p.EnsureUniqueKeys = True
'    Set ForbidSameKeys = Me
'End Function
'
'
'Public Function AllowSameKeys() As KvpC
'    p.EnsureUniqueKeys = False
'    Set AllowSameKeys = Me
'End Function


Public Function Items() As Variant

    If s.Keys.LacksItems Then
        Items = Array()
        Exit Function
    End If
    
    Items = s.Items.ToArray
    
End Function


Public Function ItemsAsSeq() As SeqC

    If s.Keys.LacksItems Then
        Set ItemsAsSeq = SeqC.Deb
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


Public Function KeysAsSeq() As SeqC
    
    If s.Keys.LacksItems Then
        Set Keys = SeqC.Deb
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

Public Function SPairs() As Variant

    Dim myKeys As SeqC: Set myKeys = Me.KeysAsSeq

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


Public Function PairsAsSeq() As SeqC

    If s.Keys.LacksItems Then
        Set PairsAsSeq = SeqC.Deb
        Exit Function
    End If
    
    Dim myS As SeqC: Set myS = SeqC.Deb
    
    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex
        myS.Add Array(s.Keys.Item(myIndex), s.Items.Item(myIndex))
    Next

    Set PairsAsSeq = myS

End Function


Public Property Get FirstIndex() As Long
    FirstIndex = p.FirstIndex
End Property


Public Property Get LastIndex() As Long
    LastIndex = s.Keys.LastIndex
End Property


Public Property Get FirstKey() As Variant

    FirstKey = Null
    
    If s.Keys.LacksItems Then
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
    
    If s.Keys.LacksItems Then
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

    If s.Keys.LacksItems Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Items.First) Then
        Set First = s.Items.First
    Else
        First = s.Items.First
    End If
    
End Property


Public Property Set First(ByVal ipItem As Variant)

    If s.Keys.LacksItems Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.Let_FirstItem", _
        "Kvp has no Keys"
    End If

    Set s.Items.First = ipItem

End Property


Public Property Let First(ByVal ipItem As Variant)

    
    If Me.LacksItems Then
        Err.Raise 17, "VBALib.KvpC.Let FirstItem", "VBALib.Kvp has no Keys"
    End If

    s.Items.First = ipItem

End Property


'
Public Property Get Last() As Variant

    Last = Null
    
    If s.Keys.LacksItems Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Items.Last) Then
        Set Last = s.Items.Last
    Else
        Last = s.Items.Last
    End If
    
End Property


Public Property Set Last(ByVal ipItem As Variant)

    If s.Keys.LacksItems Then
        Err.Raise 17, _
        "VBALib.KvpC.Set LastItem", _
        "Kvp has no Keys"
    End If

    Set s.Items.Last = ipItem
    
End Property


Public Property Let Last(ByVal ipItem As Variant)

    If s.Keys.LacksItems Then
        Err.Raise 17, _
        "VBALib.KvpC.Set LastItem", _
        "Kvp has no Keys"
    End If
    
    s.Items.Last = ipItem
    
End Property


'@Description( returns a Kvp with values as keys.  Each Value is linked to a Kvp of Key vs Index for the host value
Public Function Mirror(Optional ByVal ipMirrorType As e_MirrorType = e_MirrorType.m_ByAllValues) As KvpC

    If Me.LacksItems = 0 Then
        Set Mirror = KvpC.Deb
        Exit Function
    End If

    Select Case ipMirrorType

        Case m_ByAllValues:             Set Mirror = MirrorByAllValues
        Case m_ByFirstValue:            Set Mirror = MirrorByFirstValue
        Case Else
            Err.Raise 17 + vbObjectError, _
            "VBALib.KvpC.Mirror", _
            Fmt.Text("Mirror enum not known. Got {0}", ipMirrorType)
                
    End Select

End Function


'@Description("Returns an KvpC of two KvpC.  Item(1) is an KvpC of the first found Host value vs Host Key. .Item(2) is a copy of the original KvpC with the  first Key/Value pairs removed
Private Function MirrorByAllValues() As KvpC

    Dim myK As KvpC
    Set myK = KvpC.Deb
    
    If Me.LacksItems = 0 Then
        Set MirrorByAllValues = KvpC.Deb
        Exit Function
    End If

    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex

        If myK.LacksKey(s.Keys.Item(myIndex)) Then
            myK.Add s.Items.Item(myIndex), KvpC.Deb
        End If

        myK.Item(s.Keys.Item(myIndex)).Add s.Keys.Item(myIndex), myIndex

    Next

    Set MirrorByAllValues = myK

End Function


Private Function MirrorByFirstValue() As KvpC

    Dim myFirst As KvpC
    Set myFirst = KvpC.Deb
    
    Dim myOther As KvpC
    Set myOther = KvpC.Deb

    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex

        If myFirst.LacksKey(s.Items.Item(myIndex)) Then
            myFirst.Add s.Items.Item(myIndex), s.Keys.Item(myIndex)
        Else
            myOther.Add s.Keys.Item(myIndex), s.Items.Item(myIndex)
        End If

    Next

    Set MirrorByFirstValue = KvpC.Deb.Add(0&, myFirst).Add(1&, myOther)

End Function


''==============================
'Public Function SetOf(ByVal ipSet As e_Set, ByRef ipRHS As KvpC) As KvpC
'
'    Select Case ipSet
'
'        Case e_Set.m_Both:                   Set SetOf = SetOfCommon(ipRHS)
'        Case e_Set.m_HostOnly:               Set SetOf = SetOfHostOnly(ipRHS)
'        Case e_Set.m_ParamOnly:              Set SetOf = SetOfParamOnly(ipRHS)  ' note reversal of params
'        Case e_Set.m_NotInBoth:              Set SetOf = SetOfNotCommon(ipRHS)
'        Case e_Set.m_Unique:                 Set SetOf = SetofUnique(ipRHS)
'
'    End Select
'
'End Function
'
'Private Function SetOfBoth(ByRef ipLHS As KvpC, ByRef ipRHS As KvpC) As KvpC
'
'    Dim myResult As KvpC
'    Set myResult = KvpC.Deb
'
'    If s.Keys.LacksItems Then
'        Set SetOfBoth = KvpC.Deb
'        Exit Function
'    End If
'
'    Dim myIndex As Long
'    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex
'
'        If ipRHS.HoldsKey(s.Keys.Item(myIndex)) Then
'            myResult.Add myLHS.Key(0), myLHS.GetItem
'        End If
'    Loop While myLHS.MoveNext
'
'    Return myResult
'
'End Function
'
'
'Private Function SetOfHostOnly(ByRef ipLHS As KvpC, ByRef ipRHS As oContainer) As KvpC
'
'    Dim myResult As KvpC
'    Set myResult = KvpC.Deb
'
'    Dim myIndex As Long
'    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex
'    Do
'        If s.ipRHS.LacksKey(s.Keys.Item(myIndex)) Then
'            myResult.Add s.Keys.Item(myIndex), s.Items.Item(myIndex)
'        End If
'
'    Next
'
'    Set SetOfHostOnly = myResult
'
'End Function
'
'
''
''
''  Private Function NotBoth(ByRef ipLHS As KvpC, ByRef ipRHS As KvpC) As KvpC
''
''          'Return KvpC.Deb(Me.LHSOnly(ipLHS, ipRHS)).MergeUnique(LHSOnly(ipRHS, ipLHS)) ' note the reversal of the parameters
''          Dim myResult As KvpC = KvpC.Deb
''          If ipLHs.host.count > 0 Then
''              Dim myLHS As IterItems = IterItems(ipLHS)
''              Do
''
''                  If ipRHS.LacksKey(myLHS.Key(0)) Then
''                      myResult.Add myLHS.Key(0), myLHS.GetItem
''                  End If
''              Loop While myLHS.MoveNext
''          End If
''          If ipRHs.host.count > 0 Then
''              Dim myRHS As IterItems = IterItems(ipRHS)
''              Do
''                  If ipLHS.LacksKey(myRHS.Key(0)) Then
''                      myResult.Add myRHS.Key(0), myRHS.GetItem
''                  End If
''              Loop While myRHS.MoveNext
''          End If
''  End Function
''
''
''  Private Function CombineUnique(ByRef ipLHS As KvpC, ByRef ipRHS As KvpC) As KvpC
''      Dim myResult As KvpC = KvpC.Deb
''      Dim myI As IterItems = IterItems(ipRHS)
''      Do
''          If myResult.Count = 0 Then
''            myResult.Add myI.Key(0), myI.GetItem
''          ElseIf myResult.LacksKey(myI.Key(0)) Then
''              myResult.Add myI.Key(0), myI.GetItem
''          End If
''      Loop While myI.MoveNext
''      Return myResult
''  End Function
''
'  ' Public Function MergeUnique(ByRef ipRHS As KvpC) As KvpC
'
'  '     If ipRHs.host.count = 0 Then
'  '         Return Me
'  '     End If
'
'  '     Dim myI As IterItems = IterItems(ipRHS)
'  '     Do
'  '         DoEvents
'  '         If Me.Count = 0 Then
'  '             Me.Add myI.Key(0), myI.GetItem
'  '         ElseIf Me.LacksKey(myI.Key(0)) Then
'  '             Me.Add myI.Key(0), myI.GetItem
'  '         End If
'  '     Loop While myI.MoveNext
'  '     Return Me
'  ' End Function
