VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpHA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule
'@PredeclaredId
'@Exposed
'@Folder("seq")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========

' Based on cHashD from https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs)&s=f049d882b4fc4d27f4bc2f9bc383a53e
' cHashD, Olaf Schmidt in August 2016 (enhanced about mixed, variable KeyTypes and Remove-functions in 2020)
' Reformatted and Rubberducked by Steven Laycock July 2023
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========

Option Explicit

Private Const MY_TYPENAME                   As String = "KvpHA"
Private Const FIRST_INDEX                   As Long = 1
Private Const MIN_KVP_SIZE                  As Long = 512
Private Const MAX_KVP_SIZE                  As Long = 524288 ' 2^19
Private Const DEFAULT_KVP_SIZE              As Long = 16384
Private Const REPORT_BACK                   As Boolean = True
Private Const NOT_FOUND                     As Long = -1

'Private Const NoEntry                       As Variant = Empty
'fixed HashSlot-amount and Hash-Multiplikator
Private Const DynTakeOver                   As Long = 4
Private Const HMul                          As Long = 3727

Private Type HashTableEntry
    Count                                     As Long
    DataIdxsStat(0 To DynTakeOver - 1)        As Long
    DataIdxsDyn()                             As Long
End Type

Private Type State
     Hasher                                 As Hasher
     W()                                    As Integer
     sAW                                    As SAFEARRAY1D
     
     LastExpectedMaxCount                   As Long
     EnsureUniqueKeys                       As Boolean
     LastH                                  As Long
     DTUB                                   As Long
     HashTableSize                          As Long
     CompareMode                            As VbCompareMethod
     HashTable()                            As HashTableEntry
    'both pairing-arrays are of type Variant
     Keys()                                 As Variant
     Values()                               As Variant
    'three varsmyChecksum)                            As Long
     FCount                                 As Long
     FUB                                    As Long
     FIndexes()                             As Variant
End Type

Private s                                   As State

' we can now use
' for KvpA, KvpC and KvpHA.  Consequently for **ALL** Kvp classes
' EnsureUniqueKeys is set to false by default
' Setting EnsureUniqueKeys to true will allow a check for a
' duplicate key each time a key is added.
' From a performance perspective it may be better
' to disallow duplicate keys and to then subsequently
' extract unique keys using the UniqueKeys method.


Private Type Properties
    EnsureUniqueKeys                    As Boolean
    FirstIndex                          As Long
    Count                                  As Long

End Type

Private p                               As Properties

'@DefaultMember
Public Function Deb(Optional ByRef ipCapacity As Long = DEFAULT_KVP_SIZE) As KvpHA
Attribute Deb.VB_UserMemId = 0
    With New KvpHA
        Set Deb = .ConstructInstance(ipCapacity)
    End With
End Function

Friend Function ConstructInstance(ByRef ipCapacity As Long) As KvpHA

    p.FirstIndex = FIRST_INDEX
    s.sAW.cDims = 1
    s.sAW.cbElements = 2
    s.sAW.cLocks = 1
    s.sAW.fFeatures = &H11                         'FADF_AUTO=&H1 || FADF_FIXEDSIZE=&H10
    BindArray s.W, VarPtr(s.sAW)
    
    s.CompareMode = vbBinaryCompare
    s.EnsureUniqueKeys = False
    Reinit ipCapacity                                 'at startup we set it up, to behave ideally for up to 16384 Items
                                                     ' now  the default value for the optional constructor param
    If LWC(97) = 0 Then
        InitLWC
    End If
        
    Set ConstructInstance = Me
    
End Function



Public Function Reinit(Optional ByRef ExpectedMaxCount As Long, Optional ByVal EnsureUniqueKeys As Boolean = False) As KvpHA

    If ExpectedMaxCount <= 0 Then
        ExpectedMaxCount = MIN_KVP_SIZE
    End If
    
    If ExpectedMaxCount < 128 Then
        ExpectedMaxCount = MIN_KVP_SIZE
    End If
    
    s.LastExpectedMaxCount = ExpectedMaxCount
    s.EnsureUniqueKeys = EnsureUniqueKeys
    
    s.HashTableSize = 16
    Do Until s.HashTableSize * 2 > ExpectedMaxCount
        s.HashTableSize = s.HashTableSize * 2
    Loop
    
    If s.HashTableSize > MAX_KVP_SIZE Then
        s.HashTableSize = MAX_KVP_SIZE
    End If
    
    ReDim s.HashTable(0 To s.HashTableSize - 1)
    
    s.DTUB = s.HashTableSize  'ToDO: what to do for dtub?  was LastExpectedMaxCount
    ReDim s.Keys(0 To s.DTUB)
    ReDim s.Values(0 To s.DTUB)
    p.Count = 0
    s.FCount = 0
    s.FUB = 16
    ReDim s.FIndexes(0 To s.FUB)
    
    Set Reinit = Me
    
End Function

Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpHA

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".AddPairs"
    Guard.EnsureUniqueKeys ipKey, Me, "AddPairs"


    Dim H As Long
    Dim UB As Long
    Dim i As Long
    
    ' need to understand more about lastH
    If s.LastH Then
        H = s.LastH
        s.LastH = 0
    ElseIf s.EnsureUniqueKeys Then
        If FindIndex(ipKey, H) >= 0 Then
            Err.Raise 457
        End If
    Else
        H = -1
        FindIndex ipKey, H                         'a Value of -1 for H will skip the Index-Search, returning only H
    End If
    
    'add the new Pair, prolonging the Keys- and Values-arrays
    If s.DTUB < p.Count Then
        s.DTUB = (s.DTUB + 16) * 1.3
        ReDim Preserve s.Keys(0 To s.DTUB)
        ReDim Preserve s.Values(0 To s.DTUB)
    End If
    
    If s.FCount > 0 Then
        s.FCount = s.FCount - 1
        i = s.FIndexes(s.FCount)
    Else
        i = p.Count
    End If
    
    VariantCopyInd ByVal VarPtr(s.Values(i)), ByVal VarPtr(ipItem)
    VariantCopyInd ByVal VarPtr(s.Keys(i)), ByVal VarPtr(ipKey)
    
    'add the new DataIndex to the proper Hash-Buckets
    Select Case s.HashTable(H).Count
    
        Case Is < DynTakeOver
        
            s.HashTable(H).DataIdxsStat(s.HashTable(H).Count) = i
            
        Case DynTakeOver
        
            ReDim Preserve s.HashTable(H).DataIdxsDyn(DynTakeOver To DynTakeOver + 3)
            s.HashTable(H).DataIdxsDyn(DynTakeOver) = i
            
        Case Else
        
            UB = UBound(s.HashTable(H).DataIdxsDyn)
            
            If UB < s.HashTable(H).Count Then
                UB = (UB + 3) * 1.3
                ReDim Preserve s.HashTable(H).DataIdxsDyn(DynTakeOver To UB)
            End If
            
            s.HashTable(H).DataIdxsDyn(s.HashTable(H).Count) = i
            
    End Select
    
    s.HashTable(H).Count = s.HashTable(H).Count + 1
    
    p.Count = p.Count + 1
    
    Set Add = Me
End Function


Public Function AddPairs(ByRef ipKeys As Variant, ByRef ipItems As Variant) As KvpHA
   
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
    
    Dim myIndex As Long: myIndex = FindIndex(ipKey)
    If Guard.KeyNotFound(myIndex, ipKey, MY_TYPENAME & ".Get Item", REPORT_BACK) Then
        Exit Property
    End If
    
    If VBA.IsObject(s.Values(myIndex)) Then
        Set Item = s.Values(myIndex)
    Else
        Item = s.Values(myIndex)
    End If
    
End Property


Public Property Let Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Get Item"
    
    ' Kvp classes do not permit adding a key by assignment
    Dim myIndex As Long: myIndex = FindIndex(ipKey)
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Let Item"
    
    VariantCopy s.Values(myIndex), ipItem
    's.Values(myIndex) = ipItem
    
End Property


Public Property Set Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Get Item"
    
    ' Kvp classes do not allow adding a key by assignment
    Dim myIndex As Long: myIndex = FindIndex(ipKey)
    
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Set Item"
    
    VariantCopy s.Values(myIndex), ipItem
    'set s.Values(myIndex) = ipItem
    
End Property


Public Property Get ItemByIndex(ByVal ipIndex As Long) As Variant

    ItemByIndex = Null

    If Me.LacksItems Then
        Exit Property
    End If
    
    If Guard.IndexOutOfBounds(ipIndex, Me, "Get ItemByIndex", REPORT_BACK) Then
        Exit Property
    End If
    
    'Adjust to zero based indexing
    Dim myIndex As Long: myIndex = ipIndex - 1
    
    ' adjust for hash clashes
    If s.FCount Then
        AdjustIndex myIndex
    End If
    
    If VBA.IsObject(s.Values(myIndex)) Then
        Set ItemByIndex = s.Values(myIndex)
    Else
        ItemByIndex = s.Values(myIndex)
    End If
    'VariantCopy ItemByIndex, ByVal VarPtr(s.values(IndexZeroBased))
    
End Property


Public Property Let ItemByIndex(ByVal ipIndex As Long, ByVal ipItem As Variant)

    Guard.IndexOutOfBounds ipIndex, Me, "Let ItemByIndex"
    
    'Adjust for 1 based indexing
    Dim myIndex As Long: myIndex = ipIndex - 1
    
    'Adjust for hash clashes
    If s.FCount Then
        AdjustIndex myIndex
    End If
    
    's.Values(myIndex) = ipItem
   
    VariantCopy ByVal VarPtr(s.Values(myIndex)), ByVal VarPtr(ipItem)
    
End Property


Public Property Set ItemByIndex(ByVal ipIndex As Long, ByVal ipItem As Variant)
 
    Guard.IndexOutOfBounds ipIndex, Me, "Set ItemByIndex"
    
    'Adjust for 1 based indexing
    Dim myIndex As Long: myIndex = ipIndex - 1
    
    If s.FCount Then
        AdjustIndex myIndex
    End If
    
    Set s.Values(myIndex) = ipItem
    'VariantCopyInd ByVal VarPtr(s.Values(myIndex)), ByVal VarPtr(ipItem)
    
End Property

Public Property Get ItemByOffset(ByVal ipoffset As Long) As Variant

    ItemByOffset = Null
    
    If Me.LacksItems Then
        Exit Property
    End If
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    If Guard.IndexOutOfBounds(myIndex, Me, "Get ItemByOffset", REPORT_BACK) Then
        Exit Property
    End If
    
    'adjust to zerobased indxing
    '@Ignore AssignmentNotUsed
    myIndex = myIndex - 1
    
    ' adjust for hash clashes
    If s.FCount Then
        AdjustIndex myIndex
    End If
    
    VariantCopy ItemByOffset, ByVal VarPtr(s.Values(myIndex))
    
End Property


Public Property Let ItemByOffset(ByVal ipoffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    Guard.IndexOutOfBounds myIndex, Me, "Let ItemByIndex"
    
    ' adjust to zero based indexing
    myIndex = myIndex - 1
    
    ' adjust for hash clashes
    If s.FCount Then
        AdjustIndex myIndex
    End If
    
    'ToDo: why is variantcopyind used here when both are variants
    VariantCopyInd ByVal VarPtr(s.Values(myIndex)), ByVal VarPtr(ipItem)
    
End Property


Public Property Set ItemByOffset(ByVal ipoffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    
    Guard.IndexOutOfBounds myIndex, Me, "Set ItemByOffset"
    
    ' adjust to zero based indexing
    myIndex = myIndex - 1
    ' adjust for hash clashes
    If s.FCount Then
        AdjustIndex myIndex
    End If
    
    VariantCopyInd ByVal VarPtr(s.Values(myIndex)), ByVal VarPtr(ipItem)
    
End Property

Public Function IndexByKey(ByRef ipKey As Variant) As Long

    IndexByKey = -1
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".Indexbykey", REPORT_BACK) Then
        Exit Function
    End If
    
    Dim myIndex As Long: myIndex = FindIndex(ipKey)
    
    ' adjust for hash clashes
    If myIndex >= 0 And s.FCount > 0 Then
        AdjustIndex myIndex, True
    End If
    
    IndexByKey = myIndex
    
End Function


' No point in Get Key because it requires the Key which we already have

'@Ignore WriteOnlyProperty
Public Property Let Key(ByRef ipKey As Variant, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Let Key"
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let Key"

    Dim myIndex As Long: myIndex = FindIndex(ipKey)
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Let Key"
    
    Dim myItem As Variant: myItem = Array(Me.Item(ipKey))
    Me.Remove ipKey
    Me.Add ipNewKey, myItem(0)
   
End Property

Public Property Set Key(ByRef ipKey As Variant, ByVal ipNewKey As Variant)

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Set Key"
    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set Key"

    Dim myIndex As Long: myIndex = FindIndex(ipKey)
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Set Key"
    
    ' Keys are stored using a hash function to determine the storage location
    ' Changing the key will change the hash and hence storage location
    ' therefore we have to remove the old key (preserving the value)
    ' and insert the new Key/Item pair
    Dim myItem As Variant: myItem = Array(s.Values(myIndex))
    Me.Remove ipKey
    Me.Add ipNewKey, myItem(0)
   
End Property

Public Property Get KeyByIndex(ByRef ipIndex As Long) As Variant

    KeyByIndex = Null
 
    If Guard.IndexOutOfBounds(ipIndex, Me, "Get KeyByIndex", REPORT_BACK) Then
        Exit Property
    End If
    
    ' adjust for KvpH using 1 based indexing
    Dim myIndex As Long: myIndex = ipIndex - 1
    
    If s.FCount Then
        AdjustIndex myIndex
    End If
    
    If VBA.IsObject(s.Keys(myIndex)) Then
        Set KeyByIndex = s.Keys(myIndex)
    Else
        KeyByIndex = s.Keys(myIndex)
    End If
    
End Property

'@Description("Removes the specufued Key/Item and add NewKey/Item. NewKey will be at a new Index")
Public Property Let KeyByIndex(ByRef ipIndex As Long, ByVal ipNewKey As Variant)
Attribute KeyByIndex.VB_Description = "Removes the specufued Key/Item and add NewKey/Item. NewKey will be at a new Index"

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let KeyByIndex"
    Guard.IndexOutOfBounds ipIndex, Me, "Let KeyByIndex"
    
    ' adjust for KvpH using 1 based indexing
    Dim myIndex As Long: myIndex = ipIndex - 1
    
    Dim myItem As Variant: myItem = Array(ItemByIndex(myIndex))
    Remove KeyByIndex(myIndex)
    Me.Add ipNewKey, myItem(0)
   
End Property

'@Description("Removes the specified Key/Item and adds NewKey/Item. NewKey will be at a new Index")
Public Property Set KeyByIndex(ByRef ipIndex As Long, ByVal ipNewKey As Variant)
Attribute KeyByIndex.VB_Description = "Removes the specified Key/Item and adds NewKey/Item. NewKey will be at a new Index"

    ' adjust for KvpH using 1 based indexing
    Dim myIndex As Long: myIndex = ipIndex - 1

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set KeyByIndex"
    Guard.IndexOutOfBounds myIndex, Me, "Set KeyByIndex"
    
    Dim myItem As Variant: myItem = Array(ItemByIndex(ipIndex))
    Remove KeyByIndex(ipIndex)
    Me.Add ipNewKey, myItem(0)
   
End Property


Public Property Get KeyByOffset(ByRef ipoffset As Long) As Variant

    KeyByOffset = Null
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    If Guard.IndexOutOfBounds(myIndex, Me, "Get KeyByOffset", REPORT_BACK) Then
        Exit Property
    End If
    
    '@Ignore AssignmentNotUsed
    ' adjust to zero based indexing
    myIndex = myIndex - 1
    ' adjust for hash clashes
    If s.FCount Then
        AdjustIndex myIndex
    End If
    
    If VBA.IsObject(s.Keys(myIndex)) Then
        Set KeyByOffset = s.Keys(myIndex)
    Else
        KeyByOffset = s.Keys(myIndex)
    End If
    
    'VariantCopy KeyByIndex, ByVal VarPtr(s.Keys(myIndex))
    
End Property

'@Description("Removes the specufued Key/Item and add NewKey/Item. NewKey will be at a new Index")
Public Property Let KeyByOffset(ByRef ipoffset As Long, ByVal ipNewKey As Variant)
Attribute KeyByOffset.VB_Description = "Removes the specufued Key/Item and add NewKey/Item. NewKey will be at a new Index"

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Let KeyByOffset"
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    Guard.IndexOutOfBounds myIndex, Me, "Let KeyByOffset"
    
    ' no need to adjust to zero based indexing
    Dim myItem As Variant: myItem = Array(ItemByIndex(myIndex))
    RemoveKey KeyByIndex(myIndex)
    Me.Add ipNewKey, myItem(0)
   
End Property

'@Description("Removes the specified Key/Item and adds NewKey/Item. NewKey will be at a new Index")
Public Property Set KeyByOffset(ByRef ipoffset As Long, ByVal ipNewKey As Variant)
Attribute KeyByOffset.VB_Description = "Removes the specified Key/Item and adds NewKey/Item. NewKey will be at a new Index"

    Guard.KeyIsAdmin ipNewKey, MY_TYPENAME & ".Set KeyByOffset"
    
    Dim myIndex As Long: myIndex = Me.FirstIndex + ipoffset
    Guard.IndexOutOfBounds myIndex, Me, "Set KeyByOffset"
    
    ' no need to adjust to zero based indexing
    Dim myItem As Variant: myItem = Array(ItemByIndex(myIndex))
    RemoveKey KeyByIndex(myIndex)
    Me.Add ipNewKey, myItem(0)
   
End Property


Private Sub RemoveKey(ByRef ipKey As Variant)

    Dim H                       As Long
    Dim myIndex                 As Long
    Dim i                       As Long
    Dim j                       As Long


    If Me.LacksItems Then
        Exit Sub
    End If
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".Remove", REPORT_BACK) Then
        Exit Sub
    End If
    
    myIndex = FindIndex(ipKey, H)
    If Guard.KeyNotFound(myIndex, ipKey, "Remove", REPORT_BACK) Then
        Exit Sub
    End If
    
    If myIndex < 0 Then
        Exit Sub
    End If
    
    For i = 0 To s.HashTable(H).Count - 2
        If i < DynTakeOver Then
        
            If j = 0 Then
                If s.HashTable(H).DataIdxsStat(i) = myIndex Then
                    j = i + 1
                End If
            End If
            
            If j Then
                If j < DynTakeOver Then
                    s.HashTable(H).DataIdxsStat(i) = s.HashTable(H).DataIdxsStat(j)
                    j = j + 1
                Else
                    s.HashTable(H).DataIdxsStat(i) = s.HashTable(H).DataIdxsDyn(j)
                    j = j + 1
                End If
            End If
        Else
        
            If j = 0 Then
                If s.HashTable(H).DataIdxsDyn(i) = myIndex Then
                    j = i + 1
                End If
            End If
            
            If j Then
                s.HashTable(H).DataIdxsDyn(i) = s.HashTable(H).DataIdxsDyn(j)
                j = j + 1
            End If
            
        End If
    Next
    
    s.HashTable(H).Count = i
    
    If s.FUB < s.FCount Then
        s.FUB = s.FUB + s.FUB
        ReDim Preserve s.FIndexes(0 To s.FUB)
    End If
    
    s.FIndexes(s.FCount) = myIndex
    s.FCount = s.FCount + 1
    
    s.Keys(myIndex) = Empty
    s.Values(myIndex) = Empty
    p.Count = p.Count - 1
    
End Sub

'@Description("Removes the first found key/item pair for each key in the paramarray")
Public Function Remove(ParamArray ipParamarray() As Variant) As KvpHA
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
Public Function RemoveRange(ByRef ipRange As Variant) As KvpHA
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
        
        RemoveKey myKeys.CurItem(0)
        
ContinueLoop:
    Loop While myKeys.MoveNext
    
    Set RemoveRange = Me
    
End Function

Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpHA

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
    
    If Guard.InvalidRun(ipRun, MY_TYPENAME & ".RemoveAfter", REPORT_BACK) Then
        Exit Function
    End If
    
    Dim myRemoveIndex As Long: myRemoveIndex = IndexByKey(ipKey) + 2 '1 for after, 1 for changing to 1 based index
    If Guard.IndexOutOfBounds(myRemoveIndex, Me, "RemoveAfter", REPORT_BACK) Then
        Exit Function
    End If
    
    Dim myCount As Long
    For myCount = 1 To ipRun
        Dim myKey As Variant: myKey = Array(KeyByIndex(myRemoveIndex))
        If VBA.IsNull(myKey(0)) Then
            Exit For
        End If
        RemoveKey myKey(0)
    Next
    
    Set RemoveAfter = Me
    
End Function

Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpHA

    Set RemoveBefore = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".RemoveAfter", REPORT_BACK) Then
        Exit Function
    End If
    
    If Me.LacksKey(ipKey) Then
        Exit Function
    End If
    
    If Guard.InvalidRun(ipRun, MY_TYPENAME & ".RemoveAfter", REPORT_BACK) Then
        Exit Function
    End If
    
    Dim myRemoveIndex As Long: myRemoveIndex = IndexByKey(ipKey)  '1 for after, 1 for changing to 1 based index
    If Guard.IndexOutOfBounds(myRemoveIndex, Me, "RemoveAfter", REPORT_BACK) Then
        Exit Function
    End If
    
    Dim myCount As Long
    For myCount = 1 To ipRun
        Dim myKey As Variant: myKey = Array(KeyByIndex(myRemoveIndex))
        If VBA.IsNull(myKey(0)) Then
            Exit For
        End If
        RemoveKey myKey(0)
        myRemoveIndex = myRemoveIndex - 1
    Next
    
    Set RemoveBefore = Me
    
End Function


Public Function RemoveAll(ParamArray ipParamarray() As Variant) As KvpHA

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


Public Function Clear() As KvpHA
    Set Clear = Reinit(s.HashTableSize)
End Function


Public Function Reset() As KvpHA
    Set Reset = Reinit(s.HashTableSize)
End Function

Public Function RemoveByIndex(ParamArray ipParamarray() As Variant) As KvpHA

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
Public Function RemoveRangeByIndex(ByRef ipRange As Variant) As KvpHA
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
        
        RemoveKey KeyByIndex(myIndex)
        
ContinueLoop:
    Loop While myIndexes.MovePrev
    
    Set RemoveRangeByIndex = Me
    
End Function

Public Function Clone() As KvpHA
    Set Clone = KvpHA.Deb(s.HashTableSize).AddPairs(Me.Keys, Me.Items)
End Function

Public Function Exists(ByRef ipKey As Variant) As Boolean

    Exists = False
    
    If Guard.KeyIsAdmin(ipKey, MY_TYPENAME & ".Exists", REPORT_BACK) Then
        Exit Function
    End If
    
    If Me.LacksItems Then
        Exit Function
    End If
    
'    Dim myH As Long: myH = GetHash(ipKey)
'    If VBA.IsEmpty(s.HashSlots(myH)) Then
'        Exit Function
'    End If

    Exists = FindIndex(ipKey) >= 0
    
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
    ExistsItem = Not VBA.IsNull(KeyOf(ipItem))
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
    HoldsItems = p.Count >= p.FirstIndex
End Function

Public Function LacksItems() As Boolean
    LacksItems = p.Count < FIRST_INDEX
End Function


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property


Public Property Get HasUniqueKeys() As Boolean
    HasUniqueKeys = s.EnsureUniqueKeys
End Property

'@Description("Mutates the Kvpc in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As KvpHA
Attribute MappedIt.VB_Description = "Mutates the Kvpc in place by applying the function in IMapper to each item"

    Set MappedIt = Me
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myIndex As Long
    For myIndex = LBound(s.Keys) To UBound(s.Keys)
        If Not VBA.IsEmpty(s.Keys(myIndex)) Then
            Dim myItem As Variant
            myItem = ipMapper.ExecMapper(s.Values(myIndex))
            If VBA.IsObject(myItem(0)) Then
                Set s.Values(myIndex) = myItem(0)
            Else
                s.Values(myIndex) = myItem(0)
            End If
        End If
    Next
    
End Function


'@Description("Returns a Kvpc.Deb by applying the function in IMapper to each Item")
Public Function MapIt(ByVal ipMapper As IMapper) As KvpHA
Attribute MapIt.VB_Description = "Returns a Kvpc.Deb by applying the function in IMapper to each Item"

    Dim myK As KvpHA: Set myK = KvpHA.Deb(s.HashTableSize)
    Set MapIt = myK
    
    If Me.LacksItems Then
        Exit Function
    End If

    Dim myIndex As Long
    For myIndex = LBound(s.Keys) To UBound(s.Keys)
        If Not VBA.IsEmpty(s.Keys(myIndex)) Then
            Dim myItem As Variant
            myItem = ipMapper.ExecMapper(s.Values(myIndex))
            myK.Add s.Keys(myIndex), myItem(0)
        End If
    Next
    
End Function


'@Description("Returns a Kvpc.Deb where the ICOmparer has been applied to the Items of the original Kvp")
Public Function FilterIt(ByVal ipComparer As IComparer) As KvpHA
Attribute FilterIt.VB_Description = "Returns a Kvpc.Deb where the ICOmparer has been applied to the Items of the original Kvp"

    Dim myK As KvpHA: Set myK = KvpHA.Deb
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

Public Function Foldl _
( _
    ByRef ipFolder As IFold, _
    Optional ByRef ipAccumulator As Variant = 0, _
    Optional ByRef ipAccCmp As IComparer = Nothing, _
    Optional ByRef ipItemCmp As IComparer = Nothing, _
    Optional ByVal ipFrom As Long = 1, _
    Optional ByVal ipTo As Long = -1, _
    Optional ByVal ipStep As Long = 1 _
) As Variant

    Dim myFrom As Long
    Guard.MustBeAtLeastStartOrMinusOne ipFrom, Me.FirstIndex, "KvHA.Foldl"
    If ipFrom = -1 Then
        myFrom = Me.FirstIndex
    Else
        myFrom = ipFrom
    End If
    
    Guard.IndexOutOfBounds myFrom, Me, "KvHA.Foldl"
    
    Dim myTo As Long
    Guard.MustBeAtLeastStartOrMinusOne ipTo, myFrom, "KvHA.Foldl"
    If ipTo = -1 Then
        myTo = Me.LastIndex
    End If
    
    Guard.IndexOutOfBounds myTo, Me, "KvHA.Foldl"
    
    Dim myStep As Long
    Guard.MustBeAtLeastOne ipStep, "KvHA.Foldr"
    myStep = ipStep

    
    Dim myAccumulator() As Variant: myAccumulator = Array(ipAccumulator)
    
    Dim myItems As IterItems: Set myItems = IterItems(Me).FTS(myFrom, myTo, myStep)
    Do
        Dim myA() As Variant: myA = Array(myItems.CurItem(0))
        
        If IsNotNothing(ipItemCmp) Then
            If ipItemCmp.ExecCmp(myA(0)) Then
                GoTo ContinueDo
            End If
        End If
        
        ' exec fold returns a single item array
        myAccumulator = ipFolder.ExecFold(myAccumulator(0), myA(0))
        If IsNotNothing(ipAccCmp) Then
            If ipAccCmp.ExecCmp(myAccumulator(0)) Then
                If VBA.IsObject(myAccumulator(0)) Then
                    Set Foldl = myAccumulator(0)
                Else
                    Foldl = myAccumulator(0)
                End If
                Exit Function
            End If
        End If
        
ContinueDo:
    Loop While myItems.MoveNext

    If VBA.IsObject(myAccumulator(0)) Then
        Set Foldl = myAccumulator(0)
    Else
        Foldl = myAccumulator(0)
    End If
    
End Function

Public Function Foldr _
( _
    ByRef ipFolder As IFold, _
    Optional ByRef ipAccumulator As Variant = 0, _
    Optional ByRef ipAccCmp As IComparer = Nothing, _
    Optional ByRef ipItemCmp As IComparer = Nothing, _
    Optional ByVal ipFrom As Long = -1, _
    Optional ByVal ipTo As Long = -1, _
    Optional ByVal ipStep As Long = 1 _
) As Variant

    Dim myFrom As Long
    Guard.MustBeAtLeastStartOrMinusOne ipFrom, Me.FirstIndex, "KvHA.Foldr"
    If ipFrom = -1 Then
        myFrom = Me.FirstIndex
    Else
        myFrom = ipFrom
    End If
    
    Guard.IndexOutOfBounds myFrom, Me, "KvHA.Foldr"
    
    Dim myTo As Long
    Guard.MustBeAtLeastStartOrMinusOne ipTo, myFrom, "KvHA.Foldr"
    If ipTo = -1 Then
        myTo = Me.LastIndex
    End If
    
    Guard.IndexOutOfBounds myTo, Me, "KvHA.Foldr"
    
    Dim myStep As Long
    Guard.MustBeAtLeastOne ipStep, "KvHA. Foldr"
    myStep = ipStep
    
    
    Dim myAccumulator() As Variant: myAccumulator = Array(ipAccumulator)
    
    Dim myItems As IterItems: Set myItems = IterItems(Me).FTS(myFrom, myTo, myStep).MoveToEndIndex
    Do
        Dim myA() As Variant: myA = Array(myItems.CurItem(0))
        
        If IsNotNothing(ipItemCmp) Then
            If ipItemCmp.ExecCmp(myA(0)) Then
                GoTo ContinueDo
            End If
        End If
        
        ' exec fold returns a single item array
        myAccumulator = ipFolder.ExecFold(myAccumulator(0), myA(0))
        If IsNotNothing(ipAccCmp) Then
            If ipAccCmp.ExecCmp(myAccumulator(0)) Then
                If VBA.IsObject(myAccumulator(0)) Then
                    Set Foldr = myAccumulator(0)
                Else
                    Foldr = myAccumulator(0)
                End If
                Exit Function
            End If
        End If
        
ContinueDo:
    Loop While myItems.MovePrev

    If VBA.IsObject(myAccumulator(0)) Then
        Set Foldr = myAccumulator(0)
    Else
        Foldr = myAccumulator(0)
    End If
    
End Function

Public Function Count(Optional ByRef ipItem As Variant) As Long

    Count = -1
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If VBA.IsMissing(ipItem) Then
        Count = p.Count
        Exit Function
    End If
    
    Dim myCount As Long: myCount = 0
    
    Dim myIndex As Long
    For myIndex = LBound(s.Keys) To UBound(s.Keys)
        If Not VBA.IsEmpty(s.Keys(myIndex)) Then
            If Comparers.EQ(ipItem, s.Values(myIndex)) Then
                myCount = myCount + 1
            End If
        End If
    Next
    
    Count = myCount
    
End Function


Public Function CountIt(ByVal ipComparer As IComparer) As Long

    CountIt = -1
    
    If Me.LacksItems Then
        Exit Function
    End If
        
    Dim myCount As Long: myCount = 0
    
    Dim myIndex As Long
    For myIndex = LBound(s.Keys) To UBound(s.Keys)
        If Not VBA.IsEmpty(s.Keys(myIndex)) Then
            If ipComparer.ExecCmp(s.Values(myIndex)) Then
                myCount = myCount + 1
            End If
        End If
    Next
    
    CountIt = myCount

End Function


Public Function KeyOf(ByRef ipItem As Variant) As Variant
    
    KeyOf = Null
    
    Dim myIndex As Long
    For myIndex = LBound(s.Keys) To UBound(s.Keys)
        If Not VBA.IsEmpty(s.Keys(myIndex)) Then
            If Comparers.EQ(s.Values(myIndex), ipItem) Then
                If VBA.IsObject(s.Keys(myIndex)) Then
                    Set KeyOf = s.Keys(myIndex)
                Else
                    KeyOf = s.Keys(myIndex)
                End If
                Exit Function
            End If
        End If
    Next
    
End Function

Public Function Items() As Variant()

    If p.Count = 0 Then
        Items = Array()
        Exit Function
    End If
    
    Dim i As Long
    Dim j As Long
    
    Dim V() As Variant
    ReDim V(1 To p.Count)
    
    For i = 1 To p.Count
        Do While IsEmpty(s.Keys(j)): j = j + 1: Loop
    
        VariantCopy V(i), ByVal VarPtr(s.Values(j))
        j = j + 1
        
    Next
    
    'ReDim Preserve V(LBound(V) + 1 To UBound(V) + 1)
    Items = V
    
End Function

Public Function ItemsAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    
    Dim myItems As Variant: myItems = Me.Items
    
    Dim myItem As Variant
    For Each myItem In myItems
        myC.Add myItem
    Next
    
    Set ItemsAsColl = myC
    
End Function

Public Function ItemsAsSeq() As SeqA
    Set ItemsAsSeq = SeqA.Deb(s.HashTableSize).AddRange(Me.Items)
End Function

Public Function Keys() As Variant()

    If p.Count = 0 Then
        Keys = Array()
        Exit Function
    End If
    
    Dim i As Long
    Dim j As Long
    Dim V() As Variant
    ReDim V(1 To p.Count)
    
    For i = 1 To p.Count
    
        Do While IsEmpty(s.Keys(j)): j = j + 1: Loop
       
        VariantCopy V(i), ByVal VarPtr(s.Keys(j))
        j = j + 1
        
    Next
    
    'ReDim Preserve V(LBound(V) + 1 To UBound(V) + 1)
    Keys = V
    
End Function

Public Function KeysAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    
    Dim myKeys As Variant: myKeys = Me.Keys
    
    Dim myKey As Variant
    For Each myKey In myKeys
        myC.Add myKey
    Next
    
    Set KeysAsColl = myC
    
End Function

Public Function KeysAsSeq() As SeqL
    Set KeysAsSeq = SeqL.Deb(s.HashTableSize).AddRange(Me.Keys)
End Function


Public Function Pairs() As Variant()             'hand-out the Key-Values in a Pairs-Array

    If p.Count = 0 Then
        Pairs = Array()
        Exit Function
    End If
    
    Dim i As Long
    Dim j As Long
    
    Dim V() As Variant
    ReDim V(1 To p.Count)
    
    For i = 1 To p.Count
    
        Do While IsEmpty(s.Keys(j)): j = j + 1: Loop
        V(i) = Array(s.Keys(j), s.Values(j))
        j = j + 1
        
    Next
    
    Pairs = V
    
End Function

Public Function PairsAsColl() As Collection

    Dim myC As Collection: Set myC = New Collection
    
    Dim myPairs As Variant: myPairs = Me.Pairs
    
    Dim myPair As Variant
    For Each myPair In myPairs
        myC.Add myPair
    Next
    
    Set PairsAsColl = myC
    
End Function

Public Function PairsAsSeq() As SeqA
    Set PairsAsSeq = SeqA(Me.Pairs)
End Function

Public Property Get FirstIndex() As Long

    If p.Count < FIRST_INDEX Then
        FirstIndex = -1
    Else
        FirstIndex = p.FirstIndex
    End If

End Property

Public Property Get LastIndex() As Long

    If p.Count < FIRST_INDEX Then
        LastIndex = -1
    Else
        LastIndex = p.Count
    End If

End Property
'
''@Description("Returns the index of the second item in the kvp or -1  if the kvp has less than 2 Items")
'Public Function FBOIndex() As Long
'
'    If Me.LacksItems Then
'        FBOIndex = -1
'    ElseIf p.count = p.firstindex Then
'        FBOIndex = -1
'    Else
'        FBOIndex = p.firstindex + 1
'    End If
'
'End Function
'
'
''Description("Returns the index of the second to last item in the SeqHA, or -1 if the SeqHA has lees than 2 items")
'Public Function LBOIndex() As Long
'
'    If Me.LacksItems Then
'        LBOIndex = -1
'    ElseIf p.count = p.firstindex Then
'        LBOIndex = -1
'    Else
'        LBOIndex = p.count - 1
'    End If
'
'End Function

Public Property Get FirstKey() As Variant

    FirstKey = Null

    If Me.LacksItems Then
        Exit Property
    End If

    Dim myKey As Variant: myKey = Array(KeyByIndex(Me.FirstIndex))

    If VBA.IsObject(myKey(0)) Then
        Set FirstKey = myKey(0)
    Else
        FirstKey = myKey(0)
    End If

End Property


Public Property Get LastKey() As Variant

    LastKey = Null

    If Me.LacksItems Then
        Exit Property
    End If

    Dim myKey As Variant: myKey = Array(KeyByIndex(Me.LastIndex))

    If VBA.IsObject(myKey(0)) Then
        Set LastKey = myKey(0)
    Else
        LastKey = myKey(0)
    End If

End Property

'
'
Public Property Get First() As Variant

    First = Null

    If Me.LacksItems Then
        Exit Property
    End If

    Dim myItem As Variant: myItem = Array(ItemByIndex(Me.FirstIndex))

    If VBA.IsObject(myItem(0)) Then
        Set First = myItem(0)
    Else
        First = myItem(0)
    End If

End Property


Public Property Set First(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Set First"

    Set ItemByIndex(Me.FirstIndex) = ipItem

End Property


Public Property Let First(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Let First"

    ItemByIndex(Me.FirstIndex) = ipItem

End Property


Public Property Get Last() As Variant

    Last = Null

    If Me.LacksItems Then
        Exit Property
    End If

    Dim myItem As Variant: myItem = Array(ItemByIndex(Me.LastIndex))

    If VBA.IsObject(myItem(0)) Then
        Set Last = myItem(0)
    Else
        Last = myItem(0)
    End If

End Property


Public Property Set Last(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Set Last"

    Set ItemByIndex(Me.LastIndex) = ipItem

End Property


Public Property Let Last(ByVal ipItem As Variant)

    Guard.KeyNotFound Me.LacksItems, FIRST_INDEX, MY_TYPENAME & ".Let Last"

    ItemByIndex(Me.LastIndex) = ipItem

End Property

'Public Property Get FBOKey() As Variant
'
'    Dim myKey As Variant: myKey = Array(KeyByIndex(FIRST_INDEX + 1))
'    If VBA.IsObject(myKey(0)) Then
'
'End Property
'
'Public Property Get FBOKey() As Variant
'    Dim myKey As Variant: myKey = Array(s.Keys.LBO)
'    If VBA.IsObject(myKey(0)) Then
'        Set LBOKey = myKey(0)
'    Else
'        LBOKey = myKey(0)
'    End If
'End Property
'
'Public Property Get LBOKey() As Variant
'
'    Dim myKey As Variant: myKey = Array(s.Keys.LBO)
'    If VBA.IsObject(myKey(0)) Then
'        Set LBOKey = myKey(0)
'    Else
'        LBOKey = myKey(0)
'    End If
'
'End Property
'
'
'
'Public Property Set FBO(ByVal ipItem As Variant)
'    Set s.Items.FBO = ipItem
'End Property
'
'
'Public Property Let LBO(ByVal ipItem As Variant)
'    s.Items.FBO = ipItem
'End Property
'
'
'Public Property Get LBO() As Variant
'
'    FBO = Null
'
'    If LacksItems Then
'        Exit Property
'    ElseIf s.OrderLast.OrderPred Is Nothing Then
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
'Public Property Set LBO(ByVal ipItem As Variant)
'    Set Me.ItemByIndex(p.count - 1) = ipItem
'End Property
'
'
'Public Property Let LBO(ByVal ipItem As Variant)
'   Me.ItemByIndex(p.count - 1) = ipItem
'End Property
'
Public Property Get EnsureUniqueKeys() As Boolean
    EnsureUniqueKeys = p.EnsureUniqueKeys
End Property

Public Property Let EnsureUniqueKeys(ByVal ipEnsureUniqueKeys As Boolean)
    p.EnsureUniqueKeys = ipEnsureUniqueKeys
End Property

'@Description( returns a Kvp with values as keys.  Each Value is linked to a Kvp of Key vs Index for the host value
Public Function Mirror(Optional ByVal ipMirrorType As e_MirrorType = e_MirrorType.m_ByAllValues) As KvpHA

    If Me.LacksItems = 0 Then
        Set Mirror = KvpHA.Deb
        Exit Function
    End If

    Select Case ipMirrorType

        Case m_ByAllValues:             Set Mirror = MirrorByAllValues
        Case m_ByFirstValue:            Set Mirror = MirrorByFirstValue
        Case Else
            Err.Raise 17 + vbObjectError, _
                "VBALib.KvpHA.Mirror", _
                Fmt.Text("Mirror enum not known. Got {0}", ipMirrorType)
                
    End Select

End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'Private functions
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
'@Description("Returns an KvpH of two KvpH.  Item(1) is an KvpH of the first found Host value vs Host Key. .Item(2) is a copy of the original KvpHA with the  first Key/Value pairs removed
Private Function MirrorByAllValues() As KvpHA

    If Me.LacksItems = 0 Then
        Set MirrorByAllValues = KvpHA.Deb(s.HashTableSize)
        Exit Function
    End If
    
    Dim myK As KvpHA
    Set myK = KvpHA.Deb

    Dim myPairs As Variant: myPairs = Me.Pairs
    
    Dim myPair As Variant
    Dim myIndex As Long: myIndex = 1
    For Each myPair In myPairs
        If myK.LacksKey(myPair(1)) Then
            myK.Add myPair(1), KvpHA.Deb
        End If

        myK.Item(myPair(1)).Add myPair(0), myIndex
        myIndex = myIndex + 1
    Next
    
    Set MirrorByAllValues = myK
    
End Function


Private Function MirrorByFirstValue() As KvpHA

    Dim myFirst As KvpHA: Set myFirst = KvpHA.Deb
    Dim myOther As KvpHA: Set myOther = KvpHA.Deb

    Dim myPairs As Variant: myPairs = Me.Pairs
    
    Dim myPair As Variant
    For Each myPair In myPairs

        If myFirst.LacksKey(myPair(1)) Then
            myFirst.Add myPair(1), myPair(0)
        Else
            myOther.Add myPair(0), myPair(1)
        End If

    Next

    Set MirrorByFirstValue = KvpHA.Deb.Add(0&, myFirst).Add(1&, myOther)

End Function

Private Function FindIndex(ByRef Key As Variant, Optional ByRef H As Long) As Long 'return -1, when no Key can be found

    Dim i As Long
    Dim D As Double
    Dim L As LongPtr
    Dim F As Single
    Dim VT As VbVarType
    Dim HTUB As Long
     
    Dim C As Variant
    C = VBA.CDec(C)
     
    HTUB = s.HashTableSize - 1
    FindIndex = H
    H = HTUB                                     'init the HashValue (all bits to 1)
     
    VT = VarType(Key)
    Select Case VT
        Case vbString
        
            s.sAW.cElements1D = Len(Key)
            s.sAW.pvData = StrPtr(Key)
            
            If s.CompareMode = 0 Then
            
                For i = 0 To s.sAW.cElements1D - 1
                    H = (H + s.W(i)) * HMul And HTUB
                Next
                
                If FindIndex = -1 Then
                    Exit Function 'it's a "Hash-Only" Calculation
                End If
            
                For i = 0 To s.HashTable(H).Count - 1
                
                    If i < DynTakeOver Then
                        FindIndex = s.HashTable(H).DataIdxsStat(i)
                    Else
                        FindIndex = s.HashTable(H).DataIdxsDyn(i)
                    End If
                    
                    If VarType(s.Keys(FindIndex)) = VT Then
                        If Key = s.Keys(FindIndex) Then
                            Exit Function
                        End If
                    End If
                    
                Next
                
            Else
            
                For i = 0 To s.sAW.cElements1D - 1
                    H = (H + LWC(s.W(i))) * HMul And HTUB
                Next
                
                If FindIndex = -1 Then
                    Exit Function 'it's a "Hash-Only" Calculation
                End If
            
                For i = 0 To s.HashTable(H).Count - 1
                
                    If i < DynTakeOver Then
                        FindIndex = s.HashTable(H).DataIdxsStat(i)
                    Else
                        FindIndex = s.HashTable(H).DataIdxsDyn(i)
                    End If
                    
                    If VarType(s.Keys(FindIndex)) = VT Then
                        If StrComp(Key, s.Keys(FindIndex), s.CompareMode) = 0 Then
                            Exit Function
                        End If
                    End If
                    
                Next
                
            End If
         
        Case vbObject
        
            L = ObjPtr(Key)
            s.sAW.cElements1D = 2
            s.sAW.pvData = VarPtr(L)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If Key Is s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
         
        Case vbCurrency
        
            C = Key
            s.sAW.cElements1D = 4
            s.sAW.pvData = VarPtr(C)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            H = (H + s.W(2)) * HMul And HTUB
            H = (H + s.W(3)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If C = s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
         
        Case vbLong, vbInteger, vbByte
        
            L = Key
            s.sAW.cElements1D = 2
            s.sAW.pvData = VarPtr(L)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
         
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                Select Case VarType(s.Keys(FindIndex))
                    Case vbLong, vbInteger, vbByte
                        If L = s.Keys(FindIndex) Then
                            Exit Function
                        End If
                End Select
                
            Next
    
        Case vbDouble
        
            D = Key: s.sAW.cElements1D = 4
            s.sAW.pvData = VarPtr(D)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            H = (H + s.W(2)) * HMul And HTUB
            H = (H + s.W(3)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If D = s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
       
        Case vbDate
        
            D = Key
            s.sAW.cElements1D = 4
            s.sAW.pvData = VarPtr(D)
            
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            H = (H + s.W(2)) * HMul And HTUB
            H = (H + s.W(3)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If D = s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
         
        Case vbSingle
            F = Key
            s.sAW.cElements1D = 2
            s.sAW.pvData = VarPtr(F)
            H = (H + s.W(0)) * HMul And HTUB       'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And HTUB
            
            If FindIndex = -1 Then
                Exit Function 'it's a "Hash-Only" Calculation
            End If
            
            For i = 0 To s.HashTable(H).Count - 1
            
                If i < DynTakeOver Then
                    FindIndex = s.HashTable(H).DataIdxsStat(i)
                Else
                    FindIndex = s.HashTable(H).DataIdxsDyn(i)
                End If
                
                If VarType(s.Keys(FindIndex)) = VT Then
                    If F = s.Keys(FindIndex) Then
                        Exit Function
                    End If
                End If
                
            Next
    End Select
     
    FindIndex = -1
    
End Function


'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub AdjustIndex(ByRef Idx As Long, Optional ByVal Inverse As Boolean)

    Dim i As Long
    Dim Incr As Long
    
    'If Inverse Then Incr = -1 Else Incr = 1
    Incr = IIf(Inverse, -1, 1)
    
    For i = 0 To s.FCount - 1
        If s.FIndexes(i) <= Idx Then
            Idx = Idx + Incr
        End If
    Next
End Sub


Friend Sub CheckHashDistribution()
    Dim i As Long
    Dim Count As Long
    Dim cc As Long
    Dim Min As Long
    Dim Max As Long
    Min = &H7FFFFFFF
    For i = 0 To UBound(s.HashTable)
        Count = s.HashTable(i).Count
        If Count Then
            If Min > Count Then Min = Count
            If Max < Count Then Max = Count
            cc = cc + 1
        End If
    Next
    Debug.Print "Distribution of"; p.Count; "entries over a HashTable with"; UBound(s.HashTable) + 1; "slots:"
    Debug.Print "Used-HashSlots:"; cc
    Debug.Print "Min-Entries:"; Min
    Debug.Print "Max-Entries:"; Max
End Sub


Private Sub RemoveAllItems()
    Me.Reinit s.HashTableSize
End Sub


Private Sub RemoveAllSingleItem(ByRef ipItem As Variant)
    Do While HoldsKey(ipItem)
        RemoveKey ipItem
    Loop
End Sub

''@Description("Returns the first node containing ipItem")
'Private Function GetNodeOf(ByRef ipItem As Variant) As Variant
'
'    Set GetNodeOf = Nothing
'
'    Dim myNode As SeqHNode: Set myNode = s.HashTable(s.Hasher.OSHash(ipItem))
'
'    If NotFound(myNode) Then
'        Exit Function
'    End If
'
'    ' the item of the hash value may not match
'    Do Until Comparers.EQ(myNode.Item, ipItem)
'
'        If myNode.HashSucc Is Nothing Then
'            Exit Function
'        End If
'
'        If myNode.Hash <> myNode.HashSucc.Hash Then
'            Exit Function
'        End If
'
'        Set myNode = myNode.HashSucc
'
'    Loop
'
'    ' we now have a matching hash and item
'    Set GetNodeOf = myNode
'
'End Function

