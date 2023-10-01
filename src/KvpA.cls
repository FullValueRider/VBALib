VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Dictionary using a SeqA for Keys and Items"
'@PredeclaredId
'@Exposed
'@Folder("Kvp")
'@ModuleDescription("Dictionary using a SeqA for Keys and Items")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
Option Explicit

Private Const MY_TYPENAME                       As String = "KvpA"
Private Const FIRST_INDEX                       As Long = 1
Private Const REPORT_BACK                       As Boolean = True
Private Const DEFAULT_SIZE                      As Long = 16
Private Type State
    Keys                                        As SeqA
    Items                                       As SeqA
End Type


Private s                                       As State

' Checking for duplicate keys is a linear time process
' for KvpA, KvpC and KvpL.  Consequently for **ALL** Kvp classes
' EnsureUniqueKeys is set to false by default
' Setting EnsureUniqueKeys to true will allow a check for a
' duplicate key each time a key is added.
' From a performance perspective it may be better
' to disallow duplicate keys and to then subsequently
' extract unique keys using the UniqueKeys method.

Private Type Properties
    EnsureUniqueKeys                            As Boolean
    FirstIndex                                  As Long
End Type

Private p                                       As Properties

'@Defaultmember
Public Function Deb(Optional ByVal ipInitialSize As Long = DEFAULT_SIZE) As KvpA
Attribute Deb.VB_UserMemId = 0
    With New KvpA
        Set Deb = .ConstructInstance(ipInitialSize)
    End With
End Function


Friend Function ConstructInstance(ByVal ipInitialSize As Long) As KvpA
    p.FirstIndex = FIRST_INDEX
    Set s.Keys = SeqA(ipInitialSize)
    Set s.Items = SeqA(ipInitialSize)
    p.EnsureUniqueKeys = False
    Set ConstructInstance = Me
End Function


Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpA
   
    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".AddPairs"
    Guard.EnsureUniqueKeys ipKey, Me, "AddPairs"
    
    s.Keys.Add ipKey
    s.Items.Add ipItem
    
    Set Add = Me
   
End Function


Public Function AddPairs(ByRef ipKeys As Variant, ByRef ipItems As Variant) As KvpA
   
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

    Guard.KeyIsAdmin ipKey, MY_TYPENAME & ".Set Item"
    
    ' Kvp classes do not allow adding a key by assignment
    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    
    Guard.KeyNotFound myIndex, ipKey, MY_TYPENAME & ".Set Item"
    
    If VBA.IsObject(ipItem) Then
        Set s.Items.Item(myIndex) = ipItem
    Else
        s.Items.Item(myIndex) = ipItem
    End If
   
End Property

Public Property Get ItemByIndex(ByRef ipIndex As Long) As Variant

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
    
    If Guard.IndexOutOfBounds(myIndex, Me, "Get ItemByIndex", REPORT_BACK) Then
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

Public Function RemoveAt(ByRef ipIndex As Long) As KvpA

    Set RemoveAt = Me

    If ipIndex < s.Keys.FirstIndex Or ipIndex > s.Keys.LastIndex Then
        Exit Function
    End If
    
    s.Keys.RemoveAt ipIndex
    s.Items.RemoveAt ipIndex
    
End Function


''@Description("Removes Key/Value pairs that exist in the KvpA")
'Public Function Remove(ParamArray ipParamarray() As Variant) As KvpA
'
'    Set Remove = Me
'
'    If s.Keys.LacksItems Then
'        Exit Function
'    End If
'
'    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(ipParamarray)
'
'    If Not VBA.IsArray(myParamarray) Then
'        Err.Raise 17 + vbObjectError, _
'        "VBALib.KvpA.Remove", _
'        Fmt.Text("Expecting an array.  Got {0}", VBA.TypeName(myParamarray))
'    End If
'
'    Dim myItem As Variant
'    For Each myItem In myParamarray
'
'        If VBA.IsNull(myItem) Then
'            Err.Raise 17 + vbObjectError, _
'            "VBALib.KvpA.Remove", _
'            "Null is not a valid Key"
'        End If
'
'        Dim myIndex As Long
'        myIndex = s.Keys.IndexOf(myItem)
'        If myIndex > 0 Then
'            s.Keys.RemoveAt myIndex
'            s.Items.RemoveAt myIndex
'        End If
'    Next
'
'    Set Remove = Me
'
'End Function

''@Description("Removes Key/Value pairs that exist in the KvpA")
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
'    Set Remove = Me
'
'End Function


'@Description("Removes Key/Value pairs that exist in the KvpA")
Public Function Remove(ParamArray ipParamarray() As Variant) As KvpA
Attribute Remove.VB_Description = "Removes Key/Value pairs that exist in the KvpA"

    Set Remove = Me
    
    If Me.LacksItems Then
        Exit Function
    End If
    
    If ArrayOp.LacksItems(VBA.CVar(ipParamarray)) Then
        Exit Function
    End If

    
    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(ipParamarray)
        
    Dim myKeys As IterItems: Set myKeys = IterItems(myParamarray)
    Do
        Dim myKey As Variant: myKey = Array(myKeys.CurItem(0))
        If Guard.KeyIsAdmin(myKey(0), MY_TYPENAME & ".Remove", REPORT_BACK) Then
            GoTo ContinueLoop:
        End If
    
        Dim myIndex As Long: myIndex = s.Keys.IndexOf(myKey(0))
        If Guard.IndexOutOfBounds(myIndex, Me, "Remove", REPORT_BACK) Then
            GoTo ContinueLoop
        End If
    
        s.Keys.RemoveAt myIndex
        s.Items.RemoveAt myIndex

ContinueLoop:
    Loop While myKeys.MoveNext

    Set Remove = Me
    
End Function

'Public Function RemoveByIndex(ByRef ipIndex As Long) As KvpA
'
'    Set RemovbyIndex = Me
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

Public Function RemoveByIndex(ParamArray ipParamarray() As Variant) As KvpA

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

Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpA

    Set RemoveAfter = Me
    
    If Me.LacksItems Then
        Exit Function
    End If

    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpA.RemoveAfter", _
        "Null is not a valid Key"
    End If
    
    If s.Keys.LacksItems Then
        Exit Function
    End If
    
    If Me.LacksKey(ipKey) Then
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpA.RemoveAfter", _
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


Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpA

    Set RemoveBefore = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpA.RemoveBefore", _
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
        "VBALib.KvpA.RemoveAfter", _
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


Public Function Clear() As KvpA
    s.Keys.RemoveAll
    s.Items.RemoveAll
    Set Clear = Me
End Function


Public Function RemoveAll() As KvpA
    s.Keys.RemoveAll
    s.Items.RemoveAll
    Set RemoveAll = Me
End Function


Public Function Reset() As KvpA
    Set s.Keys = SeqA.Deb
    Set s.Items = SeqA.Deb
    Set Reset = Me
End Function


'===================================================================================
' ToDo: add insert methods
'===================================================================================

Public Function Clone() As KvpA
    Set Clone = KvpA.Deb.AddPairs(s.Keys, s.Items)
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
        "VBALib.KvpA.Exists", _
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
        "VBALib.KvpA.HoldsKey", _
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
        "VBALib.KvpA.ContainsKey", _
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
        "VBALib.KvpA.LacksKey", _
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


'@Description("Mutates the KvpA in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As KvpA
Attribute MappedIt.VB_Description = "Mutates the KvpA in place by applying the function in IMapper to each item"

    If s.Keys.LacksItems Then
        Set MappedIt = Me
        Exit Function
    End If

    s.Items.MappedIt ipMapper
    Set MappedIt = Me

End Function


'@Description("Returns a KvpA.Deb by applying the function in IMapper to each Item")
Public Function MapIt(ByVal ipMapper As IMapper) As KvpA
Attribute MapIt.VB_Description = "Returns a KvpA.Deb by applying the function in IMapper to each Item"

    If s.Keys.LacksItems Then
        Set MapIt = Me
        Exit Function
    End If


    Dim myS As SeqA: Set myS = s.Items.MapIt(ipMapper)
      
    Set MapIt = KvpA.Deb.AddPairs(s.Keys, myS)

End Function


'@Description("Returns a KvpA.Deb where the ICOmparer has been applied to the Items of the original Kvp")
Public Function FilterIt(ByVal ipComparer As IComparer) As KvpA
Attribute FilterIt.VB_Description = "Returns a KvpA.Deb where the ICOmparer has been applied to the Items of the original Kvp"

    Dim myK As KvpA: Set myK = KvpA.Deb
    
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

    Dim myA() As Variant: myA = ipReducer.ExecReduction(s.Items)
    
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
    Guard.MustBeAtLeastStartOrMinusOne ipFrom, Me.FirstIndex, "KvpA.Foldl"
    If ipFrom = -1 Then
        myFrom = Me.FirstIndex
    Else
        myFrom = ipFrom
    End If
    
    Guard.IndexOutOfBounds myFrom, Me, "KvpA.Foldl"
    
    Dim myTo As Long
    Guard.MustBeAtLeastStartOrMinusOne ipTo, myFrom, "KvpA.Foldl"
    If ipTo = -1 Then
        myTo = Me.LastIndex
    End If
    
    Guard.IndexOutOfBounds myTo, Me, "KvpA.Foldl"
    
    Dim myStep As Long
    Guard.MustBeAtLeastOne ipStep, "KvpA.Foldr"
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
    Guard.MustBeAtLeastStartOrMinusOne ipFrom, Me.FirstIndex, "KvpA.Foldr"
    If ipFrom = -1 Then
        myFrom = Me.FirstIndex
    Else
        myFrom = ipFrom
    End If
    
    Guard.IndexOutOfBounds myFrom, Me, "KvpA.Foldr"
    
    Dim myTo As Long
    Guard.MustBeAtLeastStartOrMinusOne ipTo, myFrom, "KvpA.Foldr"
    If ipTo = -1 Then
        myTo = Me.LastIndex
    End If
    
    Guard.IndexOutOfBounds myTo, Me, "KvpA.Foldr"
    
    Dim myStep As Long
    Guard.MustBeAtLeastOne ipStep, "KvpA. Foldr"
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

    Dim myCount As Long: myCount = 0
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
'Public Function ForbidSameKeys() As KvpA
'    p.EnsureUniqueKeys = True
'    Set ForbidSameKeys = Me
'End Function
'
'
'Public Function AllowSameKeys() As KvpA
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


Public Function ItemsAsSeq() As SeqA
    
    If s.Keys.LacksItems Then
        Set ItemsAsSeq = SeqA.Deb
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


Public Function KeysAsSeq() As SeqA
    
    If s.Keys.LacksItems Then
        Set KeysAsSeq = SeqA.Deb
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



Public Function PairsAsSeq() As SeqA

    If s.Keys.LacksItems Then
        Set PairsAsSeq = SeqA.Deb
        Exit Function
    End If
    
    Dim myS As SeqA: Set myS = SeqA.Deb
    
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
        "VBALib.KvpA.Set FirstItem", _
        "Kvp has no Keys"
    End If

    Set s.Items.First = ipItem

End Property


Public Property Let First(ByVal ipItem As Variant)

    If s.Keys.LacksItems Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpA.Let_FirstItem", _
        "Kvp has no Keys"
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
        "VBALib.KvpA.Set LastItem", _
        "Kvp has no Keys"
    End If

    Set s.Items.Last = ipItem
    
End Property


Public Property Let Last(ByVal ipItem As Variant)

    If s.Keys.LacksItems Then
        Err.Raise 17, _
        "VBALib.KvpA.Set LastItem", _
        "Kvp has no Keys"
    End If

    s.Items.Last = ipItem
    
End Property


'@Description( returns a Kvp with values as keys.  Each Value is linked to a Kvp of Key vs Index for the host value
Public Function Mirror(Optional ByVal ipMirrorType As e_MirrorType = e_MirrorType.m_ByAllValues) As KvpA

    If Me.LacksItems Then
        Set Mirror = KvpA.Deb
        Exit Function
    End If

    Select Case ipMirrorType

        Case m_ByAllValues:             Set Mirror = MirrorByAllValues
        Case m_ByFirstValue:            Set Mirror = MirrorByFirstValue
        Case Else
            Err.Raise 17 + vbObjectError, _
            "VBALib.KvpA.Mirror", _
            Fmt.Text("Mirror enum not known. Got {0}", ipMirrorType)
                
    End Select

End Function


'@Description("Returns an KvpA of two KvpA.  Item(1) is an KvpA of the first found Host value vs Host Key. .Item(2) is a copy of the original KvpA with the  first Key/Value pairs removed
Private Function MirrorByAllValues() As KvpA

'    If Me.LacksItems Then
'        Set MirrorByAllValues = KvpA.Deb
'        Exit Function
'    End If

    Dim myK As KvpA
    Set myK = KvpA.Deb

    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex

        If myK.LacksKey(s.Items.Item(myIndex)) Then
            myK.Add s.Items.Item(myIndex), KvpA.Deb
        End If

        myK.Item(s.Items.Item(myIndex)).Add s.Keys.Item(myIndex), myIndex

    Next

    Set MirrorByAllValues = myK

End Function


Private Function MirrorByFirstValue() As KvpA

    Dim myFirst As KvpA
    Set myFirst = KvpA.Deb
    
    Dim myOther As KvpA
    Set myOther = KvpA.Deb

    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex

        If myFirst.LacksKey(s.Items.Item(myIndex)) Then
            myFirst.Add s.Items.Item(myIndex), s.Keys.Item(myIndex)
        Else
            myOther.Add s.Keys.Item(myIndex), s.Items.Item(myIndex)
        End If

    Next

    Set MirrorByFirstValue = KvpA.Deb.Add(0&, myFirst).Add(1&, myOther)

End Function
