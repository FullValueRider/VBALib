VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Kvp is a flexible Key/Value pair dictionary targeted at VB/VBA. It facilitates data input, direct content manipulation of items and keys, and indirect manipulation of data via Map and Comparer objects"
'@IgnoreModule
'@PredeclaredId
'@Exposed
'@Folder("Kvp")
'@ModuleDescription("Kvp is a flexible Key/Value pair dictionary targeted at VB/VBA. It facilitates data input, direct content manipulation of items and keys, and indirect manipulation of data via Map and Comparer objects")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========

 '@ModuleDescription("Dictionary based on Olaf Schmidt's cHashD with remove)

'Option Explicit
'
'Const MY_TYPENAME                       As String = "KvpH"
'Const FIRST_INDEX                       As Long = 0
'
'Private Type State
'   Host                                 As cHashD
'End Type
'
'Private s                               As State
'
'
'Public Function Deb() As KvpH
'    With New KvpH
'        Set Deb = .ConstructInstance
'    End With
'End Function
'
'Friend Function ConstructInstance() As KvpH
'    Set s.Host = New cHashD
'    Set ConstructInstance = Me
'End Function
'
'''@Enumerator
''Public Function NewEnum() As IUnknown
''    Set NewEnum = s.Host.NewEnum
''End Function
'
' Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpH
'
'     If s.Host.Exists(ipKey) Then
'         Err.Raise 17 + vbObjectError, _
'             "VBALib.KvpH.Add", _
'             Fmt.Text("Duplicate key: {0}", ipKey)
'     End If
'
'     s.Host.Add ipKey, ipItem
'
'     Set Add = Me
'
' End Function
'
'' There is no equivalent SeqH to SeqA as cHashC does not preserve the order of addition due to the index of an Item being based on its Hash
'Public Function AddPairs(ByRef ipKeys As SeqA, ByRef ipItems As SeqA) As KvpH
'
'    If ipKeys.Count = 0 Then
'        Set AddPairs = Me
'        Exit Function
'    End If
'
'    Dim myIndex As Long
'    For myIndex = ipKeys.FirstIndex To ipKeys.LastIndex
'
'        If myIndex > ipItems.Count Then
'            Exit For
'        End If
'
'        If s.Host.Exists(ipKeys.Item(myIndex)) Then
'            Err.Raise 17 + vbObjectError, _
'               "VBALib.KvpH.Add", _
'               Fmt.Text("Duplicate key: Got {0}", ipKeys.Item(myIndex))
'        End If
'
'        s.Host.Add ipKeys.Item(myIndex), ipItems.Item(myIndex)
'        's.Items.Add ipItems.Item(myIndex)
'
'    Next
'
'   Set AddPairs = Me
'
'End Function
'
'Public Property Get Item(ByRef ipKey As Variant) As Variant
'
'    Dim myItem As Variant
'    myItem = Array(s.Host.Item(ipKey))
'
'    If VBA.IsObject(myItem(0)) Then
'        Set Item = myItem(0)
'    Else
'        Item = myItem(0)
'    End If
'
'End Property
'
'Public Property Let Item(ByRef ipKey As Variant, ByVal ipItem As Variant)
'   s.Host.Item(ipKey) = ipItem
'End Property
'
'Public Property Set Item(ByRef ipKey As Variant, ByVal ipItem As Variant)
'   Set s.Host.Item(ipKey) = ipItem
'End Property
'
''@Description("Removes Key/Value pairs that exist in the KvpH")
'Public Function Remove(ParamArray ipParamarray() As Variant) As KvpH
'
'    Dim myItem As Variant
'    For Each myItem In ipParamarray
'        s.Host.Remove myItem
'    Next
'
'    Set Remove = Me
'
'End Function
'
''Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpH
''
''    Dim myRemoveIndex As Long
''    myRemoveIndex = s.Host.IndexByKey(ipKey) + 1
''    If myRemoveIndex < 1 Then
''        Set RemoveAfter = Me
''        Exit Function
''    End If
''
''    If ipRun < 1 Then
''        Err.Raise 17 + vbObjectError, _
''        "VBALib.KvpH.RemoveAfter", _
''        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
''    End If
''
''    Dim myCount As Long
''    For myCount = 1 To ipRun
''        If myRemoveIndex > s.Host.LastIndex Then
''            Exit For
''        End If
''        s.Host.RemoveAt myRemoveIndex
''        s.Items.RemoveAt myRemoveIndex
''
''    Next
''
''    Set RemoveAfter = Me
''
''End Function
'
''Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpH
''
''    Dim myRemoveIndex As Long
''    myRemoveIndex = s.Host.IndexByKey(ipKey) - 1
''    If myRemoveIndex < 1 Then
''        Set RemoveBefore = Me
''        Exit Function
''    End If
''
''    If ipRun < 1 Then
''        Err.Raise 17 + vbObjectError, _
''        "VBALib.KvpH.RemoveAfter", _
''        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
''    End If
''
''    Dim myCount As Long
''    For myCount = 1 To ipRun
''        If myRemoveIndex < s.Host.FirstIndex Then
''            Exit For
''        End If
''        s.Host.RemoveAt myRemoveIndex
''        s.Items.RemoveAt myRemoveIndex
''        myRemoveIndex = myRemoveIndex - 1
''
''    Next
''
''    Set RemoveBefore = Me
''
''End Function
'
''Public Function RemoveKeys(ByRef ipContainer As Variant) As KvpH
''
''    If GroupInfo.IsNotContainer(ipContainer) Then
''        Set RemoveKeys = Me
''        Exit Function
''    End If
''
''    ' all containers can be enummerated using for each
''    Dim myItem As Variant
''    For Each myItem In ipContainer
''        Remove myItem
''    Next
''
''    Set RemoveKeys = Me
''
''End Function
'
'
'Public Function RemoveAll() As KvpH
'    s.Host.Clear
'    Set RemoveAll = Me
'End Function
'
'Public Function Clear() As KvpH
'    s.Host.Clear
'    Set Clear = Me
'End Function
'
'Public Function Reset() As KvpH
'    Set s.Host = New cHashD
'    Set Reset = Me
'End Function
'
''===================================================================================
'' There are no insert Methods for a dictionary, we can only add, remove or replace
''===================================================================================
'
'Public Function Clone() As KvpH
'    Set Clone = KvpH.Deb.AddPairs(SeqA(s.Host.Keys), SeqA(s.Host.Items))
'End Function
'
' Public Function HoldsKey(ByRef ipKey As Variant) As Boolean
'    HoldsKey = s.Host.Exists(ipKey)
'End Function
'
'Public Function Exists(ByRef ipKey As Variant) As Boolean
'    Exists = s.Host.Exists(ipKey)
'End Function
'
'Public Function ContainsKey(ByRef ipKey As Variant) As Boolean
'    ContainsKey = s.Host.Exists(ipKey)
'End Function
'
'Public Function LacksKey(ByRef ipKey As Variant) As Boolean
'    LacksKey = Not s.Host.Exists(ipKey)
'End Function
'
'
''Public Function Exists(ByRef ipItem As Variant) As Boolean
''    Exists = s.Host.Exists(ipItem)
''End Function
'
'Public Function HoldsItem(ByRef ipItem As Variant) As Boolean
'
'    Dim myItem As Variant
'    For Each myItem In s.Host.Items
'        If ComparerHelpers.Equals(myItem, ipItem) Then
'            HoldsItem = True
'            Exit Function
'        End If
'    Next
'
'    HoldsItem = False
'
'End Function
'
'Public Function Contains(ByRef ipItem As Variant) As Boolean
'    Contains = HoldsItem(ipItem)
'End Function
'
'Public Function LacksItem(ByRef ipItem As Variant) As Boolean
'    LacksItem = Not HoldsItem(ipItem)
'End Function
'
'
''Public Function Existss() As Boolean
''    Existss = s.Host.Count > 0
''End Function
'Public Function HoldsItems() As Boolean
'    HoldsItems = s.Host.Count > 0
'End Function
'Public Function LacksItems() As Boolean
'    LacksItems = s.Host.Count = 0
'End Function
'
'Public Property Get TypeName() As String
'    TypeName = MY_TYPENAME
'End Property
'
''@Description("Mutates the KvpH in place by applying the function in IMapper to each item")
'Public Function MappedIt(ByVal ipMapper As IMapper) As KvpH
'
'    Dim myKeys As Variant
'    myKeys = s.Host.Keys
'
'    Dim myKey As Variant
'    For Each myKey In myKeys
'        Dim myArray As Variant
'        myArray = ipMapper.ExecMapper(s.Host.Item(myKey))
'        If VBA.IsObject(myArray(0)) Then
'            Set s.Host.Item(myKey) = myArray(0)
'        Else
'            s.Host.Item(myKey) = myArray(0)
'        End If
'
'    Next
'
'    Set MappedIt = Me
'
'End Function
'
''@Description("Returns a new KvpH.Deb by applying the function in IMapper to each Item")
'Public Function MapIt(ByVal ipMapper As IMapper) As KvpH
'
'      Dim myK As KvpH
'      Set myK = KvpH.Deb
'
'      Dim myKeys As Variant
'      myKeys = s.Host.Keys
'
'      Dim myKey As Variant
'      For Each myKey In myKeys
'        myK.Add myKey, ipMapper.ExecMapper(s.Host.Item(myKey))(0)
'      Next
'
'      Set MapIt = myK
'
'  End Function
'
''@Description("Returns a KvpH.Deb where the ICOmparer has been applied to the Items of the original Kvp")
'Public Function FilterIt(ByVal ipComparer As IComparer) As KvpH
'
'    Dim myK As KvpH
'    Set myK = KvpH.Deb
'
'    If s.Host.Count = 0 Then
'        Set FilterIt = myK
'    End If
'
'    Dim myKeys As Variant
'    myKeys = s.Host.Keys
'
'    Dim myKey As Variant
'    For Each myKey In myKeys
'
'        If ipComparer.ExecCmp(s.Host.Item(myKey)) Then
'            myK.Add myKey, s.Host.Item(myKey)
'        End If
'
'    Next
'
'    Set FilterIt = myK
'
'End Function
'
'''@Description("Returns a KvpH.Deb where the ICOmparer has been applied to the Keys of the original Kvp")
''Public Function FilterKeys(ByVal ipComparer As IComparer) As KvpH
''
''    Dim myK As KvpH
''    Set myK = KvpH.Deb
''    If s.Host.Count = 0 Then
''        Set FilterKeys = myK
''    End If
''
''    Dim myIndex As Long
''    For myIndex = s.Host.FirstIndex To s.Host.LastIndex
''        If ipComparer.ExecCmp(s.Host.Item(myIndex)) Then
''            myK.Add s.Host.Item(myIndex), s.Items.Item(myIndex)
''        End If
''    Next
''
''    Set FilterKeys = myK
''
''End Function
'
'Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant
'
'    Dim myItem As Variant
'    myItem = Array(ipReducer.ExecReduction(s.Host.Items))
'
'    If VBA.IsObject(myItem(0)) Then
'        Set ReduceIt = myItem(0)
'    Else
'        ReduceIt = myItem(0)
'    End If
'
'End Function
'
'
'Public Function Count() As Long
'    Count = s.Host.Count
'End Function
'
'
'Public Function CountIt(ByVal ipComparer As IComparer) As Long
'
'    If s.Host.Count = 0 Then
'        CountIt = 0
'    End If
'
'    Dim myCount As Long
'    myCount = 0
'    Dim myKeys As Variant
'
'    myKeys = s.Host.Keys
'    Dim myKey As Variant
'
'    For Each myKey In myKeys
'
'        If ipComparer.ExecCmp(s.Host.Item(myKey)) Then
'            myCount = myCount + 1
'        End If
'    Next
'
'    CountIt = myCount
'
'End Function
'
''Public Function CountKeys(ByVal ipComparer As IComparer) As Long
''
''    If s.Host.Count = 0 Then
''        CountKeys = 0
''    End If
''
''    Dim myCount As Long
''    myCount = 0
''    Dim myIndex As Long
''    For myIndex = s.Host.FirstIndex To s.Host.LastIndex
''
''        If ipComparer.ExecCmp(s.Host.Item(myIndex)) Then
''            myCount = myCount + 1
''        End If
''    Next
''
''    CountKeys = myCount
''
''End Function
'
'Public Function KeyByIndex(ByVal ipIndex As Long) As Variant
'    KeyByIndex = s.Host.KeyByIndex(ipIndex)
'End Function
'
'
'Public Function KeyOf(ByVal ipItem As Variant) As Variant
'
'    Dim myKeys As Variant
'    myKeys = s.Host.Keys
'
'    Dim myKey As Variant
'    For Each myKey In myKeys
'
'        If ComparerHelpers.Equals(s.Host.Item(myKey), ipItem) Then
'
'            If VBA.IsObject(myKey) Then
'                Set KeyOf = myKey
'            Else
'                KeyOf = myKey
'            End If
'        End If
'    Next
'
'End Function
'
'
'Public Function ItemsAsSeq() As SeqH
'    Set Items = SeqH(s.Host.Items)
'End Function
'
'
'Public Function Items() As Variant
'    Items = s.Items
'End Function
'
'
'Public Function Keys() As Variant
'    Keys = s.Keys
'End Function
'
'Public Function KeysAsSeq() As SeqH
'    Items = SeqH(s.Keys.Clone)
'End Function
'
'
'
'
'
'Public Property Get FirstKey() As Variant
'
'    If s.Host.Count = 0 Then
'         Err.Raise 17 + vbObjectError, "Lib.KvpH.Get FirstKey", "Kvp has no items"
'    End If
'
'    Dim myArray As Variant
'    myArray = Array(s.Host.KeyByIndex(First_Index))
'
'    If VBA.IsObject(myArray(0)) Then
'        Set FirstKey = myArray(0)
'    Else
'        FirstKey = myArray(0)
'    End If
'
'End Property
'
'
'Public Property Get LastKey() As Variant
'
'    If s.Host.Count = 0 Then
'         Err.Raise 17 + vbObjectError, "Lib.KvpH.Get LastKey", "Kvp has no items"
'    End If
'
'    Dim myArray As Variant
'    myArray = Array(s.Host.KeyByIndex(s.Host.Count))
'
'    If VBA.IsObject(myArray(0)) Then
'        Set LastKey = myArray(0)
'    Else
'        LastKey = myArray(0)
'    End If
'
'End Property
'
'Public Property Get First() As Variant
'
'    If s.Host.Count < 1 Then
'         Err.Raise 17 + vbObjectError, "Lib.KvpH.Get FirstItem", "Kvp has no items"
'    End If
'
'    Dim myArray As Variant
'    myArray = Array(s.Host.ItemByIndex(First_Index))
'
'    If VBA.IsObject(myArray(0)) Then
'        Set First = myArray(0)
'    Else
'        First = myArray(0)
'    End If
'
'End Property
'
'
'Public Property Set First(ByVal ipItem As Variant)
'
'    If s.Host.Count = 0 Then
'        Err.Raise 17, "Lib.KvpH.Set FirstItem", "Kvp has no items"
'    End If
'
'    Set s.Host.ItemByIndex(First_Index) = ipItem
'
'End Property
'
'Public Property Let First(ByVal ipItem As Variant)
'
'    If s.Host.Count = 0 Then
'        Err.Raise 17, "Lib.KvpH.Let FirstItem", "Kvp has no items"
'    End If
'
'    s.Host.ItemByIndex(First_Index) = ipItem
'
'End Property
''
'Public Property Get Last() As Variant
'
'    If s.Host.Count = 0 Then
'        Err.Raise 17, "Lib.KvpH.Get LastItem", "Kvp has no items"
'    End If
'
'    Dim myArray As Variant
'    myArray = Array(s.Host.ItemByIndex(s.Host.Count))
'
'    If VBA.IsObject(myArray(0)) Then
'        Set Last = myArray(0)
'    Else
'        Last = myArray(0)
'    End If
'
'End Property
'
'Public Property Set Last(ByVal ipItem As Variant)
'
'    If s.Host.Count = 0 Then
'        Err.Raise 17, "Lib.KvpH.Set LastItem", "Kvp has no items"
'    End If
'
'    Set s.Host.ItemByIndex(s.Host.Count) = ipItem
'
'End Property
'
'
'Public Property Let Last(ByVal ipItem As Variant)
'
'    If s.Host.Count = 0 Then
'        Err.Raise 17, "Lib.KvpH.Let LastItem", "Kvp has no items"
'    End If
'
'    s.Host.ItemByIndex(s.Host.Count) = ipItem
'
'End Property
'
'
'
'
'''@Description( returns a Kvp with values as keys.  Each Value is linked to a Kvp of Key vs Index for the host value
''Public Function Mirror(Optional ByVal ipMirrorType As e_MirrorType = e_MirrorType.m_ByAllValues) As KvpH
''
''     If Me.LacksItems = 0 Then
''        Set Mirror = KvpH.Deb
''        Exit Function
''    End If
''
''    Select Case ipMirrorType
''
''        Case m_ByAllValues:             Set Mirror = MirrorByAllValues
''        Case m_ByFirstValue:            Set Mirror = MirrorByFirstValue
''        Case Else
''            Err.Raise 17 + vbObjectError, _
''                "VBALib.KvpH.Mirror", _
''                Fmt.Text("Mirror enum not known. Got {}", ipMirrorType)
''
''    End Select
''
''End Function
'
'
''@Description("Returns an KvpH of two KvpH.  Item(1) is an KvpH of the first found Host value vs kvp of Host Key vs Index. .Item(2) is a copy of the original KvpH with the  first Key/Value pairs removed
'Private Function MirrorByAllValues() As KvpH
'
'    Dim myK As KvpH
'    Set myK = KvpH.Deb
'
'    Dim myKeys As Variant
'    myKeys = s.Host.Keys
'
'    Dim myKey As Variant
'    For Each myKey In myKeys
'
'        Dim myArray As Variant
'        myArray = Array(s.Host.Item(myKey))
'        If myK.LacksKey(myArray(0)) Then
'            myK.Add myArray(0), KvpH.Deb
'        End If
'
'        myK.Item(myArray(0)).Add s.Host.IndexByKey(myKey), myKey
'
'    Next
'
'    Set MirrorByAllValues = myK
'
'End Function
'
'Private Function MirrorByFirstValue() As KvpH
'
'    Dim myFirst As KvpH
'    Set myFirst = KvpH.Deb
'
'    Dim myOther As KvpH
'    Set myOther = KvpH.Deb
'
'    Dim myKeys As Variant
'    myKeys = s.Host.Keys
'
'    Dim myKey As Variant
'    For Each myKey In myKeys
'
'        Dim myArray As Variant
'        myArray = Array(s.Host.Item(myKey))
'        If myFirst.LacksKey(myArray(0)) Then
'            myFirst.Add myArray(0), myKey
'        Else
'            myOther.Add myKey, myArray(0)
'        End If
'
'    Next
'
'    Set MirrorByFirstValue = KvpH.Deb.Add(1&, myFirst).Add(2&, myOther)
'
'End Function
