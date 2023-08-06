VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpL"
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

Private Const MY_TYPENAME                              As String = "KvpL"
Private Const FIRST_INDEX                      As Long = 1

' currently implemented using two seql, but could be slightly
' quicker if we implemented a Key/Item node and based kvpL on a node list
Private Type State
    Keys                                        As SeqL
    Items                                       As SeqL
End Type


Private s                              As State

Private Type Properties
    EnsureUniqueKeys                        As Boolean
    Size                                    As Long ' The dimmed size of s.HashSlots
End Type


Private p                                   As Properties


Public Function Deb() As KvpL
    With New KvpL
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As KvpL
    Set s.Keys = SeqL.Deb
    Set s.Items = SeqL.Deb
    p.EnsureUniqueKeys = True
    Set ConstructInstance = Me
End Function


''@Enumerator
'Public Function NewEnum() As IUnknown
'    Set NewEnum = s.Keys.NewEnum
'End Function
Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpL
    
    Set Add = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.Add", _
        "Null is not a valid Key"
    End If
    
    If p.EnsureUniqueKeys Then
        If s.Keys.HoldsItem(ipKey) Then
            Err.Raise 17 + vbObjectError, _
            "VBALib.KvpL.Add", _
            Fmt.Text("Duplicate key: {0}", ipKey)
        End If
    End If
    
    s.Keys.Add ipKey
    s.Items.Add ipItem
    
    Set Add = Me
    
End Function


Public Function AddPairs(ByRef ipKeys As Variant, ByRef ipItems As Variant) As KvpL
   
    Dim myKeys As IterItems: Set myKeys = IterItems(ipKeys)
    Dim myItems As IterItems: Set myItems = IterItems(ipItems)
    
    If myKeys.LacksItems Or myItems.LacksItems Then
        Set AddPairs = Me
        Exit Function
    End If
    
    Do
        
        If VBA.IsNull(myKeys.CurItem(0)) Then
            Err.Raise 17 + vbObjectError, _
            "VBALib.KvpL.Add", _
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
        
        s.Keys.Add myKeys.CurItem(0)
        s.Items.Add myItems.CurItem(0)

    Loop While myKeys.MoveNext And myItems.MoveNext
   
    Set AddPairs = Me
   
End Function


Public Property Get Item(ByRef ipKey As Variant) As Variant

    Item = Null
    
    If s.Keys.LacksItem(ipKey) Then
        Exit Property
    End If
                
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpA.Getet_Item", _
        "Null is not a valid Key"
    End If
    
    Dim myArray As Variant
    Dim myIndex As Long: myIndex = s.Keys.IndexOf(ipKey)
    myArray = Array(s.Items.Item(myIndex))
    If VBA.IsObject(myArray(0)) Then
        Set Item = myArray(0)
    Else
        Item = myArray(0)
    End If
    
End Property


Public Property Let Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    If s.Keys.LacksItem(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.Let_Item", _
        Fmt.Text("No such Key: '{0}'", ipKey)
       
    End If
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.Let_Item", _
        "Null is not a valid Key"
    End If
    
    s.Items.Item(s.Keys.IndexOf(ipKey)) = ipItem
    
End Property


Public Property Set Item(ByRef ipKey As Variant, ByVal ipItem As Variant)

    If s.Keys.LacksItem(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.Set_Item", _
        Fmt.Text("No such Key: '{0}'", ipKey)
       
    End If
                
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.Let_Item", _
        "Null is not a valid Key"
    End If
    
    Dim myKey As Long: myKey = s.Keys.IndexOf(ipKey)
    Set s.Items.Item(myKey) = ipItem
   
End Property


'@Description("Removes Key/Value pairs that exist in the KvpL")
Public Function Remove(ParamArray ipParamarray() As Variant) As KvpL
Attribute Remove.VB_Description = "Removes Key/Value pairs that exist in the KvpL"

    Set Remove = Me
    
    If s.Keys.LacksItems Then
        Exit Function
    End If
    
    Dim myParamarray As Variant: myParamarray = ArrayOp.Splat(ipParamarray)
    
    If Not VBA.IsArray(myParamarray) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.Remove", _
        Fmt.Text("Expecting an array.  Got {0}", VBA.TypeName(myParamarray))
    End If
    
    Dim myItem As Variant
    For Each myItem In ipParamarray
          
        If VBA.IsNull(myItem) Then
            Err.Raise 17 + vbObjectError, _
            "VBALib.KvpL.Remove", _
            "Null is not a valid Key"
        End If
    

        Dim myIndex As Long
        myIndex = s.Keys.IndexOf(myItem)
        If myIndex > 0 Then
            s.Keys.RemoveAt myIndex
            s.Items.RemoveAt myIndex
        End If
    Next

    Set Remove = Me
    
End Function


Public Function RemoveAt(ByRef ipIndex As Long) As KvpL

    Set RemoveAt = Me

    If ipIndex < s.Keys.FirstIndex Or ipIndex > s.Keys.LastIndex Then
        Exit Function
    End If
    
    s.Keys.RemoveAt ipIndex
    s.Items.RemoveAt ipIndex
    
End Function


Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpL

    Set RemoveAfter = Me
    
    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.RemoveAfter", _
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
        "VBALib.KvpL.RemoveAfter", _
        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
    End If
    
    Dim myRemoveIndex As Long: myRemoveIndex = s.Keys.IndexOf(ipKey) + 1
    If myRemoveIndex < 1 Or myRemoveIndex > s.Keys.Count Then
        Set RemoveAfter = Me
        Exit Function
    End If
    
    Dim myCount As Long
    For myCount = 1 To ipRun
        If myRemoveIndex > s.Keys.LastIndex Then
            Exit For
        End If
        s.Keys.RemoveAt myRemoveIndex
        s.Items.RemoveAt myRemoveIndex
        
    Next
    
    Set RemoveAfter = Me
    
End Function


Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpL

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


'Public Function RemoveKeys(ByRef ipContainer As Variant) As KvpL
'
'    If GroupInfo.IsNotContainer(ipContainer) Then
'        Set RemoveKeys = Me
'        Exit Function
'    End If
'
'    ' all containers can be enummerated using for each
'    Dim myItem As Variant
'    For Each myItem In ipContainer
'        Remove myItem
'    Next
'
'    Set RemoveKeys = Me
'
'End Function


Public Function Clear() As KvpL
    s.Keys.RemoveAll
    s.Items.RemoveAll
    Set Clear = Me
End Function


Public Function RemoveAll() As KvpL
    s.Keys.RemoveAll
    s.Items.RemoveAll
    Set RemoveAll = Me
End Function


Public Function Reset() As KvpL
    Set s.Keys = SeqL.Deb
    Set s.Items = SeqL.Deb
    Set Reset = Me
End Function


'===================================================================================
' ToDo: Add insert methods
'===================================================================================

Public Function Clone() As KvpL
    Set Clone = KvpL.Deb.AddPairs(s.Keys, s.Items)
End Function


Public Function Exists(ByRef ipKey As Variant) As Boolean

    If VBA.IsNull(ipKey) Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.Exists", _
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
        "VBALib.KvpL.HoldsKey", _
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
        "VBALib.KvpL.ContainsKey", _
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
        "VBALib.KvpL.LacksKey", _
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


'@Description("Mutates the KvpL in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As KvpL
Attribute MappedIt.VB_Description = "Mutates the KvpL in place by applying the function in IMapper to each item"

    If s.Keys.LacksItems Then
        Set MappedIt = Me
        Exit Function
    End If

    s.Items.MappedIt ipMapper
    Set MappedIt = Me

End Function


'@Description("Returns a KvpL.Deb by applying the function in IMapper to each Item")
Public Function MapIt(ByVal ipMapper As IMapper) As KvpL
Attribute MapIt.VB_Description = "Returns a KvpL.Deb by applying the function in IMapper to each Item"

    If s.Keys.LacksItems Then
        Set MapIt = Me
        Exit Function
    End If

    Dim myS As SeqL
    Set myS = s.Items.MapIt(ipMapper)
    
    Set MapIt = KvpL.Deb.AddPairs(s.Keys, myS)

End Function


'@Description("Returns a KvpL.Deb where the ICOmparer has been applied to the Items of the original Kvp")
Public Function FilterIt(ByVal ipComparer As IComparer) As KvpL
Attribute FilterIt.VB_Description = "Returns a KvpL.Deb where the ICOmparer has been applied to the Items of the original Kvp"

    Dim myK As KvpL: Set myK = KvpL.Deb
    
    If s.Keys.LacksItems Then
        Set FilterIt = myK
        Exit Function
    End If
    
    If s.Keys.Count = 0 Then
        Set FilterIt = myK
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
    
    Dim myItem As Variant
    myItem = Array(ipReducer.ExecReduction(s.Items.ToArray))
    
    If VBA.IsObject(myItem(0)) Then
        Set ReduceIt = myItem(0)
    Else
        ReduceIt = myItem(0)
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
        If ComparerHelpers.Equals(ipItem, s.Items(myIndex)) Then
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


Public Function KeyByIndex(ByVal ipIndex As Long) As Variant

    KeyByIndex = Null
    
    If s.Keys.LacksItems Then
        Exit Function
    End If
    
    If ipIndex < s.Keys.FirstIndex Or ipIndex > s.Keys.LastIndex Then
        Exit Function
    End If
    
    Dim myItem As Variant: myItem = Array(s.Keys.Item(ipIndex))
    
    If VBA.IsObject(myItem(0)) Then
        Set KeyByIndex = myItem(0)
    Else
        KeyByIndex = myItem(0)
    End If
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


Public Function HasUniqueKeys() As Boolean
    HasUniqueKeys = p.EnsureUniqueKeys
End Function


Public Function ForbidSameKeys() As KvpL
    p.EnsureUniqueKeys = True
    Set ForbidSameKeys = Me
End Function


Public Function AllowSameKeys() As KvpL
    p.EnsureUniqueKeys = False
    Set AllowSameKeys = Me
End Function


Public Function Items() As Variant
    
    If s.Keys.LacksItems Then
        Items = Array()
        Exit Function
    End If
    
    Items = s.Items.ToArray
    
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

Public Function SPairs() As Variant

    Dim myKeys As SeqL: Set myKeys = Me.KeysAsSeq

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
    FirstIndex = FIRST_INDEX
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
        "VBALib.KvpL.Set_FirstItem", _
        "Kvp has no Keys"
    End If

    Set s.Items.First = ipItem

End Property


Public Property Let First(ByVal ipItem As Variant)

    If s.Keys.LacksItems Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpL.Let_FirstItem", _
        "Kvp has no Keys"
    End If

    s.Items.First = ipItem

End Property


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
        "VBALib.KvpL.Set LastItem", _
        "Kvp has no Keys"
    End If

    Set s.Items.Last = ipItem
    
End Property


Public Property Let Last(ByVal ipItem As Variant)

    If s.Keys.LacksItems Then
        Err.Raise 17, _
        "VBALib.KvpL.Set LastItem", _
        "Kvp has no Keys"
    End If

    s.Items.Last = ipItem
    
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
            "VBALib.KvpL.Mirror", _
            Fmt.Text("Mirror enum not known. Got {0}", ipMirrorType)
                
    End Select

End Function


'@Description("Returns an KvpL of two KvpL.  Item(1) is an KvpL of the first found Host value vs Host Key. .Item(2) is a copy of the original KvpL with the  first Key/Value pairs removed
Private Function MirrorByAllValues() As KvpL

    If Me.LacksItems = 0 Then
        Set Mirror = KvpL.Deb
        Exit Function
    End If
    
    Dim myK As KvpL
    Set myK = KvpL.Deb

    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex

        If myK.LacksKey(s.Keys.Item(myIndex)) Then
            myK.Add s.Items.Item(myIndex), KvpL.Deb
        End If

        myK.Item(s.Keys.Item(myIndex)).Add s.Keys.Item(myIndex), myIndex

    Next

    Set MirrorByAllValues = myK

End Function


Private Function MirrorByFirstValue() As KvpL

    Dim myFirst As KvpL
    Set myFirst = KvpL.Deb
    
    Dim myOther As KvpL
    Set myOther = KvpL.Deb

    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex

        If myFirst.LacksKey(s.Items.Item(myIndex)) Then
            myFirst.Add s.Items.Item(myIndex), s.Keys.Item(myIndex)
        Else
            myOther.Add s.Keys.Item(myIndex), s.Items.Item(myIndex)
        End If

    Next

    Set MirrorByFirstValue = KvpL.Deb.Add(0&, myFirst).Add(1&, myOther)

End Function
