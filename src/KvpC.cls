VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KvpC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Kvp is a flexible Key/Value pair dictionary targeted at VB/VBA. It facilitates data input, direct content manipulation of items and keys, and indirect manipulation of data via Map and Comparer objects"
'@PredeclaredId
'@Exposed
'@Folder("Kvp")
'@ModuleDescription("Kvp is a flexible Key/Value pair dictionary targeted at VB/VBA. It facilitates data input, direct content manipulation of items and keys, and indirect manipulation of data via Map and Comparer objects")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========

 '@ModuleDescription("Dictionary using  Collections for Keys and Items)

 Option Explicit

 Const MY_TYPENAME                      As String = "KvpC"


 Private Type State
    Keys                                As SeqC
    Items                               As SeqC
 End Type

 Private s                              As State


 Public Function Deb() As KvpC
     With New KvpC
         Set Deb = .ConstructInstance
     End With
 End Function

 Friend Function ConstructInstance() As KvpC
     Set s.Keys = SeqC.Deb
     Set s.Items = SeqC.Deb
     Set ConstructInstance = Me
 End Function

''@Enumerator
'Public Function NewEnum() As IUnknown
'    Set NewEnum = s.Keys.NewEnum
'End Function

 Public Function Add(ByRef ipKey As Variant, ByRef ipItem As Variant) As KvpC
    
    If Me.HoldsItems Then
        If s.Keys.HoldsItem(ipKey) Then
            Err.Raise 17 + vbObjectError, _
                "VBALib.KvpC.Add", _
                Fmt.Text("Duplicate key: {0}", ipKey)
        End If
    End If
    
    s.Keys.Add ipKey
    s.Items.Add ipItem
    
    Set Add = Me
    
 End Function
 
Public Function AddPairs(ByRef ipKeys As SeqC, ByRef ipItems As SeqC) As KvpC
   
    Set AddPairs = Me
    If ipKeys.LacksItems Or ipItems.LacksItems Then
        Exit Function
    End If
    
    Dim myIndex As Long
    For myIndex = ipKeys.FirstIndex To ipKeys.LastIndex
    
        If myIndex > ipItems.Count Then
            Exit For
        End If
        
        If s.Keys.HoldsItem(ipKeys.Item(myIndex)) Then
            Err.Raise 17 + vbObjectError, _
               "VBALib.KvpC.Add", _
               Fmt.Text("Duplicate key: Got {0}", ipKeys.Item(myIndex))
        End If
        
        s.Keys.Add ipKeys.Item(myIndex)
        s.Items.Add ipItems.Item(myIndex)
          
    Next
   
End Function

Public Property Get Item(ByRef ipKey As Variant) As Variant
    Dim myArray As Variant
    Dim myIndex As Long
    myIndex = s.Keys.IndexOf(ipKey)
    myArray = Array(s.Items.Item(myIndex))
    If VBA.IsObject(myArray(0)) Then
        Set Item = myArray(0)
    Else
        Item = myArray(0)
    End If
End Property

Public Property Let Item(ByRef ipKey As Variant, ByVal ipItem As Variant)
   s.Items.Item(s.Keys.IndexOf(ipKey)) = ipItem
End Property

Public Property Set Item(ByRef ipKey As Variant, ByVal ipItem As Variant)
   Set s.Items.Item(s.Keys.IndexOf(ipKey)) = ipItem
End Property

'@Description("Removes Key/Value pairs that exist in the KvpC")
Public Function Remove(ParamArray ipParamarray() As Variant) As KvpC
Attribute Remove.VB_Description = "Removes Key/Value pairs that exist in the KvpC"

    Dim myItem As Variant
    For Each myItem In ipParamarray
        Dim myIndex As Long
        myIndex = s.Keys.IndexOf(myItem)
        If myIndex > 0 Then
            s.Keys.RemoveAt myIndex
            s.Items.RemoveAt myIndex
        End If
    Next

    Set Remove = Me
    
End Function

Public Function RemoveAfter(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpC

    Dim myRemoveIndex As Long
    myRemoveIndex = s.Keys.IndexOf(ipKey) + 1
    If myRemoveIndex < 1 Then
        Set RemoveAfter = Me
        Exit Function
    End If
    
    If ipRun < 1 Then
        Err.Raise 17 + vbObjectError, _
        "VBALib.KvpC.RemoveAfter", _
        Fmt.Text("Invalid run.  Expecting value of 1 or more.  Got {0}", ipRun)
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

Public Function RemoveBefore(ByRef ipKey As Variant, Optional ByVal ipRun As Long = 1) As KvpC

    Dim myRemoveIndex As Long
    myRemoveIndex = s.Keys.IndexOf(ipKey) - 1
    If myRemoveIndex < 1 Then
        Set RemoveBefore = Me
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
        myRemoveIndex = myRemoveIndex - 1
        
    Next
    
    Set RemoveBefore = Me
    
End Function

'Public Function RemoveKeys(ByRef ipContainer As Variant) As KvpC
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


Public Function RemoveAll() As KvpC
    s.Keys.RemoveAll
    s.Items.RemoveAll
    Set RemoveAll = Me
End Function

Public Function Clear() As KvpC
    s.Keys.RemoveAll
    s.Items.RemoveAll
    Set Clear = Me
End Function

Public Function Reset() As KvpC
    Set s.Keys = SeqC.Deb
    Set s.Items = SeqC.Deb
    Set Reset = Me
End Function

'===================================================================================
' There are no insert Methods for a dictionary, we can only add, remove or replace
'===================================================================================

Public Function Clone() As KvpC
    Set Clone = KvpC.Deb.AddPairs(s.Keys, s.Items)
End Function

 Public Function HoldsKey(ByRef ipKey As Variant) As Boolean
    HoldsKey = s.Keys.HoldsItem(ipKey)
End Function

Public Function Exists(ByRef ipKey As Variant) As Boolean
    Exists = s.Keys.HoldsItem(ipKey)
End Function

Public Function ContainsKey(ByRef ipKey As Variant) As Boolean
    ContainsKey = s.Keys.HoldsItem(ipKey)
End Function

Public Function LacksKey(ByRef ipKey As Variant) As Boolean
    LacksKey = s.Keys.LacksItem(ipKey)
End Function


Public Function HoldsItem(ByRef ipItem As Variant) As Boolean
    HoldsItem = s.Items.HoldsItem(ipItem)
End Function

Public Function ExistsItem(ByRef ipItem As Variant) As Boolean
    ExistsItem = s.Items.HoldsItem(ipItem)
End Function

Public Function Contains(ByRef ipItem As Variant) As Boolean
    Contains = s.Items.HoldsItem(ipItem)
End Function

Public Function LacksItem(ByRef ipItem As Variant) As Boolean
    LacksItem = s.Items.LacksItem(ipItem)
End Function


Public Function HoldsItems() As Boolean
    HoldsItems = s.Keys.Count > 0
End Function

Public Function LacksItems() As Boolean
    LacksItems = s.Keys.Count = 0
End Function

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
  
'@Description("Mutates the Kvpc in place by applying the function in IMapper to each item")
Public Function MappedIt(ByVal ipMapper As IMapper) As KvpC
Attribute MappedIt.VB_Description = "Mutates the Kvpc in place by applying the function in IMapper to each item"

   s.Items.MappedIt ipMapper
   Set MappedIt = Me

End Function

'@Description("Returns a Kvpc.Deb by applying the function in IMapper to each Item")
Public Function MapIt(ByVal ipMapper As IMapper) As KvpC
Attribute MapIt.VB_Description = "Returns a Kvpc.Deb by applying the function in IMapper to each Item"

      Dim myS As SeqC
      Set myS = s.Items.MapIt(ipMapper)
      
      Set MapIt = KvpC.Deb.AddPairs(s.Keys, myS)

  End Function

'@Description("Returns a Kvpc.Deb where the ICOmparer has been applied to the Items of the original Kvp")
Public Function FilterIt(ByVal ipComparer As IComparer) As KvpC
Attribute FilterIt.VB_Description = "Returns a Kvpc.Deb where the ICOmparer has been applied to the Items of the original Kvp"

    Dim myK As KvpC
    Set myK = KvpC.Deb
    
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

''@Description("Returns a Kvpc.Deb where the ICOmparer has been applied to the Keys of the original Kvp")
'Public Function FilterKeys(ByVal ipComparer As IComparer) As KvpC
'
'    Dim myK As KvpC
'    Set myK = KvpC.Deb
'    If s.Keys.Count = 0 Then
'        Set FilterKeys = myK
'    End If
'
'    Dim myIndex As Long
'    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex
'        If ipComparer.ExecCmp(s.Keys.Item(myIndex)) Then
'            myK.Add s.Keys.Item(myIndex), s.Items.Item(myIndex)
'        End If
'    Next
'
'    Set FilterKeys = myK
'
'End Function

Public Function ReduceIt(ByRef ipReducer As IReducer) As Variant

    Dim myItem As Variant
    myItem = Array(ipReducer.ExecReduction(s.Items))
    
    If VBA.IsObject(myItem(0)) Then
        Set ReduceIt = myItem(0)
    Else
        ReduceIt = myItem(0)
    End If
    
End Function


Public Function Count() As Long
    Count = s.Keys.Count
End Function


Public Function CountIt(ByVal ipComparer As IComparer) As Long

    If s.Keys.Count = 0 Then
        CountIt = 0
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

'Public Function CountKeys(ByVal ipComparer As IComparer) As Long
'
'    If s.Keys.Count = 0 Then
'        CountKeys = 0
'    End If
'
'    Dim myCount As Long
'    myCount = 0
'    Dim myIndex As Long
'    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex
'
'        If ipComparer.ExecCmp(s.Keys.Item(myIndex)) Then
'            myCount = myCount + 1
'        End If
'    Next
'
'    CountKeys = myCount
'
'End Function

Public Function KeyByIndex(ByVal ipIndex As Long) As Variant
    KeyByIndex = s.Keys.Item(ipIndex)
End Function


Public Function KeyOf(ByVal ipItem As Variant) As Variant

    Dim myIndex As Long
    myIndex = s.Items.IndexOf(ipItem)
    
    If VBA.IsObject(s.Keys.Item(myIndex)) Then
        Set KeyOf = s.Keys.Item(myIndex)
    Else
        KeyOf = s.Keys.Item(myIndex)
    End If
    
End Function


Public Function Items() As Variant
    Items = s.Items.ToArray
End Function


Public Function Keys() As Variant
    Keys = s.Keys.ToArray
End Function

Public Function ItemsAsSeq() As SeqC
    Set ItemsAsSeq = s.Items.Clone
End Function


Public Function KeysAsSeq() As SeqC
    Set KeysAsSeq = s.Keys.Clone
End Function
Public Function Pairs() As SeqC

    If s.Keys.LacksItems Then
        Set Pairs = SeqC.Deb
        Exit Function
    End If
    
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    Dim myIndex As Long
    For myIndex = s.Keys.FirstIndex To s.Keys.LastIndex
        myS.Add Array(s.Keys.Item(myIndex), s.Items.Item(myIndex))
    Next

    Set Pairs = myS

End Function


Public Property Get FirstKey() As Variant

    If s.Keys.Count = 0 Then
         Err.Raise 17 + vbObjectError, "Lib.KvpC.Get FirstKey", "Kvp has no items"
    End If
    
    If VBA.IsObject(s.Keys.First) Then
        Set FirstKey = s.Keys.First
    Else
        FirstKey = s.Keys.First
    End If
    
End Property


Public Property Get LastKey() As Variant

    If s.Keys.Count = 0 Then
         Err.Raise 17 + vbObjectError, "Lib.KvpC.Get LastKey", "Kvp has no items"
    End If
    
    If VBA.IsObject(s.Keys.Last) Then
        Set LastKey = s.Keys.Last
    Else
        LastKey = s.Keys.Last
    End If
    
End Property

Public Property Get First() As Variant

    If s.Keys.Count < 1 Then
         Err.Raise 17 + vbObjectError, "Lib.KvpC.Get FirstItem", "Kvp has no items"
    End If
    
    If VBA.IsObject(s.Items.First) Then
        Set First = s.Items.First
    Else
        First = s.Items.First
    End If
    
End Property


Public Property Set First(ByVal ipItem As Variant)

    If s.Keys.Count = 0 Then
        Err.Raise 17, "Lib.KvpC.Set FirstItem", "Kvp has no items"
    End If

    Set s.Items.First = ipItem

End Property

Public Property Let First(ByVal ipItem As Variant)

    If s.Keys.Count = 0 Then
        Err.Raise 17, "Lib.KvpC.Let FirstItem", "Kvp has no items"
    End If

    s.Items.First = ipItem

End Property
'
Public Property Get Last() As Variant

    If s.Keys.Count = 0 Then
        Err.Raise 17, "Lib.KvpC.Get LastItem", "Kvp has no items"
    End If
    
    If VBA.IsObject(s.Items.Last) Then
        Set Last = s.Items.Last
    Else
        Last = s.Items.Last
    End If
    
End Property

Public Property Set Last(ByVal ipItem As Variant)

    If s.Keys.Count = 0 Then
        Err.Raise 17, "Lib.KvpC.Set LastItem", "Kvp has no items"
    End If

    Set s.Items.Last = ipItem
    
End Property


Public Property Let Last(ByVal ipItem As Variant)

    If s.Keys.Count = 0 Then
        Err.Raise 17, "Lib.KvpC.Let LastItem", "Kvp has no items"
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
                "VBALib.KvpH.Mirror", _
                Fmt.Text("Mirror enum not known. Got {}", ipMirrorType)
                
    End Select

End Function


'@Description("Returns an KvpC of two KvpC.  Item(1) is an KvpC of the first found Host value vs Host Key. .Item(2) is a copy of the original KvpC with the  first Key/Value pairs removed
Private Function MirrorByAllValues() As KvpC

    Dim myK As KvpC
    Set myK = KvpC.Deb

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
''          If ipLHS.Count > 0 Then
''              Dim myLHS As IterItems = IterItems(ipLHS)
''              Do
''
''                  If ipRHS.LacksKey(myLHS.Key(0)) Then
''                      myResult.Add myLHS.Key(0), myLHS.GetItem
''                  End If
''              Loop While myLHS.MoveNext
''          End If
''          If ipRHS.Count > 0 Then
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
'  '     If ipRHS.Count = 0 Then
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
