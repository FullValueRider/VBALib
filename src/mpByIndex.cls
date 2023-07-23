VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpByIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Mappers")
' Originally I was going to make this respond to a wide variety of types
'  However, practically, it is best restricted to SeqC and Kvp type
' as The intent is to encourage users to convert othertypes to
' SeqC or kvp.  However this is not sufficiently flexible as it is likely we
' will want to deal with strings and linear arrays without converitng them to SeqC/kvp

Implements IMapper

Const MY_TYPENAME               As String = "mpByIndex"

Private Type State
    Index                       As Variant
    Mapper                      As IMapper
End Type

Private s                       As State


'Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Collection
'   Set IMapper_ExecMapper = ExecMapper(ipHostItem)
'End Function


Private Property Get IMapper_Typename() As String
    IMapper_Typename = MY_TYPENAME
End Property


Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant

    Dim myGroupName As String
    myGroupName = GroupInfo.Name(ipHostItem)
    Select Case myGroupName

        Case "string":                                                                  IMapper_ExecMapper = ExecMapperByString(VBA.CStr(ipHostItem))
        Case "seq":                                                                     IMapper_ExecMapper = ExecMapperByIndex1(ipHostItem)
        Case "arraylist":                                                               IMapper_ExecMapper = ExecMapperByIndex0(ipHostItem)
        Case "kvp", "dictionary", "chashd", "kvpair":                                   IMapper_ExecMapper = ExecMapperByDictionary(ipHostItem)
        ' a separate case is needed for collection due to the add/remove nature of assigning to an item
        Case "collection":                                                              IMapper_ExecMapper = ExecMapperByCollection(ipHostItem)
        Case "array":                                                                   IMapper_ExecMapper = ExecMapperByArray(ipHostItem)

        Case Else

            IMapper_ExecMapper = Array(ipHostItem)

    End Select

End Function


Private Function ExecMapperByString(ByRef ipHostItem As String) As Variant

    Dim myItem As String
    myItem = ipHostItem

    If s.Index < 1 Or s.Index > VBA.Len(myItem) Then
        ExecMapperByString = Array(myItem)
        Exit Function
    End If


    Mid$(myItem, s.Index, 1) = s.Mapper.ExecMapper(VBA.Mid$(ipHostItem, s.Index, 1))(0)

    ExecMapperByString = Array(myItem)

End Function


Private Function ExecMapperByIndex0(ByRef ipIndexFrom0 As Variant) As Variant

    If s.Index < 0 Or s.Index > ipIndexFrom0.Count - 1 Then
        ExecMapperByIndex0 = Array(ipIndexFrom0)
        Exit Function
    End If

    Dim myItem As Variant
    myItem = s.Mapper.ExecMapper(ipIndexFrom0.Item(s.Index - 1))

    If VBA.IsObject(myItem(0)) Then
        Set ipIndexFrom0.Item(s.Index) = myItem(0)
    Else
        ipIndexFrom0.Item(s.Index) = myItem(0)
    End If

    ExecMapperByIndex0 = Array(ipIndexFrom0)

End Function

Private Function ExecMapperByIndex1(ByRef ipIndexFrom1 As Variant) As Variant

    Dim myIndex As Long
    myIndex = VBA.CLng(s.Index)
    ' can't use firstindex/lastindex as the item may be a collection which doesn't support this abstraction
    If myIndex < 1 Or myIndex > ipIndexFrom1.Count Then
        ExecMapperByIndex1 = Array(ipIndexFrom1)
        Exit Function
    End If

    Dim myItem As Variant
    myItem = s.Mapper.ExecMapper(VBA.CVar(ipIndexFrom1.Item(myIndex)))

    If VBA.IsObject(myItem(0)) Then
        Set ipIndexFrom1.Item(myIndex) = myItem(0)
    Else
        ipIndexFrom1.Item(myIndex) = myItem(0)
    End If

    ExecMapperByIndex1 = Array(ipIndexFrom1)

End Function

Private Function ExecMapperByCollection(ByRef ipIndexColl As Variant) As Variant

    Dim myIndex As Long
    myIndex = VBA.CLng(s.Index)
    ' can't use firstindex/lastindex as the item may be a collection which doesn't support this abstraction
    If myIndex < 1 Or myIndex > ipIndexColl.Count Then
        ExecMapperByCollection = Array(ipIndexColl)
        Exit Function
    End If

    Dim myItem As Variant
    myItem = s.Mapper.ExecMapper(VBA.CVar(ipIndexColl.Item(myIndex)))

    ipIndexColl.Add myItem(0), After:=myIndex
    ipIndexColl.Remove myIndex

    ExecMapperByCollection = Array(ipIndexColl)

End Function

' Private Function ExecMapperByCollection(ByVal ipHostItem As Collection) As Variant
'     If s.Index < 1 Or s.Index > ipHostItem.Count Then
'         ExecMapperByCollection = Array(ipHostItem)
'         Exit Function
'     End If

'     ipHostItem.Add s.Mapper.ExecMapper(ipHostItem.Item(s.Index))(0), After:=s.Index
'     ipHostItem.Remove s.Index

'     ExecMapperByCollection = Array(ipHostItem)

' End Function


' Todo: check if passing by ref works, if not use variant copy method
Private Function ExecMapperByArray(ByRef ipHostItem As Variant) As Variant

    Dim myLbound As Long
    myLbound = LBound(ipHostItem)

    If GroupInfo.IsNotContainer(ipHostItem) And GroupInfo.IsNotString(ipHostItem) Then
        ExecMapperByArray = Array(ipHostItem)
        Exit Function
    End If

    If s.Index < myLbound Or s.Index > UBound(ipHostItem) Then
        ExecMapperByArray = Array(ipHostItem)
        Exit Function
    End If

    Dim myItem As Variant
    myItem = s.Mapper.ExecMapper(ipHostItem(s.Index))

    If VBA.IsObject(myItem(0)) Then
        Set ipHostItem(s.Index) = myItem(0)
    Else
        ipHostItem(s.Index) = myItem(0)
    End If

    ExecMapperByArray = Array(ipHostItem)

End Function


Private Function ExecMapperByDictionary(ByRef ipHostItem As Variant) As Variant

    If Not ipHostItem.HoldsKey(s.Index) Then
        ExecMapperByDictionary = Array(ipHostItem)
        Exit Function
    End If

    Dim myItem As Variant
    myItem = s.Mapper.ExecMapper(ipHostItem.Item(s.Index))

    If VBA.IsObject(myItem(0)) Then
        Set ipHostItem.Item(s.Index) = myItem(0)
    Else
        ipHostItem.Item(s.Index) = myItem(0)
    End If

    ExecMapperByDictionary = Array(ipHostItem)

End Function

'@DefaultMember
Public Function Deb(ByRef ipMapper As IMapper, Optional ByRef ipIndexKey As Variant = 1) As mpByIndex
Attribute Deb.VB_UserMemId = 0
    With New mpByIndex
        Set Deb = .ConstructInstance(ipMapper, ipIndexKey)
    End With
End Function


Friend Function ConstructInstance(ByRef ipMapper As IMapper, ByVal ipIndexKey As Variant) As mpByIndex

    Set s.Mapper = ipMapper

    If VBA.IsObject(ipIndexKey) Then
        Set s.Index = ipIndexKey
    Else
        s.Index = ipIndexKey
    End If

    Set ConstructInstance = Me

End Function

''
''Public Property Get TypeName() As String
''  TypeName = MY_TYPENAME
''End Property
