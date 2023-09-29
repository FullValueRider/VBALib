VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GroupInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Typing")


Public Enum e_Group
    m_First = 0
    m_Empty = m_First
    m_boolean
    m_number
    m_array
    m_string
    m_Admin
    m_List
    m_Dictionary
    m_ItemObject
    m_Container
    m_Other
    m_Last = m_Other
End Enum


Public Function Name(ByRef ipItem As Variant) As String
    
    Dim myTypename As String
    myTypename = VBA.LCase$(VBA.TypeName(ipItem))
    
    Select Case True
            ' We have multiple seq and kvp types distinguished by the last letter of their name
            ' so in these two cases strip off the last letter of the name
        Case VBA.Left$(myTypename, 3) = "kvp":      myTypename = "kvp"
        Case VBA.Left$(myTypename, 3) = "seq":      myTypename = "seq"
            ' We are only interested if it **IS** an array, not the type of the array
        Case VBA.IsArray(ipItem):                   myTypename = "array"
        
            ' Nothing is an object but is defined as an Admin Item
            ' Non container objects are ItemObjects
        Case VBA.IsObject(ipItem)
    
            If ipItem Is Nothing Then
                myTypename = "admin"
            End If
            
            If InStr("seq,collection,arraylist,stack,queue,chashd,dictionary,kvp,kvpair", myTypename) = 0 Then
                myTypename = "itemobject"
            End If
            
            
        Case VBA.IsEmpty(ipItem):                   myTypename = "admin"
        Case VBA.IsNull(ipItem):                    myTypename = "admin"
        Case VBA.IsError(ipItem):                   myTypename = "admin"

    End Select
    
    Name = myTypename
    
End Function


Public Function Id(ByVal ipItem As Variant) As e_Group

    'There are other groups possible, the most obvious being m_Container
    ' but these do not easily lend themselves to an id so we devolve such classifications to
    ' a specific method rather than an Id value
        
    Dim myGroupName As String
    myGroupName = Me.Name(ipItem)
    Select Case myGroupName
        
        Case "empty":                                                                           Id = e_Group.m_Empty
        Case "array":                                                                           Id = e_Group.m_array
        Case "boolean":                                                                         Id = e_Group.m_boolean
        Case "byte", "date", "decimal", "double", "integer", "long", "longlong", "single":      Id = e_Group.m_number
        Case "string":                                                                          Id = e_Group.m_string
        Case "seq", "collection", "arraylist", "stack", "queue":                                Id = e_Group.m_List
        Case "chashd", "dictionary", "kvp", "kvpair":                                           Id = e_Group.m_Dictionary
        Case "itemobject":                                                                      Id = e_Group.m_ItemObject
            ' Tempted to group Empty, null, errors and nothing under an unbralla of m_Admin
            ' so that there is a clear group for items/values that are not part of the
            ' other defined groups.
        Case "admin":                                                                           Id = e_Group.m_Admin
        Case Else:                                                                              Id = e_Group.m_Other
            Err.Raise 17 + vbObjectError, _
            "GroupInfo.Id", _
            Fmt.Text("Unknown type Type '{0}'", myGroupName)
                  
    End Select

End Function


Public Function IsBoolean(ByRef ipItem As Variant) As Boolean
    IsBoolean = Me.Id(ipItem) = e_Group.m_boolean
End Function


Public Function IsNotBoolean(ByRef ipItem As Variant) As Boolean
    IsNotBoolean = Not IsBoolean(ipItem)
End Function


Public Function IsNumber(ByRef ipItem As Variant) As Boolean
    IsNumber = Me.Id(ipItem) = e_Group.m_number
End Function


Public Function IsNotNumber(ByRef ipItem As Variant) As Boolean
    IsNotNumber = Not IsNumber(ipItem)
End Function


Public Function IsString(ByRef ipItem As Variant) As Boolean
    IsString = Me.Id(ipItem) = e_Group.m_string
End Function


Public Function IsNotString(ByRef ipItem As Variant) As Boolean
    IsNotString = Not IsString(ipItem)
End Function


Public Function IsList(ByRef ipItem As Variant) As Boolean
    IsList = Me.Id(ipItem) = e_Group.m_List
End Function


Public Function IsNotList(ByRef ipItem As Variant) As Boolean
    IsNotList = Not IsList(ipItem)
End Function

Public Function CanMap(ByRef ipItem As Variant) As Boolean
    CanMap = VBA.InStr("seq,kvp", GroupInfo.Name(ipItem)) > 0
End Function

Public Function CannotMap(ByRef ipItem As Variant) As Boolean
    CannotMap = Not CanMap(ipItem)
End Function


Public Function IsDictionary(ByRef ipItem As Variant) As Boolean
    IsDictionary = Me.Id(ipItem) = e_Group.m_Dictionary
End Function


Public Function IsNotDictionary(ByRef ipItem As Variant) As Boolean
    IsNotDictionary = Not IsDictionary(ipItem)
End Function


Public Function IsItemObject(ByRef ipItem As Variant) As Boolean
    IsItemObject = Me.Id(ipItem) = e_Group.m_ItemObject
End Function


Public Function IsNotItemObject(ByRef ipItem As Variant) As Boolean
    IsNotItemObject = Not IsItemObject(ipItem)
End Function


Public Function IsAdmin(ByRef ipItem As Variant) As Boolean
    IsAdmin = Me.Id(ipItem) = e_Group.m_Admin
End Function


Public Function IsNotAdmin(ByRef ipItem As Variant) As Boolean
    IsNotAdmin = Not IsAdmin(ipItem)
End Function


' A container Item is intended to hold many items which can be retrieved by a Location, Item Or Key.
' Notably a string is not a container as VBA does not treat strings as arrays of characters.
Public Function IsContainer(ByRef ipItem As Variant) As Boolean
    
    Select Case Id(ipItem)
        Case m_array, m_List, m_Dictionary
            IsContainer = True
        Case Else
            IsContainer = False
    End Select
    
End Function


Public Function IsNotContainer(ByRef ipItem As Variant) As Boolean
    IsNotContainer = Not IsContainer(ipItem)
End Function


Public Function IsItemByToArray(ByRef ipItem As Variant) As Boolean
    IsItemByToArray = VBA.InStr("Stack,Queue", VBA.TypeName(ipItem)) > 0
End Function


Public Function IsNotItemByToArray(ByRef ipItem As Variant) As Boolean
    IsNotItemByToArray = Not IsItemByToArray(ipItem)
End Function


Public Function IsIndexFrom1(ByRef ipItem As Variant) As Boolean
    IsIndexFrom1 = InStr("Seq,Col", VBA.Mid$(VBA.TypeName(ipItem), 1, 3)) > 0
End Function


Public Function ISNotIndexFrom1(ByRef ipItem As Variant) As Boolean
    ISNotIndexFrom1 = Not IsIndexFrom1(ipItem)
End Function


Public Function IsIndexFrom0(ByRef ipItem As Variant) As Boolean
    IsIndexFrom0 = InStr("arraylist", VBA.LCase$(VBA.TypeName(ipItem))) > 0
End Function


Public Function IsNotIndexFrom0(ByRef ipItem As Variant) As Boolean
    IsNotIndexFrom0 = Not IsIndexFrom0(ipItem)
End Function


Public Function IsArray(ByRef ipItem As Variant) As Boolean
    IsArray = Id(ipItem) = e_Group.m_array
End Function


Public Function IsNotArray(ByRef ipItem As Variant) As Boolean
    IsNotArray = Not GroupInfo.IsArray(ipItem)
End Function


''@Description("Return True for String len = 0, Array not allocated, or obj.Count  = 0", error on invalid Type for range")
'Public Function InvalidRangeItem(ByRef ipRange As Variant) As Boolean
'
'    ' An empty range item returns false.
'    ' An improper item raises an error
'
'    Select Case Me.Id(ipRange)
'
'        Case m_string:              InvalidRangeItem = False
'        Case m_array:               InvalidRangeItem = False
'        Case m_list, InvalidRangeItem = False
'
'        Case Else
'            Err.Raise 17 + vbObjectError, _
'                "SeqC.Deb: Invalid Range Type", _
'                " Type: " & VBA.TypeName(ipRange) & " is not a Range Item type"
'    End Select
'
'End Function

Private Function InvalidRangeItemUsingString(ByRef ipString As String) As Boolean
    InvalidRangeItemUsingString = IIf(VBA.Len(ipString) = 0, True, False)
End Function


Private Function InvalidRangeItemUsingArray(ByRef ipArray As Variant) As Boolean

    On Error Resume Next
    
    Dim myLbound As Long
    myLbound = LBound(ipArray, 1)
    
    Dim myUbound As Long
    myUbound = UBound(ipArray, 1)
    
    If Err.Number <> 0 Then
    
        InvalidRangeItemUsingArray = True
        On Error GoTo 0
        Exit Function
        
    End If
    
    If myUbound < myLbound Then
        InvalidRangeItemUsingArray = True
        On Error GoTo 0
        Exit Function
    End If
   
    On Error GoTo 0
    InvalidRangeItemUsingArray = False
   
End Function


Private Function InvalidRangeItemUsingCount(ByRef ipRangeItem As Variant) As Boolean
    InvalidRangeItemUsingCount = ipRangeItem.Count = 0
End Function





