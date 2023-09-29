VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stringifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule ImplicitUnboundDefaultMemberAccess
'@PredeclaredId
'@Exposed
'@Folder("VBALib")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit

Private Const TRY_METHODS_LIST                 As String = "tostring,tojson,value,name,typename"


Private Type ItemMarkup
    Left As String
    Separator As String
    Right As String
End Type


Private Type Properties

    ArrayMarkup                     As ItemMarkup
    ObjectMarkup                    As ItemMarkup
    DictionaryItemMarkup            As ItemMarkup
    ItemSeparator                   As String
   ' Initialised                     As Boolean ' by default this value will be false
End Type


Private p                           As Properties

'@DefaultMember
Public Function Deb() As Stringifier
Attribute Deb.VB_UserMemId = 0
    With New Stringifier
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As Stringifier
    ResetMarkup
   ' p.Initialised = True
    Set ConstructInstance = Me
End Function


Public Function ResetMarkup() As Stringifier
    ArrayMarkup                 'uses optional defaults
    ObjectMarkup                'uses optional defaults
    DictionaryItemMarkup        'uses optional defaults
    ItemSeparator               'uses optional defaults
    Set ResetMarkup = Me
End Function

'Public Property Get Initialised() As Boolean
'    Initialised = p.Initialised
'End Property
'
'Public Property Let Initialised(ByVal ipInitialised As Boolean)
'    p.Initialised = ipInitialised
'End Property
'Public Property Get ArrayMarkup() As ItemMarkup
'    Set ArrayMarkup = p.ArrayMarkup
'End Property
'Public Property Set ArrayMarkup(ByVal ipArrayMarkup As ItemMarkup)
'    Set p.ArrayMarkup = ipArrayMarkup
'End Property
'
'Public Property Get ObjectMarkup() As ItemMarkup
'    Set ObjectMarkup = p.ObjectMarkup
'End Property
'
'Public Property Set ObjectMarkup(ByVal ipObjectMarkup As ItemMarkup)
'    Set p.ObjectMarkup = ipObjectMarkup
'End Property
'
'Public Property Get DictionaryItemMarkup() As ItemMarkup
'    Set DictionaryItemMarkup = p.DictionaryItemMarkup
'End Property
'
'Public Property Set DictionaryItemMarkup(ByRef ipDictionaryItemMarkup As ItemMarkup)
'    Set p.DictionaryItemMarkup = ipDictionaryItemMarkup
'End Property
'
'Public Property Get ItemSeparator() As String
'    ItemSeparator = p.ItemSeparator
'End Property
'
'Public Property Let ItemMarkup(ByRef ipItemSeparator As String)
'    p.ItemSeparator = ipItemSeparator
'End Property


'Public Sub UseDefaultArrayMarkup()
'    With p.ArrayItemMarkup
'        .Left = Chars.twLSQB
'        .Separator = Chars.twComma
'        .Right = Chars.twRSQB
'    End With
'End Sub
'
'Public Sub UseDefaultSetObjectMarkup()
'    With p.ArrayItemMarkup
'        .Left = Chars.twLCUB
'        .Separator = Chars.twComma
'        .Right = Chars.twRCUB
'    End With
'End Sub
'
'Public Sub UseDefaultDictionaryItemMarkup()
'    With p.DictionaryItemMarkup
'        .Left = Chars.twSpace & Chars.twPlainSQuote
'        .Separator = Chars.twColon & Chars.twSpace
'        .Right = Chars.twPlainSQuote
'    End With
'End Sub
'
'
'Public Sub UseDefaultItemSeparator(Optional ByVal ipSeparator As Variant = Chars.twComma)
'    p.ItemSeparator = ipSeparator
'End Sub

Public Function NoMarkup() As Stringifier
    Me.ArrayMarkup vbNullString, vbNullString, vbNullString
    Me.ObjectMarkup vbNullString, vbNullString, vbNullString
    Me.DictionaryItemMarkup vbNullString, vbNullString, vbNullString
    Me.ItemSeparator vbNullString
    Set NoMarkup = Me
End Function


Public Function UseItemSeparatorOnly() As Stringifier
    Me.ArrayMarkup vbNullString, vbNullString, vbNullString
    Me.ObjectMarkup vbNullString, vbNullString, vbNullString
    Me.DictionaryItemMarkup vbNullString, vbNullString, vbNullString
    Me.ItemSeparator
    Set UseItemSeparatorOnly = Me
End Function


Public Function ArrayMarkup _
( _
    Optional ByRef ipLeft As String = Chars.twLSQB, _
    Optional ByRef ipSeparator As String = Chars.twComma, _
    Optional ByRef ipRight As String = Chars.twRSQB _
) As Stringifier

    With p.ArrayMarkup
        .Left = ipLeft
        .Separator = ipSeparator
        .Right = ipRight
    End With

    Set ArrayMarkup = Me
    
End Function

Public Function ObjectMarkup _
( _
    Optional ByRef ipLeft As String = Chars.twLCUB, _
    Optional ByRef ipSeparator As String = Chars.twComma, _
    Optional ByRef ipRight As String = Chars.twRCUB _
) As Stringifier

    With p.ObjectMarkup
        .Left = ipLeft
        .Separator = ipSeparator
        .Right = ipRight
    End With

    Set ObjectMarkup = Me
    
End Function

Public Function DictionaryItemMarkup _
( _
    Optional ByRef ipLeft As String = Chars.twSpace & Chars.twPlainDQuote, _
    Optional ByRef ipSeparator As String = Chars.twSpace, _
    Optional ByRef ipRight As String = Chars.twPlainDQuote _
) As Stringifier

    With p.DictionaryItemMarkup
        .Left = ipLeft
        .Separator = ipSeparator
        .Right = ipRight
    End With

    Set DictionaryItemMarkup = Me
    
End Function

Public Function ItemSeparator(Optional ByRef ipSeparator As String = Chars.twComma) As Stringifier
    p.ItemSeparator = ipSeparator
    Set ItemSeparator = Me
End Function

Public Function StringifyItem(ByRef ipItem As Variant) As String
  
'    If p.ArrayMarkup Is Nothing Then
'        UseDefaulArrayMarkup
'    End If
'
'    If p.ObjectMarkup Is Nothing Then
'        UseDefaultObjectMarkup
'    End If
'
'    If p.DictionaryItemMarkup Is Nothing Then
'        UseDefaultDictionaryItemMarkup
'    End If
'
    Dim myGroupId As e_Group
    myGroupId = GroupInfo.Id(ipItem)
       
    Select Case myGroupId
    
        Case e_Group.m_string:                          StringifyItem = ipItem
        Case e_Group.m_boolean, e_Group.m_number:       StringifyItem = VBA.CStr(ipItem)
        Case e_Group.m_array:                           StringifyItem = StringifyArray(ipItem)
        Case e_Group.m_List:                            StringifyItem = StringifyItemByIndex(ipItem)
        Case e_Group.m_Dictionary:                      StringifyItem = StringifyItembyKey(ipItem)
        Case e_Group.m_ItemObject:                      StringifyItem = StringifyItemObject(ipItem)
        Case e_Group.m_Admin:                           StringifyItem = StringifyAdmin(ipItem)
        Case Else
            Err.Raise 17 + vbObjectError, _
            "Stringifier.StringifyItem", _
            "Don't know how to stringify Type '" & VBA.TypeName(ipItem)
                
    End Select
            
End Function


Friend Function StringifyAdmin(ByRef ipAdmin As Variant) As String

    Select Case True
        
        Case VBA.IsError(ipAdmin)
             
            StringifyAdmin = _
            p.ObjectMarkup.Left _
            & Chars.twPlainDQuote _
            & ipAdmin.Source _
            & Chars.twPlainDQuote _
            & p.ObjectMarkup.Separator _
            & ipAdmin.Number _
            & p.ObjectMarkup.Separator _
            & Chars.twPlainDQuote _
            & ipAdmin.Description _
            & Chars.twPlainDQuote _
            & p.ObjectMarkup.Right
        
                
        Case VBA.IsEmpty(ipAdmin)
            
            StringifyAdmin = VBA.TypeName(ipAdmin)
                
                    
        Case VBA.IsNull(ipAdmin)
        
            StringifyAdmin = VBA.TypeName(ipAdmin)
                
                    
        Case VBA.VarType(ipAdmin) = vbDataObject
        
            StringifyAdmin p.ObjectMarkup.Left & "Object?" & p.ObjectMarkup.Right
                                                                    

        Case ipAdmin Is Nothing

            ' Debug.Print p.ObjectMarkup.Left
            ' Debug.Print p.ObjectMarkup.Right
            ' Debug.Print p.ObjectMarkup.Separator
            StringifyAdmin p.ObjectMarkup.Left & VBA.TypeName(ipAdmin) & p.ObjectMarkup.Right ' Nul is nothing

        Case Else

            Err.Raise _
                17 + vbObjectError, _
                "Stringifier.StringifyAdmin", _
                "Unexpected Admin Item: " & TypeName(ipAdmin)
                
                
    End Select
        
End Function


Friend Function StringifyArray(ByRef ipArray As Variant) As String
    ' The call to 'ToString' overwrites p.ArrayMarkup
    ' so we preserve a copy here
    
    If ArrayOp.LacksItems(ipArray) Then
        StringifyArray p.ArrayMarkup.Left & "Empty " & VBA.TypeName(ipArray) & p.ArrayMarkup.Right
    End If
    '    Dim myMarkup As ItemMarkup
    '    Set myMarkup = p.ArrayMarkup
    '
    Dim myS As SeqA: Set myS = SeqA.Deb
    
    Dim myItem As Variant
    For Each myItem In ipArray
        myS.Add StringifyItem(myItem)
    Next
    
    ' Set p.ArrayMarkup = myMarkup
    
    StringifyArray = p.ArrayMarkup.Left & VBA.Join(myS.ToArray, p.ArrayMarkup.Separator) & p.ArrayMarkup.Right
    
End Function


Friend Function StringifyItemByIndex(ByRef ipItemByIndex As Variant) As String
    
    Dim myItems As IterItems: Set myItems = IterItems(ipItemByIndex)
    
    If myItems.LacksItems Then
        StringifyItemByIndex = p.ObjectMarkup.Left & "Empty " & VBA.TypeName(ipItemByIndex) & p.ObjectMarkup.Right
        Exit Function
    End If
    
    Dim myS As SeqC: Set myS = SeqC.Deb
    
    Do
        Dim myItem() As Variant: myItem = Array(myItems.CurItem(0))
        myS.Add StringifyItem(myItem(0))
    Loop While myItems.MoveNext
    
    Dim myArray As Variant: myArray = myS.ToArray
    Dim myString As String: myString = VBA.Join(myArray, p.ObjectMarkup.Separator)
    StringifyItemByIndex = p.ObjectMarkup.Left & myString & p.ObjectMarkup.Right
    
End Function


Friend Function StringifyItemByArray(ByRef ipItemByArray As Variant) As String
    
    If ipItemByArray.Count = 0 Then
        StringifyItemByArray = p.ObjectMarkup.Left & "Empty " & VBA.TypeName(ipItemByArray) & p.ObjectMarkup.Right
    End If
    
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    Dim myItem As Variant
    For Each myItem In ipItemByArray.ToArray
        myS.Add StringifyItem(myItem)
    Next
    
    StringifyItemByArray = p.ObjectMarkup.Left & VBA.Join(myS.ToArray, p.ObjectMarkup.Separator) & p.ObjectMarkup.Right
    
End Function


Friend Function StringifyItembyKey(ByRef ipItemByKey As Variant) As String
    
    If ipItemByKey.Count = 0 Then
        StringifyItembyKey p.ObjectMarkup.Left & "Empty " & VBA.TypeName(ipItemByKey) & p.ObjectMarkup.Right
    End If
    
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    Dim myKey As Variant
    Dim myKeys As Variant
    myKeys = ipItemByKey.Keys
    For Each myKey In myKeys
    
        Dim myT As String
    
        myT = p.DictionaryItemMarkup.Left
        myT = myT & StringifyItem(myKey)
        myT = myT & p.DictionaryItemMarkup.Right
        myT = myT & p.DictionaryItemMarkup.Separator
        myT = myT & StringifyItem(ipItemByKey.Item(myKey))
        
        myS.Add myT                              'p.DictionaryItemMarkup.Left & ToString(myKey) & p.DictionaryItemMarkup.Right & p.DictionaryItemMarkup.Separator & ToString(ipItemByKey.item(myKey))
        
    Next
    
    StringifyItembyKey = p.ObjectMarkup.Left & VBA.Join(myS.ToArray, p.ObjectMarkup.Separator) & p.ObjectMarkup.Right
    
End Function


'@Description("Try to get string description of a non iterable object. Default result is '{<TypeName>}'"
Friend Function StringifyItemObject(ByVal ipObject As Object) As String
    '    If ipObject Is Nothing Then
    '        return p.ObjectMarkup.Left & p.ObjectMarkup.Right
    '    End If
    Dim myString As String
    If TryStringifyByDefaultMember(ipObject, myString) Then
    
        StringifyItemObject = p.ObjectMarkup.Left & myString & p.ObjectMarkup.Right
        
    End If
    
    'DefaultMember failed so try some other obvious methods/properties
    
    Dim myMethodNames As Variant
    myMethodNames = VBA.Split(TRY_METHODS_LIST, Chars.twComma)
        
    Dim myMethodName As Variant
    For Each myMethodName In myMethodNames
  
         If TryStringifyByCallByName(ipObject, myMethodName, myString) Then
        
                StringifyItemObject = p.ObjectMarkup.Left & myString & p.ObjectMarkup.Right
                Exit Function
                
         End If
    
    Next
    
    ' In desperation just return the typename wrapped in object markup
    
    StringifyItemObject = p.ObjectMarkup.Left & "Object: " & TypeName(ipObject) & p.ObjectMarkup.Right
            
End Function


Private Function TryStringifyByDefaultMember(ByVal ipObject As Object, ByRef opString As String) As Boolean

    On Error Resume Next
    opString = (ipObject)                          ' the enclosing () forces evaluation by
    TryStringifyByDefaultMember = Err.Number = 0
    On Error GoTo 0
    
End Function


'@Ignore ParameterCanBeByVal
Private Function TryStringifyByCallByName _
( _
ByRef ipObject As Object, _
ByVal ipMethodName As String, _
ByRef opString As String _
) As Boolean
    TryStringifyByCallByName = True
    
    On Error Resume Next
    opString = CallByName(ipObject, ipMethodName, VbGet)
    If Err.Number = 0 Then
        Exit Function
    End If
    
    On Error GoTo 0
    On Error Resume Next
    opString = CallByName(ipObject, ipMethodName, VbMethod)
    If Err.Number = 0 Then
        Exit Function
    End If
    On Error GoTo 0
End Function
