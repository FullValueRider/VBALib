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
'@Folder("Lib")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit

Private Const TRY_METHODS_LIST                 As String = "tostring,tojson,value,name,typename"

Private Type State

    ArrayMarkup                     As ItemMarkup
    ObjectMarkup                    As ItemMarkup
    DictionaryItemMarkup            As ItemMarkup
    ItemSeparator                   As String
    
End Type

Private s                           As State

'Private Sub Initialize_Class()
'    ResetMarkup
'End Sub

'@DefaultMember
Public Function Deb() As Stringifier
Attribute Deb.VB_UserMemId = 0
    With New Stringifier
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As Stringifier
    ResetMarkup
    Set ConstructInstance = Me
End Function


Public Function ResetMarkup() As Stringifier
    SetArrayMarkup
    SetObjectMarkup
    SetDictionaryItemMarkup
    SetItemSeparator
    Set ResetMarkup = Me
End Function


Public Function SetDictionaryItemMarkup _
( _
    Optional ByVal ipLeft As Variant = Chars.twSpace & Chars.twPlainSQuote, _
    Optional ByVal ipSeparator As Variant = Chars.twColon & Chars.twSpace, _
    Optional ByVal ipRight As Variant = Chars.twPlainSQuote _
) As Stringifier
    Set s.DictionaryItemMarkup = ItemMarkup.Deb(ipLeft, ipSeparator, ipRight)
    Set SetDictionaryItemMarkup = Me
End Function


Public Function SetArrayMarkup _
( _
    Optional ByVal ipLeft As Variant = Chars.twLSQB, _
    Optional ByVal ipSeparator As Variant = Chars.twComma, _
    Optional ByVal ipRight As Variant = Chars.twRSQB _
) As Stringifier
    Set s.ArrayMarkup = ItemMarkup.Deb(ipLeft, ipSeparator, ipRight)
    Set SetArrayMarkup = Me
End Function


Public Function SetObjectMarkup _
( _
    Optional ByVal ipLeft As Variant = Chars.twLCUB, _
    Optional ByVal ipSeparator As Variant = Chars.twComma, _
    Optional ByVal ipRight As Variant = Chars.twRCUB _
) As Stringifier
    Set s.ObjectMarkup = ItemMarkup.Deb(ipLeft, ipSeparator, ipRight)
    Set SetObjectMarkup = Me
End Function


Public Function SetItemSeparator(Optional ByVal ipSeparator As Variant = Chars.twComma) As Stringifier
    s.ItemSeparator = ipSeparator
    Set SetItemSeparator = Me
End Function


Public Function SetNoMarkup() As Stringifier
    SetArrayMarkup vbNullString, vbNullString, vbNullString
    SetObjectMarkup vbNullString, vbNullString, vbNullString
    SetItemSeparator vbNullString
    Set SetNoMarkup = Me
End Function


Public Function StringifyItem(ByRef ipItem As Variant) As String
  
    If s.ArrayMarkup Is Nothing Then SetArrayMarkup
    If s.DictionaryItemMarkup Is Nothing Then SetDictionaryItemMarkup
    If s.ObjectMarkup Is Nothing Then SetObjectMarkup
    
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
                "VBAlib.Stringifier.StringifyItem", _
                "Don't know how to stringify Type '" & VBA.TypeName(ipItem)
                
    End Select
            
End Function


Friend Function StringifyAdmin(ByRef ipAdmin As Variant) As String

    Select Case True
        
        Case VBA.IsError(ipAdmin)
             
            StringifyAdmin = _
                s.ObjectMarkup.Left _
                & Chars.twPlainDQuote _
                & ipAdmin.Source _
                & Chars.twPlainDQuote _
                & s.ObjectMarkup.Separator _
                & ipAdmin.Number _
                & s.ObjectMarkup.Separator _
                & Chars.twPlainDQuote _
                & ipAdmin.Description _
                & Chars.twPlainDQuote _
                & s.ObjectMarkup.Right
        
                
        Case VBA.IsEmpty(ipAdmin)
            
            StringifyAdmin = VBA.TypeName(ipAdmin)
                
                    
        Case VBA.IsNull(ipAdmin)
        
            StringifyAdmin = VBA.TypeName(ipAdmin)
                
                    
        Case VBA.VarType(ipAdmin) = vbDataObject
        
           StringifyAdmin s.ObjectMarkup.Left & "Object?" & s.ObjectMarkup.Right
                                                                    

        Case ipAdmin Is Nothing

            ' Debug.Print s.ObjectMarkup.Left
            ' Debug.Print s.ObjectMarkup.Right
            ' Debug.Print s.ObjectMarkup.Separator
            StringifyAdmin s.ObjectMarkup.Left & VBA.TypeName(ipAdmin) & s.ObjectMarkup.Right ' Nul is nothing

        Case Else

            Err.Raise _
                17 + vbObjectError, _
                "Stringifier.StringifyAdmin", _
                "Unexpected Item: " & TypeName(ipAdmin.Cargo)
                
                
    End Select
        
End Function


Friend Function StringifyArray(ByRef ipArray As Variant) As String
    ' The call to 'ToString' overwrites s.ArrayMarkup
    ' so we preserve a copy here
    
    If ArrayInfo.LacksItems(ipArray) Then
       StringifyArray s.ArrayMarkup.Left & s.ArrayMarkup.Right
    End If
'    Dim myMarkup As ItemMarkup
'    Set myMarkup = s.ArrayMarkup
'
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    
    
    Dim myItem As Variant
    
    For Each myItem In ipArray
        myS.Add StringifyItem(myItem)
    Next
    
   ' Set s.ArrayMarkup = myMarkup
    
    StringifyArray = s.ArrayMarkup.Left & VBA.Join(myS.ToArray, s.ArrayMarkup.Separator) & s.ArrayMarkup.Right
    
End Function


Friend Function StringifyItemByIndex(ByRef ipItemByIndex As Variant) As String
    
    If ipItemByIndex.Count = 0 Then
        StringifyItemByIndex s.ObjectMarkup.Left & s.ObjectMarkup.Right
    End If
    
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    Dim myItem As Variant
    For Each myItem In ipItemByIndex
        myS.Add StringifyItem(myItem)
    Next
    
    StringifyItemByIndex = s.ObjectMarkup.Left & VBA.Join(myS.ToArray, s.ObjectMarkup.Separator) & s.ObjectMarkup.Right
    
End Function


Friend Function StringifyItemByArray(ByRef ipItemByArray As Variant) As String
    
    If ipItemByArray.Count = 0 Then
        StringifyItemByArray = s.ObjectMarkup.Left & s.ObjectMarkup.Right
    End If
    
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    Dim myItem As Variant
    For Each myItem In ipItemByArray.ToArray
        myS.Add StringifyItem(myItem)
    Next
    
    StringifyItemByArray = s.ObjectMarkup.Left & VBA.Join(myS.ToArray, s.ObjectMarkup.Separator) & s.ObjectMarkup.Right
    
End Function


Friend Function StringifyItembyKey(ByRef ipItemByKey As Variant) As String
    
    If ipItemByKey.Count = 0 Then
        StringifyItembyKey s.ObjectMarkup.Left & s.ObjectMarkup.Right
    End If
    
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    Dim myKey As Variant
    For Each myKey In ipItemByKey.Keys
    
        Dim myT As String
    
        myT = s.DictionaryItemMarkup.Left _
            & StringifyItem(myKey) _
            & s.DictionaryItemMarkup.Right _
            & s.DictionaryItemMarkup.Separator _
            & StringifyItem(ipItemByKey.Item(myKey))
        
        myS.Add myT 's.DictionaryItemMarkup.Left & ToString(myKey) & s.DictionaryItemMarkup.Right & s.DictionaryItemMarkup.Separator & ToString(ipItemByKey.item(myKey))
        
    Next
    
    StringifyItembyKey = s.ObjectMarkup.Left & VBA.Join(myS.ToArray, s.ObjectMarkup.Separator) & s.ObjectMarkup.Right
    
End Function


'@Description("Try to get string description of a non iterable object. Default result is '{<TypeName>}'"
Friend Function StringifyItemObject(ByVal ipObject As Object) As String
'    If ipObject Is Nothing Then
'        return s.ObjectMarkup.Left & s.ObjectMarkup.Right
'    End If
    Dim myString As String
    If TryStringifyByDefaultMember(ipObject, myString) Then
    
        StringifyItemObject = s.ObjectMarkup.Left & myString & s.ObjectMarkup.Right
        
    End If
    
    'DefaultMember failed so try some other obvious methods/properties
    
    Dim myMethodNames As Variant
    myMethodNames = VBA.Split(TRY_METHODS_LIST, Chars.twComma)
        
    Dim myMethodName As Variant
    For Each myMethodName In myMethodNames
  
        Select Case True
        
            Case TryStringifyByCallByName(ipObject, myMethodName(0), VbGet, myString)
        
                StringifyItemObject = s.ObjectMarkup.Left & myString & s.ObjectMarkup.Right
                Exit Function
                
            Case TryStringifyByCallByName(ipObject, myMethodName(0), VbMethod, myString)
        
                StringifyItemObject = s.ObjectMarkup.Left & myString & s.ObjectMarkup.Right
                Exit Function
                
        End Select
    
    Next
    
    ' In desperation just return the typename wrapped in object markup
    
    StringifyItemObject = s.ObjectMarkup.Left & "Object: " & TypeName(ipObject) & s.ObjectMarkup.Right
            
End Function


Private Function TryStringifyByDefaultMember(ByVal ipObject As Object, ByRef opString As String) As Boolean

    On Error Resume Next
    opString = ipObject ' may need to be ipObject()
    TryStringifyByDefaultMember = Err.Number = 0
    On Error GoTo 0
    
End Function

'@Ignore ParameterCanBeByVal
Private Function TryStringifyByCallByName _
( _
    ByRef ipObject As Object, _
    ByVal ipMethodName As String, _
    ByVal ipMethodType As VbCallType, _
    ByRef opString As String _
) As Boolean

    On Error Resume Next
    opString = CallByName(ipObject, ipMethodName, ipMethodType)
    TryStringifyByCallByName = Err.Number = 0
    On Error GoTo 0
End Function
