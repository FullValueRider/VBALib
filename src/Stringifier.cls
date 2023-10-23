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

Private Const TRY_METHODS_LIST                 As String = "tostring,tojson,value"

'The ability to produce stringified output with Type information
'is a useful tool for debugging and testing.  The enumeration below
Public Enum e_WithTypes

    ' Do not provide any type information
    m_None
    m_First = e_WithTypes.m_None
    
    ' Prefix Array and containers with their Type
    ' e.g. Array(1,2,3,4,5) would give
    ' Variant(): [1,2,3,4,5,6]
    ' and
    ' Array(1,Array(2,3), 4,5) would give
    ' Variant():[1,Variant(): [2, 3],4,5]
     m_Outer
     
    ' Do not prefix the type of Arrays or container types.
    ' e.g. Array(1,2,3,4,5) would give
    ' [Integer: 1,Integer: 2,Integer: 3,Integer: 4,Integer: 5]
    ' and
    ' Array(1,Array(2,3), 4,5) would give
    ' [ Integer: 1,[Integer: 2, Integer: 3],Integer: 4, Integer: 5]
    m_Inner
    
    ' prefix all  variables by th eType of the variable
    ' e.g Array(1,2,3,4,5) would give
    ' Variant(): [Integer: 1, Integer: 2, Integer: 3, Integer: 4, Integer: 5]
    m_Last
    m_All = e_WithTypes.m_Last
End Enum

Private Type Properties

    ArrayMarkup                     As ElemMarkup
    DictionaryItemMarkup            As ElemMarkup
    ItemMarkup                      As ElemMarkup
    ObjectMarkup                    As ElemMarkup
    TypeMarkup                      As ElemMarkup
    
    WithTypes                       As e_WithTypes

End Type


Private p                           As Properties
    

Private Sub Class_Initialize()
	'Debug.Print "Class Initialize was called"
    Set p.ArrayMarkup = ElemMarkup.Deb
    Set p.DictionaryItemMarkup = ElemMarkup.Deb
    Set p.ItemMarkup = ElemMarkup.Deb
    Set p.ObjectMarkup = ElemMarkup.Deb
    Set p.TypeMarkup = ElemMarkup.Deb
    Me.ResetMarkup
    p.WithTypes = e_WithTypes.m_None
End Sub


'@DefaultMember
Public Function Deb(Optional ByVal ipWithTypes As e_WithTypes = e_WithTypes.m_None) As Stringifier
Attribute Deb.VB_UserMemId = 0
    With New Stringifier
        Set Deb = .ConstructInstance(ipWithTypes)
    End With
End Function

Friend Function ConstructInstance(ByVal ipWithTypes As e_WithTypes) As Stringifier
	
	Allow.InRange ipWithTypes, e_WithTypes.m_First, e_WithTypes.m_Last, "Stringifier"
    
    Set p.ArrayMarkup = ElemMarkup.Deb
    Set p.DictionaryItemMarkup = ElemMarkup.Deb
    Set p.ItemMarkup = ElemMarkup.Deb
    Set p.ObjectMarkup = ElemMarkup.Deb
    Set p.TypeMarkup = ElemMarkup.Deb
    ResetMarkup
    p.WithTypes = ipWithTypes
    Set ConstructInstance = Me
End Function

Public Function ResetMarkup() As Stringifier
    SetArrayMarkup                  'uses optional defaults
    SetDictionaryItemMarkup         ' uses optional defaults
    SetObjectMarkup                 'uses optional defaults
    SetItemMarkup                   'uses optional defaults
    SetTypeMarkup                  ' uses optional defaults
    Set ResetMarkup = Me
End Function


Public Function NoMarkup() As Stringifier
    SetArrayMarkup vbNullString, vbNullString, vbNullString
    SetObjectMarkup vbNullString, vbNullString, vbNullString
    SetDictionaryItemMarkup vbNullString, vbNullString, vbNullString
    SetItemMarkup vbNullString, vbNullString, vbNullString
    SetTypeMarkup vbNullString, vbNullString, vbNullString
    Set NoMarkup = Me
End Function


Public Function UseItemMarkupOnly() As Stringifier
    Me.SetArrayMarkup vbNullString, vbNullString, vbNullString
    Me.SetObjectMarkup vbNullString, vbNullString, vbNullString
    Me.SetDictionaryItemMarkup vbNullString, vbNullString, vbNullString
    Me.SetItemMarkup vbNullString, vbNullString, vbNullString
    Me.SetTypeMarkup vbNullString, Chars.twComma, vbNullString
    Set UseItemMarkupOnly = Me
End Function


Public Function SetArrayMarkup _
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

    Set SetArrayMarkup = Me
    
End Function

Public Function SetObjectMarkup _
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

    Set SetObjectMarkup = Me
    
End Function

Public Function SetTypeMarkup _
( _
    Optional ByRef ipLeft As String = vbnullstring, _
    Optional ByRef ipSeparator As String = Chars.twColon & chars.twSpace, _
    Optional ByRef ipRight As String = vbNullstring _
) As Stringifier

    With p.TypeMarkup
        .Left = ipLeft
        .Separator = ipSeparator
        .Right = ipRight
    End With

    Set SetTypeMarkup = Me
    
End Function

Public Function SetDictionaryItemMarkup _
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

    Set SetDictionaryItemMarkup = Me
    
End Function

Public Function SetItemMarkup _
   ( _
       Optional ByRef ipLeft As String = Chars.twSpace & Chars.twPlainDQuote, _
       Optional ByRef ipSeparator As String = Chars.twSpace, _
       Optional ByRef ipRight As String = Chars.twPlainDQuote _
   ) As Stringifier
    With p.ItemMarkup
        .Left = ipLeft
        .Separator = ipSeparator
        .Right = ipRight
    End With
        
    Set SetItemMarkup = Me
End Function

'@Description("Sets the level of type information included in the string")
Public Function WithTypes(Optional ByVal ipWithTypes As e_WithTypes = e_WithTypes.m_Inner) As Stringifier
	
    Allow.InRange ipWithTypes, e_WithTypes.m_First, e_WithTypes.m_Last, "Stringifier.WithTypes"
    
    p.WithTypes = ipWithTypes
    Set WithTypes = Me
    
End Function


Public Function StringifyItem(ByRef ipItem As Variant) As String
  
    Dim myGroupId As e_Group
    myGroupId = GroupInfo.Id(ipItem)
       
    Select Case myGroupId
     
        Case e_Group.m_String
        
            Select Case p.WithTypes
                Case m_Inner, m_All
                    StringifyItem = p.TypeMarkup.Left & VBA.TypeName(ipItem) & p.TypeMarkup.Separator & p.TypeMarkup.Right & ipItem
                Case Else
            	    StringifyItem = ipItem
            End Select
            
        Case e_Group.m_Boolean, e_Group.m_Number
        
            Select Case p.WithTypes
                Case m_Inner, m_All
                    StringifyItem = p.TypeMarkup.Left & VBA.TypeName(ipItem) & p.TypeMarkup.Separator & p.TypeMarkup.Right & ipItem
                Case Else
                    StringifyItem = ipItem
            End Select
                   
        Case e_Group.m_array:                           StringifyItem = StringifyArray(ipItem)
        Case e_Group.m_ItemByIndex:                     StringifyItem = StringifyItemByIndex(ipItem)
        Case e_Group.m_ItemByKey:                       StringifyItem = StringifyItembyKey(ipItem)
        Case e_Group.m_ItemObject:                      StringifyItem = StringifyItemObject(ipItem)
        Case e_Group.m_Admin:                           StringifyItem = StringifyAdmin(ipItem)
        Case Else
            Err.Raise 17 + vbObjectError, _
            "Stringifier.StringifyItem", _
            "Don't know how to stringify Type '" & VBA.TypeName(ipItem)
                
    End Select
            
End Function


Friend Function StringifyAdmin(ByRef ipAdmin As Variant) As String
    
    Dim myResult As String
    Select Case p.WithTypes
        Case m_Inner, m_All
            myResult = p.TypeMarkup.Left & VBA.TypeName(ipAdmin) & p.TypeMarkup.Separator & p.TypeMarkup.Right
    End Select
    
    Select Case True
        
        Case VBA.IsError(ipAdmin)
             
            myResult = _
            myResult _
            & p.ObjectMarkup.Left _
            & ipAdmin.Source _
            & p.ObjectMarkup.Separator _
            & ipAdmin.Number _
            & p.ObjectMarkup.Separator _
            & ipAdmin.Description _
            & p.ObjectMarkup.Right
        
                
        Case VBA.IsEmpty(ipAdmin)
            
            myResult = myResult & VBA.TypeName(ipAdmin)
                
                    
        Case VBA.IsNull(ipAdmin)
        
            myResult = myResult & VBA.TypeName(ipAdmin)
                
                    
        Case VBA.VarType(ipAdmin) = vbDataObject
        
            myResult = myResult & p.ObjectMarkup.Left & "Instance of " & VBA.TypeName(ipAdmin) & p.ObjectMarkup.Right
                                                                    

        Case ipAdmin Is Nothing

            myResult = myResult & p.ObjectMarkup.Left & VBA.TypeName(ipAdmin) & p.ObjectMarkup.Right ' Nul is nothing

        Case Else

            Err.Raise _
                17 + vbObjectError, _
                "Stringifier.StringifyAdmin", _
                "Unexpected Admin Item: " & TypeName(ipAdmin)
                
    End Select
    
    StringifyAdmin = myResult
    
End Function


Friend Function StringifyArray(ByRef ipArray As Variant) As String
    
    Dim myResult As String
    Select Case p.WithTypes
        Case m_Outer, m_All
            myResult = p.TypeMarkup.Left & VBA.TypeName(ipArray) & p.TypeMarkup.Separator & p.TypeMarkup.Right
    End Select
    
    If ArrayOp.LacksItems(ipArray) Then
        StringifyArray = myResult & p.ArrayMarkup.Left & "Empty " & p.ArrayMarkup.Right
        Exit Function
    End If
 
    Dim myS As SeqA: Set myS = SeqA.Deb
    
    Dim myItem As Variant
    For Each myItem In ipArray
        myS.Add StringifyItem(myItem)
    Next
    
    StringifyArray = myResult & p.ArrayMarkup.Left & VBA.Join(myS.ToArray, p.ArrayMarkup.Separator) & p.ArrayMarkup.Right
    
End Function


Friend Function StringifyItemByIndex(ByRef ipItemByIndex As Variant) As String
    
    Dim myResult As String
    Select Case p.WithTypes
        Case m_Outer, m_All
            myResult = p.TypeMarkup.Left & VBA.TypeName(ipItemByIndex) & p.TypeMarkup.Separator & p.TypeMarkup.Right
    End Select
    
    Dim myItems As IterItems: Set myItems = IterItems(ipItemByIndex)
    
    If myItems.LacksItems Then
        StringifyItemByIndex = myResult & p.ObjectMarkup.Left & "Empty " & VBA.TypeName(ipItemByIndex) & p.ObjectMarkup.Right
        Exit Function
    End If
    
    Dim myS As SeqC: Set myS = SeqC.Deb
    
    Do
        Dim myItem() As Variant: myItem = Array(myItems.CurItem(0))
        myS.Add StringifyItem(myItem(0))
    Loop While myItems.MoveNext
    
    Dim myArray As Variant: myArray = myS.ToArray
    Dim myString As String: myString = VBA.Join(myArray, p.ObjectMarkup.Separator)
    StringifyItemByIndex = myResult & p.ObjectMarkup.Left & myString & p.ObjectMarkup.Right
    
End Function


Friend Function StringifyItemByArray(ByRef ipItemByArray As Variant) As String
    
    Dim myResult As String
    Select Case p.WithTypes
        Case m_Outer, m_All
            myResult = p.TypeMarkup.Left & VBA.TypeName(ipItemByArray) & p.TypeMarkup.Separator & p.TypeMarkup.Right
    End Select
    
    If ipItemByArray.Count = 0 Then
        StringifyItemByArray = myResult & p.ObjectMarkup.Left & "Empty " & VBA.TypeName(ipItemByArray) & p.ObjectMarkup.Right
        Exit Function
    End If
    
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    Dim myItem As Variant
    For Each myItem In ipItemByArray.ToArray
        myS.Add StringifyItem(myItem)
    Next
    
    StringifyItemByArray = myResult & p.ObjectMarkup.Left & VBA.Join(myS.ToArray, p.ObjectMarkup.Separator) & p.ObjectMarkup.Right
    
End Function


Friend Function StringifyItembyKey(ByRef ipItemByKey As Variant) As String
    
    Dim myResult As String
    Select Case p.WithTypes
        Case m_Outer, m_All
            myResult = p.TypeMarkup.Left & VBA.TypeName(ipItemByKey) & p.TypeMarkup.Separator & p.TypeMarkup.Right
    End Select
    
    If ipItemByKey.Count = 0 Then
        StringifyItembyKey = myResult & p.ObjectMarkup.Left & "Empty " & VBA.TypeName(ipItemByKey) & p.ObjectMarkup.Right
        Exit Function
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
        
        myS.Add myT                              'p.DictionaryElemMarkup.Left & ToString(myKey) & p.DictionaryElemMarkup.Right & p.DictionaryElemMarkup.Separator & ToString(ipItemByKey.item(myKey))
        
    Next
    
    StringifyItembyKey = myResult & p.ObjectMarkup.Left & VBA.Join(myS.ToArray, p.ObjectMarkup.Separator) & p.ObjectMarkup.Right
    
End Function


'@Description("Try to get string description of a non iterable object. Default result is '{<TypeName>}'"
Friend Function StringifyItemObject(ByVal ipObject As Object) As String
    
    Dim myResult As String
    Select Case p.WithTypes
        Case m_Inner, m_All
            myResult = p.TypeMarkup.Left & VBA.TypeName(ipObject) & p.TypeMarkup.Separator & p.TypeMarkup.Right
    End Select
    
    Dim myString As String
    If TryStringifyByDefaultMember(ipObject, myString) Then
    
        StringifyItemObject = myResult & p.ObjectMarkup.Left & myString & p.ObjectMarkup.Right
        Exit Function
    End If
    
    'DefaultMember failed so try some other obvious methods/properties
    
    Dim myMethodNames As Variant
    myMethodNames = VBA.Split(TRY_METHODS_LIST, Chars.twComma)
        
    Dim myMethodName As Variant
    For Each myMethodName In myMethodNames
  
         If TryStringifyByCallByName(ipObject, myMethodName, myString) Then
        
                StringifyItemObject = myResult & p.ObjectMarkup.Left & myString & p.ObjectMarkup.Right
                Exit Function
                
         End If
    
    Next
    
    ' In desperation just return the typename wrapped in object markup
    
    StringifyItemObject = myResult & p.ObjectMarkup.Left & "Object: " & TypeName(ipObject) & p.ObjectMarkup.Right
            
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
    opString = CallByName(ipObject, ipMethodName, vbGet)
    If Err.Number = 0 Then
        Exit Function
    End If
    
    On Error GoTo 0
    On Error Resume Next
    opString = CallByName(ipObject, ipMethodName, vbMethod)
    If Err.Number = 0 Then
        Exit Function
    End If
    On Error GoTo 0
End Function
