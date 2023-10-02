VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Format"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Class Fmt enables substitution of variables and control characters in strings using fields delineated by {} markers."
'@PredeclaredId
'@Exposed
'@Folder("Strings")
'12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
'=========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C


'@ModuleDescription("Class Fmt enables substitution of variables and control characters in strings using fields delineated by {} markers.")

' This module enables fields within strings which represent variables or
' control characters.
'
' Variables are indicated by {x} where x is a positive integer.
' e.g.
'       Fmt("this string {0} {1}", "Hello", 9)
'
' gives 'this string Hello 9'
'
' Layout fields represent control characters and are of the form {zz(x)}
    
'  x  = an optional positive integer to indicate the number of times the control character is repeated.
     
'   zz can be
'       nl = new line
'       nt = newline(s) followed by a tab (x applies to the newline only)
'       tb = tab
'       sq = single ordinary quote
'       dq = double ordinary quote
'       so = single smart quote - open
'       sc = single smart quote - close
'       do = double smart quote - open
'       dc = double smart quote - close
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit


Private Const MY_TYPENAME                       As String = "Format"
Private Type State
    
'    ArrayMarkupIsSet                            As Boolean
'    ObjectMarkupIsSet                           As Boolean
'    DictionaryItemMarkupIsSet                   As Boolean
'    ItemSeparatorIsSet                          As Boolean
'
    ' Cannot be a collection because we need the keys array
    FormatFields                                As KvpC
    ToString                                    As Stringifier
    
End Type

Private s                                       As State

'Private Type Properties
'        ' The next object hold the characters used to delineate different types of objects/Arrays
'    ToString                                     As Stringifier
'End Type


'Field openers - definitions
Private Const NEWLINES                          As String = "{nl"
Private Const TABS                              As String = "{tb"
Private Const NEWLINES_TAB                      As String = "{nt" ' 1 or more newlines followed by one Tab
Private Const DQUOTE                            As String = "{dq"
Private Const SQUOTE                            As String = "{sq"
Private Const SMART_SQUOTE_OPEN                 As String = "{so"
Private Const SMART_SQUOTE_CLOSE                As String = "{sc"
Private Const SMART_DQUOTE_OPEN                 As String = "{do"
Private Const SMART_DQUOTE_CLOSE                As String = "{dc"

' Field closures - definitions
Private Const FIELD_COUNT_IS_NONE               As String = "}"
Private Const FIELD_COUNT_IS_ZERO               As String = "0}"
Private Const FIELD_COUNT_IS_ONE                As String = "1}"


'@DefaultMember
Public Function Deb() As Format
Attribute Deb.VB_UserMemId = 0
    With New Format
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As Format
    Set s.ToString = Stringifier.Deb  ' we don't appear to be calling this method.
    s.ToString.ResetMarkup
    Set ConstructInstance = Me
    
End Function

Private Sub SetFormatFields()

    Set s.FormatFields = KvpC.Deb
         
    With s.FormatFields
         
        .Add ipKey:=NEWLINES, ipItem:=vbCrLf     ' {nl
        .Add ipKey:=TABS, ipItem:=vbTab          ' {tb
        .Add ipKey:=NEWLINES_TAB, ipItem:=vbCrLf ' {nt
        .Add ipKey:=DQUOTE, ipItem:=Chars.twPlainDQuote ' {dq
        .Add ipKey:=SQUOTE, ipItem:=Chars.twPlainSQuote ' {sq
        .Add ipKey:=SMART_SQUOTE_OPEN, ipItem:=Chars.twLSmartSQuote ' {so
        .Add ipKey:=SMART_SQUOTE_CLOSE, ipItem:=Chars.twRSmartSQuote ' {sc
        .Add ipKey:=SMART_DQUOTE_OPEN, ipItem:=Chars.twLSMartDQuote ' {do
        .Add ipKey:=SMART_DQUOTE_CLOSE, ipItem:=Chars.twRSmartDQuote ' {dc
             
    End With
    
End Sub
'@Description("Sends the formatted output to Debug.Print")
Public Function Dbg(ByRef ipFormatTemplate As String, ParamArray ipParamarray() As Variant) As String
Attribute Dbg.VB_Description = "Sends the formatted output to Debug.Print"
        
    Dim myFPA As Variant
    myFPA = ipParamarray
    myFPA = ArrayOp.Splat(myFPA)


    ' Dim myIndex As Long
    ' For myIndex = LBound(myFPA) To UBound(myFPA)
    '     myFpa(myIndex) =
    ' Next

    Dim myString As String
    myString = TextByVar(ipFormatTemplate, myFPA)
    
    Debug.Print myString
    Dbg = myString
  
End Function


'@Description("Returns a string representing the paramArray")
Public Function Text(ByRef ipFormatTemplate As String, ParamArray ipParamarray() As Variant) As String
Attribute Text.VB_Description = "Returns a string representing the paramArray"
    Dim myFPA As Variant
    myFPA = ipParamarray
    myFPA = ArrayOp.Splat(myFPA)
    Text = TextByVar(ipFormatTemplate, myFPA)
End Function


Private Function TextByVar(ByRef ipFormatTemplate As String, ByRef ipParamarray As Variant) As String
    
'    If Not s.ToString.Initialised Then  ' when predeclaredId
'        Set s.ToString = Stringifier.Deb
'    End If

'    If Not s.ArrayMarkupIsSet Then
'        s.ToString.SetArrayMarkup
'    End If
'
'    If Not s.DictionaryItemMarkupIsSet Then
'        s.ToString.SetDictionaryItemMarkup
'    End If
'
'
'    If Not s.ObjectMarkupIsSet Then
'        s.ToString.SetObjectMarkup
'    End If
'
'    If Not s.ItemSeparatorIsSet Then
'        s.ToString.SetItemSeparator
'    End If
    
    If VBA.LenB(ipFormatTemplate) = 0 Then
        TextByVar = ipFormatTemplate
        Exit Function
    End If
        
    If ArrayOp.LacksItems(ipParamarray) Then
        TextByVar = ipFormatTemplate
        Exit Function
    End If
    
    'ipParamarray = ArrayOp.Splat(ipParamarray)
    '
    '    Dim myToString As Stringifier
    '    Set myToString = Stringifier.Deb
    '
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    Dim myItem As Variant
    For Each myItem In ipParamarray
        myS.Add s.ToString.StringifyItem(myItem)
    Next

    If myS.LacksItems Then
        TextByVar = ipFormatTemplate
        Exit Function
    End If
    
    Dim myReturn  As String
    myReturn = ReplaceFormatFieldWithZeroCountByvbNullString(ipFormatTemplate)
    myReturn = ReplaceFormatFieldWithNoCountByFormatFieldWithCountOfOne(myReturn)
    myReturn = ReplaceFormatFieldWithCountByControlCharacters(myReturn)
    myReturn = ReplaceVariableFieldByVariableString(myReturn, myS)
    
'    s.ArrayMarkupIsSet = False
'    s.ObjectMarkupIsSet = False
'    s.DictionaryItemMarkupIsSet = False
'    s.ItemSeparatorIsSet = False
'
    TextByVar = myReturn
    
End Function


' Friend Function GetFormattedParameters(ByVal ipFormatTemplate As String, ByVal ipFPA As Variant) As String

'     Dim myReturn  As String
'     myReturn = ReplaceFormatFieldWithZeroCountByvbNullString(ipFormatTemplate)
'     myReturn = ReplaceFormatFieldWithNoCountByFormatFieldWithCountOfOne(myReturn)
'     myReturn = ReplaceFormatFieldWithCountByControlCharacters(myReturn)
'     myReturn = ReplaceVariableFieldsToStringRepresentations(myReturn, ipFPA)
        
'     GetFormattedParameters = myReturn
        
' End Function

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Methods below need to be made public for the test module to run correctly
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@Description("Replace formatting fields of the form '{WithZeroCount}' with NoString")
Friend Function ReplaceFormatFieldWithZeroCountByvbNullString(ByVal ipLayoutTemplate As String) As String
Attribute ReplaceFormatFieldWithZeroCountByvbNullString.VB_Description = "Replace formatting fields of the form '{WithZeroCount}' with NoString"

    Dim myLayoutTemplate As String
    Dim myItem As Variant

    myLayoutTemplate = ipLayoutTemplate
    Dim myKeys As Variant
    If s.FormatFields Is Nothing Then
        SetFormatFields
    End If
    myKeys = s.FormatFields.Keys
    For Each myItem In myKeys
        
        myLayoutTemplate = VBA.Replace(myLayoutTemplate, myItem & FIELD_COUNT_IS_ZERO, Chars.twNullStr)
            
    Next

    ReplaceFormatFieldWithZeroCountByvbNullString = myLayoutTemplate
        
End Function


'@Description("Replace non numbered LayoutObj.Format fields {XX} to {XX1}")
'@Ignore AssignedByValParameter
Friend Function ReplaceFormatFieldWithNoCountByFormatFieldWithCountOfOne(ByVal ipFormatTemplate As String) As String
Attribute ReplaceFormatFieldWithNoCountByFormatFieldWithCountOfOne.VB_Description = "Replace non numbered LayoutObj.Format fields {XX} to {XX1}"

    Dim myItem As Variant

    Dim myKeys As Variant
    myKeys = s.FormatFields.Keys
    For Each myItem In myKeys
        
        ipFormatTemplate = VBA.Replace(ipFormatTemplate, myItem & FIELD_COUNT_IS_NONE, myItem & FIELD_COUNT_IS_ONE)
            
    Next
        
    ReplaceFormatFieldWithNoCountByFormatFieldWithCountOfOne = ipFormatTemplate
        
End Function


'@Description("return the number in a format field")
Friend Function GetRepeatCountForFormatField(ByRef ipFormatTemplate As String, ByVal ipFormatField As String) As Long
Attribute GetRepeatCountForFormatField.VB_Description = "return the number in a format field"

    Dim myFormatFieldRepeatLocation         As Long
    Dim myRepeatCount                       As String

    myFormatFieldRepeatLocation = InStr(ipFormatTemplate, ipFormatField) + Len(ipFormatField)
        
    Do While VBA.Mid$(ipFormatTemplate, myFormatFieldRepeatLocation, 1) Like "#"
            
        myRepeatCount = myRepeatCount & VBA.Mid$(ipFormatTemplate, myFormatFieldRepeatLocation, 1)
        myFormatFieldRepeatLocation = myFormatFieldRepeatLocation + 1
            
    Loop
        
    GetRepeatCountForFormatField = CLng(myRepeatCount)
        
End Function


'@Description("Replace {XXn} layout field to fomatting characters 'XX' repeated 'n' times")
Friend Function ReplaceFormatFieldWithCountByControlCharacters(ByRef ipFormatTemplate As String) As String
Attribute ReplaceFormatFieldWithCountByControlCharacters.VB_Description = "Replace {XXn} layout field to fomatting characters 'XX' repeated 'n' times"

    Dim myKey                               As Variant
    Dim myReplace                           As String
    Dim myField                             As String
    Dim myCount                             As Long

    'Dim myItm As Variant
    Dim myKeys As Variant
    myKeys = s.FormatFields.Keys
    For Each myKey In myKeys
                
        If InStr(ipFormatTemplate, myKey) > 0 Then
            
            myCount = GetRepeatCountForFormatField(ipFormatTemplate, myKey)
            myField = myKey & CStr(myCount) & "}"
            myReplace = GetFormattingFieldReplacementString(myKey, myCount)
            ipFormatTemplate = VBA.Replace(ipFormatTemplate, myField, myReplace)
                
        End If
        
    Next
                
    ReplaceFormatFieldWithCountByControlCharacters = ipFormatTemplate
        
End Function


'@Description("Returns a string of formatting characters in line with the formatting tag")
Friend Function GetFormattingFieldReplacementString(ByVal ipFormatString As String, ByVal ipRepeatCount As Long) As String
Attribute GetFormattingFieldReplacementString.VB_Description = "Returns a string of formatting characters in line with the formatting tag"
        
    Dim myReturn As String
    ' cannot use vba.string$ here because it only returns multiples of the first character in the substring, ie only the vbcr in vbcrlf
    myReturn = Strs.Repeat(s.FormatFields.Item(ipFormatString), ipRepeatCount)
    
    If ipFormatString = NEWLINES_TAB Then
        
        myReturn = myReturn & vbTab
        
    End If
        
    GetFormattingFieldReplacementString = myReturn
        
End Function


'@Description("Replace each ocurrence of '{<number>}' with the corresponding stringified item from the parameters list")
Friend Function ReplaceVariableFieldByVariableString _
( _
ByVal ipFormatTemplate As String, _
ByVal ipSeq As SeqC _
) As String
Attribute ReplaceVariableFieldByVariableString.VB_Description = "Replace each ocurrence of '{<number>}' with the corresponding stringified item from the parameters list"

    Dim myReturn                        As String

    myReturn = ipFormatTemplate
    
    Dim myIndex As Long
    For myIndex = ipSeq.FirstIndex To ipSeq.LastIndex
    
        Dim myReplace As String
        ' Remember that {} fields start at 0
        myReplace = Chars.twLCUB & VBA.CStr(myIndex - 1) & Chars.twRCUB
        myReturn = VBA.Replace(myReturn, myReplace, ipSeq.Item(myIndex))

    Next
        
    ReplaceVariableFieldByVariableString = myReturn
        
End Function


Public Function ResetMarkup() As Format
    s.ToString.ArrayMarkup
    s.ToString.ObjectMarkup
    s.ToString.DictionaryItemMarkup
    s.ToString.ItemSeparator
    Set ResetMarkup = Me
    
End Function


Public Function DictionaryItemMarkup _
( _
    Optional ByVal ipLeft As String = Chars.twSpace & Chars.twPlainSQuote, _
    Optional ByVal ipSeparator As String = Chars.twColon & Chars.twSpace, _
    Optional ByVal ipRight As String = Chars.twPlainSQuote _
) As Format
    s.ToString.DictionaryItemMarkup ipLeft, ipSeparator, ipRight
    Set DictionaryItemMarkup = Me
End Function


Public Function ArrayMarkup _
( _
    Optional ByVal ipLeft As String = Chars.twLSQB, _
    Optional ByVal ipSeparator As String = Chars.twComma, _
    Optional ByVal ipRight As String = Chars.twRSQB _
) As Format
    s.ToString.ArrayMarkup ipLeft, ipSeparator, ipRight
    Set ArrayMarkup = Me
End Function

Public Function ObjectMarkup _
( _
    Optional ByVal ipLeft As String = Chars.twLCUB, _
    Optional ByVal ipSeparator As String = Chars.twComma, _
    Optional ByVal ipRight As String = Chars.twRCUB _
) As Format
    s.ToString.ObjectMarkup ipLeft, ipSeparator, ipRight
    Set ObjectMarkup = Me
End Function


Public Function ItemSeparator(Optional ByVal ipSeparator As String = Chars.twComma) As Format
    s.ToString.ItemSeparator ipSeparator
    Set ItemSeparator = Me
End Function


Public Function ArrayMarkupSeparatorOnly() As Format
    s.ToString.ArrayMarkup ipLeft:=vbNullString, ipRight:=vbNullString
    Set ArrayMarkupSeparatorOnly = Me
End Function
Public Function ObjectMarkupSeparatorOnly() As Format
    s.ToString.ObjectMarkup ipLeft:=vbNullString, ipRight:=vbNullString
    Set ObjectMarkupSeparatorOnly = Me
End Function


Public Function DictionaryItemMarkupSeparatorOnly() As Format
    s.ToString.DictionaryItemMarkup vbNullString, Chars.twComma, vbNullString
    Set DictionaryItemMarkupSeparatorOnly = Me
End Function
Public Function UseItemSeparatorOnly() As Format

    s.ToString.UseItemSeparatorOnly
    Set UseItemSeparatorOnly = Me
    
End Function

Public Function NoMarkup() As Format

    s.ToString.NoMarkup
    Set NoMarkup = Me
    
End Function

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
    