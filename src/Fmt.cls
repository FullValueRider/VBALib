VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Class Fmt enables substitution of variables and control characters in strings using fields delineated by {} markers."
'@PredeclaredId
'@Exposed
'@Folder("StringyStuff")
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

Private Type State
    
    ArrayMarkupIsSet                            As Boolean
    ObjectMarkupIsSet                           As Boolean
    DictionaryItemMarkupIsSet                   As Boolean
    ItemSeparatorIsSet                          As Boolean
    
    ' The next object hold the characters used to delineate different types of objects/Arrays
    ArrayMarkup                                 As ItemMarkup
    ObjectMarkup                                As ItemMarkup
    DictionaryItemMarkup                        As ItemMarkup
    ItemSeparator                               As String
     
    ' Cannot be a collection bcause we need the keys collection
    FormatFields                                As KvpC
    ToString                                    As Stringifier
    
End Type


Private s                                       As State


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


 Public Sub Class_Initialize()
 
    s.ArrayMarkupIsSet = False
    s.ObjectMarkupIsSet = False
    s.DictionaryItemMarkupIsSet = False
    s.ItemSeparatorIsSet = False

     If Not s.FormatFields Is Nothing Then
         Exit Sub
     End If
     
     Set s.FormatFields = KvpC.Deb
         
     With s.FormatFields
         
         .Add ipKey:=NEWLINES, ipItem:=vbCrLf                              ' {nl
         .Add ipKey:=TABS, ipItem:=vbTab                                   ' {tb
         .Add ipKey:=NEWLINES_TAB, ipItem:=vbCrLf                          ' {nt
         .Add ipKey:=DQUOTE, ipItem:=Chars.twPlainDQuote                   ' {dq
         .Add ipKey:=SQUOTE, ipItem:=Chars.twPlainSQuote                   ' {sq
         .Add ipKey:=SMART_SQUOTE_OPEN, ipItem:=Chars.twLSmartSQuote       ' {so
         .Add ipKey:=SMART_SQUOTE_CLOSE, ipItem:=Chars.twRSmartSQuote      ' {sc
         .Add ipKey:=SMART_DQUOTE_OPEN, ipItem:=Chars.twLSMartDQuote       ' {do
         .Add ipKey:=SMART_DQUOTE_CLOSE, ipItem:=Chars.twRSmartDQuote      ' {dc
             
     End With
     
 End Sub


Public Function Deb() As Fmt
    With New Fmt
        Set Deb = ConstructInstance
    End With
End Function

Friend Function ConstructInstance() As Fmt
    Set ConstructInstance = Me
End Function



'@Description("Sends the formatted output to Debug.Print")
Public Function Dbg(ByRef ipFormatTemplate As String, ParamArray ipParamArray() As Variant) As String
Attribute Dbg.VB_Description = "Sends the formatted output to Debug.Print"
        
    
    Dim myString As String
    myString = TextByVar(ipFormatTemplate, VBA.CVar(ipParamArray))
    
    Debug.Print myString
    Dbg = myString
  
End Function
    

'@Description("Returns a string representing the paramArray")
Public Function Text(ByRef ipFormatTemplate As String, ParamArray ipParamArray() As Variant) As String
Attribute Text.VB_Description = "Returns a string representing the paramArray"
    Text = TextByVar(ipFormatTemplate, VBA.CVar(ipParamArray))
End Function
        
    
Private Function TextByVar(ByRef ipFormatTemplate As String, ByRef ipParamArray As Variant) As String
    
    Set s.ToString = Stringifier.Deb

    If Not s.ArrayMarkupIsSet Then
        s.ToString.SetArrayMarkup
    End If
    
    If Not s.DictionaryItemMarkupIsSet Then
        s.ToString.SetDictionaryItemMarkup
    End If
    
    
    If Not s.ObjectMarkupIsSet Then
        s.ToString.SetObjectMarkup
    End If
    
    If Not s.ItemSeparatorIsSet Then
        s.ToString.SetItemSeparator
    End If
    
    If VBA.LenB(ipFormatTemplate) = 0 Then
        TextByVar = ipFormatTemplate
        Exit Function
    End If
        
    If ArrayOp.LacksItems(ipParamArray) Then
        TextByVar = ipFormatTemplate
        Exit Function
    End If
    
    ArrayOp.Splat ipParamArray
'
'    Dim myToString As Stringifier
'    Set myToString = Stringifier.Deb
'
    Dim myS As SeqC
    Set myS = SeqC.Deb
    
    Dim myItem As Variant
    For Each myItem In ipParamArray
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
    
    s.ArrayMarkupIsSet = False
    s.ObjectMarkupIsSet = False
    s.DictionaryItemMarkupIsSet = False
    s.ItemSeparatorIsSet = False
   
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
    ByVal ipseq As SeqC _
) As String
Attribute ReplaceVariableFieldByVariableString.VB_Description = "Replace each ocurrence of '{<number>}' with the corresponding stringified item from the parameters list"

    Dim myReturn                        As String

    myReturn = ipFormatTemplate
    
    Dim myIndex As Long
    For myIndex = ipseq.FirstIndex To ipseq.LastIndex
    
        Dim myReplace As String
        ' Remember that {} fields start at 0
        myReplace = Chars.twLCUB & VBA.CStr(myIndex - 1) & Chars.twRCUB
        myReturn = VBA.Replace(myReturn, myReplace, ipseq.Item(myIndex))

    Next
        
    ReplaceVariableFieldByVariableString = myReturn
        
End Function


Public Function ResetMarkup() As Fmt
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
) As Fmt
    Set s.DictionaryItemMarkup = ItemMarkup.Deb(ipLeft, ipSeparator, ipRight)
    Set SetDictionaryItemMarkup = Me
End Function


Public Function SetArrayMarkup _
( _
    Optional ByVal ipLeft As Variant = Chars.twLSQB, _
    Optional ByVal ipSeparator As Variant = Chars.twComma, _
    Optional ByVal ipRight As Variant = Chars.twRSQB _
) As Fmt
    Set s.ArrayMarkup = ItemMarkup.Deb(ipLeft, ipSeparator, ipRight)
    s.ArrayMarkupIsSet = True
    Set SetArrayMarkup = Me
End Function


Public Function SetObjectMarkup _
( _
    Optional ByVal ipLeft As Variant = Chars.twLCUB, _
    Optional ByVal ipSeparator As Variant = Chars.twComma, _
    Optional ByVal ipRight As Variant = Chars.twRCUB _
) As Fmt
    Set s.ObjectMarkup = ItemMarkup.Deb(ipLeft, ipSeparator, ipRight)
    s.ObjectMarkupIsSet = True
    Set SetObjectMarkup = Me
End Function


Public Function SetItemSeparator(Optional ByVal ipSeparator As Variant = Chars.twComma) As Fmt
    s.ItemSeparator = ipSeparator
    s.ItemSeparatorIsSet = True
    Set SetItemSeparator = Me
End Function


Public Function SetNoMarkup() As Fmt

    SetArrayMarkup vbNullString, vbNullString, vbNullString
    SetDictionaryItemMarkup vbNullString, vbNullString, vbNullString
    SetObjectMarkup vbNullString, vbNullString, vbNullString
    SetItemSeparator vbNullString
    
    Set SetNoMarkup = Me
    
End Function
