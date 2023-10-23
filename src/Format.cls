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

    ' Cannot be a collection because we need the keys array
    FormatFields                                As KvpC
    ToString                                    As Stringifier
    ToSize                                      As Long
    PadStr                                      As String
    PadRight                                    As Boolean
    PadLeft                                     As Boolean
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


Private Const NO_PADDING                        As Boolean = False

'@DefaultMember
Public Function Deb() As Format
Attribute Deb.VB_UserMemId = 0
    With New Format
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As Format
    Set s.ToString = Stringifier.Deb
    s.ToString.ResetMarkup
    s.PadLeft = NO_PADDING
    s.PadRight = NO_PADDING
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
'
    If s.PadLeft Then
        TextByVar = Strs.PadLeft(myReturn, s.ToSize, s.PadStr)
    ElseIf s.PadRight Then
        TextByVar = Strs.PadRight(myReturn, s.ToSize, s.PadStr)
    Else
        TextByVar = myReturn
    End If
    
    s.PadLeft = NO_PADDING
    s.PadRight = NO_PADDING
    
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
'@Description("Replace formatting fields of the form '{X0}' with vbNullString")
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
    s.ToString.SetArrayMarkup
    s.ToString.SetObjectMarkup
    s.ToString.SetDictionaryItemMarkup
    s.ToString.SetItemMarkup
    s.ToString.SetTypeMarkup
    Set ResetMarkup = Me
    
End Function


Public Function SetDictionaryItemMarkup _
( _
    Optional ByVal ipLeft As String = Chars.twSpace & Chars.twPlainSQuote, _
    Optional ByVal ipSeparator As String = Chars.twColon & Chars.twSpace, _
    Optional ByVal ipRight As String = Chars.twPlainSQuote _
) As Format
    s.ToString.SetDictionaryItemMarkup ipLeft, ipSeparator, ipRight
    Set SetDictionaryItemMarkup = Me
End Function


Public Function SetArrayMarkup _
( _
    Optional ByVal ipLeft As String = Chars.twLSQB, _
    Optional ByVal ipSeparator As String = Chars.twComma, _
    Optional ByVal ipRight As String = Chars.twRSQB _
) As Format
    s.ToString.SetArrayMarkup ipLeft, ipSeparator, ipRight
    Set SetArrayMarkup = Me
End Function

Public Function SetObjectMarkup _
( _
    Optional ByVal ipLeft As String = Chars.twLCUB, _
    Optional ByVal ipSeparator As String = Chars.twComma, _
    Optional ByVal ipRight As String = Chars.twRCUB _
) As Format
    s.ToString.SetObjectMarkup ipLeft, ipSeparator, ipRight
    Set SetObjectMarkup = Me
End Function


Public Function SetItemMarkup _
   ( _
       Optional ByVal ipLeft As String = Chars.twLCUB, _
       Optional ByVal ipSeparator As String = Chars.twComma, _
       Optional ByVal ipRight As String = Chars.twRCUB _
   ) As Format
    s.ToString.SetItemMarkup ipLeft, ipSeparator, ipRight
    Set SetItemMarkup = Me
End Function

Public Function TypeMarkup _
   ( _
       Optional ByVal ipLeft As String = Chars.twLCUB, _
       Optional ByVal ipSeparator As String = Chars.twComma, _
       Optional ByVal ipRight As String = Chars.twRCUB _
   ) As Format
    s.ToString.SetItemMarkup ipLeft, ipSeparator, ipRight
    Set TypeMarkup = Me
End Function


Public Function SetArrayMarkupSeparatorOnly() As Format
    s.ToString.SetArrayMarkup ipLeft:=vbNullString, ipRight:=vbNullString
    Set SetArrayMarkupSeparatorOnly = Me
End Function

Public Function SetObjectMarkupSeparatorOnly() As Format
    s.ToString.SetObjectMarkup ipLeft:=vbNullString, ipRight:=vbNullString
    Set SetObjectMarkupSeparatorOnly = Me
End Function


Public Function SetDictionaryItemMarkupSeparatorOnly() As Format
    s.ToString.SetDictionaryItemMarkup vbNullString, Chars.twComma, vbNullString
    Set SetDictionaryItemMarkupSeparatorOnly = Me
End Function

Public Function UseItemMarkupOnly() As Format
    s.ToString.UseItemMarkupOnly
    Set UseItemMarkupOnly = Me
End Function

Public Function NoMarkup() As Format
    s.ToString.NoMarkup
    Set NoMarkup = Me
End Function


Public Function PadRight(ByRef ipToSize As Long, Optional ByRef ipPadStr As String = Chars.twSpace) As Format
    Set PadRight = Me
    If s.PadLeft Then
        Exit Function
    End If
    s.ToSize = ipToSize
    s.PadStr = ipPadStr
    s.PadRight = True
End Function

Public Function PadLeft(ByRef ipToSize As Long, Optional ByRef ipPadStr As String = Chars.twSpace) As Format
    Set PadLeft = Me
    If s.PadRight Then
        Exit Function
    End If
    s.ToSize = ipToSize
    s.PadStr = ipPadStr
    s.PadLeft = True
End Function

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
    
