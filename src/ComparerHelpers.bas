Attribute VB_Name = "ComparerHelpers"
'@Folder("Helpers")
Option Explicit


' Comparisons are managed as follows
' Numbers   VBA default comparison
' Strings   VBA default comparison
' Boolean   Error for MT, MTEQ, LT, LTEQ because we don't allow booleans to be ordinal
' Containers = Convert to string using Fmt.Text then compare the strings
' Other objects, try to generate a defaultmember value, else compare by objptr
' Admin types, compare string version of thier type name

'A comparison of a string and a number can give strange results due to numbers being coerced to strings or vice versa
' Therefore VBALib defines the comparison of a string and a number as false


Public Function GetItemAsComparerValue(ByRef ipItem As Variant) As Variant

    Dim myResult As Variant

    Select Case True
    
        Case GroupInfo.IsContainer(ipItem):                        myResult = Fmt.Text("{0}", ipItem)
            ' Admin needs to be before Object because nothing is an object
        Case GroupInfo.IsAdmin(ipItem):                            myResult = VBA.TypeName(ipItem)
        Case VBA.IsObject(ipItem):
            
            On Error Resume Next
            myResult = ipItem
            
            If Err.Number <> 0 Then
                myResult = VBA.ObjPtr(ipItem)
            End If
            
            On Error GoTo 0
            ' Booleans, Strings and Numbers
        Case Else:                                                  myResult = ipItem
            
    End Select
    
    GetItemAsComparerValue = myResult

End Function


Public Function StringNumberComparison(ByRef ipReference As Variant, ByRef ipItem As Variant) As Boolean

    If VBA.VarType(ipReference) = vbString And GroupInfo.IsNumber(ipItem) Then
        StringNumberComparison = True
    ElseIf VBA.VarType(ipItem) = vbString And GroupInfo.IsNumber(ipReference) Then
        StringNumberComparison = True
    Else
        StringNumberComparison = False
    End If
    
End Function


' In vba it is not possible to compare a non object with an object
' so to be stricter, and to give a result, rather than an error,
' the following comparisons for equality are defined)
Public Function Equals(ByRef ipLHS As Variant, ByRef ipRhs As Variant) As Boolean
    
    If VBA.IsObject(ipLHS) And VBA.IsObject(ipRhs) Then
        Equals = ipLHS Is ipRhs
        '    ElseIf (ipLHS Is Nothing) And (ipRHS Is Nothing) Then
        '        Equals = True
    ElseIf VBA.IsEmpty(ipLHS) And VBA.IsEmpty(ipRhs) Then
        Equals = True
    ElseIf GroupInfo.IsString(ipLHS) And GroupInfo.IsString(ipRhs) Then
        Equals = ipLHS = ipRhs
    ElseIf GroupInfo.IsNumber(ipLHS) And GroupInfo.IsNumber(ipRhs) Then
        Equals = ipLHS = ipRhs
    ElseIf GroupInfo.IsBoolean(ipLHS) And GroupInfo.IsBoolean(ipRhs) Then
        Equals = ipLHS = ipRhs
    Else
        Equals = False
    End If
    
End Function


