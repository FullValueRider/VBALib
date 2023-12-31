VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''@IgnoreModule VariableNotUsed, AssignmentNotUsed
'@Folder("Lib")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@PredeclaredId
'@Exposed
Option Explicit

    
Private Type InternalConstants
    
    Location                                    As String
    
End Type

Private i                                       As InternalConstants

Public Sub Class_Initialize()
    
    If VBA.Len(i.Location) = 0 Then
        
        i.Location = "VBALib.Types."
        
    End If
    
    Guard InvalidNew, Me.IsNotSameInstance(Me, Types), i.Location & "Class_Initialize"
    
End Sub
    


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'  Class Intellisense
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Property Get Group() As TypesGrouped
    Set Group = TypesGrouped
End Property

Public Property Get Name() As TypesNames
    Set Name = TypesNames
End Property

Public Property Get Iterable() As TypesIterables
    Set Iterable = TypesIterables
End Property
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'  Class Intellisense forwarding
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@Description("Returns true if the list of parameters are all arrays))
Public Function IsArray(ParamArray ipArgs() As Variant) As Boolean
    IsArray = pvIsArrayByCVar(CVar(ipArgs))
End Function

'@Description("Returns True if none of the list of parameters is an array)")
Public Function IsNotArray(ParamArray ipArgs() As Variant) As Boolean
Attribute IsNotArray.VB_Description = "Returns True if none of the list of parameters is an array)"
    IsNotArray = Not pvIsArrayByCVar(CVar(ipArgs))
End Function

Public Function pvIsArrayByCVar(ByVal ipArgs As Variant) As Boolean
    
    pvIsArrayByCVar = False
    If Arrays.HasNoItems(ipArgs) Then Exit Function
    
            
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Arrays.HasNoItems(myArg) Then Exit Function
        
    Next
    
    pvIsArrayByCVar = True
    
End Function

'@Description("Returns True if an object has a 'Count' method")
Public Function HasCountMethod(ByVal ipItem As Variant) As Boolean
Attribute HasCountMethod.VB_Description = "Returns True if an object has a 'Count' method"
    
    '@Ignore VariableNotUsed
    Dim myCount As Long
    Dim myError As Long
    On Error Resume Next
    '@Ignore AssignmentNotUsed
    myCount = ipItem.Count
    myError = Err.Number
    Err.Clear
    On Error GoTo 0

    HasCountMethod = myError = 0
    
End Function

'@Description("Returns True if an object does not have a count method")
Public Function LacksCountMethod(ByVal ipItem As Variant) As Boolean
Attribute LacksCountMethod.VB_Description = "Returns True if an object does not have a count method"
    LacksCountMethod = Not HasCountMethod(ipItem)
End Function

Public Function HasCloneMethod(ByVal ipTestItem As Variant) As Boolean
    
    Dim myError As Long
    '@Ignore VariableNotUsed
    Dim myObj As Object
    On Error Resume Next
    '@Ignore AssignmentNotUsed
    Set myObj = ipTestItem.Clone
    myError = Err.Number
    Err.Clear
    On Error GoTo 0
        
    HasCloneMethod = myError = 0
    
End Function

Public Function LacksCloneMethod(ByVal ipTestItem As Variant) As Boolean
    LacksCloneMethod = Not HasCloneMethod(ipTestItem)
End Function



'@Description("Returns True if an object has an Item Method")
Public Function HasItemMethod(ByVal ipItem As Variant) As Result
Attribute HasItemMethod.VB_Description = "Returns True if an object has an Item Method"
    
    ' A reasonable assumption is that an object with an item method also has a count method
    
    Dim myresult As Result
    Set myresult = Result.Deb
    Set HasItemMethod = myresult
    If Not VBA.IsObject(ipItem) Then
    
        myresult.Reason = InvalidType
        Exit Function
        
    End If
    
    If LacksCountMethod(ipItem) Then
    
       ' myresult.reson GuardClause
        Exit Function
    
    End If
    ' Can't get an item if it dowsn't exist
    If ipItem.Count = 0 Then Exit Function
        
        
    
    ' Using Array is the simplest way of recieving an object without an isobject test.
    ' The assignment below is to get the error number only and tests if the item method exists
    '@Ignore VariableNotUsed
    Dim myItem As Variant
    '@Ignore AssignmentNotUsed
    myItem = Array(ipItem.Item(ipItem.Count - 1))
    Dim myErrNo As Long
    myErrNo = Err.Number
    
    On Error GoTo 0

    HasItemMethod = myErrNo = 0
    
End Function

'@Description("Returns True if the object does not have an Item method)")
Public Function LacksItemMethod(ByVal ipItem As Variant) As Boolean
Attribute LacksItemMethod.VB_Description = "Returns True if the object does not have an Item method)"
    LacksItemMethod = Not HasItemMethod(ipItem)
End Function

'@Description("Returns True if any parameter is empty or is ismissing")
Public Function IsMissing(ParamArray ipArgs() As Variant) As Boolean
Attribute IsMissing.VB_Description = "Returns True if any parameter is empty or is ismissing"
    IsMissing = pvIsMissingByCVar(CVar(ipArgs))
End Function

'@Description"Returns True if all parameters are not empty and are not missing")
Public Function IsNotMissing(ParamArray ipArgs() As Variant) As Boolean
    IsNotMissing = Not pvIsMissingByCVar(ipArgs)
End Function

Public Function pvIsMissingByCVar(ByVal ipArgs As Variant) As Boolean

    pvIsMissingByCVar = False
    If VBA.IsEmpty(ipArgs) Then Exit Function
    '@Ignore IsMissingOnInappropriateArgument
    If VBA.IsMissing(ipArgs) Then Exit Function
        
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        '@Ignore IsMissingWithNonArgumentParameter
        If Not VBA.IsMissing(myArg) Then Exit Function
        
    Next
    
    pvIsMissingByCVar = True
    
End Function


Public Function IsEmpty(ParamArray ipArgs() As Variant) As Boolean
    IsEmpty = pvIsEmptyByCVar(CVar(ipArgs))
End Function

Public Function IsNotEmpty(ParamArray ipArgs() As Variant) As Boolean
    IsNotEmpty = Not pvIsEmptyByCVar(CVar(ipArgs))
End Function


Public Function pvIsEmptyByCVar(ByRef ipArgs As Variant) As Boolean

    pvIsEmptyByCVar = False
    If VBA.IsEmpty(ipArgs) Then Exit Function
    '@Ignore IsMissingOnInappropriateArgument
    If VBA.IsMissing(ipArgs) Then Exit Function
            
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Me.Group.IsIterable(myArg) Then
        
        If Me.Iterable.HasItems(myArg) Then Exit Function
            
        Else
        
        If Not VBA.IsEmpty(myArg) Then Exit Function
        
        End If
    
    Next
    
    pvIsEmptyByCVar = True
    
End Function
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Tests using groups of types
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Function IsSameType(ByVal ipRequiredType As Variant, ParamArray ipArgs() As Variant) As Boolean
    IsSameType = pvIsSameTypeByCVAr(ipRequiredType, ipArgs)
End Function

Public Function IsNotSameType(ByVal ipRequiredType As Variant, ParamArray ipArgs() As Variant) As Boolean
    IsNotSameType = Not pvIsSameTypeByCVAr(ipRequiredType, ipArgs)
End Function

Public Function pvIsSameTypeByCVAr(ByVal ipRequiredType As Variant, ByVal ipArgs As Variant) As Boolean

    pvIsSameTypeByCVAr = False
    If Arrays.HasNoItems(ipArgs) Then Exit Function
            
    Dim myRequiredType As String
    myRequiredType = TypeName(ipRequiredType)
    Dim myItem As Variant
    For Each myItem In ipArgs
        ' cannot use vartype as this cannot distinguish object types
        
        If InStr(myRequiredType, TypeName(myItem)) = 0 Then Exit Function
        
    Next
    
    pvIsSameTypeByCVAr = True
    
End Function


Public Function IsObject(ParamArray ipArgs() As Variant) As Boolean
    IsObject = pvIsObjectByCVar(CVar(ipArgs))
End Function

Public Function IsNotObject(ParamArray ipArgs() As Variant) As Boolean
    IsNotObject = Not pvIsObjectByCVar(CVar(ipArgs))
End Function

Public Function pvIsObjectByCVar(ByVal ipArgs As Variant) As Boolean
    
    pvIsObjectByCVar = False
    If VBA.IsEmpty(ipArgs) Then Exit Function
    '@Ignore IsMissingOnInappropriateArgument
    If VBA.IsMissing(ipArgs) Then Exit Function
            
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Not VBA.IsObject(myArg) Then Exit Function
        
    Next
    
    pvIsObjectByCVar = True
    
End Function


Public Function IsObjectArray(ParamArray ipArgs() As Variant) As Boolean
    IsObjectArray = pvIsObjectArrayByCVar(CVar(ipArgs))
End Function

Public Function IsNotObjectArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotObjectArray = Not pvIsObjectArrayByCVar(CVar(ipArgs))
End Function


Public Function pvIsObjectArrayByCVar(ByVal ipArgs As Variant) As Boolean
    
    pvIsObjectArrayByCVar = False
    If VBA.IsEmpty(ipArgs) Then Exit Function
    '@Ignore IsMissingOnInappropriateArgument
    If VBA.IsMissing(ipArgs) Then Exit Function
            
    Dim myArg As Variant
    For Each myArg In ipArgs
        
        If InStr(Me.Name.AsArray.OfObject, TypeName(myArg)) = 0 Then Exit Function
            
    Next
    
    pvIsObjectArrayByCVar = True
    
End Function


'Public Function HasDefaultItemMethod(ByVal ipobject As Variant) As Boolean
'
'    Dim myTLI As TLI.TLIApplication
'    Set myTLI = New TLI.TLIApplication
'
'    Dim myInterface As TLI.TypeInfo
'    Set myInterface = myTLI.InterfaceInfoFromObject(ipobject)
'
''    Dim myMemberInfo as
''    HasDefaultItemMethod = False
''    If Types.IsNotObject(ipobject) Then
'End Function
    


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'  Tests for specific types.
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@Description("Returns true if the types of all the values in the parameter list are found in the Types string")
Public Function IsTypeByName(ByVal ipTypesasStringOrIterable As Variant, ParamArray ipArgs() As Variant) As Boolean
Attribute IsTypeByName.VB_Description = "Returns true if the types of all the values in the parameter list are found in the Types string"
    IsTypeByName = pvIsTypeByNameByCVar(ipTypesasStringOrIterable, ipArgs)
End Function

Public Function IsNotTypeByName(ByVal ipTypesasStringOrIterable As Variant, ParamArray ipArgs() As Variant) As Boolean
    IsNotTypeByName = Not IsTypeByName(ipTypesasStringOrIterable, ipArgs)
End Function

Private Function IsTypeByNameByCVar(ByVal ipTypes As String, ParamArray ipArgs() As Variant) As Boolean
    IsTypeByNameByCVar = Not pvIsTypeByNameByCVar(ipTypes, ipArgs)
End Function


Public Function pvIsTypeByNameByCVar(ByVal ipTypesasStringOrIterable As Variant, ByVal ipArgs As Variant) As Boolean
    
    pvIsTypeByNameByCVar = False
    If VBA.IsArray(ipArgs) Then
        
        If Arrays.HasNoItems(ipArgs) Then Exit Function
        
    End If
    
        
    Dim myIterable As Variant
    If VBA.VarType(ipTypesasStringOrIterable) = vbString Then
            
        Set myIterable = Strings.ToSubStrLyst(ipTypesasStringOrIterable, Char.comma)
        
        
    ElseIf Me.Group.IsIterableItemsByEnum(ipTypesasStringOrIterable) Then
        
        Set myIterable = ipTypesasStringOrIterable
        
    Else
        
        Exit Function
            
    End If
    

    Dim myItem As Variant
    For Each myItem In ipArgs
        
        If myIterable.LacksItem(TypeName(myItem)) Then Exit Function
    
    Next
    
    pvIsTypeByNameByCVar = True
    
End Function


'@Description("Returns true is all the values in the Parameter array have the VarType ipVarType, Cannot differentiate between objects")
Public Function IsVarType(ByVal ipVarType As VbVarType, ParamArray ipArgs() As Variant) As Boolean
Attribute IsVarType.VB_Description = "Returns true is all the values in the Parameter array have the VarType ipVarType, Cannot differentiate between objects"
    IsVarType = pvIsVarTypeByCVar(ipVarType, ipArgs)
End Function

Public Function IsNotVarType(ByVal ipVarType As VbVarType, ParamArray ipArgs() As Variant) As Boolean
    IsNotVarType = Not pvIsVarTypeByCVar(ipVarType, ipArgs)
End Function

Public Function pvIsVarTypeByCVar(ByVal ipVarType As VbVarType, ByVal ipArgs As Variant) As Boolean

    'ipArgs is derived from CVar(ParamArray) and consequently will always be an array
    pvIsVarTypeByCVar = False
    If VBA.IsEmpty(ipArgs) Then Exit Function
    '@Ignore IsMissingOnInappropriateArgument
    If VBA.IsMissing(ipArgs) Then Exit Function
            
    Dim myItem As Variant
    For Each myItem In ipArgs
    
        If Not VBA.VarType(myItem) = ipVarType Then Exit Function
        
    Next
    
    pvIsVarTypeByCVar = True

End Function

Public Function IsBoolean(ParamArray ipArgs() As Variant) As Boolean
    IsBoolean = pvIsTypeByNameByCVar(Me.Name.OfBoolean, ipArgs)
End Function

Public Function IsNotBoolean(ParamArray ipArgs() As Variant) As Boolean
    IsNotBoolean = Not pvIsTypeByNameByCVar(Me.Name.OfBoolean, ipArgs)
End Function

Public Function IsBooleanArray(ParamArray ipArgs() As Variant) As Boolean
    IsBooleanArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfBoolean, ipArgs)
End Function

Public Function IsNotBooleanArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotBooleanArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfBoolean, ipArgs)
End Function



Public Function IsByte(ParamArray ipArgs() As Variant) As Boolean
    IsByte = pvIsTypeByNameByCVar(Me.Name.OfByte, ipArgs)
End Function

Public Function IsNotByte(ParamArray ipArgs() As Variant) As Boolean
    IsNotByte = Not pvIsTypeByNameByCVar(Me.Name.OfByte, ipArgs)
End Function

Public Function IsByteArray(ParamArray ipArgs() As Variant) As Boolean
    IsByteArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfByte, ipArgs)
End Function

Public Function IsNotByteArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotByteArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfByte, ipArgs)
End Function



Public Function IsCurrency(ParamArray ipArgs() As Variant) As Boolean
    IsCurrency = pvIsTypeByNameByCVar(Me.Name.OfCurrency, ipArgs)
End Function

Public Function IsNotCurrency(ParamArray ipArgs() As Variant) As Boolean
    IsNotCurrency = Not pvIsTypeByNameByCVar(Me.Name.OfCurrency, ipArgs)
End Function

Public Function IsCurrencyArray(ParamArray ipArgs() As Variant) As Boolean
    IsCurrencyArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfCurrency, ipArgs)
End Function

Public Function IsNotCurrencyArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotCurrencyArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfCurrency, ipArgs)
End Function



'Public Function IsDataObject(ParamArray ipArgs() As Variant) As Boolean
'    IsDataObject = pvIsTypeByNameByCVar(Types.Name.OfCurrency, ipArgs)
'End Function
'
'Public Function IsNotDataObject(ParamArray ipArgs() As Variant) As Boolean
'    IsNotDataObject = Not IsVarTypeByCVar(CVar(ipArgs), vbDataObject)
'End Function
'
'Public Function IsDataObjectArray(ParamArray ipArgs() As Variant) As Boolean
'    IsDataObjectArray = IsVarTypeByCVar(CVar(ipArgs), vbDataObject + vbArray)
'End Function
'
'Public Function IsNotDataObjectArray(ParamArray ipArgs() As Variant) As Boolean
'    IsNotDataObjectArray = Not IsVarTypeByCVar(CVar(ipArgs), vbDataObject + vbArray)
'End Function



Public Function IsDate(ParamArray ipArgs() As Variant) As Boolean
    IsDate = pvIsTypeByNameByCVar(Me.Name.OfDate, ipArgs)
End Function

Public Function IsNotDate(ParamArray ipArgs() As Variant) As Boolean
    IsNotDate = Not pvIsTypeByNameByCVar(Me.Name.OfDate, ipArgs)
End Function

Public Function IsDateArray(ParamArray ipArgs() As Variant) As Boolean
    IsDateArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfDate, ipArgs)
End Function

Public Function IsNotDateArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotDateArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfDate, ipArgs)
End Function


Public Function IsDecimal(ParamArray ipArgs() As Variant) As Boolean
    IsDecimal = pvIsTypeByNameByCVar(Me.Name.OfDecimal, ipArgs)
End Function

Public Function IsNotDecimal(ParamArray ipArgs() As Variant) As Boolean
    IsNotDecimal = Not pvIsTypeByNameByCVar(Me.Name.OfDecimal, ipArgs)
End Function

Public Function IsDecimalArray(ParamArray ipArgs() As Variant) As Boolean
    IsDecimalArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfDecimal, ipArgs)
End Function

Public Function IsNotDecimalArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotDecimalArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfDecimal, ipArgs)
End Function


Public Function IsDouble(ParamArray ipArgs() As Variant) As Boolean
    IsDouble = pvIsTypeByNameByCVar(Me.Name.OfDouble, ipArgs)
End Function

Public Function IsNotDouble(ParamArray ipArgs() As Variant) As Boolean
    IsNotDouble = Not pvIsTypeByNameByCVar(Me.Name.OfDouble, ipArgs)
End Function

Public Function IsDoubleArray(ParamArray ipArgs() As Variant) As Boolean
    IsDoubleArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfDouble, ipArgs)
End Function

Public Function IsNotDoubleArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotDoubleArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfDouble, ipArgs)
End Function





Public Function IsError(ParamArray ipArgs() As Variant) As Boolean
    IsError = pvIsTypeByNameByCVar(Me.Name.OfError, ipArgs)
End Function

Public Function IsNotError(ParamArray ipArgs() As Variant) As Boolean
    IsNotError = Not pvIsTypeByNameByCVar(Me.Name.OfError, ipArgs)
End Function

Public Function IsErrorArray(ParamArray ipArgs() As Variant) As Boolean
    IsErrorArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfError, ipArgs)
End Function

Public Function IsNotErrorArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotErrorArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfError, ipArgs)
End Function


Public Function IsNothing(ParamArray ipArgs() As Variant) As Boolean
    IsNothing = pvIsTypeByNameByCVar(Me.Name.OfNothing, ipArgs)
End Function

Public Function IsNotNothing(ParamArray ipArgs() As Variant) As Boolean
    IsNotNothing = Not pvIsTypeByNameByCVar(Me.Name.OfNothing, ipArgs)
End Function

Public Function IsNothingArray(ParamArray ipArgs() As Variant) As Boolean
    IsNothingArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfNothing, ipArgs)
End Function

Public Function IsNotNothingArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotNothingArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfNothing, ipArgs)
End Function


Public Function IsNull(ParamArray ipArgs() As Variant) As Boolean
    IsNull = pvIsTypeByNameByCVar(Me.Name.OfNull, ipArgs)
End Function

Public Function IsNotNull(ParamArray ipArgs() As Variant) As Boolean
    IsNotNull = Not pvIsTypeByNameByCVar(Me.Name.OfNull, ipArgs)
End Function

Public Function IsNullArray(ParamArray ipArgs() As Variant) As Boolean
    IsNullArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfNull, ipArgs)
End Function

Public Function IsNotNullArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotNullArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfNull, ipArgs)
End Function


Public Function IsInteger(ParamArray ipArgs() As Variant) As Boolean
    IsInteger = pvIsTypeByNameByCVar(Me.Name.OfInteger, ipArgs)
End Function

Public Function IsNotInteger(ParamArray ipArgs() As Variant) As Boolean
    IsNotInteger = Not pvIsTypeByNameByCVar(Me.Name.OfInteger, ipArgs)
End Function

Public Function IsIntegerArray(ParamArray ipArgs() As Variant) As Boolean
    IsIntegerArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfInteger, ipArgs)
End Function

Public Function IsNotIntegerArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotIntegerArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfInteger, ipArgs)
End Function


Public Function IsLong(ParamArray ipArgs() As Variant) As Boolean
    IsLong = pvIsTypeByNameByCVar(Me.Name.OfLong, ipArgs)
End Function

Public Function IsNotLong(ParamArray ipArgs() As Variant) As Boolean
    IsNotLong = Not pvIsTypeByNameByCVar(Me.Name.OfLong, ipArgs)
End Function

Public Function IsLongArray(ParamArray ipArgs() As Variant) As Boolean
    IsLongArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfLong, ipArgs)
End Function

Public Function IsNotLongArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotLongArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfLong, ipArgs)
End Function


Public Function IsLongLong(ParamArray ipArgs() As Variant) As Boolean
    IsLongLong = pvIsTypeByNameByCVar(Me.Name.OfLongLong, ipArgs)
End Function

Public Function IsNotLongLong(ParamArray ipArgs() As Variant) As Boolean
    IsNotLongLong = Not pvIsTypeByNameByCVar(Me.Name.OfLongLong, ipArgs)
End Function

Public Function IsLongLongArray(ParamArray ipArgs() As Variant) As Boolean
    IsLongLongArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfLongLong, ipArgs)
End Function

Public Function IsNotLongLongArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotLongLongArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfLongLong, ipArgs)
End Function


'@Description(" True if the object type is object.  Which is different to asking if a variable is an object")
Public Function IsObjectObject(ParamArray ipArgs() As Variant) As Boolean
Attribute IsObjectObject.VB_Description = " True if the object type is object.  Which is different to asking if a variable is an object"
    IsObjectObject = pvIsTypeByNameByCVar(Me.Name.OfObject, ipArgs)
End Function

Public Function IsNotObjectObject(ParamArray ipArgs() As Variant) As Boolean
    IsNotObjectObject = Not pvIsTypeByNameByCVar(Me.Name.OfObject, ipArgs)
End Function

Public Function IsObjectObjectArray(ParamArray ipArgs() As Variant) As Boolean
    IsObjectObjectArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfObject, ipArgs)
End Function

Public Function IsNotArrayOfObjectObjects(ParamArray ipArgs() As Variant) As Boolean
    IsNotArrayOfObjectObjects = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfObject, ipArgs)
End Function


Public Function IsSingle(ParamArray ipArgs() As Variant) As Boolean
    IsSingle = pvIsTypeByNameByCVar(Me.Name.OfSingle, ipArgs)
End Function

Public Function IsNotSingle(ParamArray ipArgs() As Variant) As Boolean
    IsNotSingle = Not pvIsTypeByNameByCVar(Me.Name.OfSingle, ipArgs)
End Function

Public Function IsSingleArray(ParamArray ipArgs() As Variant) As Boolean
    IsSingleArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfSingle, ipArgs)
End Function

Public Function IsNotSingleArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotSingleArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfSingle, ipArgs)
End Function


Public Function IsString(ParamArray ipArgs() As Variant) As Boolean
    IsString = pvIsTypeByNameByCVar(Me.Name.OfString, ipArgs)
End Function

Public Function IsNotString(ParamArray ipArgs() As Variant) As Boolean
    IsNotString = Not pvIsTypeByNameByCVar(Me.Name.OfString, ipArgs)
End Function

Public Function IsStringArray(ParamArray ipArgs() As Variant) As Boolean
    IsStringArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfString, ipArgs)
End Function

Public Function IsNotStringArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotStringArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfString, ipArgs)
End Function


Public Function IsUnknown(ParamArray ipArgs() As Variant) As Boolean
    IsUnknown = pvIsTypeByNameByCVar(Me.Name.OfUnknown, ipArgs)
End Function

Public Function IsNotUnknown(ParamArray ipArgs() As Variant) As Boolean
    IsNotUnknown = Not pvIsTypeByNameByCVar(Me.Name.OfUnknown, ipArgs)
End Function

Public Function IsUnknownArray(ParamArray ipArgs() As Variant) As Boolean
    IsUnknownArray = pvIsTypeByNameByCVar(Me.Name.AsArray.OfUnknown, ipArgs)
End Function

Public Function IsNotUnknownArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotUnknownArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.OfUnknown, ipArgs)
End Function


' Its not possible to test for a variable being a variant in vba
Public Function IsVariantArray(ParamArray ipArgs() As Variant) As Boolean
    IsVariantArray = pvIsTypeByNameByCVar(Me.Name.AsArray.ofVariant, ipArgs)
End Function

Public Function IsNotVariantArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotVariantArray = Not pvIsTypeByNameByCVar(Me.Name.AsArray.ofVariant, ipArgs)
End Function


Public Function IsKVPair(ParamArray ipArgs() As Variant) As Boolean
    IsKVPair = pvIsTypeByNameByCVar(Me.Name.OfKVPair, ipArgs)
End Function

Public Function IsNotKVPair(ParamArray ipArgs() As Variant) As Boolean
    IsNotKVPair = Not pvIsTypeByNameByCVar(Me.Name.OfKVPair, ipArgs)
End Function
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'  Miscellaneous
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Function ToVarTypePrimitive(ByVal ipType As VbVarType, ByVal ipValue As Variant) As Variant

    Guard InvalidType, Me.Group.IsNotPrimitive, i.Location & "ToVarType", Array("Primitive", TypeName(ipValue), "ipValue")
        
    Dim myresult As Variant
    'With Types.Name
    
        Select Case ipType
        
            Case vbBoolean
            
                myresult = VBA.CBool(ipValue)
                
            Case vbByte
            
                myresult = VBA.CByte(ipValue)
            
            Case vbCurrency
            
                myresult = VBA.CCur(ipValue)
            
            Case vbDate
            
                myresult = VBA.CDate(ipValue)
                
            Case vbDouble
            
                myresult = VBA.CDbl(ipValue)
        
            Case vbDecimal
            
                myresult = VBA.CDec(ipValue)
            
            Case vbInteger
            
                myresult = VBA.CInt(ipValue)
            
            Case vbLong
            
                myresult = VBA.CLng(ipValue)
            
            Case vbLongLong
            
                myresult = VBA.CLngLng(ipValue)
                        
            Case vbSingle
            
                myresult = VBA.CSng(ipValue)
                
            Case vbString
            
                myresult = VBA.CStr(ipValue)
                
        End Select
        
    ' End With
    
    ToVarTypePrimitive = myresult
    
End Function

Public Function GetDefaultValue(ByVal ipItem As Variant) As Variant
        
    
    Select Case True
        
        Case Me.Group.IsPrimitive(ipItem)
        
            If Me.IsString(ipItem) Then
                
                    GetDefaultValue = vbNullString
                    
            ElseIf Me.IsBoolean(ipItem) Then
                
                    GetDefaultValue = False
                    
            Else
            
                GetDefaultValue = vbNullString
            
            End If
            
        Case VBA.IsObject(ipItem)
            
            Set GetDefaultValue = Nothing
            
        Case VBA.IsEmpty(ipItem)
        
            GetDefaultValue = Empty
            
            
        Case Me.IsTypeByName(Me.Name.AsArray.ofVariant, ipItem)
        
            GetDefaultValue = Empty
            
            
        Case IsNull(ipItem)
        
            GetDefaultValue = Null
        
    End Select
    
End Function

Public Function IsSameInstance(ByVal ipInstance As Object, ByVal ipReference As Object) As Boolean
    IsSameInstance = ipInstance Is ipReference
    
End Function

Public Function IsNotSameInstance(ByVal ipInstance As Object, ByVal ipReference As Object) As Boolean
        IsNotSameInstance = Not IsSameInstance(ipInstance, ipReference)
End Function
