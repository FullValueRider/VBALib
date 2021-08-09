VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Types"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule VariableNotUsed, AssignmentNotUsed
'@Folder("Lib")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'@PredeclaredId
'@Exposed
Option Explicit
 
 
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'  Class Intellisense forwarding
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


Public Property Get Arrays() As TypesArrays
    Set Arrays = TypesArrays
End Property
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'  Class Intellisense forwarding
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Function IsArray(ParamArray ipArgs() As Variant) As Boolean
    IsArray = pvIsArrayByCVar(CVar(ipArgs))
End Function

Public Function IsNotArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotArray = Not pvIsArrayByCVar(CVar(ipArgs))
End Function

Public Function pvIsArrayByCVar(ByVal ipArgs As Variant) As Boolean
    
    pvIsArrayByCVar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Not VBA.IsArray(myArg) Then Exit Function
        
    Next
    
    pvIsArrayByCVar = True
    
End Function


Public Function HasCountMethod(ByVal ipItem As Variant) As Boolean

    HasCountMethod = False
    
    Dim mycount As Long
    On Error Resume Next
    '@Ignore AssignmentNotUsed
    mycount = ipItem.Count
    
    Dim myErrNo As Long
    myErrNo = Err.Number
    
    On Error GoTo 0

    HasCountMethod = myErrNo = 0
    
End Function

Public Function LacksCountMethod(ByVal ipItem As Variant) As Boolean
    LacksCountMethod = Not HasCountMethod(ipItem)
End Function

Public Function HasItemMethod(ByVal ipEntity As Variant) As Boolean

    HasItemMethod = False
    If ipEntity.Count = 0 Then Exit Function
    '@Ignore VariableNotUsed
    Dim myItem As Variant
    On Error Resume Next
    myItem = Array(ipEntity.Item(ipEntity.Count - 1))
    Dim myErrNo As Long
    myErrNo = Err.Number
    
    On Error GoTo 0

    HasItemMethod = myErrNo = 0
    
End Function

Public Function LacksItemMethod(ByVal ipItem As Variant) As Boolean
    LacksItemMethod = Not HasItemMethod(ipItem)
End Function


Public Function IsMissing(ParamArray ipArgs() As Variant) As Boolean
    IsMissing = pvIsMissingByCVar(CVar(ipArgs))
End Function

Public Function IsNotMissing(ParamArray ipArgs() As Variant) As Boolean
    IsNotMissing = Not pvIsMissingByCVar(ipArgs)
End Function

Public Function pvIsMissingByCVar(ByVal ipArgs As Variant) As Boolean

    pvIsMissingByCVar = False
    
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

Public Function IsEmptyArray(ParamArray ipArgs() As Variant) As Boolean
    IsEmptyArray = pvIsEmptyByCVar(CVar(ipArgs))
End Function

Public Function IsNotEmptyArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotEmptyArray = pvIsEmptyByCVar(CVar(ipArgs))
End Function

Public Function pvIsEmptyByCVar(ByRef ipArgs As Variant) As Boolean

    pvIsEmptyByCVar = False

    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Types.Group.IsIterable(myArg) Then
        
           If Types.Iterable.HasItems(myArg) Then Exit Function
            
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
Public Function IsSameType(ParamArray ipArgs() As Variant) As Boolean
    IsSameType = pvIsSameTypeByCVAr(CVar(ipArgs))
End Function

Public Function IsNotSameType(ParamArray ipArgs() As Variant) As Boolean
    IsNotSameType = Not pvIsSameTypeByCVAr(CVar(ipArgs))
End Function

Public Function pvIsSameTypeByCVAr(ByVal ipArgs As Variant) As Boolean

    pvIsSameTypeByCVAr = False
    
    Dim myVarType As VbVarType
    myVarType = VarType(ipArgs(0))
    
    Dim myItem As Variant
    For Each myItem In ipArgs
    
        If Not (VarType(myItem) = myVarType) Then Exit Function
        
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
    
    Dim myArg As Variant
    For Each myArg In ipArgs
        
        If InStr(Types.Name.AsArray.OfObject, TypeName(myArg)) = 0 Then Exit Function
            
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
Public Function IsTypeByName(ByVal ipTypes As String, ParamArray ipArgs() As Variant) As Boolean
Attribute IsTypeByName.VB_Description = "Returns true if the types of all the values in the parameter list are found in the Types string"
    IsTypeByName = pvIsTypeByNameByCVar(ipTypes, ipArgs)
End Function

Public Function IsTypeByNameByCVar(ByVal ipTypes As String, ParamArray ipArgs() As Variant) As Boolean
    IsTypeByNameByCVar = Not pvIsTypeByNameByCVar(ipTypes, ipArgs)
End Function

' ParamArrays cannot be passed as a paramarray as the declaring method gets a Variant()
' so we cast Variant() (the result of a parmeter array) to a variant
' Methods Postfixed by 'byCPA' use a 'cast parameter array'
Public Function pvIsTypeByNameByCVar(ByVal ipTypes As String, ByVal ipArgs As Variant) As Boolean
' CPA stands for Collapsed parameter array
' allow for the case of a single argument provided for ipargsarray
    Dim myArgs As Variant
    If Not VBA.IsArray(ipArgs) Then
    
        myArgs = Array(ipArgs)
        
    Else
    
        myArgs = ipArgs
        
    End If

    pvIsTypeByNameByCVar = False
    
    Dim myItem As Variant
    For Each myItem In myArgs
        Debug.Print TypeName(myItem)
        
        If VBA.InStr(ipTypes, TypeName(myItem)) = 0 Then Exit Function
    
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
    
    Dim myItem As Variant
    For Each myItem In ipArgs
    
        If Not VBA.VarType(myItem) = ipVarType Then Exit Function
        
    Next
    
    pvIsVarTypeByCVar = True

End Function

Public Function IsBoolean(ParamArray ipArgs() As Variant) As Boolean
    IsBoolean = pvIsTypeByNameByCVar(Types.Name.OfBoolean, ipArgs)
End Function

Public Function IsNotBoolean(ParamArray ipArgs() As Variant) As Boolean
    IsNotBoolean = Not pvIsTypeByNameByCVar(Types.Name.OfBoolean, ipArgs)
End Function

Public Function IsBooleanArray(ParamArray ipArgs() As Variant) As Boolean
    IsBooleanArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfBoolean, ipArgs)
End Function

Public Function IsNotBooleanArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotBooleanArray = Not pvIsTypeByNameByCVar(Types.Name.AsArray.OfBoolean, ipArgs)
End Function



Public Function IsByte(ParamArray ipArgs() As Variant) As Boolean
    IsByte = pvIsTypeByNameByCVar(Types.Name.OfByte, ipArgs)
End Function

Public Function IsNotByte(ParamArray ipArgs() As Variant) As Boolean
    IsNotByte = Not pvIsTypeByNameByCVar(Types.Name.OfByte, ipArgs)
End Function

Public Function IsByteArray(ParamArray ipArgs() As Variant) As Boolean
    IsByteArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfByte, ipArgs)
End Function

Public Function IsNotByteArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotByteArray = Not pvIsTypeByNameByCVar(Types.Name.AsArray.OfByte, ipArgs)
End Function



Public Function IsCurrency(ParamArray ipArgs() As Variant) As Boolean
    IsCurrency = pvIsTypeByNameByCVar(Types.Name.OfCurrency, ipArgs)
End Function

Public Function IsNotCurrency(ParamArray ipArgs() As Variant) As Boolean
    IsNotCurrency = Not pvIsTypeByNameByCVar(Types.Name.OfCurrency, ipArgs)
End Function

Public Function IsCurrencyArray(ParamArray ipArgs() As Variant) As Boolean
    IsCurrencyArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfCurrency, ipArgs)
End Function

Public Function IsNotCurrencyArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotCurrencyArray = Not pvIsTypeByNameByCVar(Types.Name.AsArray.OfCurrency, ipArgs)
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
    IsDate = pvIsTypeByNameByCVar(Types.Name.OfDate, ipArgs)
End Function

Public Function IsNotDate(ParamArray ipArgs() As Variant) As Boolean
    IsNotDate = Not pvIsTypeByNameByCVar(Types.Name.OfDate, ipArgs)
End Function

Public Function IsDateArray(ParamArray ipArgs() As Variant) As Boolean
    IsDateArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfDate, ipArgs)
End Function

Public Function IsNotDateArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotDateArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfDate, ipArgs)
End Function


Public Function IsDecimal(ParamArray ipArgs() As Variant) As Boolean
    IsDecimal = pvIsTypeByNameByCVar(Types.Name.OfDecimal, ipArgs)
End Function

Public Function IsNotDecimal(ParamArray ipArgs() As Variant) As Boolean
    IsNotDecimal = Not pvIsTypeByNameByCVar(Types.Name.OfDecimal, ipArgs)
End Function

Public Function IsDecimalArray(ParamArray ipArgs() As Variant) As Boolean
    IsDecimalArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfDecimal, ipArgs)
End Function

Public Function IsNotDecimalArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotDecimalArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfDecimal, ipArgs)
End Function


Public Function IsDouble(ParamArray ipArgs() As Variant) As Boolean
    IsDouble = pvIsTypeByNameByCVar(Types.Name.OfDouble, ipArgs)
End Function

Public Function IsNotDouble(ParamArray ipArgs() As Variant) As Boolean
    IsNotDouble = Not pvIsTypeByNameByCVar(Types.Name.OfDouble, ipArgs)
End Function

Public Function IsDoubleArray(ParamArray ipArgs() As Variant) As Boolean
    IsDoubleArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfDouble, ipArgs)
End Function

Public Function IsNotDoubleArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotDoubleArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfDouble, ipArgs)
End Function





Public Function IsError(ParamArray ipArgs() As Variant) As Boolean
    IsError = pvIsTypeByNameByCVar(Types.Name.OfError, ipArgs)
End Function

Public Function IsNotError(ParamArray ipArgs() As Variant) As Boolean
    IsNotError = Not pvIsTypeByNameByCVar(Types.Name.OfError, ipArgs)
End Function

Public Function IsErrorArray(ParamArray ipArgs() As Variant) As Boolean
    IsErrorArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfError, ipArgs)
End Function

Public Function IsNotErrorArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotErrorArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfError, ipArgs)
End Function


Public Function IsNothing(ParamArray ipArgs() As Variant) As Boolean
    IsNothing = pvIsTypeByNameByCVar(Types.Name.OfNothing, ipArgs)
End Function

Public Function IsNotNothing(ParamArray ipArgs() As Variant) As Boolean
    IsNotNothing = Not pvIsTypeByNameByCVar(Types.Name.OfNothing, ipArgs)
End Function

Public Function IsNothingArray(ParamArray ipArgs() As Variant) As Boolean
    IsNothingArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfNothing, ipArgs)
End Function

Public Function IsNotNothingArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotNothingArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfNothing, ipArgs)
End Function


Public Function IsNull(ParamArray ipArgs() As Variant) As Boolean
    IsNull = pvIsTypeByNameByCVar(Types.Name.OfNull, ipArgs)
End Function

Public Function IsNotNull(ParamArray ipArgs() As Variant) As Boolean
    IsNotNull = Not pvIsTypeByNameByCVar(Types.Name.OfNull, ipArgs)
End Function

Public Function IsNullArray(ParamArray ipArgs() As Variant) As Boolean
    IsNullArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfNull, ipArgs)
End Function

Public Function IsNotNullArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotNullArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfNull, ipArgs)
End Function


Public Function IsInteger(ParamArray ipArgs() As Variant) As Boolean
    IsInteger = pvIsTypeByNameByCVar(Types.Name.OfInteger, ipArgs)
End Function

Public Function IsNotInteger(ParamArray ipArgs() As Variant) As Boolean
    IsNotInteger = Not pvIsTypeByNameByCVar(Types.Name.OfInteger, ipArgs)
End Function

Public Function IsIntegerArray(ParamArray ipArgs() As Variant) As Boolean
    IsIntegerArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfInteger, ipArgs)
End Function

Public Function IsNotIntegerArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotIntegerArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfInteger, ipArgs)
End Function


Public Function IsLong(ParamArray ipArgs() As Variant) As Boolean
    IsLong = pvIsTypeByNameByCVar(Types.Name.OfLong, ipArgs)
End Function

Public Function IsNotLong(ParamArray ipArgs() As Variant) As Boolean
    IsNotLong = Not pvIsTypeByNameByCVar(Types.Name.OfLong, ipArgs)
End Function

Public Function IsLongArray(ParamArray ipArgs() As Variant) As Boolean
    IsLongArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfLong, ipArgs)
End Function

Public Function IsNotLongArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotLongArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfLong, ipArgs)
End Function


Public Function IsLongLong(ParamArray ipArgs() As Variant) As Boolean
    IsLongLong = pvIsTypeByNameByCVar(Types.Name.OfLongLong, ipArgs)
End Function

Public Function IsNotLongLong(ParamArray ipArgs() As Variant) As Boolean
    IsNotLongLong = Not pvIsTypeByNameByCVar(Types.Name.OfLongLong, ipArgs)
End Function

Public Function IsLongLongArray(ParamArray ipArgs() As Variant) As Boolean
    IsLongLongArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfLongLong, ipArgs)
End Function

Public Function IsNotLongLongArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotLongLongArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfLongLong, ipArgs)
End Function


'@Description(" True if the object type is object.  Which is different to asking if a variable is an object")
Public Function IsObjectObject(ParamArray ipArgs() As Variant) As Boolean
Attribute IsObjectObject.VB_Description = " True if the object type is object.  Which is different to asking if a variable is an object"
    IsObjectObject = pvIsTypeByNameByCVar(Types.Name.OfObject, ipArgs)
End Function

Public Function IsNotObjectObject(ParamArray ipArgs() As Variant) As Boolean
    IsNotObjectObject = Not pvIsTypeByNameByCVar(Types.Name.OfObject, ipArgs)
End Function

Public Function IsObjectObjectArray(ParamArray ipArgs() As Variant) As Boolean
    IsObjectObjectArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfObject, ipArgs)
End Function

Public Function IsNotArrayOfObjectObjects(ParamArray ipArgs() As Variant) As Boolean
    IsNotArrayOfObjectObjects = Not pvIsTypeByNameByCVar(Types.Name.AsArray.OfObject, ipArgs)
End Function


Public Function IsSingle(ParamArray ipArgs() As Variant) As Boolean
    IsSingle = pvIsTypeByNameByCVar(Types.Name.OfSingle, ipArgs)
End Function

Public Function IsNotSingle(ParamArray ipArgs() As Variant) As Boolean
    IsNotSingle = Not pvIsTypeByNameByCVar(Types.Name.OfSingle, ipArgs)
End Function

Public Function IsSingleArray(ParamArray ipArgs() As Variant) As Boolean
    IsSingleArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfSingle, ipArgs)
End Function

Public Function IsNotSingleArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotSingleArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfSingle, ipArgs)
End Function


Public Function IsString(ParamArray ipArgs() As Variant) As Boolean
    IsString = pvIsTypeByNameByCVar(Types.Name.OfString, ipArgs)
End Function

Public Function IsNotString(ParamArray ipArgs() As Variant) As Boolean
    IsNotString = pvIsTypeByNameByCVar(Types.Name.OfString, ipArgs)
End Function

Public Function IsStringArray(ParamArray ipArgs() As Variant) As Boolean
    IsStringArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfString, ipArgs)
End Function

Public Function IsNotStringArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotStringArray = Not pvIsTypeByNameByCVar(Types.Name.AsArray.OfString, ipArgs)
End Function


Public Function IsUnknown(ParamArray ipArgs() As Variant) As Boolean
    IsUnknown = pvIsTypeByNameByCVar(Types.Name.OfUnknown, ipArgs)
End Function

Public Function IsNotUnknown(ParamArray ipArgs() As Variant) As Boolean
    IsNotUnknown = Not pvIsTypeByNameByCVar(Types.Name.OfUnknown, ipArgs)
End Function

Public Function IsUnknownArray(ParamArray ipArgs() As Variant) As Boolean
    IsUnknownArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfUnknown, ipArgs)
End Function


Public Function IsNotUnknownArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotUnknownArray = pvIsTypeByNameByCVar(Types.Name.AsArray.OfUnknown, ipArgs)
End Function
' Its not possible to test for a variable being a variant in vba
Public Function IsVariantArray(ParamArray ipArgs() As Variant) As Boolean
    IsVariantArray = pvIsTypeByNameByCVar(Types.Name.AsArray.ofVariant, ipArgs)
End Function

Public Function IsNotVariantArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotVariantArray = pvIsTypeByNameByCVar(Types.Name.AsArray.ofVariant, ipArgs)
End Function


''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
''
''  Miscellaneous
''
''========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Function ToNative(ByVal ipType As String, ByVal ipValue As Variant) As Variant

    If Types.Group.IsNotPrimitive(ipValue) Then

        Err.Raise _
            17, _
            "VBALib.COnvertor.ToPrimitiveType", _
            Fmt.Txt("Can't convert type '{0}' to type '{1}'", TypeName(ipValue), ipType)
            
        Exit Function
            
    End If
    
    With Types.Name
    
        Select Case ipType
        
            Case .OfBoolean
            
                ToNative = VBA.CBool(ipValue)
                
            Case .OfByte
            
                ToNative = VBA.CByte(ipValue)
            
            Case .OfCurrency
            
                ToNative = VBA.CCur(ipValue)
            
            Case .OfDate
            
                ToNative = VBA.CDate(ipValue)
                
            Case .OfDouble
            
                ToNative = VBA.CDbl(ipValue)
        
            Case .OfDecimal
            
                ToNative = VBA.CDec(ipValue)
            
            Case .OfInteger
            
                ToNative = VBA.CInt(ipValue)
            
            Case .OfLong
            
                ToNative = VBA.CLng(ipValue)
            
            Case .OfLongLong
            
                ToNative = VBA.CLngLng(ipValue)
                        
            Case .OfSingle
            
                ToNative = VBA.CSng(ipValue)
                
            Case .OfString
            
                ToNative = VBA.CStr(ipValue)
                
        End Select
        
    End With
    
End Function
