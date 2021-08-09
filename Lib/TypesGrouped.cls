VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypesGrouped"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Type groupings
' Admin: Types that represent a value but which are not defined as primitives
'         - a mix of primitive and object types
'
' Short: Types whose values are stored in the variable descriptor
' see https://bytecomb.com/vba-internals-whats-in-a-variable/
'
' Numbers: Shorts which can be incremented or decremented by 1
'
' Primitives: Types that hold a single value which is not accessed by dot notation
'
' Iterable: prefix to indicate an array or other 'collection' object type
'
' ByIndex           Type allows Var(integerIndex) or Var.Item(integerIndex)
' ItemsByEnum       Types which allow ForEach enumerations to return items
' KeysByKey         Types which allow ForEach enumerations to return keys
' ItemsByMethod     Types which do not use .Item(Key), .Item(integerIndex) or (integerIndex) to return items (e.g. queues and stacks)
'
' Iterables         Types that can be enumerated using a For Each construct.
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit
'@PredeclaredId
'@Exposed

Private Type Properties
    Admin                                   As Lyst
    Numbers                                 As Lyst
    Shorts                                  As Lyst
    Primitives                              As Lyst
    IterableNumbers                         As Lyst
    IterableShorts                          As Lyst
    IterablePrimitives                      As Lyst
    IterableByIndex                         As Lyst  ' Allows items in an object can be referenced as 'objectinstance(integer index)'
    IterableItemsByEnum                     As Lyst  ' items in object can be iterated using for each
    IterableByMethod                        As Lyst  ' items can be retrieved using a specific method
    IterableKeysByEnum                      As Lyst  ' Has a Keys method and 'for each' gives the key
    Iterables                               As Lyst
    
    
End Type

Private p                                   As Properties


Private Sub Class_Initialize()

    ' Populate the group lists only if the instance is the predecalred instance
    If Me Is TypesGrouped Then
    
        If p.Numbers Is Nothing Then
    
            Set p.Numbers = VBALib.Strings.ToSubStrLyst(ShortNumberTypes)
            Set p.Shorts = VBALib.Strings.ToSubStrLyst(ShortTypes)
            Set p.Primitives = VBALib.Strings.ToSubStrLyst(PrimitiveTypes)
            Set p.IterableNumbers = VBALib.Strings.ToSubStrLyst(IterableNumberTypes)
            Set p.IterableShorts = VBALib.Strings.ToSubStrLyst(IterableShortTypes)
            Set p.IterablePrimitives = VBALib.Strings.ToSubStrLyst(IterablePrimitiveTypes)
            Set p.IterableByIndex = VBALib.Strings.ToSubStrLyst(IterableByIndexTypes)
            Set p.IterableItemsByEnum = VBALib.Strings.ToSubStrLyst(IterableItemsByEnumTypes)
            Set p.IterableByMethod = VBALib.Strings.ToSubStrLyst(IterableByMethodTypes)
            Set p.IterableKeysByEnum = VBALib.Strings.ToSubStrLyst(IterableKeysByEnumTypes)
            Set p.Iterables = VBALib.Strings.ToSubStrLyst(IterableTypes)
            Set p.Admin = VBALib.Strings.ToSubStrLyst(AdminTypes)
            
        End If
        
    End If
    
End Sub

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'  Property get to emulate constants so that we can use functions during the
'  assembly of the group strings
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Private Property Get AdminTypes() As String
    AdminTypes = "Empty,Error,Nothing,Null,Unknown"
End Property

Private Property Get ShortNumberTypes() As String
    ShortNumberTypes = "Byte,Currency,Date,Decimal,Double,Integer,Long,LongLong,LongPtr,Single"
End Property

Private Property Get ShortTypes() As String
    ShortTypes = "Boolean" & "," & ShortNumberTypes
End Property

Private Property Get PrimitiveTypes() As String
    PrimitiveTypes = "String" & "," & ShortTypes
End Property

Private Property Get IterableNumberTypes() As String
    IterableNumberTypes = VBA.Replace(VBA.Trim$(ShortNumberTypes), Char.comma, Char.LParen & Char.RParen & Char.comma) & Char.LParen & Char.RParen
End Property

Private Property Get IterableShortTypes() As String
    IterableShortTypes = VBA.Replace(VBA.Trim$(ShortTypes), Char.comma, Char.LParen & Char.RParen & Char.comma) & Char.LParen & Char.RParen
End Property

Private Property Get IterablePrimitiveTypes() As String
    IterablePrimitiveTypes = VBA.Replace(VBA.Trim$(PrimitiveTypes), Char.comma, Char.LParen & Char.RParen & Char.comma) & Char.LParen & Char.RParen
End Property

Private Property Get IterableByIndexTypes() As String
    IterableByIndexTypes = Types.Name.AsArray.ofVariant & Char.comma & Types.Name.OfArrayList & Char.comma & Types.Name.OfLyst & Char.comma & Types.Name.OfCollection & Char.comma & IterablePrimitiveTypes
End Property

Private Property Get IterableByMethodTypes() As String
    IterableByMethodTypes = Types.Name.OfQueue & Char.comma & Types.Name.OfStack
End Property

Private Property Get IterableItemsByEnumTypes() As String
    IterableItemsByEnumTypes = IterableByIndexTypes & Char.comma & Types.Name.OfQueue & Char.comma & Types.Name.OfStack
End Property

Private Property Get IterableKeysByEnumTypes() As String
    IterableKeysByEnumTypes = Types.Name.OfScriptingDictionary & Char.comma & Types.Name.OfKvp
End Property

Private Property Get IterableTypes() As String
    IterableTypes = IterableByIndexTypes & Char.comma & IterableKeysByEnumTypes & Char.comma & IterableByMethodTypes
End Property

' Additional groups may be defined depending on user requirements.

'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'  Type groups- returns lists (as Lyst object) of names of types in groups
'  Public at present but it may be less confusing if these functions are private
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Property Get Admin() As Lyst
    Set Admin = p.Admin
End Property

Public Property Get Numbers() As Lyst
    Set Numbers = p.Numbers
End Property

Public Property Get Shorts() As Lyst
    Set Shorts = p.Shorts
End Property

Public Property Get Primitives() As Lyst
    Set Primitives = p.Primitives
End Property

Public Property Get IterableNumbers() As Lyst
    Set IterableNumbers = p.IterableNumbers
End Property

Public Property Get IterableShorts() As Lyst
    Set IterableShorts = p.IterableShorts
End Property

Public Property Get IterablePrimitives() As Lyst
    Set IterablePrimitives = p.IterablePrimitives
End Property

Public Property Get IterableByIndex() As Lyst
    Set IterableByIndex = p.IterableByIndex
End Property

Public Property Get IterableByMethod() As Lyst
    Set IterableByMethod = p.IterableByMethod
End Property

Public Property Get IterableItemsByEnum() As Lyst
    Set IterableItemsByEnum = p.IterableItemsByEnum
End Property

Public Property Get IterableKeysByEnum() As Lyst
    Set IterableKeysByEnum = p.IterableKeysByEnum
End Property

Public Property Get Iterables() As Lyst
    Set Iterables = p.Iterables
End Property


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
'  Tests for group types - basically lots of sugar
'  My opinion is that IsX(Var) has less cognitive load that X.HoldsValue(Typename(Var))
'  Likewise I prefer 'IsNoX' to 'Not IsX'
'  Multiple variables can be tested for a type by one call
'  VBA functions (e.g. IsArray) should be preferred when the context allows.
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' ToDo: add try functions to return ResultLyst where ResultLyst.value is a lyst
' of two lysts
'   Item(0) are of Type
'   Item(1) are not of Type
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Public Function IsAdmin(ParamArray ipArgs() As Variant) As Boolean
    IsAdmin = pvIsAdminByCvar(CVar(ipArgs))
End Function

Public Function IsNotAdmin(ParamArray ipArgs() As Variant) As Boolean
    IsNotAdmin = Not pvIsAdminByCvar(CVar(ipArgs))
End Function

Public Function pvIsAdminByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsAdminByCvar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
        'Debug.Print
        If Types.Group.Admin.LacksValue(TypeName(myArg)) Then Exit Function
    
    Next
    
    pvIsAdminByCvar = True
    
End Function



Public Function IsNumber(ParamArray ipArgs() As Variant) As Boolean
    IsNumber = pvIsNumberByCvar(CVar(ipArgs))
End Function

Public Function IsNotNumber(ParamArray ipArgs() As Variant) As Boolean
    IsNotNumber = Not pvIsNumberByCvar(CVar(ipArgs))
End Function

Public Function pvIsNumberByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsNumberByCvar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
        
        If Types.Group.Numbers.LacksValue(TypeName(myArg)) Then Exit Function
    
    Next
    
    pvIsNumberByCvar = True
    
End Function


Public Function IsShort(ParamArray ipArgs() As Variant) As Boolean
    IsShort = pvIsShortByCvar(CVar(ipArgs))
End Function

Public Function IsNotShort(ParamArray ipArgs() As Variant) As Boolean
    IsNotShort = Not pvIsShortByCvar(CVar(ipArgs))
End Function

Public Function pvIsShortByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsShortByCvar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Types.Group.Shorts.LacksValue(TypeName(myArg)) Then Exit Function
    
    Next
    
    pvIsShortByCvar = True
    
End Function


Public Function IsPrimitive(ParamArray ipArgs() As Variant) As Boolean
    IsPrimitive = pvIsPrimitiveByCvar(CVar(ipArgs))
End Function

Public Function IsNotPrimitive(ParamArray ipArgs() As Variant) As Boolean
    IsNotPrimitive = Not pvIsPrimitiveByCvar(CVar(ipArgs))
End Function

Public Function pvIsPrimitiveByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsPrimitiveByCvar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Types.Group.Primitives.LacksValue(TypeName(myArg)) Then Exit Function
    
    Next
    
    pvIsPrimitiveByCvar = True
    
End Function


Public Function IsIterableNumber(ParamArray ipArgs() As Variant) As Boolean
    IsIterableNumber = pvIsIterableNumberByCvar(CVar(ipArgs))
End Function

Public Function IsNotIterableNumber(ParamArray ipArgs() As Variant) As Boolean
    IsNotIterableNumber = Not pvIsIterableNumberByCvar(CVar(ipArgs))
End Function

Public Function pvIsIterableNumberByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsIterableNumberByCvar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Types.Group.IterableNumbers.LacksValue(TypeName(myArg)) Then Exit Function
    
    Next
    
    pvIsIterableNumberByCvar = True
    
End Function



Public Function IsIterableShort(ParamArray ipArgs() As Variant) As Boolean
    IsIterableShort = pvIsIterableShortByCvar(CVar(ipArgs))
End Function

Public Function IsNotIterableShort(ParamArray ipArgs() As Variant) As Boolean
    IsNotIterableShort = Not pvIsIterableShortByCvar(CVar(ipArgs))
End Function

Public Function pvIsIterableShortByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsIterableShortByCvar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Types.Group.IterableShorts.LacksValue(TypeName(myArg)) Then Exit Function
    
    Next
    
    pvIsIterableShortByCvar = True
    
End Function


Public Function IsIterablePrimitive(ParamArray ipArgs() As Variant) As Boolean
    IsIterablePrimitive = pvIsIterablePrimitiveByCvar(CVar(ipArgs))
End Function

Public Function IsNotIterablePrimitive(ParamArray ipArgs() As Variant) As Boolean
    IsNotIterablePrimitive = Not pvIsIterablePrimitiveByCvar(CVar(ipArgs))
End Function

Public Function pvIsIterablePrimitiveByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsIterablePrimitiveByCvar = False
    
    Dim myItem As Variant
    For Each myItem In ipArgs
    
        If Types.Group.IterablePrimitives.LacksValue(TypeName(myItem)) Then Exit Function
    
    Next
    
    pvIsIterablePrimitiveByCvar = True
    
End Function


Public Function IsIterableByIndex(ParamArray ipArgs() As Variant) As Boolean
    IsIterableByIndex = pvIsIterableByIndexByCVar(CVar(ipArgs))
End Function

Public Function IsNotIterableByIndex(ParamArray ipArgs() As Variant) As Boolean
    IsNotIterableByIndex = Not pvIsIterableByIndexByCVar(CVar(ipArgs))
End Function

Public Function pvIsIterableByIndexByCVar(ByVal ipArgs As Variant) As Boolean

    pvIsIterableByIndexByCVar = False
    
    Dim myItem As Variant
    For Each myItem In ipArgs
    
        If Types.Group.IterableByIndex.LacksValue(TypeName(myItem)) Then Exit Function
    
    Next
    
    pvIsIterableByIndexByCVar = True
    
End Function


Public Function IsIterableByMethod(ParamArray ipArgs() As Variant) As Boolean
    IsIterableByMethod = pvIsIterableByMethodByCvar(CVar(ipArgs))
End Function

Public Function IsNotIterableByMethod(ParamArray ipArgs() As Variant) As Boolean
    IsNotIterableByMethod = Not pvIsIterableByMethodByCvar(CVar(ipArgs))
End Function

Public Function pvIsIterableByMethodByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsIterableByMethodByCvar = False
    
    Dim myItem As Variant
    For Each myItem In ipArgs
    
        If Types.Group.IterableByMethod.LacksValue(TypeName(myItem)) Then Exit Function
    
    Next
    
    pvIsIterableByMethodByCvar = True
    
End Function


Public Function IsIterableItemsByEnum(ParamArray ipArgs() As Variant) As Boolean
    IsIterableItemsByEnum = pvIsIterableItemsByEnumByCvar(CVar(ipArgs))
End Function

Public Function IsNotIterableItemsByEnum(ParamArray ipArgs() As Variant) As Boolean
    IsNotIterableItemsByEnum = Not pvIsIterableItemsByEnumByCvar(CVar(ipArgs))
End Function

Public Function pvIsIterableItemsByEnumByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsIterableItemsByEnumByCvar = False
    
    Dim myItem As Variant
    For Each myItem In ipArgs
    
        If Types.Group.IterableItemsByEnum.LacksValue(TypeName(myItem)) Then Exit Function
    
    Next
    
    pvIsIterableItemsByEnumByCvar = True
    
End Function


Public Function IsIterableKeysByEnum(ParamArray ipArgs() As Variant) As Boolean
    IsIterableKeysByEnum = pvIsIterableByKeyByCvar(CVar(ipArgs))
End Function

Public Function IsNotIterableByKey(ParamArray ipArgs() As Variant) As Boolean
    IsNotIterableByKey = Not pvIsIterableByKeyByCvar(CVar(ipArgs))
End Function

Public Function pvIsIterableByKeyByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsIterableByKeyByCvar = False
    
    Dim myItem As Variant
    For Each myItem In ipArgs
    
        If IterableKeysByEnum.LacksValue(TypeName(myItem)) Then Exit Function
    
    Next
    
    pvIsIterableByKeyByCvar = True
    
End Function

Public Function IsIterable(ParamArray ipArgs() As Variant) As Boolean
    IsIterable = pvIsIterableByCvar(CVar(ipArgs))
End Function

Public Function IsNotIterable(ParamArray ipArgs() As Variant) As Boolean
    IsNotIterable = Not pvIsIterableByCvar(CVar(ipArgs))
End Function

Public Function pvIsIterableByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsIterableByCvar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
        
        If Not VBA.IsArray(myArg) And p.Iterables.LacksValue(TypeName(myArg)) Then Exit Function
    
    Next
    
    pvIsIterableByCvar = True
    
End Function
