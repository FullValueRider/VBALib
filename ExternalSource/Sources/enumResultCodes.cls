VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "enumResultCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@Folder("Enums")
Option Explicit

Private Type InternalConstants

    Location                    As String
    
    
End Type
    
Private ic                      As InternalConstants
    


Public Enum ResultCode
    
    Unspecified = CustomErrors.Base
    NotOkay = -1
    Okay = 0
    
    'TypeIssues
    InvalidMethod
    'NoCountMethod
    NotListArray
    IsNotNumber
    IsNotString
    NotSameType
    
    Keys
    Items
    'Class issues
    UsePredeclaredIdInstance
    InvalidNew
    UseNewInstance
    
    InvalidInstanceCreation
    DisallowedOverwrite
    
    DoubleInitialisation
    NullReference
    
    ' Dictionary Issues
    AutoKeyResetNotAllowed
    InvalidKey
    
    ' Iterable issues
    NotIterableType
    InvalidIterable
    NotIterableByIndex
    InvalidType
    HasNoItems
    
    'Array Issues
    'ArrayNotInitialised
    IsNotArray
    NotTableArray
    InvalidArrayMarkup
    IndexOutOfBounds
    InvalidRank
    UnexpectedItemInParamArray
    
    'Parameter issues
    UnexpectedEmpty
    StringCannotBeEmpty
    
    IndexLessThanFirstIndex
    IndexMoreThanLastIndex
    IndexWillOverrun
    DestinationTooSmall
    KeyNotFound
    InvalidInsertCountZero
    
    
    
    
    'Enumeration issues
    ActionNotExecutable
    NotEnumerationMember
    DisAllowedEnumerationMember
    
End Enum


Private Type State
    
    Actions                 As Scripting.Dictionary
    Strings                 As Scripting.Dictionary
    
End Type

Private s                   As State


Private Sub Class_Initialise()
    
    ic.Location = "VBALib.EnumGuardClauses"
    Guard InvalidNew, Types.IsSameInstance(Me, Enums.GuardClauses), ic.Location & "Class_Initialize"
    SetupDictionaries
    
End Sub


Private Sub SetupDictionaries()
    
    Set s.Strings = CreateObject("Scripting.Dictionary")
    With s.Strings
    
        '.Add IsMissing, "No enumeration member was provided"
        .Add Unspecified, "An unspecified error has occurred"
        .Add NotSameType, "The types do not match '{0}','{1}"
        .Add InvalidMethod, "Object '{0}' does not have a '{1}' method"
        '.Add NoCountMethod, "Object '{0}' has no Count method"
        .Add IsNotNumber, "Expected a number"
        .Add IsNotString, "Expected a string"
        .Add UnexpectedEmpty, "Unexpected Empty in '{0}'"
        .Add DisallowedOverwrite, "The current instance of '{0}' cannot be overridden by a new instance"
        
        .Add IsNotArray, "Expecting an array: Found '{0}'"
        '.Add InvalidDimensions, "Not a '{0}' dimensional array"
        '.Add NotSingleDimension, "Item '{0}' is not a single dimension array"
        .Add InvalidRank, "Type '{0}' does not have a rank of '{0}'"
        
        .Add Keys, vbNullString
        .Add Items, vbNullString
        '.Add ArrayNotInitialised, "Array is not initialised"
        
        .Add AutoKeyResetNotAllowed, "Second attempty at initialising Autokey"
        
        .Add InvalidKey, "The specified key Type '{0}' does not mast the Host Kvp Key Type '{1}'"
        
        .Add NotIterableType, "Type '{0}' is not iterable"
        .Add InvalidIterable, "'{0}' is not an Iterable Type"
        .Add NotIterableByIndex, "'{0}' is not iterable by Index"
        .Add InvalidArrayMarkup, "The Array Markup was invalid"
        .Add HasNoItems, "The iterable {0} does not contain any items"
        .Add InvalidType, "Type error: expecting {0} found '{1} in '{2}''"
        
        .Add IndexLessThanFirstIndex, "Value is less than the first index"
        .Add IndexMoreThanLastIndex, "Value is more then the last index"
        .Add IndexWillOverrun, "Index will exceed size"
        .Add KeyNotFound, "Key of '{0}' was not found"
        .Add InvalidInsertCountZero, "Cannot insert if count is zero"
        
        .Add DestinationTooSmall, "Cannot copy {0} items to array of {1}"
        .Add UsePredeclaredIdInstance, "Method should be invoked from the default/predeclared instance of this class."
        .Add UseNewInstance, "Method should be invoked from a new instance of this class."
        .Add InvalidNew, "New should no be used for with this predecalredIs '{0}'"
        .Add InvalidInstanceCreation, "Use the PredeclaredId of this class only"
        
        .Add DoubleInitialisation, "Raises a run-time error if the specified object reference is already set."
        .Add NullReference, "Object reference '{0}' cannot be Nothing."
        .Add StringCannotBeEmpty, "String cannot be empty."
        
        .Add ActionNotExecutable, "The enumeration action is not an executable object"
        .Add NotEnumerationMember, "Not a member of Enumeration '{0}'"
        .Add DisAllowedEnumerationMember, "Method does not allow enummeration member at index {0}"
        .Add UnexpectedItemInParamArray, "Don't know how to parse Type '{0}'"
    End With
    
End Sub

Public Function Strings() As Variant
    Strings = s.Strings.Items
End Function


Public Function Members() As Variant
    Members = s.Strings.Keys
End Function


Public Function Actions() As Variant
    Actions = s.Actions.Items
End Function


Public Function Ord(ByVal ipEnum As ResultCode) As Long
    
    Dim myItem As Variant
    Dim myOrd As Long
    myOrd = 1
    For Each myItem In s.Strings.Keys
        
        If myItem = ipEnum Then
            
            Ord = myOrd
            Exit Function
            
        Else
        
            myOrd = myOrd + 1
            
        End If
        
    Next
    
End Function

Public Property Get AsEnum(ByVal ipEnum As ResultCode) As ResultCode
    AsEnum = ipEnum
End Property

' Public Function CanExec(ByVal ipEnum As GuardClause) As Boolean
'     CanExec = VBA.IsObject(s.Actions.Item(ipEnum))
' End Function

Public Property Get ToString(ByVal ipEnum As ResultCode) As String
    Guard ResultCode.NotEnumerationMember, Enums.GuardClauses.IsMember(ipEnum), ic.Location & "ToString", Array(ic.Location)
    ToString = s.Strings.Item(ipEnum)
End Property
    
' Public Property Get Exec(ByVal ipEnum As Inspections, ByVal ipTableArray As Variant) As Variant
'     ' if Guard.ActionNotExecutable() s.Actions.Item(ipEnum), s.Location & "Exec") then guard.raiseerror
'     If Not VBA.IsObject(s.Actions.Item(ipEnum)) Then Vba.information.err.raise 17, s.location & "Exec", "The enumeration does not have an associated Action")
'     Set Exec = s.Actions.Item(ipEnum).exec(ipTableArray)
' End Property

Public Property Get IsMember(ByVal ipValue As Long) As Boolean
    If s.Strings Is Nothing Then SetupDictionaries
    IsMember = s.Strings.Exists(ipValue)
End Property

Public Property Get Name() As String
    Name = TypeName(Me)
End Property

Public Property Get Count() As Long
    If s.Strings Is Nothing Then SetupDictionaries
    Count = s.Strings.Count
End Property
