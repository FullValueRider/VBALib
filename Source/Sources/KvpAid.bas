Attribute VB_Name = "KvpAid"
Option Explicit


Public Enum ErrorReporter

    ToDebugPrint
    ToVBAErr
    
End Enum


Public Enum Indexer

    UseKey
    UseIndex
    
End Enum




' Public Enum KvpError

'     KeyNotSupported
'     AutoKeyAlreadySet
'     AutoKeyExists
'     KeysArrayEmpty
'     KeysMustBeNumber
'     SinglePairExpected
'     KeyTypeMismatch
'     OneInputNotArray
'     OneDimensionOnly
'     IterableSizeMismatch
'     TableExpected
'     OpNotSupported
'     SeparatorNotString
'     IndexOutOfBounds
'     KeyNotFound
'     KeyExists
'     NumberExpected
'     InvalidKVPair
'     UnknownSignature
'     InvalidTableKeysEnum
'     InvalidSignature
'     KeysIterableNot1D
'     KeysAreNotIterable
'     UnknownKeysOrientation
'     UnexpectedIterable
'     AdminTypeIsInvalidKey
'     UnexpectedKeyType
'     InvalidEnumMember
'     InvalidAdjustAmount
    
' End Enum


'End Enum

Public Const DefaultAdjustAmount       As Long = 1

Private Type state

    Msg                                 As Scripting.Dictionary
    Signatures                          As Scripting.Dictionary
    
End Type

Private s                               As state


' Public Function Item(ByVal ipError As KvpError) As String
    
'     If s.Msg Is Nothing Then PopulateMsg
'     If s.Msg.Count = 0 Then PopulateMsg
    
'     Item = s.Msg.Item(ipError)

' End Function


Private Sub PopulateMsg()

    ' Set s.Msg = New Scripting.Dictionary
    
    ' With s.Msg
    
    '     .Add Key:=KeyNotSupported, Item:="The Kvp class does not support keys of '{0}'.  Consider using AutoKeyByArray"
    '     .Add Key:=AutoKeyAlreadySet, Item:="The autokey has already been set and is currently {0}"
    '     .Add Key:=AutoKeyExists, Item:="AutoKey already exists"
    '     .Add Key:=KeysArrayEmpty, Item:="The Keys array is empty"
    '     .Add Key:=KeysMustBeNumber, Item:="Keys must be number types i.e.{nl2}{0}"
    '     .Add Key:=SinglePairExpected, Item:="At least one of the inputs is an array.  Consider using Pairs form"
    '     .Add Key:=KeyTypeMismatch, Item:="The Kvp keys are of Type '{0}'{nl}The provided key is of type '{1}'"
    '     .Add Key:=OneInputNotArray, Item:="One of the inputs is not an array{nl2}Keys is Type '{0}'{nl}Values is Type '{1}'"
    '     .Add Key:=OneDimensionOnly, Item:="Array size: One of the inputs is not a one dimension array"
    '     .Add Key:=IterableSizeMismatch, Item:="Size mismatch: Keys has '{0}' items{nl}Values has '{1}' items"
    '     .Add Key:=TableExpected, Item:="A Table (2D array) was expected"
    '     .Add Key:=OpNotSupported, Item:="{0} is not supported for Type '{1}'"
    '     .Add Key:=SeparatorNotString, Item:="The seperator must be a string"
    '     .Add Key:=IndexOutOfBounds, Item:="Index of '{0}' is not in range '{1}' to '{2}'"
    '     .Add Key:=KeyNotFound, Item:="The Key '{0}' was not found"
    '     .Add Key:=NumberExpected, Item:="A value of Type '{0}' was provided{nl}A number type was expected, i.e. one of{nl2}{1}"
    '     .Add Key:=KeyExists, Item:="The key '{0}' is already present"
    '     .Add Key:=InvalidKVPair, Item:="The KVPair is invalid, a variable of Type '{0}' was found"
    '     .Add Key:=UnknownSignature, Item:="The signature'{0}' is not known'"
    '     .Add Key:=InvalidTableKeysEnum, Item:="The value of the ipKeysElement '{0}' is not a valid Member of IAutoKey.TableKeys"
    '     .Add Key:=InvalidSignature, Item:="The combination of types is not permitted for the requsted operation:{nl2} '{0}'"
    '     .Add Key:=KeysIterableNot1D, Item:="The Keys iterable (when an array) is required to have only one dimension"
    '     .Add Key:=KeysAreNotIterable, Item:="The type for the keys iterable is not an iterable type: '{0}'"
    '     .Add key:=UnknownKeysOrientation, Item:="The value for ipKeysOrentation is not in the KeysOreientation enumeration"
    '     .Add Key:= UnexpectedIterable, item:="Unexpected iterable type '{0}'"
    '     .Add Key:= AdminTypeIsInvalidKey, Item:="Admin types are invalid as Keys '{0}'"
    '     .Add key:=UnexpectedKeyType, Item:="Unexpected Key type '{0}'"
    '     .Add Key:=InvalidEnumMember, item:="The value '{0}' is not a member of Enum '{1}'}"
    '     .Add Key:=InvalidAdjustAmount, item:= "Invalid Type for Inc/Dec  adjust amount '{0}'"
        
    ' End With
    
End Sub

