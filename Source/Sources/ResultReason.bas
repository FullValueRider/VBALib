Attribute VB_Name = "ResultReason"
Option Explicit

Public Enum ReasonCode
    
    rcfail = -1
    rcSuccess = 0
    rcNoItemMethod
    rcNoCountMethod
    rcNotSingleDim
    rcToArray
    rcKeys
    rcItems
    rcNotArray
    rcArrayNotInitialised
    rcInvalidRank
    rcNotIterable
    rcNotTableArray
    rcInvalidArrayMarkup
    rcHasNoItems
    
End Enum

Private Type State

    ReasonCodeStr                   As Scripting.Dictionary

End Type

Private s                           As State

Public Function AsEnum(ByVal ipReason As ReasonCode) As ReasonCode
    AsEnum = ipReason
End Function


Public Function AsString(ByVal ipReason As ReasonCode) As String

    If s.ReasonCodeStr Is Nothing Then PopulateReasonCodeStr
    AsString = s.ReasonCodeStr.Item(ipReason)
    
End Function

Public Sub PopulateReasonCodeStr()

    Set s.ReasonCodeStr = New Scripting.Dictionary

    With s.ReasonCodeStr
    
        .Add rcfail, "Failed"
        .Add rcSuccess, "Succeeded"
        .Add rcNoItemMethod, "Object '{0}' has no Item method"
        .Add rcNoCountMethod, "Object '{0}' has no Count method"
        .Add rcNotSingleDim, "Iterable '{0}' is not a single dimension array"
        .Add rcToArray, vbNullString
        .Add rcKeys, vbNullString
        .Add rcItems, vbNullString
        .Add rcNotArray, vbNullString
        .Add rcArrayNotInitialised, "Array is not initialised"
        .Add rcInvalidRank, "The array does not have dimension '{0}'"
        .Add rcNotIterable, "The Parameter is not an Iterable Type"
        .Add rcNotTableArray, "Array dimensions are not 2D"
        .Add rcInvalidArrayMarkup, "The Array Markup was invalid"
        .Add rcHasNoItems, "The iterable does not cintain any items"
    End With
    
End Sub
