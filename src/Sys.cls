VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A location for global helper methods"
'@PredeclaredId
'@Exposed
'@Folder("VBALib")
Option Explicit


'@ModuleDescription("A location for global helper methods")
'@Description("Revises the FPA (forwarded ParamArray) to FPA(0) if FPA contains a single item of the Enumarable group, a number or a string.")
'@Ignore ProcedureCanBeWrittenAsFunction

' One of the many issues with VBA is that it supports methods with a variable number of parameters via the ParamArray.,
' However, there is no 'splat' operater so Paramarrays cannot be forwarded to methods expecting a paramArray
' To get around this limitation we use the resolve paramarray method which assumes that if the FPA (forwarded ParamArray)
' contains a single item then the intent is to interpret FPA(0) as that single item.
' On the occasion that single item processing is not required the Item shouls be encapsulated as an array
' e.g. Method(Item), use Method(Array(Item)).
'ResolveParamArray returns true if the paramarray was reduced  to a single Item, false if not.
Public Function ResolveParamArray(ByRef iopFPA As Variant) As Boolean
Attribute ResolveParamArray.VB_Description = "Revises the FPA (forwarded ParamArray) to FPA(0) if FPA contains a single item of the Enumarable group, a number or a string."
    ResolveParamArray = False
    ' the FPA must be an allocated array i.e. not <Array>() to be processable
    If ArrayInfo.LacksItems(iopFPA) Then
        Exit Function
    End If
    
   ' if the FPA contains only one item then
   ' Lbound = Ubound
   ' so if iopFPA has more than one item its a true paramarray
    If LBound(iopFPA, 1) <> UBound(iopFPA, 1) Then
        Exit Function
    End If
    
    ' The paramArray contains a single Item so we convert the array to the single Item
    If VBA.IsObject(iopFPA(0)) Then
        Set iopFPA = iopFPA(0)
    Else
        iopFPA = iopFPA(0)
    End If
    
    ResolveParamArray = True
    
End Function
