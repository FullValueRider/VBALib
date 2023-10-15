VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Number"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A class to represent the generic case of a number"
'@PredeclaredId
'@Exposed
Option Explicit

'@ModuleDescription("A class to represent the generic case of a number")

Private Type Properties
    Value                                           As Variant
End Type

Private p                                           As Properties


Public Function Initialise(ByVal ipNumber As Variant) As Number
    If GroupInfo.IsNotNumber(ipNumber) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.Number", _
            Fmt.Text("Expecting a number. Got {0}", VBA.TypeName(ipNumber))
    End If
    
    p.Value = ipNumber
    
    Set Initialise = Me
    
End Function

'@DefaultMember
Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = p.Value
End Property


Public Property Let Value(ByVal ipNumber As Variant)
    ' the inner parentheses enforces evaluation of ipNumber as an expression
    ' thus if a 'Number' instance is passed in we get the Value of the instance
    ' ToDo: consider adding this action to the GroupInfo.isnumber method.
    ' i.e. add a specific test for tye of 'Number'
    If GroupInfo.IsNotNumber((ipNumber)) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.Number", _
            Fmt.Text("Expecting a number. Got {0}", VBA.TypeName(ipNumber))
    End If
    p.Value = ipNumber
End Property
