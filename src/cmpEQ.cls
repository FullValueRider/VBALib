VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpEQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Comparers")
'@ModuleDescription("True if the ipHostitem is equal to the reference")Option Explicit
Option Explicit

Implements IComparer

Const MY_TYPENAME                               As String = "cmpEQ"

Private Type State
    Reference()                                 As Variant
End Type


Private s                                       As State


Private Property Get IComparer_TypeName() As String
    IComparer_TypeName = MY_TYPENAME
End Property


Private Function IComparer_ExecCmp(ByRef ipHostItem As Variant) As Boolean
    IComparer_ExecCmp = Comparers.EQ(ipHostItem, s.Reference(0))
End Function

'@Description("Takes a single value which is used to reset the reference value used for Comparison")
Private Function IComparer_Reference(ParamArray ipReference() As Variant) As IComparer
Attribute IComparer_Reference.VB_Description = "Takes a single value which is used to reset the reference value used for Comparison"
    s.Reference = ipReference
    Set IComparer_Reference = Me
End Function

'@DefaultMember
' The paramarray allows mutiple parameters to be provided if required
' The IComparer interface allows the name of the parameter to be
' anything so where more than one value is required this should be
' captured by the name used for Paramarray.
' e.g. cmpIn uses ipLowHighBounds
Public Function Deb(ParamArray ipReference() As Variant) As cmpEQ
Attribute Deb.VB_UserMemId = 0
    With New cmpEQ
        Set Deb = .ConstructInstance(VBA.CVar(ipReference))
    End With
End Function


Friend Function ConstructInstance(ByRef ipReference As Variant) As cmpEQ
    s.Reference = ipReference
    Set ConstructInstance = Me
End Function
'
