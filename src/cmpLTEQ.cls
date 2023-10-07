VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpLTEQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Comparers")
    
Implements IComparer

Const MY_TYPENAME                                       As String = "cmpLTEQ"

Private Type State
    Reference()                                         As Variant
End Type


Private s                                               As State


Private Property Get IComparer_TypeName() As String
    IComparer_TypeName = MY_TYPENAME
End Property


Private Function IComparer_ExecCmp(ByRef ipHostItem As Variant) As Boolean
        IComparer_ExecCmp = Comparers.LTEQ(ipHostItem, s.Reference(0))
End Function

Private Sub IComparer_Reference(ParamArray ipReference() As Variant)
    VariantCopy s.Reference, ipReference
End Sub

'@DefaultMember
Public Function Deb(ParamArray ipReference() As Variant) As cmpLTEQ
Attribute Deb.VB_UserMemId = 0
    With New cmpLTEQ
        Set Deb = .ConstructInstance(VBA.CVar(ipReference))
    End With
End Function


Friend Function ConstructInstance(ByRef ipReference As Variant) As cmpLTEQ
    s.Reference = Array(ipReference(0))
    Set ConstructInstance = Me
End Function
