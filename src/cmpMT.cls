VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpMT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "True if the ipHostitem is more than the reference"
'@PredeclaredId
'@Exposed
'@Folder("Comparers")Option Explicit
'@ModuleDescription("True if the ipHostitem is more than the reference")
Option Explicit

Implements IComparer


Const MY_TYPENAME                               As String = "cmpMT"

Private Type State
    Reference                                   As Variant
End Type


Private s                                       As State


Private Property Get IComparer_TypeName() As String
    IComparer_TypeName = MY_TYPENAME
End Property


Private Function IComparer_ExecCmp(ByRef ipHostItem As Variant) As Boolean
        IComparer_ExecCmp = Comparers.MT(ipHostItem, s.Reference(0))
End Function

Private Function IComparer_Reference(ParamArray ipReference() As Variant) As IComparer
   s.Reference = ipReference
   Set IComparer_Reference = Me
End Function

'@DefaultMember
Public Function Deb(ParamArray ipReference() As Variant) As cmpMT
Attribute Deb.VB_UserMemId = 0
    With New cmpMT
        Set Deb = .ConstructInstance(VBA.CVar(ipReference))
    End With
End Function


Friend Function ConstructInstance(ByVal ipReference As Variant) As cmpMT
    s.Reference = ipReference
    Set ConstructInstance = Me
End Function
