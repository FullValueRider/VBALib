VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpMTEQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Comparers")
    
Implements IComparer


Const MY_TYPENAME                   As String = "cmpMTEQ"

Private Type State
    Reference                       As Variant
End Type


Private s                           As State


Private Property Get IComparer_TypeName() As String
    IComparer_TypeName = MY_TYPENAME
End Property


Private Function IComparer_ExecCmp(ByRef ipHostItem As Variant) As Boolean
    IComparer_ExecCmp = Comparers.MTEQ(ipHostItem, s.Reference(0))
End Function


'@DefaultMember
Public Function Deb(ByVal ipReference As Variant) As cmpMTEQ
Attribute Deb.VB_UserMemId = 0
    With New cmpMTEQ
        Set Deb = .ConstructInstance(ipReference)
    End With
End Function


Friend Function ConstructInstance(ByVal ipReference As Variant) As cmpMTEQ
    s.Reference = Array(ipReference)
    Set ConstructInstance = Me
End Function
