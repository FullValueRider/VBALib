VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpLT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Comparers")
    
Implements IComparer

Const MY_TYPENAME               As String = "cmpLT"

Private Type State
    Reference                   As Variant
    RefType                     As VbVarType
End Type


Private s                       As State


Private Property Get IComparer_TypeName() As String
    IComparer_TypeName = MY_TYPENAME
End Property


Private Function IComparer_ExecCmp(ByRef ipHostItem As Variant) As Boolean
    
    Dim myHostItem As Variant
    myHostItem = GetItemAsComparerValue(ipHostItem)
    
    If StringNumberComparison(s.Reference, myHostItem) Then
        IComparer_ExecCmp = False
    Else
        IComparer_ExecCmp = myHostItem < s.Reference
    End If
    
End Function


'@DefaultMember
Public Function Deb(ByVal ipReference As Variant) As cmpLT
Attribute Deb.VB_UserMemId = 0
    With New cmpLT
        Set Deb = .ConstructInstance(ipReference)
    End With
End Function


Friend Function ConstructInstance(ByVal ipReference As Variant) As cmpLT

    If VarType(ipReference) = vbBoolean Then
        Err.Raise "cmpLT", _
        "LT is not a valid comparision for booleans"
    End If
    
    s.Reference = GetItemAsComparerValue(ipReference)
    s.RefType = VBA.VarType(s.Reference)
    Set ConstructInstance = Me
    
End Function
