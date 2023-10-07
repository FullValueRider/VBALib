VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpNotIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Comparers")
Option Explicit

Implements IComparer

Const MY_TYPENAME                               As String = "cmpIn"

Private Type State
    Low                                         As IComparer
    High                                        As IComparer
End Type


Private s                                       As State


Private Property Get IComparer_TypeName() As String
    IComparer_TypeName = MY_TYPENAME
End Property


Private Function IComparer_ExecCmp(ByRef ipHostItem As Variant) As Boolean
    IComparer_ExecCmp = True
    If s.Low.ExecCmp(ipHostItem) Then
        Exit Function
    End If
    
    If s.High.ExecCmp(ipHostItem) Then
        Exit Function
    End If
    
    IComparer_ExecCmp = False
    
End Function

Private Sub IComparer_Reference(ParamArray ipLowHighRefs() As Variant)
    s.Low.Value = ipLowHighRefs(0)
    s.High.Value = ipLowHighRefs(1)
End Sub

'@DefaultMember
Public Function Deb(ParamArray ipLowHighRefs() As Variant) As cmpIn
Attribute Deb.VB_UserMemId = 0
    With New cmpIn
        Set Deb = .ConstructInstance(VBA.CVar(ipLowHighRefs))
    End With
End Function


Friend Function ConstructInstance(ByRef ipLowHighRefsArray As Variant) As cmpNotIn
    Set s.Low = cmpLT(ipLowHighRefsArray(0))
    Set s.High = cmpMT(ipLowHighRefsArray(1))
    Set ConstructInstance = Me
End Function
