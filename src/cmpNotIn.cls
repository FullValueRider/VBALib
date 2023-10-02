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

'@DefaultMember
Public Function Deb(ByRef ipLowValue As Variant, ByRef ipHighValue As Variant) As cmpIn
Attribute Deb.VB_UserMemId = 0
    With New cmpIn
        Set Deb = .ConstructInstance(ipLowValue, ipHighValue)
    End With
End Function


Friend Function ConstructInstance(ByRef ipLowValue As Variant, ByRef ipHighValue As Variant) As cmpNotIn
    Set s.Low = cmpLT(ipLowValue)
    Set s.High = cmpMT(ipHighValue)
    Set ConstructInstance = Me
End Function
