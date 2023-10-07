VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fnCmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Reducers")
Option Explicit

Implements IFold

Private Const MY_TYPENAME                               As String = "fnCmp"

'@DefaultMember
Public Function Deb() As fnCmp
Attribute Deb.VB_UserMemId = 0
    With New fnCmp
        Set Deb = .ConstructInstance
    End With
End Function

Friend Function ConstructInstance() As fnCmp
    Set ConstructInstance = Me
End Function

Private Function IFold_ExecFold(ByVal ipAccumulator As Variant, ByRef ipHostItem As Variant) As Variant

    If ipAccumulator.ExecCmp(ipHostItem) Then
        ipAccumulator = ipHostItem
    End If
    
    IFold_ExecFold = Array(ipAccumulator)
    
End Function


Private Property Get IFold_TypeName() As String
    IFold_TypeName = MY_TYPENAME
End Property
