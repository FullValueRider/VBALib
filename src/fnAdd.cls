VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "fnAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Reducers")
Option Explicit

Implements IFold

Private Const MY_TYPENAME                               As String = "fnAdd"

'@DefaultMember
Public Function Deb() As fnAdd
Attribute Deb.VB_UserMemId = 0
    With New fnAdd
        Set Deb = .ConstructInstance
    End With
End Function

Friend Function ConstructInstance() As fnAdd
    Set ConstructInstance = Me
End Function

Private Function IFold_ExecFold(ByRef ipAccumulator As Variant, ByRef ipHostItem As Variant) As Variant
' Its likely we can expand this function so that + can be applied to non numeeric types
' but this maynot be a good idea compared to a specific function
    IFold_ExecFold = Array(ipAccumulator + ipHostItem)
End Function


Private Property Get IFold_TypeName() As String
    IFold_TypeName = MY_TYPENAME
End Property
