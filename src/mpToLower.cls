VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpToLower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder("VBALib.Mappers")
    
Implements IMapper

Const MY_TYPENAME                   As String = "mpToLower"


Private Property Get IMapper_Typename() As String
    IMapper_Typename = MY_TYPENAME
End Property

Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
        
    If VBA.VarType(ipHostItem) <> vbString Then
        IMapper_ExecMapper = Array(ipHostItem)
    Else
        IMapper_ExecMapper = Array(VBA.LCase$(ipHostItem))
    End If
        
End Function

'@DefaultMember
Public Function Deb() As mpToLower
Attribute Deb.VB_UserMemId = 0
    With New mpToLower
        Set Deb = .ConstructInstance
    End With
End Function

Friend Function ConstructInstance() As mpToLower
    Set ConstructInstance = Me
End Function
