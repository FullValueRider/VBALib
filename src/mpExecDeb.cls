VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpExecDeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Mappers")
Implements IMapper

Const MY_TYPENAME                       As String = "mpDeb"

Private Type State
    PredeclaredId                       As Object
End Type

Private s                               As State

'@DefaultMember
Public Function Deb(ByVal ipPredeclaredId As Object) As mpExecDeb
    With New mpExecDeb
        Set Deb = .ConstructInstance(ipPredeclaredId)
    End With
End Function


Friend Function ConstructInstance(ByVal ipPredeclaredId As Object) As mpExecDeb
        
    Set s.PredeclaredId = ipPredeclaredId
    Set ConstructInstance = Me
        
End Function

Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
    IMapper_ExecMapper = Array(s.PredeclaredId.Deb(ipHostItem))
End Function

Private Property Get IMapper_Typename() As String
    IMapper_Typename = MY_TYPENAME
End Property

'Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Collection
'    Set IMapper_ExecMapper = ExecMapper(ipHostItem)
'End Function



