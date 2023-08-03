VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpReplaceNull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Replaces null values with a default value"
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Mappers")
Implements IMapper
'@ModuleDescription("Replaces null values with a default value")
Const MY_TYPENAME                   As String = "mpReplaceNull"

Private Type State
    Default                         As Variant
End Type


Private s                           As State


'@DefaultMember
Public Function Deb(Optional ByRef ipDefault As Variant = "Null") As mpReplaceNull
Attribute Deb.VB_UserMemId = 0
    With New mpReplaceNull
        Set Deb = .ConstructInstance(ipDefault)
    End With
End Function


Friend Function ConstructInstance(ByVal ipDefault As Variant) As mpReplaceNull
        
    s.Default = ipDefault
    Set ConstructInstance = Me
        
End Function


Private Property Get IMapper_Typename() As String
    IMapper_Typename = MY_TYPENAME
End Property


''@Description("")
'Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
'    Set IMapper_ExecMapper = ExecMapper(ipHostItem)
'End Function

'@Description("Replaces null with a default value. The optional default value returned for null is 'Null'")
Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
Attribute IMapper_ExecMapper.VB_Description = "Replaces null with a default value. The optional default value returned for null is 'Null'"
        
    If VBA.IsNull(ipHostItem) Then
        IMapper_ExecMapper = Array(s.Default)
    Else
        IMapper_ExecMapper = Array(ipHostItem)
    End If
        
End Function
