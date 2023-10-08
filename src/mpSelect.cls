VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@folder("Mappers")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit
    
Implements IMapper


Private Const MY_TYPENAME                       As String = "mpSelect"

Private Type State
    Index                                       As Long
End Type

Private s                                       As State


Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
    
    If GroupInfo.IsString(ipHostItem) Then
        '@Ignore UntypedFunctionUsage
        IMapper_ExecMapper = Array(VBA.Mid(ipHostItem, s.Index, 1))
    Else
        IMapper_ExecMapper = Array(ipHostItem(s.Index))
    End If
        
End Function

'@DefaultMember
Public Function Deb(Optional ByVal ipIndex As Long) As mpSelect
Attribute Deb.VB_UserMemId = 0
    With New mpSelect
        Set Deb = .ConstructInstance(ipIndex)
    End With
End Function

Friend Function ConstructInstance(ByVal ipIndex As Variant) As mpSelect
    
    s.Index = ipIndex
    Set ConstructInstance = Me
        
End Function


Private Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property
