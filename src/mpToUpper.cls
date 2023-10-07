VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpToUpper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder("Mappers")
    
Implements IMapper

Const MY_TYPENAME                   As String = "mpToUpper"

'Private Function ExecMapper(ByRef ipHostItem As Variant) As Collection
'    Set IMapper_ExecMapper = ExecMapper(ipHostItem)
'End Function


Private Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property


Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
        
    If GroupInfo.IsString(ipHostItem) Then
        '@Ignore UntypedFunctionUsage
        IMapper_ExecMapper = Array(VBA.UCase(ipHostItem))
    Else
        IMapper_ExecMapper = Array(ipHostItem)
    End If
End Function


Public Function Deb() As mpToUpper
    With New mpToUpper
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As mpToUpper
    Set ConstructInstance = Me
End Function


'Public Function TypeName() As String
'    TypeName = MY_TYPENAME
'End Function
