VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpCharsSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
'@PredeclaredId
'@Folder("Mappers")
'@ModuleDescription"Splits a string to a SeqC of characters)
Option Explicit

Implements IMapper

Const MY_TYPENAME               As String = "mpCharsSeq"

Private Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property


Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant

    If GroupInfo.IsString(ipHostItem) Then
        IMapper_ExecMapper = Array(SeqC(ipHostItem))
    Else
        IMapper_ExecMapper = Array(ipHostItem)
    End If
        
End Function


'@DefaultMember
Public Function Deb() As mpCharsSeq
Attribute Deb.VB_UserMemId = 0
    With New mpCharsSeq
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As mpCharsSeq
    Set ConstructInstance = Me
End Function
