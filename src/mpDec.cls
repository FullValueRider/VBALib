VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpDec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Mappers")
Implements IMapper

Const MY_TYPENAME                   As String = "mpDec"

Private Type State
    Decrement                       As Variant
End Type


Private s                           As State


'@DefaultMember
Public Function Deb(Optional ByVal ipDecrement As Variant = 1) As mpDec
Attribute Deb.VB_UserMemId = 0
    With New mpDec
        Set Deb = .ConstructInstance(ipDecrement)
    End With
End Function


Friend Function ConstructInstance(ByVal ipDecrement As Variant) As mpDec
        
    s.Decrement = ipDecrement
    Set ConstructInstance = Me
        
End Function


Private Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property


Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
    
    If GroupInfo.IsNumber(ipHostItem) Then
        IMapper_ExecMapper = Array(ipHostItem - s.Decrement)
    Else
        IMapper_ExecMapper = Array("Nan")
    End If
            
End Function
