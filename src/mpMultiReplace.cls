VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpMultiReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Mappers")

Implements IMapper

Const MY_TYPENAME                   As String = "mpMultiReplace"

Private Type State

    FindReplaceItems                As Variant
        
End Type


Private s                           As State


Private Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property


Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant

    If GroupInfo.IsNotString(ipHostItem) Then
        IMapper_ExecMapper = Array(ipHostItem)
    Else
        IMapper_ExecMapper = Array(Strs.MultiReplacer(VBA.CStr(ipHostItem), s.FindReplaceItems))
    End If
        
End Function


'@Description("Takes a jagged array of find/replacepairs")
'@DefaultMember
Public Function Deb(ByRef ipFindReplaceItems As Variant) As mpMultiReplace
Attribute Deb.VB_Description = "Takes a jagged array of find/replacepairs"
Attribute Deb.VB_UserMemId = 0
    With New mpMultiReplace
        Set Deb = .ConstructInstance(ipFindReplaceItems)
    End With
End Function


Friend Function ConstructInstance(ByRef ipFindReplaceItems As Variant) As mpMultiReplace
        
    s.FindReplaceItems = ipFindReplaceItems
    Set ConstructInstance = Me
        
End Function
