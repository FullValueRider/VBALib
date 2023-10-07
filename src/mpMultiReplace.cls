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


Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant

    If GroupInfo.IsNotString(ipHostItem) Then
        IMapper_ExecMapper = Array(ipHostItem)
    Else
        IMapper_ExecMapper = Array(Strs.MultiReplacer(VBA.CStr(ipHostItem), s.FindReplaceItems))
    End If
        
End Function



'@DefaultMember
Public Function Deb(ParamArray ipFindReplaceItems() As Variant) As mpMultiReplace
Attribute Deb.VB_UserMemId = 0
    With New mpMultiReplace
        Set Deb = .ConstructInstance(VBA.CVar(ipFindReplaceItems))
    End With
End Function


Friend Function ConstructInstance(ByRef ipFindReplaceItems As Variant) As mpMultiReplace
        
    If ArrayOp.HoldsOneItem(ipFindReplaceItems) Then
        If GroupInfo.IsItemByKey(ipFindReplaceItems(0)) Then
            Set s.FindReplaceItems = ipFindReplaceItems(0)
        Else
            s.FindReplaceItems = ipFindReplaceItems(0)
        End If
    Else
        Set s.FindReplaceItems = ArrayOp.Splat(ipFindReplaceItems)
    End If
      
    Set ConstructInstance = Me
        
End Function
