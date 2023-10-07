VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpLookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Mappers")
Option Explicit

Implements IMapper

Private Const MY_TYPENAME                                   As String = "mpLookup"

Private Type State
    Lookup                                                  As Variant ' a kvp or dictionary
End Type

Private s                                                   As State

'@Description("return the item in the lookup dictionary")
Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
Attribute IMapper_ExecMapper.VB_Description = "return the item in the lookup dictionary"
    IMapper_ExecMapper = Array(s.Lookup.Item(ipHostItem))
End Function

'@DefaultMember
Public Function Deb(ByVal ipLookup As Variant) As mpLookup
Attribute Deb.VB_UserMemId = 0
    
    If GroupInfo.IsNotItemByKey(ipLookup) Then
        Err.Raise 17 + vbObjectError, "mpLookup", Fmt.Text("Expecting an ItemByKey object. Found '{0}'  ", VBA.TypeName(ipLookup))
    End If
    
    With New mpLookup
        Set Deb = .ConstructInstance(ipLookup)
    End With
    
End Function

Friend Function ConstructInstance(ByVal ipLookup As Variant) As mpLookup
    Set s.Lookup = ipLookup
    Set ConstructInstance = Me
End Function

Public Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property
