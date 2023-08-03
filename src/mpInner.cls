VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpInner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder("Mappers")

Implements IMapper

Const MY_TYPENAME           As String = "mpInner"

Private Type State
    Mapper                  As IMapper
End Type


Private s                   As State



'Private Function IMapper_ExecMapper(ByRefipHostItem As Variant) As Collection
'    IMapper_ExecMapper = ExecMapper(ipHostItem)
'End Function


Private Property Get IMapper_Typename() As String
    IMapper_Typename = MY_TYPENAME
End Property


'@Description("if the Hostitem is a SeqC or kvp then we reapply mapit with the mpInner instance")
Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
Attribute IMapper_ExecMapper.VB_Description = "if the Hostitem is a SeqC or kvp then we reapply mapit with the mpInner instance"
    ' myItem allows easier debugging inspection
    Dim myItem As Variant
    Select Case GroupInfo.Id(ipHostItem)
        
        Case m_Dictionary, m_List
            
            ' mapit returns a SeqC not a single item array so we need set
            ' An mpInner instance was passed in by the original call to mapit
            ' so we don't need to reapply mpInner in the call to mapit below
            ' e.g. using .Mapit(mpInner(s.mapper))
            ' we can just use the Me instance
            Set myItem = ipHostItem.MapIt(Me)
            IMapper_ExecMapper = Array(myItem)
            
        Case Else
            ' Execmapper returns a single item array
            myItem = s.Mapper.ExecMapper(ipHostItem)
            IMapper_ExecMapper = Array(myItem(0))
        
    End Select
    
End Function


'@DefaultMember
Public Function Deb(ByVal ipMapper As IMapper) As mpInner
Attribute Deb.VB_UserMemId = 0
    With New mpInner
        Set Deb = .ConstructInstance(ipMapper)
    End With
End Function


Friend Function ConstructInstance(ByRef ipMapper As IMapper) As mpInner
    Set s.Mapper = ipMapper
    Set ConstructInstance = Me
End Function
