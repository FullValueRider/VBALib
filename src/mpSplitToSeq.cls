VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpSplitToSeq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@ModuleDescription"Splits a string to a SeqC of substrings)
'@Exposed
'@PredeclaredId
'@Folder("Mappers")

Const MY_TYPENAME               As String = "mpSplitToSeq"
Private Type State

    Separator                   As String
        
End Type


Private s                       As State

    
Implements IMapper

'Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
'    Set IMapper_ExecMapper = ExecMapper(ipHostItem)
'End Function


Private Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property


Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant

    If GroupInfo.IsString(ipHostItem) Then
        
        Set IMapper_ExecMapper = Array(SeqC.Deb.AddRange(VBA.Split(ipHostItem, s.Separator)))
            
    Else
            
        IMapper_ExecMapper = Array(ipHostItem)
        
    End If
        
End Function


'@DefaultMember
Public Function Deb(Optional ByVal ipSeparator As String = ",") As mpSplitToSeq
Attribute Deb.VB_UserMemId = 0
    
    With New mpSplitToSeq

        Set Deb = .ConstructInstance(ipSeparator)
            
    End With
    
End Function


Friend Function ConstructInstance(ByVal ipSeparator As String) As mpSplitToSeq
        
    s.Separator = ipSeparator
    Set ConstructInstance = Me
        
End Function
