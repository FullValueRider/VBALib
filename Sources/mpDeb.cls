VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpDeb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Mappers")
Implements IMapper
    
Private Type State
        
    PredeclaredId                       As Object
        
        
End Type

Private s                               As State

Public Function Deb(ByRef ipPredeclaredId As Object) As mpDeb
        
    With New mpDeb
        
        Set Deb = .ReadyToUseInstance(ipPredeclaredId)
            
    End With
        
End Function

Friend Function ReadyToUseInstance(ByRef ipPredeclaredId As Object) As mpDeb
        
    Set s.PredeclaredId = ipPredeclaredId
    Set ReadyToUseInstance = Me
        
End Function

'@Description("Adds the value in ipArgsArray(0) to each item in the Lyst.  If the host item is not a number the string NAN is returned")
Private Function IMapper_ExecMapper(ByVal ipHostItem As Variant) As Variant
Attribute IMapper_ExecMapper.VB_Description = "Adds the value in ipArgsArray(0) to each item in the Lyst.  If the host item is not a number the string NAN is returned"
    Set IMapper_ExecMapper = ExecMapper(ipHostItem)
End Function

Public Function ExecMapper(ByVal ipHostItem As Variant) As Variant
        
    Set ExecMapper = s.PredeclaredId.Deb(ipHostItem)
       
        
End Function
