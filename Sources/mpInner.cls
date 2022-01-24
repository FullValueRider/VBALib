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
'@Folder("VBALib.Mappers")

Private Type State

    Mapper                  As Object
        
End Type

Private s                   As State

Implements IMapper

'@Description("if the lyst item is a Kvp or Lyst calls the map method with ipmapper")
Private Function IMapper_ExecMapper(ByVal ipHostItem As Variant) As Variant
Attribute IMapper_ExecMapper.VB_Description = "if the lyst item is a Kvp or Lyst calls the map method with ipmapper"

    Dim myResult As Variant
    myResult = Array(ExecMapper(ipHostItem))
    If VBA.IsObject(myResult(0)) Then
    
        Set IMapper_ExecMapper = ExecMapper(ipHostItem)
    
    Else
    
        IMapper_ExecMapper = ExecMapper(ipHostItem)
    
    End If
    
End Function

Public Function ExecMapper(ByVal ipHostItem As Variant) As Variant

    Dim myResult As Variant
    Select Case TypeName(ipHostItem)
        
        Case "Kvp", "Lyst"
                
            myResult = Array(ipHostItem.MapIt(mpInner.Deb(s.Mapper)))
                    
        Case Else
            
            myResult = Array(s.Mapper.ExecMapper(ipHostItem))
        
    End Select
    
    If VBA.IsObject(myResult(0)) Then
    
        Set ExecMapper = myResult(0)
        
    Else
    
        ExecMapper = myResult(0)
        
    End If
        
End Function

'@DefaultMember
Public Function Deb(ByVal ipMapper As Object) As mpInner
Attribute Deb.VB_UserMemId = 0
    
    With New mpInner
        
        Set Deb = .ReadyToUseInstance(ipMapper)
            
    End With

End Function

Friend Function ReadyToUseInstance(ByRef ipMapper As Object) As mpInner
        
    Set s.Mapper = ipMapper
    Set ReadyToUseInstance = Me
            
End Function
