VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("VBALib")

Private Const location      As String = "VBALib.Logger."
Private Type State
    
    Info                    As Stack
    
End Type

Private s                   As state


    
Private Sub Class_Initialize()
        Guard GuardClause.InvalidNew, Types.IsNotSameInstance(Me, Logger), location & "Class_Initialize"
End Sub

Public Sub Push(ParamArray ipArgs() As Variant)
    
    s.Info.push Types.Iterable.ToLyst(VBA.CVar(ipArgs))
    
End Sub

Public Function Pull() As Variant
    
    Set Pull = s.Info.pop
    
End Function

Public Sub Remove()
    
    s.Info.pop
    
End Sub

Public Function Peek() As Variant
    
    Set Peek = s.Info.peek
    
End Function
   
