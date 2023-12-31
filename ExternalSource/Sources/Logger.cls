VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib")

Private Const Location      As String = "VBALib.Logger."
Private Type State
    
    Info                    As Stack
    
End Type

Private s                   As State


    
Private Sub Class_Initialize()
        Guard ResultCode.InvalidNew, Types.IsNotSameInstance(Me, Logger), Location & "Class_Initialize"
End Sub

Public Sub Push(ParamArray ipArgs() As Variant)
    
    s.Info.Push Types.Iterable.ToLyst(VBA.CVar(ipArgs))
    
End Sub

Public Function Pull() As Variant
    
    Set Pull = s.Info.Pop
    
End Function

Public Sub Remove()
    
    s.Info.Pop
    
End Sub

Public Function Peek() As Variant
    
    Set Peek = s.Info.Peek
    
End Function
   
