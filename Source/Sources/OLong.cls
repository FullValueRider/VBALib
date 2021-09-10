VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OLong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Exposed
Option Explicit
'@PredeclaredId
'Exposed

Private Type Properties

    value               As Long
    
End Type

Private p               As Properties

Public Function Deb(ByVal ipValue As Variant) As OLong

    With New OLong
    
        Set Deb = .ReadyToUseInstance(ipValue)
        
    End With
    
End Function

Public Function ReadyToUseInstance(ByVal ipValue As Variant) As OLong

    Set ReadyToUseInstance = Nothing
    On Error Resume Next
    p.value = CLng(ipValue)
    If Err.Number <> 0 Then
    
        On Error GoTo 0
        Exit Function
        
    End If
    
    On Error GoTo 0
    Set ReadyToUseInstance = Me
    
End Function

'@DefaultMember
Public Property Get value() As Long
Attribute value.VB_UserMemId = 0
    value = p.value
End Property

Public Property Let value(ByVal ipValue As Long)
    p.value = ipValue
End Property


Public Function ToString() As String
    ToString = CStr(p.value)
End Function


Public Function ToLyst() As Lyst
    Set ToLyst = Lyst.Deb(Array(p.value))
End Function


Public Function ToArray() As Variant
    ToArray = Array(p.value)
End Function

Public Property Get GetType() As VbVarType
    GetType = vbLong
End Property
