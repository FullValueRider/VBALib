VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAutoKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit
'@Folder("AutoKeys")
'@Exposed
'@Interface

Public Function Clone() As IAutoKey
End Function

Public Property Get CurrentKey() As Variant
End Property

Public Property Let CurrentKey(ByVal Value As Variant)
End Property

Public Property Get FirstUseKey() As Variant
End Property

Public Property Let FirstUseKey(ByVal ipValue As Variant)
End Property

Public Property Get FirstUseKeyType() As String
End Property

Public Function GetKeySequence(ByVal ipCount As Long) As Lyst
End Function

Public Function GetNextKey() As Variant
End Function

Public Function Name() As String
End Function

Public Function Res() As Result
End Function





