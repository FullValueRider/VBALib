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

Friend Function Clone() As IAutoKey
End Function

Public Function GetCurrentKey() As Variant
End Function

Friend Function SetCurrentKey(ByVal Value As Variant) As IAutoKey
End Function

Public Function GetFirstUseKey() As Variant
End Function

Public Function SetFirstUseKey(ByVal ipValue As Variant) As IAutoKey
End Function

Public Property Get FirstUseKeyType() As String
End Property

Public Function NextKeys(ByVal ipCount As Long) As Lyst
End Function

Public Function NextKey() As Variant
End Function

Public Property Get Name() As String
End Property






