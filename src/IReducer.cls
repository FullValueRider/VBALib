VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IReducer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Interface
'@Folder("Interfaces")
Option Explicit


' IReducers take the whole SeqC or Items of a Kvp so we can be specific and specify SeqC as the parameter type
Public Function ExecReduction(ByRef ipForEach As Variant) As Variant
End Function


Public Property Get TypeName() As String
End Property


