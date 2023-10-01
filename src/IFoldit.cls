VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFoldit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
'@Exposed
'@Folder("Reducers")
Option Explicit

Public Function ExecFold(ByRef ipAccumulator As Variant, ByRef ipHostItem As Variant) As Variant
End Function

Public Property Get TypeName() As String
End Property
