VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Folder("Interfaces")
Option Explicit

'@Interface


Public Function ExecCmp(ByRef ipHostItem As Variant) As Boolean
End Function


Public Property Get TypeName() As String
End Property


'Public Function ExecCount(ByRef ipHostItems As SeqC) As Long
'End Function
