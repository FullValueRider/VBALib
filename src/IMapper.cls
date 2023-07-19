VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
Option Explicit

'@Interface
Public Function ExecMapper(ByRef ipHostItem As Variant) As Variant
End Function
 
Public Property Get TypeName() As String
End Property
