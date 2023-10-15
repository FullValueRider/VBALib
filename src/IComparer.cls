VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Exposed
'@Interface
'@Folder("Interfaces")
Option Explicit

Public Function ExecCmp(ByRef ipHostItem As Variant) As Boolean
End Function

' The reference method was added to support the foldl and foldr reducers
' which require a method and an accumulator
' To userstand what to provide as a reference consult  Deb method
' of the individual Comparer
Public Function Reference(ParamArray ipReference() As Variant) As IComparer
End Function

Public Property Get TypeName() As String
End Property
