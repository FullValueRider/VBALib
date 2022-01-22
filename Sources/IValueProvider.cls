VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IValueProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'https://stackoverflow.com/questions/52261264/implement-python-like-generator-using-custom-enumerator-in-vba

' IValueProvider: Provide values.
Option Explicit


' Return True if there are more values
'@Ignore EmptyMethod
Public Function HasMore() As Boolean
End Function

' Return the next value
'@Ignore EmptyMethod
Public Function GetNext() As Variant
End Function
