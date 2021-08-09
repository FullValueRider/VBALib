VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Helpers")
'@PredeclaredId
'@Exposed

Option Explicit

Public Property Get MaxLong() As Long
    MaxLong = &H7FFFFFFF
End Property


Public Property Get MinLong() As Long
    MinLong = &HFFFFFFFF
End Property


