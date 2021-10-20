VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("VBALib.Enums")
'@PredeclaredId
'@exposed


Option Explicit

' Public Property Get GuardErrors() As EnumGuardErrors
'     Set GuardErrors = EnumGuardErrors
' End Property

Public Property Get GuardClauses() As enumResultCodes
    Set GuardClauses = enumResultCodes
End Property

Public Property Get TableToLystActions() As EnumTableToLystActions
    Set TableToLystActions = EnumTableToLystActions
End Property
    
