VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdMinNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Reducers")
Option Explicit

Implements IReducer

Const MY_TYPENAME                   As String = "rdMinNum"


Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property


Public Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant
    
    Dim myMin As Variant
    myMin = Globals.maxDecimal
    
    Dim myItems As IterItems: Set myItems = IterItems(ipHostItems)
    Do
    
        If myMin > myItems.CurItem(0) Then
            myMin = myItems.CurItem(0)
        End If
                
    Loop While myItems.MoveNext

    IReducer_ExecReduction = myMin
    
End Function


'@DefaultMember
Public Function Deb() As rdMinNum
Attribute Deb.VB_UserMemId = 0
    With New rdMinNum
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As rdMinNum
    Set ConstructInstance = Me
End Function
