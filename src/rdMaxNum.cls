VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdMaxNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Reducers")
Option Explicit

Implements IReducer

Const MY_TYPENAME                               As String = "rdMaxNum"


Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property


Public Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant
    
    Dim myMax As Variant: myMax = Globals.MinDecimal
    
    Dim myItems As IterItems: Set myItems = IterItems(ipHostItems)
    Do
        If myMax < myItems.CurItem(0) Then
            myMax = myItems.CurItem(0)
        End If
    Loop While myItems.MoveNext

    IReducer_ExecReduction = Array(myMax)
    
End Function


'@DefaultMember
Public Function Deb() As rdMaxNum
Attribute Deb.VB_UserMemId = 0
    With New rdMaxNum
        Set Deb = .ConstructInstance
    End With
End Function


Friend Function ConstructInstance() As rdMaxNum
    Set ConstructInstance = Me
End Function
