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
'@Folder("VBALib.Reducers")
Option Explicit

Implements IReducer

Const MY_TYPENAME                   As String = "rdMinNum"

Public Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property


Public Function IReducer_ExecReduction(ByRef ipHostItems As SeqC) As Variant
    
    Dim myMin As Variant
    myMin = Globals.MaxDecimal
    
    Dim myIndex As Long
    For myIndex = ipHostItems.FirstIndex To ipHostItems.LastIndex
    
        If myMin > ipHostItems.Item(myIndex) Then
            myMin = ipHostItems.Item(myIndex)
        End If
                
    Next

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
