VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdMaxNum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Reducers")
Option Explicit

Implements IReducer

Const MY_TYPENAME                   As String = "rdMaxNum"


Public Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property


Public Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant
    
    '@Ignore ObsoleteTypeHint
    Dim myMax As Variant
    myMax = Globals.MinDecimal
    
    Dim myIndex As Long
    For myIndex = ipHostItems.FirstIndex To ipHostItems.LastIndex
    
        If myMax < ipHostItems.Item(myIndex) Then
            myMax = ipHostItems.Item(myIndex)
        End If
                
    Next

    IReducer_ExecReduction = myMax
    
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
