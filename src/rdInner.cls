VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdInner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Reducers")
Option Explicit
 
Implements IReducer
    
Private Const MY_TYPENAME               As String = "rdInner"


Private Type State
    Reducer                             As IReducer
End Type


Private s                               As State


Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property


'@Description "Calls Mapit on the lowest level item supporting mapit on The hostItem"
Private Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant
Attribute IReducer_ExecReduction.VB_Description = "Calls Mapit on the lowest level item supporting mapit on The hostItem"

    ' If the Items of ipHostItems support Mapping then forward to rdInner
    ' else send to reducer
    Dim myItems As IterItems: Set myItems = IterItems(ipHostItems)
    Do
        Dim myA() As Variant: myA = Array(myItems.CurItem(0))
        If GroupInfo.CannotMap(myA(0)) Then
            IReducer_ExecReduction = s.Reducer.ExecReduction(ipHostItems)
            Exit Function
        Else
            ' ipHostItem is an object that supports mapping
            myItems.CurItem(0) = myItems.CurItem(0).ReduceIt(Me)
        End If
    Loop While myItems.MoveNext
        
    IReducer_ExecReduction = Array(ipHostItems)
    
End Function


'@DefaultMember
Public Function Deb(ByVal ipReducer As IReducer) As rdInner
Attribute Deb.VB_UserMemId = 0
    With New rdInner
        Set Deb = .ConstructInstance(ipReducer)
    End With
End Function


Friend Function ConstructInstance(ByVal ipReducer As IReducer) As rdInner
    Set s.Reducer = ipReducer
    Set ConstructInstance = Me
End Function


' Returns the first three characters of the type name, avoids issues with SeqCA, SeqCC, SeqCT, SeqCH, SeqCL etc
Public Function Shortname(ByRef ipItem As Variant) As String
    Shortname = VBA.LCase$(VBA.Mid$(VBA.TypeName(ipItem), 1, 3))
End Function
