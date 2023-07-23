VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Returns the sum of the numbers in the SeqC.  Non-numbers are ignored"
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Reducers")
'@ModuleDescription("Returns the sum of the numbers in the SeqC.  Non-numbers are ignored")
Option Explicit

Implements IReducer

Const MY_TYPENAME                   As String = "rdSum"

Private Type State
    Comparer                        As IComparer
End Type

Private s                           As State


'@Description("Computes the sum of the items")
Private Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant
Attribute IReducer_ExecReduction.VB_Description = "Computes the sum of the items"
    IReducer_ExecReduction = ExecReduction(ipHostItems)
End Function

Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property


Public Function ExecReduction(ByRef ipHostItems As Variant) As Variant

    Dim mySum As Variant
    mySum = VBA.CDec(0)
    
    'Dim myItem As Variant
    Dim myIndex As Long
    If s.Comparer Is Nothing Then
    
        For myIndex = ipHostItems.FirstIndex To ipHostItems.LastIndex
            If GroupInfo.IsNumber(ipHostItems.Item(myIndex)) Then
                mySum = mySum + ipHostItems.Item(myIndex)
            End If
        Next
       
    Else
        
        For myIndex = ipHostItems.FirstIndex To ipHostItems.LastIndex
            If s.Comparer.ExecCmp(ipHostItems.Item(myIndex)) Then
                mySum = mySum + ipHostItems.Item(myIndex)
            End If
        Next
        
    End If

    ExecReduction = mySum

End Function

Public Function Deb(Optional ByRef ipComparer As IComparer = Nothing) As rdSum
    With New rdSum
        Set Deb = .ConstructInstance(ipComparer)
    End With
End Function

Friend Function ConstructInstance(ByRef ipComparer As IComparer) As rdSum

    Set s.Comparer = ipComparer
    Set ConstructInstance = Me

End Function
