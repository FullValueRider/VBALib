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
'@Folder("Reducers")
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

'@Description("Returns a decimal containong the sum of the Items.  An optional comparer may be used to emulate sumIf")
Public Function ExecReduction(ByRef ipHostItems As Variant) As Variant
Attribute ExecReduction.VB_Description = "Returns a decimal containong the sum of the Items.  An optional comparer may be used to emulate sumIf"

    Dim mySum As Variant:  mySum = VBA.CDec(0)
    
    Dim myItems As IterItems: Set myItems = IterItems(ipHostItems)
    Do
        Dim myItem As Variant: myItem = Array(myItems.CurItem(0))
        If GroupInfo.IsNumber(myItem(0)) Then
            If s.Comparer Is Nothing Then
                mySum = mySum + myItem(0)
            Else
                If s.Comparer.ExecCmp(myItem(0)) Then
                    mySum = mySum + myItem(0)
                End If
            End If
        End If
                
    Loop While myItems.MoveNext

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
