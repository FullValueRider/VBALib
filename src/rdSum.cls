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


''@Description("Computes the sum of the items")
'Private Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant
'    IReducer_ExecReduction = ExecReduction(ipHostItems)
'End Function


Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property

'@Description("Returns the sum of the numerical values as a variant containing a Decimal.  An optional comparer may be used to implement sumIf")
Public Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant
Attribute IReducer_ExecReduction.VB_Description = "Returns the sum of the numerical values as a variant containing a Decimal.  An optional comparer may be used to implement sumIf"

    ' RDsum is used to sum any type of number so we use a decimal type as the accumulator
    Dim mySum As Variant:  mySum = VBA.CDec(0)
    
    ' Reducers have also been implemented for arrays so we use the most
    ' generic method we have availbel for enumerating ipHostItems
    Dim myItems As IterItems: Set myItems = IterItems(ipHostItems)
    Do
        Dim myA() As Variant: myA = Array(myItems.CurItem(0))
        If GroupInfo.IsNumber(myA(0)) Then
            If s.Comparer Is Nothing Then
                mySum = mySum + myA(0)
            Else
                If s.Comparer.ExecCmp(myA(0)) Then
                    mySum = mySum + myA(0)
                End If
            End If
        End If
                
    Loop While myItems.MoveNext

    IReducer_ExecReduction = Array(mySum)

End Function

'@DefaultMember
Public Function Deb(Optional ByRef ipComparer As IComparer = Nothing) As rdSum
Attribute Deb.VB_UserMemId = 0
    With New rdSum
        Set Deb = .ConstructInstance(ipComparer)
    End With
End Function


Friend Function ConstructInstance(ByRef ipComparer As IComparer) As rdSum

    Set s.Comparer = ipComparer
    Set ConstructInstance = Me

End Function
