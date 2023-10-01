VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdCountIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Filters")
Option Explicit

Implements IReducer

Const MY_TYPENAME                                   As String = "rdCountIt"

Private Type State
    Comparer                                        As IComparer
End Type

Private s                                           As State

'@DefaultMember
Public Function Deb(ByRef ipComparer As IComparer) As rdCountIt
Attribute Deb.VB_UserMemId = 0
    With New rdCountIt
        Set Deb = .ConstructInstance(ipComparer)
    End With
End Function


Friend Function ConstructInstance(ByRef ipComparer As IComparer) As rdCountIt
        
    Set s.Comparer = ipComparer
    Set ConstructInstance = Me
        
End Function


Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property

'@Description("Count the occasions where the comparer is true")
Public Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant
Attribute IReducer_ExecReduction.VB_Description = "Count the occasions where the comparer is true"
        
    Dim myCounter As Long: myCounter = 0
    
    Dim myItems As IterItems: Set myItems = IterItems(ipHostItems)
    Do
        Dim myA() As Variant: myA = Array(myItems.CurItem(0))
        Dim myB As Boolean: myB = s.Comparer.ExecCmp(myA(0))
        If myB Then
            myCounter = myCounter + 1
        End If
    Loop While myItems.MoveNext
   
    IReducer_ExecReduction = Array(VBA.CVar(myCounter))
        
End Function
