VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdFoldLeft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Filters")
Option Explicit

' The rdFoldLeft class is an attempt at replicating the Haskell Foldl function
' https://www.davidculley.com/blog/haskells-foldl-and-foldr-explained/
' In Haskell terms consider the list [1,2,3,4,5]
' rdFoldRight( f acc) means
' (1 f (2 f (3 f (4 f acc)))) or (f (f (f (f acc 1) 2) 3) 4
' and rdFoldLeft( f acc ) means
' ((((acc f 1) f 2) f 3) f 4) or f 1 (f 2 (f 3 (f 4 acc)))

' in these VBA versions we also allow
' an optional limit for the acc
' and an optional filter for the individual items
' optionally to specify the subset of the list on which to operate (from To Step)

Implements IReducer

Const MY_TYPENAME                                   As String = "rdFoldLeft"

Private Type State
    Accumulator                                     As Variant
    FoldLeft                                        As IFold
    AccComparer                                     As IComparer
    ItemComparer                                    As IComparer
    From                                            As Long
    To                                              As Long
    Step                                            As Long
End Type

Private s                                           As State

'@DefaultMember
Public Function Deb _
( _
    ByRef ipFold As IFold, _
    ByRef ipAccumulator As Variant, _
    ByRef ipAccCmp As IComparer, _
    ByRef ipItemCmp As IComparer, _
    Optional ByRef ipFrom As Long = -1, _
    Optional ByRef ipTo As Long = -1, _
    Optional ByRef ipStep As Long = 1 _
) As rdFoldLeft
Attribute Deb.VB_UserMemId = 0
    With New rdFoldLeft
        Set Deb = .ConstructInstance(ipFold, ipAccumulator, ipAccCmp, ipItemCmp, ipFrom, ipTo, ipStep)
    End With
End Function


Friend Function ConstructInstance _
( _
    ByRef ipFold As IFold, _
    ByRef ipAccumulator As Variant, _
    ByRef ipAccCmp As IComparer, _
    ByRef ipItemCmp As IComparer, _
    ByRef ipFrom As Long, _
    ByRef ipTo As Long, _
    ByRef ipStep As Long _
) As rdFoldLeft
        
    Set s.FoldLeft = ipFold
    
    If VBA.IsObject(s.Accumulator) Then
        Set s.Accumulator = ipAccumulator
    Else
        s.Accumulator = ipAccumulator
    End If
    
    Set s.AccComparer = ipAccCmp
    Set s.ItemComparer = ipItemCmp
    
    Guard.MustBeAtLestStartOrMinusOne ipFrom, 1, "rdFoldLeft", " "
    s.From = ipFrom
    
    Guard.MustBeAtLeastStartOrMinusOne ipTo, s.From, "rdFoldLeft", " "
    s.To = ipTo
    
    Guard.MustBeAtLeastOne ipTo, "rdFoldLeft", " "
    s.Step = ipStep
    
    Set ConstructInstance = Me
    
End Function

Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property

Private Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant

    Dim myItems As IterItems: Set myItems = IterItems(ipHostItems)
    
    If s.To = -1 Then
        s.To = myItems.LastIndex
    End If
    
    myItems.FTS s.From, s.To, s.Step
    Do
        Dim myA() As Variant: myA = Array(myItems.CurItem(0))
        
        If IsNotNothing(s.ItemComparer) Then
            If Not s.ItemComparer.ExecCmp(myA(0)) Then
                GoTo ContinueDo
            End If
        End If
        
        ' exec fold returns a single item array
        s.Accumulator = s.FoldLeft.ExecFold(s.Accumulator, myA(0))(0)
        If IsNotNothing(s.AccComparer) Then
            If s.AccComparer.ExecCmp(s.Accumulator) Then
                IReducer_ExecReduction = Array(s.Accumulator)
                Exit Function
            End If
        End If
        
ContinueDo:
    Loop While myItems.MoveNext
    
    IReducer_ExecReduction = Array(s.Accumulator)
    
End Function
   







