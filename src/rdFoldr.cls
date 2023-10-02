VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdFoldr"
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

Const MY_TYPENAME                                   As String = "rdFoldl"

Private Type State
    SetupAccumulator                                As Variant
    FoldFn                                          As IFold
    AccComparer                                     As IComparer
    ItemComparer                                    As IComparer
    SetupFrom                                       As Long
    SetUpTo                                         As Long
    SetUpStep                                       As Long
    
    Accumulator                                     As Long
    From                                            As Long
    To                                              As Long
    Step                                              As Long
End Type

Private s                                           As State

'@DefaultMember
Public Function Deb _
( _
    ByRef ipFold As IFold, _
    Optional ByRef ipAccumulator As Variant = 0, _
    Optional ByRef ipAccCmp As IComparer = Nothing, _
    Optional ByRef ipItemCmp As IComparer = Nothing, _
    Optional ByRef ipFrom As Long = -1, _
    Optional ByRef ipTo As Long = -1, _
    Optional ByRef ipStep As Long = 1 _
) As rdFoldr
Attribute Deb.VB_UserMemId = 0
    With New rdFoldr
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
) As rdFoldr
        
    Set s.FoldFn = ipFold
    
    If VBA.IsObject(s.Accumulator) Then
        Set s.SetupAccumulator = ipAccumulator
    Else
        s.SetupAccumulator = ipAccumulator
    End If
    
    Set s.AccComparer = ipAccCmp
    Set s.ItemComparer = ipItemCmp
    
    ' validation of from,to,step is deredder to exec reduction as
    ' this method is where we know the host properties
    ' rdFoldl may be called many times with different ipHostitems using the same initalisation value
    ' therefore it is important to make a copy of the values that will change during the
    ' application of the reducer to iphost items.  This is why the accumulator,from,to and step have
    ' setup duplicates.
    s.SetupFrom = ipFrom
    s.SetUpTo = ipTo
    s.SetUpStep = ipStep
    
    Set ConstructInstance = Me
    
End Function

Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property

Private Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant

     ' iphostitems may be an array so use first/lastindex provided by Iteritems
    Dim myItems As IterItems: Set myItems = IterItems(ipHostItems)
    
    'Todo, improvide validation of from, to ,step
    s.Accumulator = s.SetupAccumulator
    If s.SetupFrom = -1 Then
        s.From = myItems.FirstIndex
    Else
        s.From = s.From
    End If
    
    If s.SetUpTo = -1 Then
        s.To = myItems.LastIndex
    Else
        s.To = IIf(s.SetUpTo > myItems.LastIndex, myItems.LastIndex, s.SetUpTo)
    End If
    
    If s.SetUpStep < 1 Then
        s.Step = 1
    Else
        s.Step = s.SetUpStep
    End If
    
    'Debug.Print myFrom, myTo, myStep
    myItems.FTS(s.From, s.To, s.Step).MoveToEndIndex
    Do
        Dim myA() As Variant: myA = Array(myItems.CurItem(0))
        
        If IsNotNothing(s.ItemComparer) Then
            If Not s.ItemComparer.ExecCmp(myA(0)) Then
                GoTo ContinueDo
            End If
        End If
        
        ' exec fold returns a single item array
        s.Accumulator = s.FoldFn.ExecFold(s.Accumulator, myA(0))(0)
        If IsNotNothing(s.AccComparer) Then
            If s.AccComparer.ExecCmp(s.Accumulator) Then
                IReducer_ExecReduction = Array(s.Accumulator)
                Exit Function
            End If
        End If
        
ContinueDo:
    Loop While myItems.MovePrev
    
    IReducer_ExecReduction = Array(s.Accumulator)
    
End Function
