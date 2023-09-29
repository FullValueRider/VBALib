VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdMaxGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Reducers")
''@IgnoreModule
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Option Explicit

Implements IReducer

Const MY_TYPENAME                  As String = "rdMaxGroup"

Private Type State
    'Comparer                        As IComparer

    Index                           As Variant
    Max                             As Variant
    'IndexOfMax                      As Variant
End Type

Private s                           As State

'@DefaultMember
Public Function Deb(Optional ByRef ipIndex As Variant = Empty) As rdMaxGroup
Attribute Deb.VB_UserMemId = 0
    With New rdMaxGroup
        Set Deb = .ConstructInstance(ipIndex)
    End With
End Function

Public Function ConstructInstance(ByRef ipIndex As Variant) As rdMaxGroup
    s.Max = Empty
    s.Index = ipIndex
    Set ConstructInstance = Me
End Function

Private Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant


    ' hostitems will be some sort of grouped values
    Dim myI As IterItems: Set myI = IterItems(ipHostItems)
    Do
        DoEvents
        Dim myItem As Variant: myItem = myI.CurItem(0)

        If VBA.IsEmpty(s.Max) Then
            'Sys.Assign s.Max, myItem

            If VBA.IsObject(myItem) Then
                Set s.Max = myItem

            Else
                    s.Max = myItem

            End If
            'Set s.IndexOfMax = 1
            GoTo ContinueDo
        End If

        If GroupInfo.IsString(myItem) Then

            '@Ignore UntypedFunctionUsage
            '@Ignore UntypedFunctionUsage
            s.Max = IIf(VBA.Mid(myItem, s.Index, 1) > VBA.Mid(s.Max, s.Index, 1), s.Max, myItem)
            's.IndexOfMax = IIf(VBA.Mid(myItem, s.Index, 1) > VBA.Mid(s.Max, s.Index, 1), s.IndexOfMax, myI.Key(0))
        Else
            If VBA.IsObject(myItem) Then
                Set s.Max = IIf(s.Max(s.Index) > myItem(s.Index), s.Max, myItem)

            Else
                    s.Max = IIf(s.Max(s.Index) > myItem(s.Index), s.Max, myItem)
            End If
             's.IndexOfMax = IIf(s.Max(s.Index) > myItem(s.Index), s.IndexOfMax, myI.Key(0))
        End If
ContinueDo:
    Loop While myI.MoveNext

    IReducer_ExecReduction = Array(s.Max)

End Function


Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property
