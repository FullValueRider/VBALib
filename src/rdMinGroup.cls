VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdMinGroup"
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

Const MY_TYPENAME                               As String = "rdMinGroup"

Private Type State
    'Comparer                                    As IComparer
    Index                                       As Variant
    Min                                         As Variant
    'IndexOfMin                                  As Variant
End Type

Private s                                       As State

'@DefaultMember
Public Function Deb(Optional ByRef ipIndex As Variant = Empty) As rdMinGroup
Attribute Deb.VB_UserMemId = 0
    With New rdMinGroup
        Set Deb = .ConstructInstance(ipIndex)
    End With
End Function

Public Function ConstructInstance(ByRef ipIndex As Variant) As rdMinGroup
    s.Min = Empty
    s.Index = ipIndex
    Set ConstructInstance = Me
End Function

Private Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant

    ' hostitems will be some sort of grouped values
    Dim myI As IterItems: Set myI = IterItems(ipHostItems)
    Do
        DoEvents
        Dim myItem As Variant: myItem = myI.CurItem(0)

        If VBA.IsEmpty(s.Min) Then
            'Sys.Assign s.Min, myItem

            If VBA.IsObject(myItem) Then
                Set s.Min = myItem

            Else
                    s.Min = myItem

            End If
            's.IndexOfMin = 1
            GoTo ContinueDo:
        End If

        If GroupInfo.IsString(myItem) Then

            '@Ignore UntypedFunctionUsage
            '@Ignore UntypedFunctionUsage
            s.Min = IIf(VBA.Mid(s.Min, s.Index, 1) < VBA.Mid(myItem, s.Index, 1), s.Min, myItem)
            's.IndexOfMin = IIf(VBA.Mid(s.Min, s.Index, 1) < VBA.Mid(myItem, s.Index, 1), s.IndexOfMin, myI.Key(0))
        Else
            If VBA.IsObject(myItem) Then
                Set s.Min = IIf(s.Min(s.Index) < myItem(s.Index), s.Min, myItem)
            Else
                s.Min = IIf(s.Min(s.Index) < myItem(s.Index), s.Min, myItem)
            End If
            's.IndexOfMin = IIf(s.Min(s.Index) < myItem(s.Index), s.IndexOfMin, myI.Key(0))
        End If

ContinueDo:
    Loop While myI.MoveNext

    IReducer_ExecReduction = Array(s.Min)

End Function


Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property
