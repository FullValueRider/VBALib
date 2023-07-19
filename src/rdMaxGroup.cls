VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdMaxGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Reducers")
'@IgnoreModule
Option Explicit

'Implements IReducer

'Private Type State
'    Comparer                        As IComparer
'End Type
'
'Private s                           As State
'
'    Const MY_TYPENAME                  As String = "rdMaxGroup"
'
'    Private Type State
'        Index                           As Variant
'        Max                             As Variant
'        IndexOfMax                      As Variant
'    End Type
'
'    Private s                           As State
'
'    '@DefaultMember
'    Public Function Deb(ByVal ipIndex As Variant = Empty) As rdMaxGroup
'        With New rdMaxGroup
'            Return .ConstructInstance(ipIndex)
'        End With
'    End Function
'
'    Public Function ConstructInstance(ByRef ipIndex As Variant) As rdMaxGroup
'        s.Max = Empty
'        s.Index = ipIndex
'        Return Me
'    End Function
'
'    Public Function ExecReduction(ByRef ipHostItems As Variant) As Variant Implements IReducer.ExecReduction
'
'
'        ' hostitems will be some sort of grouped values
'        Dim myI As IterItems = IterItems(ipHostItems)
'        Do
'                DoEvents
'            Dim myItem As Variant = myI.GetItem
'
'            If VBA.IsEmpty(s.Max) Then
'                'Sys.Assign s.Max, myItem
'
'                If VBA.IsObject(myItem) Then
'                    Set s.Max = myItem
'
'                Else
'                        s.Max = myItem
'
'                End If
'                Set s.IndexOfMax = 1
'                Continue Do
'            End If
'
'            If Var.IsString(myItem) Then
'
'                s.Max = IIf(VBA.Mid(myItem, s.Index, 1) > VBA.Mid(s.Max, s.Index, 1), s.Max, myItem)
'                s.IndexOfMax = IIf(VBA.Mid(myItem, s.Index, 1) > VBA.Mid(s.Max, s.Index, 1), s.IndexOfMax, myI.Key(0))
'            Else
'                If VBA.IsObject(myItem) Then
'                    Set s.Max = IIf(s.Max(s.Index) > myItem(s.Index), s.Max, myItem)
'
'                Else
'                        s.Max = IIf(s.Max(s.Index) > myItem(s.Index), s.Max, myItem)
'                End If
'                 s.IndexOfMax = IIf(s.Max(s.Index) > myItem(s.Index), s.IndexOfMax, myI.Key(0))
'            End If
'
'        Loop While myI.MoveNext
'
'        Return Array(s.IndexOfMax, s.Max)
'
'    End Function
'
'
'    Public Property Get TypeName() As String Implements IReducer.TypeName
'        Return MY_TYPENAME
'    End Property


