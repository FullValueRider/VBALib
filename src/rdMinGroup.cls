VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdMinGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Reducers")
'@IgnoreModule
Option Explicit
 
'Implements IReducer
'
'Private Type State
'    Comparer                        As IComparer
'End Type
'
'Private s                           As State
'
'    Const MY_TYPENAME                  As String = "rdMinGroup"
'
'    Private Type State
'        Index                           As Variant
'        Min                             As Variant
'        IndexOfMin                      As Variant
'    End Type
'
'    Private s                           As State
'
'    '@DefaultMember
'    Public Function Deb(ByVal ipIndex As Variant = Empty) As rdMinGroup
'        With New rdMinGroup
'            Return .ConstructInstance(ipIndex)
'        End With
'    End Function
'
'    Public Function ConstructInstance(ByRef ipIndex As Variant) As rdMinGroup
'        s.Min = Empty
'        s.Index = ipIndex
'        Return Me
'    End Function
'
'    Public Function ExecReduction(ByRef ipHostItems As Variant) As Variant Implements IReducer.ExecReduction
'
'        ' hostitems will be some sort of grouped values
'        Dim myI As IterItems = IterItems(ipHostItems)
'        Do
'                DoEvents
'            Dim myItem As Variant = myI.GetItem
'
'            If VBA.IsEmpty(s.Min) Then
'                'Sys.Assign s.Min, myItem
'
'                If VBA.IsObject(myItem) Then
'                    Set s.Min = myItem
'
'                Else
'                        s.Min = myItem
'
'                End If
'                s.IndexOfMin = 1
'                 Continue Do
'            End If
'
'            If Var.IsString(myItem) Then
'
'                s.Min = IIf(VBA.Mid(s.Min, s.Index, 1) < VBA.Mid(myItem, s.Index, 1), s.Min, myItem)
'                s.IndexOfMin = IIf(VBA.Mid(s.Min, s.Index, 1) < VBA.Mid(myItem, s.Index, 1), s.IndexOfMin, myI.Key(0))
'            Else
'                If VBA.IsObject(myItem) Then
'                    Set s.Min = IIf(s.Min(s.Index) < myItem(s.Index), s.Min, myItem)
'                Else
'                        s.Min = IIf(s.Min(s.Index) < myItem(s.Index), s.Min, myItem)
'                End If
'                s.IndexOfMin = IIf(s.Min(s.Index) < myItem(s.Index), s.IndexOfMin, myI.Key(0))
'            End If
'
'
'        Loop While myI.MoveNext
'
'        Return Array(s.IndexOfMin, s.Min)
'
'    End Function
'
'
'    Public Property Get TypeName() As String Implements IReducer.TypeName
'        Return MY_TYPENAME
'    End Property


