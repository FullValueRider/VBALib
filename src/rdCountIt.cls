VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdCountIt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 '@PredeclaredId
 '@Exposed
 '@Folder("VBALib.Filters")
Option Explicit

Implements IReducer

Const MY_TYPENAME                   As String = "rdCountIt"

Private Type State
    Comparer                        As IComparer
End Type

Private s                           As State

Public Function Deb(ByRef ipComparer As IComparer) As rdCountIt
        
    With New rdCountIt
        
        Set Deb = .ConstructInstance(ipComparer)
            
    End With
        
End Function

Friend Function ConstructInstance(ByRef ipComparer As IComparer) As rdCountIt
        
    Set s.Comparer = ipComparer
    Set ConstructInstance = Me
        
End Function

'@Description("Adds the value in ipArgsArray(0) to each item in the SeqC.  If the host item is not a number the string NAN is returned")
Private Function IReducer_ExecReduction(ByRef ipHostItems As SeqC) As Variant
Attribute IReducer_ExecReduction.VB_Description = "Adds the value in ipArgsArray(0) to each item in the SeqC.  If the host item is not a number the string NAN is returned"
    IReducer_ExecReduction = ExecReduction(ipHostItems)
End Function

Private Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property

Public Function ExecReduction(ByRef ipHostItems As SeqC) As Variant
        
    Dim myCounter As Long
    myCounter = 0
        
   Dim myItem As Variant
   For Each myItem In ipHostItems
   
        If s.Comparer.ExecCmp(myItem) Then
            myCounter = myCounter + 1
        End If
   
   Next
   
   ExecReduction = myCounter
        
End Function
