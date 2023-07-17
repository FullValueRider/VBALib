VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdInner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Reducers")
Option Explicit
 
Implements IReducer

'Private s                               As State
    
Private Const MY_TYPENAME               As String = "rdInner"


Private Type State
    Reducer                 As IReducer
End Type

Private s                   As State


Public Property Get IReducer_TypeName() As String
    IReducer_TypeName = MY_TYPENAME
End Property

 '@Description("Calls Mapit on the lowest level item supporting mapit on The hostItem")
Public Function IReducer_ExecReduction(ByRef ipHostItems As Variant) As Variant
Attribute IReducer_ExecReduction.VB_Description = "Calls Mapit on the lowest level item supporting mapit on The hostItem"

    Dim myItem As Variant
    For Each myItem In ipHostItems
    
        Dim myName As String
        myName = Shortname(myItem)
        
        Select Case myName
        
            Case "seq"
            
                IReducer_ExecReduction = myItem.ReduceIt(Me)
                
            Case "kvp"
            
                IReducer_ExecReduction = SeqC(myItem.Items).ReduceIt(Me)
                
            Case Else
            
                IReducer_ExecReduction = s.Reducer.ExecReduction(ipHostItems)
                
        End Select
        
    Next
        
End Function

'@DefaultMember
Public Function Deb(ByVal ipReducer As IReducer) As rdInner
Attribute Deb.VB_UserMemId = 0
    With New rdInner
        Set Deb = .ConstructInstance(ipReducer)
    End With
End Function

Friend Function ConstructInstance(ByVal ipReducer As IReducer) As rdInner
    Set s.Reducer = ipReducer
    Set ConstructInstance = Me
End Function

' Returns the first three characters of the type name, avoids issues with SeqCA, SeqCC, SeqCT, SeqCH, SeqCL etc
Public Function Shortname(ByRef ipItem As Variant) As String
    Shortname = VBA.LCase$(VBA.Mid$(VBA.TypeName(ipItem), 1, 3))
End Function
