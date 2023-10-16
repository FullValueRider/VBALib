VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpNotIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "True if the ipHostitem is not within the low/high bounds"
'@PredeclaredId
'@Exposed
'@Folder("Comparers")
'@ModuleDescription("True if the ipHostitem is not within the low/high bounds")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
Option Explicit

Implements IComparer

Const MY_TYPENAME                               As String = "cmpIn"

Private Enum e_RangeBounds
    m_Low = 0
    m_High
End Enum

Private Type State
    Reference()                                 As Variant
End Type


Private s                                       As State


Private Property Get IComparer_TypeName() As String
    IComparer_TypeName = MY_TYPENAME
End Property


Private Function IComparer_ExecCmp(ByRef ipHostItem As Variant) As Boolean

    IComparer_ExecCmp = True
    
    If s.Reference(m_Low).ExecCmp(ipHostItem) Then
        Exit Function
    End If
    
    If s.Reference(m_High).ExecCmp(ipHostItem) Then
        Exit Function
    End If
    
    IComparer_ExecCmp = False
    
End Function

Private Function IComparer_Reference(ParamArray ipLowHighBounds() As Variant) As IComparer
    'At initialisation cmpIntake two values which are encapsulated
    ' into the Low and High comparers, hence we need to acces the reference method
    ' for the IComparers stored in s.Reference
    s.Reference(m_Low).Reference = ipLowHighBounds(m_Low)
    s.Reference(m_High).Reference = ipLowHighBounds(m_High)
    Set IComparer_Reference = Me
End Function

'@DefaultMember
'See comments for cmpEQ
Public Function Deb(ParamArray ipLowHighBounds() As Variant) As cmpNotIn
Attribute Deb.VB_UserMemId = 0
    With New cmpNotIn
        Set Deb = .ConstructInstance(VBA.CVar(ipLowHighBounds))
    End With
End Function


Friend Function ConstructInstance(ByRef ipLowHighBounds As Variant) As cmpNotIn
    s.Reference = Array(cmpLT(ipLowHighBounds(m_Low)), cmpMT(ipLowHighBounds(m_High)))
    Set ConstructInstance = Me
End Function
