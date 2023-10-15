VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Comparers")
'@ModuleDescription("True if the ipHostitem is within the low/high bounds")Option Explicit
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

    IComparer_ExecCmp = False
    
    If s.Reference(m_Low).ExecCmp(ipHostItem) Then
        Exit Function
    End If
    
    If s.Reference(m_High).ExecCmp(ipHostItem) Then
        Exit Function
    End If
    
    IComparer_ExecCmp = True
    
End Function

Private Function IComparer_Reference(ParamArray ipLowHighBounds() As Variant) As IComparer
    'At initialisation cmpIntake two values which are encapsulated
    ' into the Low and High comparers, hence we need to acces the reference method
    ' for the IComparers stored in s.Reference
    s.Reference(m_Low).Reference ipLowHighBounds(m_Low)
    s.Reference(m_High).Reference ipLowHighBounds(m_High)
    Set IComparer_Reference = Me
End Function

'@DefaultMember
' See comments in cmpEQ
Public Function Deb(ParamArray ipLowHighBounds() As Variant) As cmpIn
Attribute Deb.VB_UserMemId = 0
    With New cmpIn
        Set Deb = .ConstructInstance(VBA.CVar(ipLowHighBounds))
    End With
End Function


Friend Function ConstructInstance(ByRef ipLowHighBounds As Variant) As cmpIn
    s.Reference = Array(cmpLT(ipLowHighBounds(e_RangeBounds.m_Low)), cmpMT(ipLowHighBounds(e_RangeBounds.m_High)))
    Set ConstructInstance = Me
End Function
