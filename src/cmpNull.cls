VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cmpNull"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Returns false if the input item is null"
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Comparers")
'@ModuleDescription("Returns false if the input item is null")
Implements IComparer

Const MY_TYPENAME                   As String = "cmpNull"

'Private Type State
'    Reference                       As Variant
'    RefType                         As VbVarType
'End Type
'
'Private s                           As State
'

Private Property Get IComparer_TypeName() As String
    IComparer_TypeName = MY_TYPENAME
End Property

'@Description("Returns false if the input is null. A bit counterintuitive but works well with the filterIt method for removing null values")
Private Function IComparer_ExecCmp(ByRef ipHostItem As Variant) As Boolean
Attribute IComparer_ExecCmp.VB_Description = "Returns false if the input is null. A bit counterintuitive but works well with the filterIt method for removing null values"
    
      IComparer_ExecCmp = Not VBA.IsNull(ipHostItem)
       
End Function


'@DefaultMember
Public Function Deb() As cmpNull
Attribute Deb.VB_UserMemId = 0
    With New cmpNull
        Set Deb = .ConstructInstance()
    End With
End Function


Friend Function ConstructInstance() As cmpNull

'    If VBA.IsNull(ipReference) Then
'        Err.Raise 17 + vbObjectError, _
'            "VBALib.cmpNull", _
'            "Good grief!!!, The replacement value for NUll - it really really really shouldn't be NUll"
'    End If
    
    Set ConstructInstance = Me
    
End Function
