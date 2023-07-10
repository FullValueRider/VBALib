VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpSetIf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Mappers")
    
Implements IMapper

Const MY_TYPENAME               As String = "mpSetIf"

Private Type State
  
  Comparer                      As IComparer
  Item                          As Variant
  
End Type


Private s                       As State


'twinbasic requirement
Private Sub Class_Initialize()
  
  If s.Comparer Is Nothing Then
    
    Set s.Comparer = cmpEQ.Deb(0)
    
  End If
  
End Sub


Public Function Deb(Optional ByRef ipComparer As IComparer, Optional ByVal ipReset As Variant = 0) As mpSetIf
    With New mpSetIf
        Set Deb = .ConstructInstance(ipComparer, ipReset)
    End With
End Function

Friend Function ConstructInstance(ByRef ipComparer As IComparer, ByVal ipReset As Variant) As mpSetIf
    
    If s.Comparer Is Nothing Then
      
      Set s.Comparer = cmpEQ.Deb(0)
      
    Else
      
      Set s.Comparer = ipComparer
      
    End If
    
    s.Item = ipReset
    
    Set ConstructInstance = Me
        
End Function

Private Property Get IMapper_Typename() As String
    IMapper_Typename = MY_TYPENAME
End Property

'Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
'    IMapper_ExecMapper = ExecMapper(ipHostItem)
'End Function

'@Description("Converts tthe items in the SeqC to Long Integers.  If the input item cannot be converted the string "NaN" is returned")
Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
    IMapper_ExecMapper = Array(IIf(s.Comparer.ExecCmp(ipHostItem), s.Item, ipHostItem))
End Function
