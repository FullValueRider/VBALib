VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpSplit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder("VBALib.Mappers")

Implements IMapper

Const MY_TYPENAME                   As String = "mpSplit"

Private Type State
    
    Separator                       As String
    Count                           As Long
    Compare                         As VbCompareMethod
        
End Type

Private s                           As State
    
    
Private Property Get IMapper_Typename() As String
    IMapper_Typename = MY_TYPENAME
End Property

Private Sub Class_Initialize()
    ' Allows mpSPlit to be used without calling .Deb
    s.Separator = " "
    s.Count = -1
    s.Compare = vbBinaryCompare
End Sub

Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
        
    If GroupInfo.IsNotString(ipHostItem) Then
            
        IMapper_ExecMapper = Array(ipHostItem)
        Exit Function
            
            
    End If
        
    IMapper_ExecMapper = Array(SeqC(VBA.Split(ipHostItem, s.Separator, s.Count, s.Compare)))

End Function

'@Defaultmember
Public Function Deb(Optional ByVal ipSeparator As String = " ", Optional ByVal ipCount As Long = -1, Optional ByVal ipCompare As VbCompareMethod = vbBinaryCompare) As mpSplit
Attribute Deb.VB_UserMemId = 0
    With New mpSplit
        Set Deb = .ConstructInstance(ipSeparator, ipCount, ipCompare)
    End With
End Function

Friend Function ConstructInstance(ByVal ipSeparator As String, ByVal ipCount As Long, ByVal ipCompare As VbCompareMethod) As mpSplit
        
    s.Separator = ipSeparator
    s.Count = ipCount
    s.Compare = ipCompare
    Set ConstructInstance = Me
        
End Function
