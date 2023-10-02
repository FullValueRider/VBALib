VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpTrim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'@PredeclaredId
'@Exposed
'@Folder("Mappers")


Implements IMapper

Const MY_TYPENAME               As String = "mpTrim"

Private Type State
    TrimChars                   As String
End Type


Private s                       As State


Private Function Imapper_ExecMapper(ByRef ipHostItem As Variant) As Variant

    If GroupInfo.IsString(ipHostItem) Then
        Imapper_ExecMapper = Array(Strs.Trimmer(VBA.CStr(ipHostItem), s.TrimChars))
    Else
        Imapper_ExecMapper = Array(ipHostItem)
    End If
        
    
    
End Function


Public Function Deb(Optional ByRef ipTrimChars As String = Chars.twSpace) As mpTrim
    With New mpTrim
        Set Deb = .ConstructInstance(ipTrimChars)
    End With
End Function


Friend Function ConstructInstance(ByRef ipTrimChars As String) As mpTrim
        
    If VBA.Len(ipTrimChars) = 0 Then
        s.TrimChars = Chars.twSpace
    Else
        s.TrimChars = ipTrimChars
    End If
   
    Set ConstructInstance = Me
        
End Function


Private Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property


