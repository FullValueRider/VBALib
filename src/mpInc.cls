VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Mappers")
Implements IMapper

Const MY_TYPENAME                   As String = "mpInc"

Private Type State
    Increment                       As Variant
End Type


Private s                           As State


'@DefaultMember
Public Function Deb(Optional ByVal ipIncrement As Variant = 1) As mpInc
Attribute Deb.VB_UserMemId = 0
    With New mpInc
        Set Deb = .ConstructInstance(ipIncrement)
    End With
End Function


Friend Function ConstructInstance(ByVal ipIncrement As Variant) As mpInc
        
    s.Increment = ipIncrement
    Set ConstructInstance = Me
        
End Function


Private Property Get IMapper_TypeName() As String
    IMapper_TypeName = MY_TYPENAME
End Property


''@Description("")
'Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
'    Set IMapper_ExecMapper = ExecMapper(ipHostItem)
'End Function

Private Function Imapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
        
    '@Ignore VariableNotAssigned
    
    Select Case True
    
        Case GroupInfo.IsString(ipHostItem)
        
            Dim myString As String
            myString = ipHostItem
            Dim myIndex As Long
            For myIndex = 1 To s.Increment
                myString = Strs.Inc(myString)
            Next
            Imapper_ExecMapper = Array(myString)
            
            
        Case GroupInfo.IsNumber(ipHostItem)
           
            Imapper_ExecMapper = Array(ipHostItem + s.Increment)
            
            
        Case Else
            
            Imapper_ExecMapper = Array(ipHostItem)
            
    End Select
    
End Function
