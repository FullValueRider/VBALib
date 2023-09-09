VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpMath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("Mappers")
'@IgnoreModule
Implements IMapper

Const MY_TYPENAME               As String = "mpMath"

Public Enum Fx
  
    vbAbs
    Atn
    Atn2
    Ceil
    Cos
    Exp
    Factorial
    Floor
    IsPrime
    Ln
    Log
    Rnd
    TriangularNumber
    Power
    Round
    vbSgn
    Sin
    Root
    Tan
  
End Enum


'twinbasic requirement

Private Type State
  
    '@Ignore KeywordsUsedAsMember
    Function                      As Fx
    OptionalParam                 As Variant
  
  
End Type


Private s                     As State


Private Property Get IMapper_Typename() As String
    IMapper_Typename = MY_TYPENAME
End Property


'@Description("Applies the specified function. Check with the equivalent VBA function to determine if the option argument should be used")
Public Function Deb(ByVal ipFx As Fx, Optional ByVal ipOptionalParam As Variant = 0) As mpMath
        
    With New mpMath
        
        Set Deb = .ConstructInstance(ipFx, ipOptionalParam)
            
    End With
        
End Function


Friend Function ConstructInstance(ByVal ipFx As Fx, ByVal ipOptionalParam As Variant) As mpMath
  
    s.OptionalParam = ipOptionalParam
    s.Function = ipFx
    Set ConstructInstance = Me
        
End Function


'Private Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
'    IMapper_ExecMapper = ExecMapper(ipHostItem)
'End Function


Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
     
    Select Case s.Function
    
        Case vbAbs: IMapper_ExecMapper = Array(VBA.Abs(ipHostItem))
        Case Atn: IMapper_ExecMapper = Array(VBA.Atn(ipHostItem))
        Case Cos: IMapper_ExecMapper = Array(VBA.Cos(ipHostItem))
        Case Exp: IMapper_ExecMapper = Array(VBA.Exp(ipHostItem))
        Case Round: IMapper_ExecMapper = Array(VBA.Round(ipHostItem, s.OptionalParam))
        Case Power: IMapper_ExecMapper = Array(ipHostItem ^ s.OptionalParam)
        Case vbSgn: IMapper_ExecMapper = Array(VBA.Sgn(ipHostItem))
        Case Sin: IMapper_ExecMapper = Array(VBA.Sin(ipHostItem))
        Case Tan: IMapper_ExecMapper = Array(VBA.Tan(ipHostItem))
    
            ' Function defined, or redefined, in Class Maths
        Case Root: IMapper_ExecMapper = Array(ipHostItem ^ (1 / s.OptionalParam))
        Case IsPrime: IMapper_ExecMapper = Array(Maths.IsPrime(ipHostItem))
        Case Factorial: IMapper_ExecMapper = Array(Maths.Factorial(ipHostItem))
        Case Ln: IMapper_ExecMapper = Array(Maths.Ln(ipHostItem))
        Case Log: IMapper_ExecMapper = Array(Maths.Log(ipHostItem))
        Case TriangularNumber: IMapper_ExecMapper = Array(Maths.TriangularNumber(ipHostItem))
    
    
    End Select
        
End Function


