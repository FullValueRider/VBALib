VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayMarkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder("Lib.Fmt")
Private Type Properties

    Separator                   As String
    LeftParen                   As String
    RightParen                  As String
    
End Type

Private p                       As Properties


'@DefaultMember
Public Function Deb _
( _
    Optional ByVal ipSeparator As String = ",", _
    Optional ByVal ipLParen As String = "[", _
    Optional ByVal ipRParen As String = "]" _
) As ArrayMarkup
Attribute Deb.VB_UserMemId = 0

    With New ArrayMarkup
    
        Set Deb = .NewFilledInstance(ipSeparator, ipLParen, ipRParen)
        
    End With
    
End Function

Public Function NewFilledInstance _
( _
    Optional ByVal ipSeperator As String, _
    Optional ByVal ipLParen As String, _
    Optional ByVal ipRParen As String _
) As ArrayMarkup
    
    p.Separator = ipSeperator
    p.LeftParen = ipLParen
    p.RightParen = ipRParen
    
    Set NewFilledInstance = Me
    
End Function

Public Property Get Separator() As String
    Separator = p.Separator
End Property


Public Property Get LParen() As String
    LParen = p.LeftParen
End Property


Public Property Get RParen() As String
    RParen = p.RightParen
End Property

Public Function Clone() As ArrayMarkup

    Dim myAM As ArrayMarkup
    Set myAM = ArrayMarkup.Deb(p.Separator, p.LeftParen, p.RightParen)
   
    Set Clone = myAM
    
End Function

Public Function ToString() As String
    ToString = Fmt.Txt("{ {0} , {1} , {2} }", p.Separator, p.LeftParen, p.RightParen)
End Function
