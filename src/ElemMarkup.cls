VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ElemMarkup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@IgnoreModule SuspiciousPredeclaredInstanceAccess
'@PredeclaredId
'@Exposed
'@Folder("VBALib.Fmt")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Option Explicit

Const MY_NAME                   As String = "ElemMarkup"
Private Type Properties

    Left                        As String
    Separator                   As String
    Right                       As String
    
End Type

Private p                       As Properties


'@DefaultMember
Public Function Deb _
( _
Optional ByRef ipLeft As String = Chars.twLParen, _
Optional ByRef ipSeparator As String = Chars.twComma, _
Optional ByRef ipRight As String = Chars.twRParen _
) As ElemMarkup
Attribute Deb.VB_UserMemId = 0
    With New ElemMarkup
        Set Deb = .ConstructInstance(ipLeft, ipSeparator, ipRight)
    End With
End Function


Public Function ConstructInstance _
( _
ByRef ipLeft As String, _
ByRef ipSeparator As String, _
ByRef ipRight As String _
) As ElemMarkup
    
    p.Separator = ipSeparator
    p.Left = ipLeft
    p.Right = ipRight
    
    Set ConstructInstance = Me
    
End Function


Public Property Get Separator() As String
    Separator = p.Separator
End Property

Public Property Let Separator(ByRef ipSeparator As String)
    p.Separator = ipSeparator
End Property


Public Property Get Left() As String
    Left = p.Left
End Property
Public Property Let Left(ByRef ipLeft As String)
    p.Left = ipLeft
End Property


Public Property Get Right() As String
    Right = p.Right
End Property

Public Property Let Right(ByRef ipRight As String)
    p.Right = ipRight
End Property


Public Function Clone() As ElemMarkup
    Set Clone = ElemMarkup.Deb(p.Left, p.Separator, p.Right)
End Function


Public Function ToString() As String
    ToString = p.Left & "," & p.Separator & "," & p.Right
End Function


Public Property Get Name() As String
    Name = MY_NAME
End Property


Public Function ToArray() As Variant
    ToArray = Array(p.Left, p.Separator, p.Right)
End Function
