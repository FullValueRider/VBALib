VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ItemMarkup"
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

Const MY_NAME                   As String = "ItemMarkup"
Private Type Properties

    Separator                   As String
    Left                        As String
    Right                       As String
    
End Type


Private p                       As Properties


'@DefaultMember
Public Function Deb _
( _
Optional ByVal ipLeft As String = Chars.twLParen, _
Optional ByVal ipSeparator As String = Chars.twComma, _
Optional ByVal ipRight As String = Chars.twRParen _
) As ItemMarkup
Attribute Deb.VB_UserMemId = 0
    With New ItemMarkup
        Set Deb = .ConstructInstance(ipLeft, ipSeparator, ipRight)
    End With
    
End Function


Public Function ConstructInstance _
( _
ByVal ipLeft As String, _
ByVal ipSeparator As String, _
ByVal ipRight As String _
) As ItemMarkup
    
    p.Separator = ipSeparator
    p.Left = ipLeft
    p.Right = ipRight
    
    Set ConstructInstance = Me
    
End Function


Public Property Get Separator() As String
    Separator = p.Separator
End Property


Public Property Get Left() As String
    Left = p.Left
End Property


Public Property Get Right() As String
    Right = p.Right
End Property


Public Function Clone() As ItemMarkup
    Set Clone = ItemMarkup.Deb(p.Left, p.Separator, p.Right)
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
