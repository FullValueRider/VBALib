VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MoveXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Lib")
Option Explicit

Private Const MY_TYPENAME                  As String = "MoveXY"
Private Type Properties
    x                                       As Long
    y                                       As Long
End Type

Private p                                   As Properties

'@DefaultMember
Public Function Deb(ByVal ipX As Long, ByVal ipY As Long) As MoveXY
Attribute Deb.VB_UserMemId = 0
    With New MoveXY
        Set Deb = .ConstructInstance(ipX, ipY)
    End With
End Function

Friend Function ConstructInstance(ByVal ipX As Long, ByVal ipY As Long) As MoveXY

    If ipX < -1 Or ipX > 1 Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.Direction.Let X", _
            Fmt.Text(" Value for MoveXY.X should be -1, 0 or 1, Got {0}", ipX)
    End If

    If ipY < -1 Or ipY > 1 Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.Direction.Let Y", _
            Fmt.Text(" Value for MoveXY.Y should be -1, 0 or 1, Got {0}", ipY)
    End If
    
    p.x = ipX
    p.y = ipY
    
    Set ConstructInstance = Me
    
End Function

Public Property Get x() As Long
    x = p.x
End Property



Public Property Get y() As Long
    y = p.y
End Property


Public Property Get SouthEast() As MoveXY
        Set SouthEast = MoveXY(1, -1)
End Property

Public Property Get South() As MoveXY
    Set South = MoveXY(0, -1)
End Property

Public Property Get SouthWest() As MoveXY
    Set SouthWest = MoveXY(-1, -1)
End Property

Public Property Get West() As MoveXY
    Set West = MoveXY(-1, 0)
End Property

Public Property Get NorthWest() As MoveXY
    Set NorthWest = MoveXY(-1, 1)
End Property

Public Property Get North() As MoveXY
    Set North = MoveXY(0, 1)
End Property

Public Property Get NorthEast() As MoveXY
    Set NorthEast = MoveXY(1, 1)
End Property

Public Property Get East() As MoveXY
    Set East = MoveXY(1, 0)
End Property


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
