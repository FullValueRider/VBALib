VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PointXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A class for two dimensional points"
'@predeclaredId
'@Exposed
'@Folder("Lib")
'@ModuleDescription("A class for two dimensional points")
Option Explicit
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
Private Const MY_TYPENAME                       As String = "PointXY"

'ToDO  Add bounds and forbidden locations in the same manner as keypad.
Public Enum e_AdjacentDirection ' Defined clockwise
    m_First = 0
    m_North = m_First
    m_NE
    m_East
    m_SE
    m_South
    m_SW
    m_West
    m_NW
    m_Last = m_NW
    
End Enum


Public Enum e_AdjacentSet
    m_First
    m_4WaysNESW = m_First
    m_4WaysDiagonal
    m_8Ways
    m_Last = m_8Ways
End Enum

Public Enum e_AdjacentRotation
    m_First
    m_Clockwise = m_First
    m_Anticlockwise
    m_Last = m_Anticlockwise
End Enum

Private Type Properties
    X                                           As Long
    Y                                           As Long
End Type

Private p                                       As Properties


'@DefaultMember
'Variant is used as a surrogate for any number
Public Function Deb(Optional ByRef ipX As Variant = 0, Optional ByRef ipY As Variant = 0) As PointXY
Attribute Deb.VB_UserMemId = 0
    With New PointXY
        Set Deb = .ConstructInstance(ipX, ipY)
    End With
End Function

Friend Function ConstructInstance(ByRef ipX As Variant, ByRef ipY As Variant) As PointXY

    Allow.IsNumber ipX, "PointXY.ConstructInstance"
    Allow.IsNumber ipY, "PointXY.ConstructInstance"
    
    p.X = ipX
    p.Y = ipY
    
    Set ConstructInstance = Me
    
End Function

Public Property Get X() As Long
    X = p.X
End Property

Public Property Let X(ByVal ipX As Long)

    Allow.IsNumber ipX, "PointXY.Let_X"
    
    p.X = ipX
    
End Property

Public Property Get Y() As Long
    Y = p.Y
End Property

Public Property Let Y(ByVal ipY As Long)
    
    Allow.IsNumber ipY, "PointXY.Let_Y"

    p.Y = ipY
    
End Property


Public Function ToString() As String
    ToString = Fmt.Text("{0},{1}", p.X, p.Y)
End Function

'@Description("Return a two item array with x at index 1 and y at index 2")
Public Function ToArray() As Variant
Attribute ToArray.VB_Description = "Return a two item array with x at index 1 and y at index 2"
    Dim myA() As Variant
    ReDim myA(1 To 2)
    myA(1) = p.X
    myA(2) = p.Y
    ToArray = myA
End Function

Public Function ToKVPair() As KVPair
    Set ToKVPair = KVPair(p.X, p.Y)
End Function


'@Description("Returns a SeqA of the adjacent cooridinates starting.  The starting point, subset and rotation direction can be selected. Defaults North, Eightways, Clockwise")
Public Function AdjacentCoords _
( _
    Optional ByVal ipAdjacentSet As e_AdjacentSet = e_AdjacentSet.m_8Ways, _
    Optional ByVal ipStartCoord As e_AdjacentDirection = e_AdjacentDirection.m_North, _
    Optional ByVal ipRotation As e_AdjacentRotation = m_Clockwise _
) As SeqA
Attribute AdjacentCoords.VB_Description = "Returns a SeqA of the adjacent cooridinates starting.  The starting point, subset and rotation direction can be selected. Defaults North, Eightways, Clockwise"

    Dim myS As SeqA: Set myS = SeqA.Deb
    
    If ipAdjacentSet = m_4WaysNESW Or ipAdjacentSet = m_8Ways Then myS.Add Me.North
    If ipAdjacentSet = m_4WaysDiagonal Or ipAdjacentSet = m_8Ways Then myS.Add Me.NorthEast
    If ipAdjacentSet = m_4WaysNESW Or ipAdjacentSet = m_8Ways Then myS.Add Me.East
    If ipAdjacentSet = m_4WaysDiagonal Or ipAdjacentSet = m_8Ways Then myS.Add Me.SouthEast
    If ipAdjacentSet = m_4WaysNESW Or ipAdjacentSet = m_8Ways Then myS.Add Me.South
    If ipAdjacentSet = m_4WaysDiagonal Or ipAdjacentSet = m_8Ways Then myS.Add Me.SouthWest
    If ipAdjacentSet = m_4WaysNESW Or ipAdjacentSet = m_8Ways Then myS.Add Me.West
    If ipAdjacentSet = m_4WaysDiagonal Or ipAdjacentSet = m_8Ways Then myS.Add Me.NorthWest
    
    If ipStartCoord <> m_North Then
        myS.RotLeft ipStartCoord
    End If
    
    If ipRotation = m_Anticlockwise Then
        myS.Reversed.RotRight
    End If

    Set AdjacentCoords = myS
    
End Function


'@Description("Returns a SeqA of offset cooridinates.The starting point, subset and rotation direction can be selected. Defaults North, Eightways, Clockwise")
Public Function AdjacentOffsets _
( _
    Optional ByVal ipAdjacentSet As e_AdjacentSet = e_AdjacentSet.m_8Ways, _
    Optional ByVal ipStartCoord As e_AdjacentDirection = e_AdjacentDirection.m_North, _
    Optional ByVal ipRotation As e_AdjacentRotation = m_Clockwise _
) As SeqA
Attribute AdjacentOffsets.VB_Description = "Returns a SeqA of offset cooridinates.The starting point, subset and rotation direction can be selected. Defaults North, Eightways, Clockwise"
    
    Dim my00 As PointXY: Set my00 = PointXY()
    
    Dim myS As SeqA: Set myS = SeqA.Deb
    
    If ipAdjacentSet = m_4WaysNESW Or ipAdjacentSet = m_8Ways Then myS.Add my00.North
    If ipAdjacentSet = m_4WaysDiagonal Or ipAdjacentSet = m_8Ways Then myS.Add my00.NorthEast
    If ipAdjacentSet = m_4WaysNESW Or ipAdjacentSet = m_8Ways Then myS.Add my00.East
    If ipAdjacentSet = m_4WaysDiagonal Or ipAdjacentSet = m_8Ways Then myS.Add my00.SouthEast
    If ipAdjacentSet = m_4WaysNESW Or ipAdjacentSet = m_8Ways Then myS.Add my00.South
    If ipAdjacentSet = m_4WaysDiagonal Or ipAdjacentSet = m_8Ways Then myS.Add my00.SouthWest
    If ipAdjacentSet = m_4WaysNESW Or ipAdjacentSet = m_8Ways Then myS.Add my00.West
    If ipAdjacentSet = m_4WaysDiagonal Or ipAdjacentSet = m_8Ways Then myS.Add my00.NorthWest
    
    If ipStartCoord <> m_North Then
        myS.RotLeft ipStartCoord
    End If
    
    If ipRotation = m_Anticlockwise Then
        myS.Reversed.RotRight
    End If

    Set AdjacentOffsets = myS

End Function


Public Property Get North() As PointXY
    Set North = PointXY(p.X, p.Y + 1)
End Property

Public Property Get NorthEast() As PointXY
    Set NorthEast = PointXY(p.X + 1, p.Y + 1)
End Property

Public Property Get East() As PointXY
    Set East = PointXY(p.X + 1, p.Y)
End Property

Public Property Get SouthEast() As PointXY
        Set SouthEast = PointXY(p.X + 1, p.Y - 1)
End Property

Public Property Get South() As PointXY
    Set South = PointXY(p.X, p.Y - 1)
End Property

Public Property Get SouthWest() As PointXY
    Set SouthWest = PointXY(p.X - 1, p.Y - 1)
End Property

Public Property Get West() As PointXY
    Set West = PointXY(p.X - 1, p.Y)
End Property

Public Property Get NorthWest() As PointXY
    Set NorthWest = PointXY(p.X - 1, p.Y + 1)
End Property

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
