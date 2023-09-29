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
Public Enum e_AdjacentCoord
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


Public Enum e_WhichCoords
    m_First
    m_NESW = m_First
    m_Diagonals
    m_All
    m_Last = m_All
End Enum

Public Enum e_AdjacentDirection
    m_First
    m_clockwise = m_First
    m_Anticlockwise
    m_Last = m_Anticlockwise
End Enum

Private Type Properties
    x                                           As Long
    y                                           As Long
End Type

Private p                                       As Properties


'@DefaultMember
Public Function Deb(Optional ByRef ipX As Variant = 0, Optional ByRef ipY As Variant = 0) As PointXY
Attribute Deb.VB_UserMemId = 0
    With New PointXY
        Set Deb = .ConstructInstance(ipX, ipY)
    End With
End Function

Friend Function ConstructInstance(ByRef ipX As Variant, ByRef ipY As Variant) As PointXY
    p.x = VBA.CLng(ipX)
    p.y = VBA.CLng(ipY)
    Set ConstructInstance = Me
End Function

Public Property Get x() As Long
    x = p.x
End Property

Public Property Let x(ByVal ipX As Long)
'    If ipX < p.BoundsArray1To4(m_MaxX) Or ipX > p.BoundsArray1To4(m_MaxX) Then
'        AlertOutOfBOunds
'    End If
'
'    p.Moved = IIf(p.X <> ipX, True, False)

    p.x = ipX
    
End Property

Public Property Get y() As Long
    y = p.y
End Property

Public Property Let y(ByVal ipY As Long)
    
'    If ipY < p.BoundsArray1To4(m_MinY) Or ipY > p.BoundsArray1To4(m_MaxY) Then
'        AlertOutOfBOunds
'    End If
'
'    p.Moved = IIf(p.Y <> ipY, True, False)

    p.y = ipY
    
End Property



Public Function ToString() As String
    ToString = VBA.CStr(p.x) & Chars.twComma & VBA.CStr(p.y)
End Function

'@Description("Return a two item array with x at index 1 and y at index 2")
Public Function ToArray() As Variant
Attribute ToArray.VB_Description = "Return a two item array with x at index 1 and y at index 2"
    Dim myA(1 To 2) As Long
    myA(1) = p.x
    myA(2) = p.y
    ToArray = myA
End Function

Public Function ToKVPair() As KVPair
    Set ToKVPair = KVPair(p.x, p.y)
End Function

'ToDo: update the Adjacent method to be able to start from any coord and go clockwise or anticlockwise.
'@Description("Returns a SeqA of cooridinate starting at lowest,rightest point for e_perimType then proceeding anticlockwise")
Public Function AdjacentCoords _
( _
    Optional ByVal ipPerimType As e_WhichCoords = e_WhichCoords.m_All, _
    Optional ByVal ipStartCoord As e_AdjacentCoord = e_AdjacentCoord.m_North, _
    Optional ByVal ipDirection As e_AdjacentDirection = m_clockwise _
) As SeqA
Attribute AdjacentCoords.VB_Description = "Returns a SeqA of cooridinate starting at lowest,rightest point for e_perimType then proceeding anticlockwise"

    ' Note: this method return a list of 8 coordinates
    ' irrespective of the bounds or list of forbidden points
    
    ' the convention of starting at the SE corner and proceeding anticlockwise
    ' comes from Advent Of Code 2017 Day03 Part 02
    Dim myS As SeqA: Set myS = SeqA.Deb
    
    ' add the points in order starting at the SE location
    If ipPerimType = m_NESW Or ipPerimType = m_All Then myS.Add Me.North
    If ipPerimType = m_Diagonals Or ipPerimType = m_All Then myS.Add Me.NorthEast
    If ipPerimType = m_NESW Or ipPerimType = m_All Then myS.Add Me.East
    If ipPerimType = m_Diagonals Or ipPerimType = m_All Then myS.Add Me.SouthEast
    If ipPerimType = m_NESW Or ipPerimType = m_All Then myS.Add Me.South
    If ipPerimType = m_Diagonals Or ipPerimType = m_All Then myS.Add Me.SouthWest
    If ipPerimType = m_NESW Or ipPerimType = m_All Then myS.Add Me.West
    If ipPerimType = m_Diagonals Or ipPerimType = m_All Then myS.Add Me.NorthWest
    
    If ipStartCoord <> m_North Then
        myS.RotLeft ipStartCoord
    End If
    
    If ipDirection = m_Anticlockwise Then
        myS.Reversed.RotRight
    End If

    Set AdjacentCoords = myS
    
End Function

'ToDo: update the AdjacentOffsets method to be able to start from any coord and go clockwise or anticlockwise.
'@Description("Returns a SeqA of offset cooridinates starting at lowest,rightest point for e_perimType then proceeding anticlockwise")
Public Function AdjacentOffsets _
( _
    Optional ByVal ipPerimType As e_WhichCoords = e_WhichCoords.m_All, _
    Optional ByVal ipStartCoord As e_AdjacentCoord = e_AdjacentCoord.m_North, _
    Optional ByVal ipDirection As e_AdjacentDirection = m_clockwise _
) As SeqA
Attribute AdjacentOffsets.VB_Description = "Returns a SeqA of offset cooridinates starting at lowest,rightest point for e_perimType then proceeding anticlockwise"
    ' Note: this method return a list of 8 coordinates
    ' irrespective of the bounds or list of forbidden points
    '
    ' the convention of starting at the SE corner and proceeding anticlockwise
    ' comes from Advent Of Code 2017 Day03 Part 02
    Dim myS As SeqA: Set myS = SeqA.Deb
    
    ' add the points in order starting at the SE location
    If ipPerimType = m_NESW Or ipPerimType = m_All Then myS.Add PointXY(0, 1)
    If ipPerimType = m_Diagonals Or ipPerimType = m_All Then myS.Add PointXY(1, 1)
    If ipPerimType = m_NESW Or ipPerimType = m_All Then myS.Add PointXY(1, 0)
    If ipPerimType = m_Diagonals Or ipPerimType = m_All Then myS.Add PointXY(1, -1)
    If ipPerimType = m_NESW Or ipPerimType = m_All Then myS.Add PointXY(0, -1)
    If ipPerimType = m_Diagonals Or ipPerimType = m_All Then myS.Add PointXY(-1, -1)
    If ipPerimType = m_NESW Or ipPerimType = m_All Then myS.Add PointXY(-1, 0)
    If ipPerimType = m_Diagonals Or ipPerimType = m_All Then myS.Add PointXY(-1, 1)
    
    If ipStartCoord <> m_North Then
        myS.RotLeft ipStartCoord
    End If
    
    If ipDirection = m_Anticlockwise Then
        myS.Reversed.RotRight
    End If

    Set AdjacentOffsets = myS

End Function


Public Property Get North() As PointXY
    Set North = PointXY(p.x, p.y + 1)
End Property

Public Property Get NorthEast() As PointXY
    Set NorthEast = PointXY(p.x + 1, p.y + 1)
End Property

Public Property Get East() As PointXY
    Set East = PointXY(p.x + 1, p.y)
End Property

Public Property Get SouthEast() As PointXY
        Set SouthEast = PointXY(p.x + 1, p.y - 1)
End Property

Public Property Get South() As PointXY
    Set South = PointXY(p.x, p.y - 1)
End Property

Public Property Get SouthWest() As PointXY
    Set SouthWest = PointXY(p.x - 1, p.y - 1)
End Property

Public Property Get West() As PointXY
    Set West = PointXY(p.x - 1, p.y)
End Property

Public Property Get NorthWest() As PointXY
    Set NorthWest = PointXY(p.x - 1, p.y + 1)
End Property

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property



