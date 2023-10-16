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
    m_last = m_NW
End Enum


Public Enum e_AdjacentSet
    m_First
    m_4WaysNESW = m_First
    m_4WaysDiagonal
    m_8Ways
    m_last = m_8Ways
End Enum

Public Enum e_AdjacentRotation
    m_First
    m_Clockwise = m_First
    m_Anticlockwise
    m_last = m_Anticlockwise
End Enum

Public Enum e_AdjacentType
    m_First
    m_Absolute = m_First
    m_Relative
    m_last = m_Relative
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

' The adjacent set may be, (N,NE,E,SE,S,SW,W,NW), (N,S,E,W) or (NE,SE,SW,NW)
' The start direction may be any of the eight possible directions
' The set of adjacent coordinated may be ordered clockwise or anticlockwise
' The set of adjacent coordinates my be Absolute or Relative to the current X,Y location

'@Description("Returns a SeqA of the adjacent coordinates.  The starting direction, subset, Type and rotation direction can be selected. Defaults to  Eightways, North, Clockwise, Absolute")
Public Function AdjacentCoords _
( _
    Optional ByVal ipAdjacentSet As e_AdjacentSet = e_AdjacentSet.m_8Ways, _
    Optional ByVal ipStartDirection As e_AdjacentDirection = e_AdjacentDirection.m_North, _
    Optional ByVal ipRotation As e_AdjacentRotation = m_Clockwise, _
    Optional ByVal ipAdjacentType As e_AdjacentType = m_Absolute _
) As SeqA
Attribute AdjacentCoords.VB_Description = "Returns a SeqA of the adjacent coordinates.  The starting direction, subset, Type and rotation direction can be selected. Defaults to  Eightways, North, Clockwise, Absolute"

    Dim myLoc As String: myLoc = MY_TYPENAME & ".AdjacentCoords"
    Allow.InRange ipAdjacentSet, e_AdjacentSet.m_First, e_AdjacentSet.m_last, myLoc
    Allow.InRange ipStartDirection, e_AdjacentDirection.m_First, e_AdjacentDirection.m_last, myLoc
    Allow.InRange ipRotation, e_AdjacentRotation.m_First, e_AdjacentDirection.m_last, myLoc
    Allow.InRange ipAdjacentType, e_AdjacentType.m_First, e_AdjacentType.m_last, myLoc

    Dim myS As SeqA: Set myS = SeqA()
    
    Dim myDirection As e_AdjacentDirection:
    For myDirection = e_AdjacentDirection.m_First To e_AdjacentDirection.m_last
        myS.Add AdjacentCoord(myDirection, ipAdjacentType)
    Next
        
    ' there is no need to adjust for ipAdjacentSet due to start location
    ' as the first point will always be the correct next available
    ' point in the direction of rotation based on the code below
    If ipStartDirection <> m_North Then
        myS.RotLeft ipStartDirection
    End If
    
    If ipRotation = m_Anticlockwise Then
        myS.Reversed.RotRight
    End If
    
    ' Class instances of the Group type ItemObject (See GroupInfo class)
    ' will be compared based their string output if available.
    ' Hence the removes below are valid because it will be done
    ' based on comparison of the string output of the content, not the address of the object
    If ipAdjacentSet = m_4WaysNESW Then
        myS.Remove AdjacentCoord(m_NE, ipAdjacentType)
        myS.Remove AdjacentCoord(m_SE, ipAdjacentType)
        myS.Remove AdjacentCoord(m_SW, ipAdjacentType)
        myS.Remove AdjacentCoord(m_NW, ipAdjacentType)
    End If
    
    If ipAdjacentSet = m_4WaysDiagonal Then
        myS.Remove AdjacentCoord(m_North, ipAdjacentType)
        myS.Remove AdjacentCoord(m_East, ipAdjacentType)
        myS.Remove AdjacentCoord(m_South, ipAdjacentType)
        myS.Remove AdjacentCoord(m_West, ipAdjacentType)
    End If
    
    Set AdjacentCoords = myS
    
End Function


Public Function AdjacentCoord(ByRef ipDirection As e_AdjacentDirection, Optional ByVal ipAdjacentType As e_AdjacentType = m_Absolute) As PointXY

    Dim myLoc As String: myLoc = MY_TYPENAME & ".AdjacentCoord"
    Allow.InRange ipDirection, e_AdjacentDirection.m_First, e_AdjacentDirection.m_last, myLoc
    Allow.InRange ipAdjacentType, e_AdjacentType.m_First, e_AdjacentType.m_last, myLoc
    
    Dim myPoint As PointXY
    
    If ipAdjacentType = m_Absolute Then
        Set myPoint = Me
    Else
        Set myPoint = PointXY(0, 0)
    End If
    
    Select Case ipDirection
    
        Case m_North:   Set AdjacentCoord = PointXY(myPoint.X, myPoint.Y + 1)
        Case m_NE:      Set AdjacentCoord = PointXY(myPoint.X + 1, myPoint.Y + 1)
        Case m_East:    Set AdjacentCoord = PointXY(myPoint.X + 1, myPoint.Y)
        Case m_SE:      Set AdjacentCoord = PointXY(myPoint.X + 1, myPoint.Y - 1)
        Case m_South:   Set AdjacentCoord = PointXY(myPoint.X, myPoint.Y - 1)
        Case m_SW:      Set AdjacentCoord = PointXY(myPoint.X - 1, myPoint.Y - 1)
        Case m_West:    Set AdjacentCoord = PointXY(myPoint.X - 1, myPoint.Y)
        Case m_NW:      Set AdjacentCoord = PointXY(myPoint.X - 1, myPoint.Y + 1)
            
    End Select
    
End Function


Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property
