VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TrackedXY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Lib")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Option Explicit

Private Const MY_TYPENAME                   As String = "TrackedXY"
Private Const MOVE_FORWARD                  As String = "f"
Private Const MOVE_BACKWARD                 As String = "b"


Public Enum e_Heading
    
    m_North = 0
    m_NE
    m_East
    m_SE
    m_South
    m_SW
    m_West
    m_NW
    
End Enum


Public Enum e_Bounds
    m_First = 1
    m_Minx = m_First
    m_MinY
    m_Maxx
    m_MaxY
    m_Last = m_MaxY
End Enum

' The TrackedXY class represents an X,Y location which can be moved
' and which tracks where it has been and when it crossed its own trail.
'
' Bounds can be set to limit the area in which the point can move.
'
' Forbidden locations can also be set into which the point cannot move.
'
' The point knows which way it is facing so allows
' forward and backwards movement.
'
' The move methods requires a direction and distance.
' At the end of a move the heading will be the move direction.

' Fowards and Backwards methods can be used, which will move the point
' along the direction of the heading by the specified number of steps.

' The Moves Kvp contains x,y offets for movements in the eight directions
' of the compass as well as forward and backwards.
'
' The forwards and backwards moves are updated each time the heading is
' changed (e.g by a turn command or a move command).
'
' The turn command turns left or right by 45 degrees or 90 degrees depending
' on the Movement (Fourways or EightWays) (NESW vs N,NE,E,etc)

Public Enum e_MovementType
    FourWays
    EightWays
End Enum

Private Type Properties
    ' the current x,y position
    Location                                    As PointXY
    ' the direction we are facing for forward/backwards movements
    Heading                                     As e_Heading
    HeadingTextToEnum                           As KvpA
    ' Dictionary of current heading vs new heading for a left turn
    TurnLeft                                    As KvpA
    ' DIctionary of current heading vs new heading for a right turn
    TurnRight                                   As KvpA             ' Old heading vs new heading for eightway turns
    Movement                                    As e_MovementType
    Moves                                       As KvpA             ' heading vs x,y offset for a single step in the direction of the heading
    ' Trail provides a record of the x,y coordinates in the order in which they were visited
    Trail                                       As SeqHC    ' x,y point (as string) visited in order of passage
    'Revisits is a seq of xy coording vs a sequence of when that coordinate was revisited.
    Revisits                                    As KvpHL    ' String:X,Y location vs SeqA of trail.count when visited
   ' Intersection            As String
    ' saves converting the string at Trail.Item(1) back to a point
    Origin                                      As PointXY
    
    ' Moved is a boolean which is set to true if the movement
    ' could not be fully realised.  for multi step moves
    ' indicated that the specified move could not be completed
    ' i.e the location may have only partially moved.
    Moved                                       As Boolean
    ' defines an area outside of which the point may not go
    ' a seqA of SeqA which defines a rectangular area
    Area                                        As SeqA
    ' Bounds outside which the point cannot move
    ' bounds can be set by providing an area or by using the
    ' set bounds method. If an area is set then the setbounds is ignored
    BoundsArray1To4                             As Variant
    ' true if bounds are in force
    BoundsInUse                                 As Boolean
    ' a sequence of x,y strings defining points which cannot be visited.
    ForbiddenCoords                             As SeqT
    ' a list of values in area which may not be visited
    ForbiddenValues                             As SeqT
    Intersection                              As Boolean
    
End Type

Private p As Properties
 
'@DefaultMember
Public Function Deb _
( _
    Optional ByRef ipPoint As PointXY = Nothing, _
    Optional ByRef ipArea As SeqA = Nothing, _
    Optional ByRef ipForbiddenCoords As SeqT = Nothing, _
    Optional ByRef ipForbiddenvalues As SeqT = Nothing, _
    Optional ByRef ipBoundsArray As Variant = Empty _
) As TrackedXY
Attribute Deb.VB_UserMemId = 0
    With New TrackedXY
        Set Deb = .ConstructInstance(ipPoint, ipArea, ipForbiddenCoords, ipForbiddenvalues, ipBoundsArray)
    End With
End Function


Public Function ConstructInstance _
( _
    Optional ByRef ipPoint As PointXY = Nothing, _
    Optional ByRef ipArea As SeqA = Nothing, _
    Optional ByRef ipForbiddenCoords As SeqT = Nothing, _
    Optional ByRef ipForbiddenvalues As SeqT = Nothing, _
    Optional ByRef ipBounds As Variant = Empty _
) As TrackedXY

    ' Initialise internal variables first where possible
    p.Heading = e_Heading.m_North
    p.BoundsInUse = False
    p.Moved = False
    
    InitialiseHeadings
    InitialiseMoves
    InitialiseTurns
    
    Set p.Trail = SeqHC.Deb
    Set p.Revisits = KvpHL.Deb
    
    ' Now process parameters leaving pointXT to the last
    Set p.Area = ipArea
    If IsNotNothing(p.Area) Then
        
        p.BoundsArray1To4(m_Minx) = p.Area.First.FirstIndex
        p.BoundsArray1To4(m_MinY) = p.Area.FirstIndex
        p.BoundsArray1To4(m_Maxx) = p.Area.First.LastIndex
        p.BoundsArray1To4(m_MaxY) = p.Area.LastIndex
        p.BoundsInUse = True
        
    End If
    
    Set p.ForbiddenCoords = ipForbiddenCoords
    Set p.ForbiddenValues = ipForbiddenvalues
    
    If Not p.BoundsInUse Then
        If Not VBA.IsEmpty(ipBounds) Then
            If UBound(ipBounds) - LBound(ipBounds) + 1 = 4 Then
                p.BoundsArray1To4 = ipBounds
                ReDim p.BoundsArray1To4(1 To 4)
                p.BoundsInUse = True
            Else
                Err.Raise 17 + vbObjectError, _
                    "VBALib.TrackedPoint", _
                    Fmt.Text("The ipBounds array should contain 4 elements.  Found {0}", ArrayOp.Count(ipBounds))
            End If
        End If
    End If
    
    ' we can finally set the currrent location, checking that it is withing bounds first.
    If IsNothing(ipPoint) Then
        If p.BoundsInUse Then
            Set p.Location = PointXY(p.BoundsArray1To4(m_Minx), p.BoundsArray1To4(m_MinY))
        Else
            Set p.Location = PointXY(0, 0)
        End If
    Else
        Set p.Location = ipPoint
    End If
    
    Set p.Origin = PointXY(p.Location.X, p.Location.Y)
    
    'If bounds are set check that ipPoint is withing the bounded area
    If p.BoundsInUse Then
        With ipPoint
            If .X < p.BoundsArray1To4(m_Minx) Or .X > p.BoundsArray1To4(m_Maxx) Or .Y < p.BoundsArray1To4(m_MinY) Or .Y > p.BoundsArray1To4(m_MaxY) Then
                Err.Raise 17 + vbObjectError, _
                    "VBALib.TrackedPoint", _
                    "ipPoint is not within the bounded area"
            End If
        End With
    End If
    
    
    
    ' now we know the X,Y location we can add it as the first coordinate in p.Trail
    p.Trail.Add p.Location.ToString
    
    Set ConstructInstance = Me

End Function
 
Private Sub InitialiseHeadings()

    Set p.HeadingTextToEnum = KvpA.Deb
    With p.HeadingTextToEnum
        .Add "^", m_North
        .Add ">", m_East
        .Add "v", m_South
        .Add "<", m_West
        
        .Add "n", m_North
        .Add "e", m_East
        .Add "s", m_South
        .Add "w", m_West
        
        .Add "north", m_North
        .Add "east", m_East
        .Add "south", m_South
        .Add "west", m_West
        
        .Add "u", m_North
        .Add "r", m_East
        .Add "d", m_South
        .Add "l", m_West
        
        .Add "up", m_North
        .Add "right", m_East
        .Add "down", m_South
        .Add "left", m_West
        
        .Add "0", m_North
        .Add "1", m_East
        .Add "2", m_South
        .Add "3", m_West
        
        .Add "ne", m_NE
        .Add "se", m_SE
        .Add "sw", m_SW
        .Add "nw", m_NW
        
        .Add "northeast", m_NE
        .Add "southeast", m_SE
        .Add "southwest", m_SW
        .Add "northwest", m_NW
        
        ' The directions below will be updated based on any turn instruction
        .Add "f", m_North
        .Add "forward", m_North
        .Add "Forwards", m_North
        
        .Add "b", m_South
        .Add "back", m_South
        .Add "backwards", m_South
    End With
End Sub

Private Sub InitialiseMoves()

    Set p.Moves = KvpA.Deb
    With p.Moves
    
        .Add "^", MoveXY.North
        .Add ">", MoveXY.East
        .Add "v", MoveXY.South
        .Add "<", MoveXY.West
        
        .Add "n", MoveXY.North
        .Add "e", MoveXY.East
        .Add "s", MoveXY.South
        .Add "w", MoveXY.West
        
        .Add "north", MoveXY.North
        .Add "east", MoveXY.East
        .Add "south", MoveXY.South
        .Add "west", MoveXY.West
        
        .Add "u", MoveXY.North
        .Add "r", MoveXY.East
        .Add "d", MoveXY.South
        .Add "l", MoveXY.West
        
        .Add "up", MoveXY.North
        .Add "right", MoveXY.East
        .Add "down", MoveXY.South
        .Add "left", MoveXY.West
        
        .Add "0", MoveXY.North
        .Add "1", MoveXY.East
        .Add "2", MoveXY.South
        .Add "3", MoveXY.West
        
        .Add "ne", MoveXY.NorthEast
        .Add "se", MoveXY.SouthEast
        .Add "sw", MoveXY.SouthWest
        .Add "nw", MoveXY.NorthWest
        
        .Add "northeast", MoveXY.NorthEast
        .Add "southeast", MoveXY.SouthEast
        .Add "southwest", MoveXY.SouthWest
        .Add "northwest", MoveXY.NorthWest
        
        ' The directions below will be updated based on any turn instruction
        .Add "f", MoveXY.North
        .Add "forward", MoveXY.North
        .Add "Forwards", MoveXY.North
        
        .Add "b", MoveXY.South
        .Add "back", MoveXY.South
        .Add "backwards", MoveXY.South
               
    End With
    
End Sub

Private Sub InitialiseTurns()

    Set p.TurnLeft = KvpA.Deb
    
    With p.TurnLeft
        .Add e_Heading.m_North, e_Heading.m_NW
        .Add e_Heading.m_NW, e_Heading.m_West
        .Add e_Heading.m_West, e_Heading.m_SW
        .Add e_Heading.m_SW, e_Heading.m_South
        .Add e_Heading.m_South, e_Heading.m_SE
        .Add e_Heading.m_SE, e_Heading.m_East
        .Add e_Heading.m_East, e_Heading.m_NE
        .Add e_Heading.m_NE, e_Heading.m_North
    End With
    
    
    Set p.TurnRight = KvpA.Deb
    
    With p.TurnRight
        .Add e_Heading.m_North, e_Heading.m_NE
        .Add e_Heading.m_NE, e_Heading.m_East
        .Add e_Heading.m_East, e_Heading.m_SE
        .Add e_Heading.m_SE, e_Heading.m_South
        .Add e_Heading.m_South, e_Heading.m_SW
        .Add e_Heading.m_SW, e_Heading.m_West
        .Add e_Heading.m_West, e_Heading.m_NW
        .Add e_Heading.m_NW, e_Heading.m_North
    End With
    
End Sub
Public Property Get Movement() As e_MovementType
    Movement = p.Movement
End Property

Public Property Let Movement(ByVal ipMovementType As e_MovementType)
    p.Movement = ipMovementType
End Property

Public Property Get Moves() As KvpA
    Set Moves = p.Moves
End Property

Public Function Turn(ByRef ipTurn As String) As TrackedXY
    
    Set Turn = Me
    Dim myTurn As String: myTurn = VBA.LCase$(ipTurn)
    If VBA.InStr("l,r,left,right", myTurn) = 0 Then
        Exit Function
    End If
    
    If VBA.InStr("l,left", myTurn) > 0 Then
        p.Heading = p.TurnLeft.Item(p.Heading)
        If p.Movement = e_AdjacentSet.m_4WayNESW Then
            p.Heading = p.TurnLeft.Item(p.Heading)
        End If
        
        p.Moves.Item(MOVE_FORWARD) = p.Moves.Item(p.Heading)
        p.Moves.Item(MOVE_BACKWARD) = p.Moves.Item((p.Heading + 4) Mod 8)
        Exit Function
    End If
    
    p.Heading = p.TurnRight.Item(p.Heading)
    If p.Movement = PointXY.m_4WayNESW Then
        p.Heading = p.TurnRight.Item(p.Heading)
    End If
    
    p.Moves.Item(MOVE_FORWARD) = p.Moves.Item(p.Heading)
    p.Moves.Item(MOVE_BACKWARD) = p.Moves.Item((p.Heading + 4) Mod 8)
    p.HeadingTextToEnum.Item(MOVE_FORWARD) = p.Heading
    p.HeadingTextToEnum.Item(MOVE_BACKWARD) = (p.Heading + 4) Mod 8
    
End Function

Public Property Set Moves(ByVal ipMoves As KvpA)
    Set p.Moves = ipMoves
End Property

Public Function Move(ByRef ipDirection As String, Optional ByVal ipSteps As Long = 1) As TrackedXY

    Set Move = Me
    
    If p.Moves.LacksKey(VBA.LCase$(ipDirection)) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.TrackedXY.Move", _
            Fmt.Text("Unknown direction for Move. Got {0}", ipDirection)
    Else
        p.Heading = p.HeadingTextToEnum.Item(ipDirection)
        p.HeadingTextToEnum.Item(MOVE_FORWARD) = p.Heading
        p.HeadingTextToEnum.Item(MOVE_BACKWARD) = (p.Heading + 4) Mod 8
        
        p.Moves.Item(MOVE_FORWARD) = p.HeadingTextToEnum.Item(MOVE_FORWARD)
        p.Moves.Item(MOVE_BACKWARD) = p.HeadingTextToEnum.Item(MOVE_BACKWARD)
    End If
    
    If ipSteps = 0 Then
        Exit Function
    End If
    
    ' Setting the heading or turning will automatically
    ' update the forward and backward move values
    Dim myMover As MoveXY
    
    If ipSteps > 0 Then
        Set myMover = p.Moves.Item(MOVE_FORWARD)
    Else
        Set myMover = p.Moves.Item(MOVE_BACKWARD)
    End If
    
    Dim myX As Long: myX = p.Location.X
    Dim myY As Long: myY = p.Location.Y
    
    Dim myStep As Long
    For myStep = 1 To VBA.Abs(ipSteps)
        
        p.Moved = False
    
        myX = myX + myMover.X
        myY = myY + myMover.Y
        
        Dim myCoord As String: myCoord = VBA.CStr(myX) & Chars.twComma & VBA.CStr(myY)
        
        ' check if new location is out of bounds
        If p.BoundsInUse Then
            If myX < p.BoundsArray1To4(m_Minx) Or myX > p.BoundsArray1To4(m_Maxx) Or myY < p.BoundsArray1To4(m_MinY) Or myY > p.BoundsArray1To4(m_MaxY) Then
                Exit Function
            End If
        End If
        
        ' is it a forbidden location
        If p.ForbiddenCoords.HoldsItems Then
            If p.ForbiddenCoords.HoldsItem(myCoord) Then
                Exit Function
            End If
        End If
        
        ' is location a forbidden value
        If Not VBA.IsEmpty(p.Area) Then
            If p.ForbiddenValues.HoldsItems Then
                If p.ForbiddenValues.HoldsItem(p.Area.Item(myY).Item(myX)) Then
                    Exit Function
                End If
            End If
        End If
        
        ' hurrah we can update location
        p.Moved = True
        p.Location.X = myX
        p.Location.Y = myY
        
        p.Intersection = False
        
        If p.Trail.HoldsItem(myCoord) Then
            p.Intersection = True
            
            If p.Trail.HoldsItem(myCoord) Then
            
                p.Revisits.Add myCoord, SeqA.Deb
                p.Intersection = True
                
            End If
            
            p.Revisits.Item(myCoord).Add p.Trail.Count
        
        End If
        
    Next
    
End Function


Public Function Location() As PointXY
    Set Location = p.Location
End Function


Public Function Steps() As Long
    Steps = p.Trail.Count - 1
End Function


Public Function Trail() As SeqHC
    Set Trail = p.Trail
End Function


Public Function Revisits() As KvpHL
    Set Revisits = p.Revisits
End Function


Public Function AtOrigin() As Boolean
    AtOrigin = (Me.Location.X = Me.Origin.X) And (Me.Location.Y = Me.Origin.Y)
End Function

Public Function Manhatten() As Long
    Manhatten = Abs(p.Location.X - p.Origin.X) + Abs(p.Location.Y - p.Origin.Y)
End Function


Public Function Origin() As PointXY
    Set Origin = p.Origin
End Function

Public Function Heading() As e_Heading
    Heading = p.Heading
End Function

Public Property Get Bounds() As Long()
    Bounds = p.BoundsArray1To4
End Property

Public Property Let Bounds(ByRef ipBoundsArray1To4() As Long)
    p.BoundsArray1To4 = VBA.CVar(ipBoundsArray1To4)
    
    If VBA.IsEmpty(ipBoundsArray1To4) Then
        p.BoundsInUse = False
        Exit Property
    End If
    
    If p.Location.X < p.BoundsArray1To4(m_Minx) Or p.Location.X > p.BoundsArray1To4(m_Maxx) Or p.Location.Y < p.BoundsArray1To4(m_MinY) Or p.Location.Y > p.BoundsArray1To4(m_MaxY) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.PointXT.Let Bounds", _
            Fmt.Text("The new Bounds of {0} do not include the current coordinates ({1},{2})", ipBoundsArray1To4, p.Location.X, p.Location.Y)
    End If

End Property

'Public Sub SetBounds(ByVal ipMinx As Long, ByVal ipMinY As Long, ByVal ipMaxX As Long, ByVal ipMaxY As Long)
'    p.MinX = ipMinx
'    p.MinY = ipMinY
'    p.MaxX = ipMaxX
'    p.MaxY = ipMaxY
'End Sub

Public Function BoundsInUse() As Boolean
    BoundsInUse = p.BoundsInUse
End Function

Public Property Get ForbiddenCoords() As SeqT
    Set ForbiddenCoords = p.ForbiddenCoords
End Property

Public Property Set ForbiddenCoords(ByVal ipForbiddenCoords As SeqT)
    Set p.ForbiddenCoords = ipForbiddenCoords
    
    If ipForbiddenCoords Is Nothing Then
        p.ForbiddenCoords.Clear
        Exit Property
    End If
    
    If p.ForbiddenCoords.HoldsItem(VBA.CStr(p.Location.X) & Chars.twComma & VBA.CStr(p.Location.Y)) Then
        Err.Raise 17 + vbObjectError, _
             "VBALib.PointXY.Set Forbidden", _
            Fmt.Text("The new list of forbidden coordinates includes the current coordinate({0},{1})", p.Location.X, p.Location.Y)
    End If
    
    Set p.ForbiddenCoords = ipForbiddenCoords
    
End Property

Public Function Moved() As Boolean
    Moved = p.Moved
End Function
Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property

Private Sub AlertOutOfBOunds()
    Err.Raise 17 + vbObjectError, _
        "VBALib.PointXY: Out of bounds"
End Sub

'Public Function ISMove(ByRef ipMove As Mover, Optional ipSteps As Long = 1) As PointXY
'
'    Set Move = Me
'
'
'    p.Moved = False
'
'    Dim myX As Long: myX = p.X + ipMove.X
'    Dim myY As Long: myY = p.Y + ipMove.Y
'
'    If p.BoundsInUse And (myX < p.Bounds(m_MinX) Or myX > p.Bounds(m_MaxX) Or myY < p.Bounds(m_MinY) Or myY > p.Bounds(m_MaxY)) Then
'        Exit Function
'    End If
'
'    If p.Forbidden.HoldsItems Then
'        If p.Forbidden.HoldsItem(VBA.CStr(myX) & Chars.twComma & VBA.CStr(myY)) Then
'            Exit Function
'        End If
'    End If
'
'    p.Moved = True
'    p.X = myX
'    p.Y = myY
'
'End Function
'
