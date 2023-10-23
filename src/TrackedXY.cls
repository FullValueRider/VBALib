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
' A tracked version of point XY.
' When Eightway movement is specified, a move along a diagonal counts as a 
' single step rather than two steps (! step horizontal, 1 step vertical).
' For this reason, the track only records the location after a step.
' Thus moving 3 steps NE from location 0,0 gives a track of {{0,0},{1,1},{2,2}{3,3}}
' Move NE is 1 move and two steps
Option Explicit

Private Const MY_TYPENAME                   As String = "TrackedXY"
' Move Forward and Backward are instructions rather than headings
' They are allowed in lie of headngs but as we just reuse the current heading

Private Const MOVE_FORWARD                  As String = "f"
Private Const MOVE_BACKWARD                 As String = "b"


Public Enum e_Heading
    
    m_North = 1
    m_First = m_North
    m_NE
    m_East
    m_SE
    m_South
    m_SW
    m_West
    m_NW
    m_Last = m_NW
   
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
' and which tracks where it has been and when it crossed its own Track.
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

Public Enum e_AllowedHeadings
    m_FourwaysNESW = 0
    m_First = e_AllowedHeadings.m_FourwaysNESW
    m_FourWaysDiagonals
    m_Eightways
    m_Last = e_AllowedHeadings.m_Eightways
End Enum

Private Type Properties

   ' Seq of e_Heading.  Seq used because of rotLeft and rotRight methods (See Turn)
    Heading                                     As SeqA
    'Backward                                    As SeqA
    
    AllowedHeadings                             As e_AllowedHeadings
    HeadingTextToEnum                           As KvpA
    Compass                                     As SeqA
    
    ' A set of PointXY containing the offsets to move for
    ' p.heading (p.heading is used as the index into Moves)
    Moves                                       As SeqA
    
    ' A record of the x,y coordinates in the order in which they were visited
    Track                                       As SeqHC
    
    ' xy coordinates vs seq of the step count when the coordinate was revisited.
    Revisits                                    As KvpHL
  
    ' Count of steps taken after last move command
    StepsTaken                                       As Long
    
    ' True if the move command results in a change of Heading
    HasTurned                                      As Boolean
'
'    ' rectangular area outside of which the point may not go
'    ' a seqA of SeqA which defines a rectangular area
'    Area As SeqA

    ' Bounds outside which the point cannot move
    ' set bounds method. If an area is set then the setbounds is ignored
    
    BoundsArray1To4                             As Variant
    ' true if bounds are in force
    BoundsInUse                                 As Boolean
    
    ' a sequence of x,y strings defining points which cannot be visited.
    ForbiddenCoords                             As SeqT
 
    'set to true if the current X,y location has been visited previously
    Intersection                                As Boolean
    
End Type

Private p As Properties
 
'@DefaultMember
Public Function Deb _
( _
    Optional ByRef ipPoint As PointXY = Nothing, _
    Optional ByRef ipForbiddenCoords As SeqT = Nothing, _
    Optional ByRef ipBoundsArray As Variant = Empty, _
    Optional ByRef ipAllowedHeadings As e_AllowedHeadings = m_EightWays _
) As TrackedXY ' Optional ByRef ipForbiddenvalues As SeqT = Nothing, _ ptional ByRef ipArea As SeqA = Nothing, _
    With New TrackedXY
Attribute Deb.VB_UserMemId = 0
    With New TrackedXY
        Set Deb = .ConstructInstance(ipPoint, ipForbiddenCoords, ipBoundsArray, ipAllowedHeadings)', ipForbiddenvalues
    End With
End Function


Public Function ConstructInstance _
( _
    ByRef ipPoint As PointXY, _
    ByRef ipForbiddenCoords As SeqT, _
    ByRef ipBounds As Variant, _
    ByRef ipAllowedHeadings As e_AllowedHeadings _
) As TrackedXY 'Optional ByRef ipForbiddenvalues As SeqT = Nothing, _,  'Optional ByRef ipArea As SeqA = Nothing, _

    If IsNothing(ipPoint) Then
        Set ipPoint = PointXY(0, 0)
    End If
    
    If IsNothing(ipForbiddenCoords) Then
        Set p.ForbiddenCoords = SeqT.Deb
    Else
        Set p.ForbiddenCoords = ipForbiddenCoords
    End If
    
    If p.ForbiddenCoords.HoldsItem(ipPoint) Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.TrackedXY", _
            Fmt.Text("The initial point '{0}' is a forbidden Point", ipPoint)
    End If
    
    p.BoundsInUse = False
    If ArrayOp.HoldsItems(ipBounds) Then
        If ArrayOp.Count(ipBounds) = 4 Then
            p.BoundsArray1To4 = ipBounds
            ReDim p.BoundsArray1To4(1 To 4)
            p.BoundsInUse = True
        Else
            Err.Raise 17 + vbObjectError, _
                "VBALib.TrackedPoint", _
                Fmt.Text("The ipBounds array should contain 4 elements.  Found {0}", ArrayOp.Count(ipBounds))
        End If
    End If

    Allow.InRange ipAllowedHeadings, e_AllowedHeadings.m_First, e_AllowedHeadings.m_Last, "TrackedXY.ConstructInstance"
    p.AllowedHeadings = ipAllowedHeadings
    
    p.StepsTaken = 0
    p.HasTurned = False
    
    InitialiseHeadingTextToEnum
    
    ' We generate a set of adjacent coordinates as PointXY instances
    ' we use the m_Relative argument to get offsets from the current xy position
    ' These offsets are used to move the current xy position 1 step in the heading direction
    ' p.location.x = p.location.x + p.moves.Item(m_North).X
    ' p.location.y = p.location.y + p.moves.Item(m_North).Y
    Set p.Moves = ipPoint.AdjacentCoords(ipAdjacentType:=m_Relative)
    
    ' Seq are used so that we can use the rotleft and rotright methods
    Set p.Heading = SeqA(m_North, m_NE, m_East, m_SE, m_South, m_SW, m_West, m_NW)
    'Set p.Backward = SeqA(m_South, m_SW, m_West, m_NW, m_North, m_NE, m_East, m_SE)

    Set p.Track = SeqHC.Deb.AddItems(ipPoint)
    Set p.Revisits = KvpHL.Deb
    
    Set ConstructInstance = Me

End Function
 
Private Sub InitialiseHeadingTextToEnum()

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

       ' The enums below will be ignored but the keys are needed in the Move Method
       .Add "f", -1
       .Add "forward", -1
       .Add "Forwards", -1

       .Add "b", -2
       .Add "back", -2
       .Add "backwards", -2
    End With
End Sub

'@("Description("Returns the movement enum (m_EightWays, m_FourWaysNESW or m_FourWaysDiagaonals")
Public Property Get AllowedHeadings() As e_AllowedHeadings
    AllowedHeadings = p.AllowedHeadings
End Property

'@Description("Set the allowed directions for movement.  The next clockwise direction is selected when fourways directions do not match the current heading")
Public Property Let AllowedHeadings(ByVal ipAllowedHeadings As e_AllowedHeadings)
Attribute AllowedHeadings.VB_Description = "Set the allowed directions for movement.  The next clockwise direction is selected when fourways directions do not match the current heading"

    Allow.InRange ipAllowedHeadings, e_Heading.m_First, e_Heading.m_Last, MY_TYPENAME & ".Let_Move"
    
    ' If the alowed headings are fourways and the current heading
    ' is not one of the fourways set we turn clockwise 1 step
    Select Case ipAllowedHeadings
        
        Case m_FourwaysNESW
        
            Select Case p.Heading.First
                Case m_NE, m_SE, m_SW, m_NW
                    p.Heading.RotLeft
                    
            End Select
            
                
        Case m_FourWaysDiagonals
        
            Select Case p.Heading.First
                Case m_North, m_East, m_South, m_West
                    p.Heading.RotLeft
                    
            End Select
            
    End Select
    
End Property


' Private Function BackwardsHeading(ByVal ipHeading As e_Heading) As e_Heading

'     ipHeading = ipHeading + 4
'     If ipHeading > m_EightWays_Last Then
'         ipHeading = ipHeading - m_EightWays_Last
'     End If
'     BackwardsHeading = ipHeading
    
' End Function

Public Property Get Moves() As SeqA
    Set Moves = p.Moves
End Property

Public Property Set Moves(ByVal ipMoves As SeqA)
    Set p.Moves = ipMoves
End Property

Public Function Turn(ByRef ipTurn As String) As TrackedXY
    
    Set Turn = Me
    Dim myTurn As String: myTurn = VBA.Trim$ (VBA.LCase$(ipTurn))
    Allow.WhenTrue VBA.InStr("l,r,left,right", myTurn) <> 0, MY_TYPENAME & ".Turn", Fmt.Text("Invalid Turn command. Got {0}", ipTurn)
    
    Dim myEightWaysPoints As Long: myEightWaysPoints = IIf(p.AllowedHeadings = m_Eightways, 1, 2)
    
    Select Case myTurn
    
        Case "l", "left":            p.Heading.RotRight myEightWaysPoints
        Case "r", "right":            p.Heading.RotLeft myEightWaysPoints
            
    End Select
    
    p.HasTurned = True
          
End Function


'@Description("Moves the point the specified number of steps in the Specified direction.  If no heading is specified the point moves forward. ")
Public Function Move(Optional ByVal ipSteps As Long = 1, Optional ByRef ipHeading As String = MOVE_FORWARD) As TrackedXY

    Set Move = Me
    Dim myBackward As Boolean
    ' Convert heading text to heading enum
    ' and set the myBackwards flag
    Dim myHeadingStr As String: myHeadingStr = VBA.Trim$(VBA.LCase$(ipHeading))
    Dim myHeadingEnum As e_Heading
    Select Case myHeadingStr
    
        Case "f", "Forward", "forwards"
            myHeadingEnum = p.Heading.First
            myBackward = False
        
        Case "b", "back", "backward", "backwards"
            myHeadingEnum = p.Heading.First
            myBackward = True
            
        Case Else
        
            Allow.WhenTrue p.HeadingTextToEnum.HoldsKey(myHeadingStr), MY_TYPENAME & ".Move", Fmt.Text("Unknown heading.  Got {0}", ipHeading)
        
            myHeadingEnum = p.HeadingTextToEnum.Item(myHeadingStr)
            
            'Rotate Headings to match currentheadings
            ' so that the new direction is preserved for the next move/turn
            ' Also set the hasTurned flag
            If myHeadingEnum = p.Heading.First Then
                p.HasTurned = False
            Else
                p.Heading.RotLeft myHeadingEnum - p.Heading.First
                p.HasTurned = True
            End If
            
    End Select
    
    ' Determine if myHeading is one of the headings allowed by p.Allowedeadings
    ' if it isn't do one turn clockwise to align p.heading with p.allowedheadings
    ' this extra rotation is only needed if p.AllowedHeadings is Fourways
    Select Case p.AllowedHeadings
        
        Case e_AllowedHeadings.m_FourwaysNESW
        
            Select Case myHeadingStr
                Case m_NE, m_SE, m_SW, m_NW
                    p.Heading.RotLeft
                    
            End Select
        
        Case e_AllowedHeadings.m_FourWaysDiagonals
            
            Select Case myHeadingEnum
                 Case m_North, m_East, m_South, m_West
                     p.Heading.RotLeft
                 
            End Select
        
    End Select
    
    If ipSteps = 0 Then
        Exit Function
    End If
    
    Dim myStep As Long
    For myStep = 1 To VBA.Abs(ipSteps)
        
        Dim myX As Long
        Dim myY As Long
        If myBackward Then
        
            myX = p.Track.Last.x - p.Moves.Item(p.Heading.First).x
            myY = p.Track.Last.Y - p.Moves.Item(p.Heading.First).Y
        Else
            myX = p.Track.Last.x + p.Moves.Item(myHeadingEnum).x
            myY = p.Track.Last.Y + p.Moves.Item(myHeadingEnum).y
        End If
         
        'if the next point is out of bounds then exit
        If Not WithinBounds(myX, myY) Then
            Exit Function
        End If
        
        ' check location is not a forbidden point
        Dim myPoint As PointXY: Set myPoint = PointXY(myX, myY)
        ' is it a forbidden location
        If p.ForbiddenCoords.HoldsItem(myPoint) Then
            Exit Function
        End If
        
       
        ' check if we are at an intersection
        p.Intersection = False
        
        If p.Track.HoldsItem(myPoint) Then
            p.Intersection = True
            
            If p.Revisits.LacksItem(myPoint) Then
            
                p.Revisits.Add myPoint, SeqA.Deb
                
            End If
            
            p.Revisits.Item(myPoint).Add p.Track.LastIndex + 1
        
        End If
        
        ' we can now update SETPS MOVED
        p.StepsTaken = p.StepsTaken + 1

        ' now we can finally add myPpoint to the track
        p.Track.Add myPoint

    Next
    
End Function


Private Function WithinBounds(ByRef ipX As Long, ByRef ipY As Long) As Boolean
	
    WithinBounds = True

    If Not p.BoundsInUse Then
    	Exit Function
    End If
    
    If ipX < p.BoundsArray1To4(m_Minx) Or ipX > p.BoundsArray1To4(m_Maxx) Or ipY < p.BoundsArray1To4(m_MinY) Or ipY > p.BoundsArray1To4(m_MaxY) Then
    	WithinBounds = False
        Exit Function
    End If
   
End Function


Public Function Location() As PointXY
    Set Location = p.Track.Last
End Function


Public Function Steps() As Long
    Steps = p.Track.Count - 1
End Function


Public Function Track() As SeqHC
    Set Track = p.Track
End Function


Public Function Revisits() As KvpHL
    Set Revisits = p.Revisits
End Function


Public Function AtOrigin() As Boolean
    AtOrigin = p.Track.First.ToString = p.Track.Last.ToString
End Function

'The manhatten distance between the origin and the current location or the provided location")
Public Function Manhatten(Optional ByRef ipPoint As PointXY = Nothing) As Long
    
    If IsNothing(ipPoint) Then
        Manhatten = p.Track.First.Manhatten(p.Track.Last)
    Else
        Manhatten = (p.Track.First.Manhatten(ipPoint))
    End If
    
End Function


Public Function Origin() As PointXY
    Set Origin = p.Track.First
End Function

Public Function Heading() As e_Heading
    Heading = p.Heading.First
End Function

Public Property Get Bounds() As Long()
    Bounds = p.BoundsArray1To4
End Property

Public Property Let Bounds(Optional ByRef ipBoundsArray1To4 As Variant = Empty)
	
    If VBA.IsEmpty(ipBoundsArray1To4) Then
        p.BoundsInUse = False
        Exit Property
    End If
    
    If LBound(ipBoundsArray1To4) <> 1 And UBound(ipBoundsArray1To4) <> 4 Then
        Err.Raise 17 + vbObjectError, _
            MY_TYPENAME & ".Let_Bounds", _
            Fmt.Text("Expecting bound of 1 and 4.  Got {0} and {1}", LBound(ipBoundsArray1To4), UBound(ipBoundsArray1To4))
    End If
    
    p.BoundsArray1To4 = VBA.CVar(ipBoundsArray1To4)
    
    Dim myCount As Long:myCount = 0
    Dim myPoints As IterItems: Set myPoints = IterItems(p.Track)
    Do
        Dim myPoint As PointXY: Set myPoint = myPoints.CurItem(0)

        If myPoint.X < p.BoundsArray1To4(m_Minx) Or myPoint.X > p.BoundsArray1To4(m_Maxx) Or myPoint.Y < p.BoundsArray1To4(m_MinY) Or myPoint.Y > p.BoundsArray1To4(m_MaxY) Then
            myCount = myCount + 1
        End If
    Loop While myPoints.MoveNext
    
    If myCount > 0 Then
        Err.Raise 17 + vbObjectError, _
            "VBALib.PointXT.Let Bounds", _
            Fmt.Text("The new Bounds exclude {0} points in the current track", myCount)
    End If

End Property

Public Function BoundsInUse() As Boolean
    BoundsInUse = p.BoundsInUse
End Function

Public Property Get ForbiddenCoords() As SeqT
    Set ForbiddenCoords = p.ForbiddenCoords
End Property

''@Description("Adds forbiddenpoint if not already in the Track. True if point added.")
Public Function AddForbiddenCoord(ByRef ipPoint As PointXY) As Boolean
    
    AddForbiddenCoord = False
    If p.Track.HoldsItem(ipPoint) Then
        Exit Function
    End If
    p.ForbiddenCoords.Add ipPoint
    AddForbiddenCoord = True
End Function

Public Property Set ForbiddenCoords(ByVal ipForbiddenCoords As SeqT)
    
    If IsNothing(ipForbiddenCoords) Then
        p.ForbiddenCoords.Clear
        Exit Property
    End If
    
    If p.Track.SetOf(m_Common, ipForbiddenCoords).HoldsItems Then
        Err.Raise 17 + vbObjectError, _
             "VBALib.PointXY.Set Forbidden", _
            "The new list of forbidden coordinates includes coordinates in the current track"
    End If
    
    Set p.ForbiddenCoords = ipForbiddenCoords
    
End Property

Public Function StepsTaken() As Long
    StepsTaken = p.StepsTaken
End Function

Public Function HasTurned() As Boolean
    HasTurned = p.HasTurned
end function

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property