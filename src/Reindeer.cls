VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Reindeer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Helpers")
Option Explicit

Private Const FLYING As Boolean = True
Private Const NOT_FLYING As Boolean = False

Public Enum ReindeerInfo
    
    Name = 1
    FlightSpeed
    FlightTime
    RestTime
    
End Enum
    

Private Type State

    Activity                    As Boolean
    Name                        As String
    FlightSpeed                 As Long
    FlightTime                  As Long
    RestTime                    As Long
    
End Type

Private s                       As State

'@DefaultMember
Public Function Deb(ByRef ipInfo As SeqC) As Reindeer
Attribute Deb.VB_UserMemId = 0
    With New Reindeer
        Set Deb = .ConstructInstance(ipInfo)
    End With
End Function

Friend Function ConstructInstance(ByRef ipInfo As SeqC) As Reindeer
    
    s.Name = ipInfo.Item(ReindeerInfo.Name)
    s.FlightSpeed = VBA.CLng(ipInfo.Item(ReindeerInfo.FlightSpeed))
    s.FlightTime = VBA.CLng(ipInfo.Item(ReindeerInfo.FlightTime))
    s.RestTime = VBA.CLng(ipInfo.Item(ReindeerInfo.RestTime))
    s.Activity = FLYING
    
    Set ConstructInstance = Me
    
End Function

Public Function Distance(ByRef ipTime As Long) As Long
    
    Dim myresult As Long: myresult = (ipTime \ (s.FlightTime + s.RestTime)) * s.FlightSpeed * s.FlightTime
    
    Dim myRemainder As Long: myRemainder = ipTime Mod (s.FlightTime + s.RestTime)
    
    If myRemainder < s.FlightTime Then
    
        myresult = myresult + myRemainder * s.FlightSpeed
    
    Else
        
        myresult = myresult + s.FlightTime * s.FlightSpeed
        
    End If
    
    Distance = myresult
    
End Function
