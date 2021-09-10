VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib")

Const mcLocation                    As String = "VBALib.ResultLong."

Private Type Properties

    Location                        As string
    Reason                          As GuardClause
    ReasonArgs                      As variant
    ReasonMsg                       As String
    Status                          As Boolean
    Value                           As Variant
    
End Type

Private p                           As Properties


Public Sub Class_Initialise()
    Guard InvalidNew, Types.IsNotSameInstance(Me, ResultLong), mcLocation & "ResultLong", (TypeName(Me))
End Sub

'@Description("Returns a new instance prepoulated with fail values")
Public Function Deb() As Result


    Guard UsePredeclaredIdInstance, Types.IsNotSameInstance(Me, ResultLong), mcLocation & "Class_Initialize", TypeName(Me)
    
    If Me Is ResultLong Then
    
        With New Result
        
            Set Deb = .ReadyToUseInstance()
        
        End With
    
    End If
    
End Function

Public Function ReadyToUseInstance() As Result

    Guard UsePredeclaredIdInstance, Types.IsNotSameInstance(Me, ResultLong), mcLocation & "Class_Initialize", TypeName(Me)
    
    Set ReadyToUseInstance = Me
    
End Function
    
Public Function SetStatus(ByVal ipStatus As Boolean) As Result
    p.Status = ipStatus
    Set SetStatus = Me
End Function

Public Property Get Status() As Boolean
    Status = p.Status
End Property

Public Property Let Status(ByVal ipStatus As Boolean)
    p.Status = ipStatus
End Property


Public Function SetReason(ByVal ipReason As GuardClause) As Result
    p.Reason = ipReason
    Set SetReason = Me
End Function

Public Property Get Reason() As GuardClause
    Reason = p.Reason
End Property

Public Property Let Reason(ByVal ipReason As GuardClause)
    p.Reason = ipReason
End Property


Public Function SetReasonMsg(ByVal ipReasonMsg As GuardClause) As Result
    p.ReasonMsg = ipReasonMsg
    Set SetReasonMsg = Me
End Function

Public Property Get ReasonMsg() As String
    ReasonMsg = p.ReasonMsg
End Property

Public Property Let ReasonMsg(ByVal ipReasonMsg As String)
    p.ReasonMsg = ipReasonMsg
End Property


Public Function SetValue(ByVal ipValue As Variant) As Result
    p.Value = ipValue
    Set SetValue = Me
End Function
'@DefaultMember
Public Property Get Value() As Variant
    Value = p.Value
End Property

Public Property Let Value(ByVal ipValue As Variant)
    p.Value = CLng(ipValue)
End Property


Public Function SetReasonArgs(ParamArray ipReasonArgs() As Variant) As Result
    p.ReasonArgs = ipReasonArgs
    Set SetReasonArgs = Me
End Function

Public Property Get ReasonArgs() As Variant
    ReasonArgs = p.ReasonArgs
End Property

Public Property Let ReasonArgs(ByVal ipReasonArgs As Variant)
    p.ReasonArgs = ipReasonArgs
End Property


Public Function SetLocation(ByVal ipLocation As String) As Result
    p.Location = ipLocation
    Set SetLocation = Me
End Function

Public Property Get Location() As String
    Location = p.Location
End Property

Public Property Let Location(ByVal ipLocation As String)
    p.Location = ipLocation
End Property


Public Function IsNotOkay() As Boolean
    IsNotOkay = Not p.Status
End Function

Public Function IsOkay() As Boolean
    IsOkay = p.Status
End Function

Public Sub Throw()

    Dim myMessage As String
    If VBA.IsEmpty(p.ReasonArgs) Then
        
        myMessage = p.ReasonMsg
        
    Else
        
        myMessage = Fmt.TxtArr(p.ReasonMsg, p.ReasonArgs)
        
    End If
    
    Guard p.Reason, True, p.Location, myMessage
    
End Sub

Public Function SetToNotOkay _
( _
    Optional ByVal ipReason As GuardClause = NotOkay, _
    Optional ByVal ipLocation As String = vbNullString, _
    Optional ByVal ipReasonMsg As String = vbNullString, _
    Optional ByVal ipReasonArgs As Variant _
) As Result

    p.Reason = ipReason
    p.Status = False
    p.Value = -1
    
    If VBA.Len(ipLocation) > 0 Then p.Location = ipLocation
        
    p.ReasonMsg = IIf(VBA.Len(ipReasonMsg) = 0, Enums.GuardClauses.ToString(ipReason), ipReasonMsg)
      
    Dim myargs As Variant
    myargs = IIf(Not VBA.IsMissing(ipReasonArgs), ipReasonArgs, p.ReasonArgs)
    
    p.ReasonMsg = Fmt.TxtArr(p.ReasonMsg, ipReasonArgs)
        
    Set SetToNotOkay = Me
        
End Function

Public Sub SetOkay(Optional ByVal ipValue As Long = -1)
    
    p.Status = True
    p.Reason = IsOkay
    p.ReasonMsg = "Okay"
    p.Value = ipValue
        
End Sub

Public Function Clone() As Result
    
    Dim myResult As ResultLong
    Set myResult = ResultLong.Deb
    
    With myResult
    
        .Status = p.Status
        .Reason = p.Reason
        .ReasonMsg = p.ReasonMsg
        .Value = p.Value
        .ReasonArgs = p.ReasonArgs
        .Location = p.Location
        
    End With
    
End Function