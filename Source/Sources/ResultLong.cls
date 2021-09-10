VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResultLong"
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
    Value                           As Long
    
End Type

Private p                           As Properties


Public Sub Class_Initialise()
    Guard InvalidNew, Types.IsNotSameInstance(Me, ResultLong), mcLocation & "ResultLong", (TypeName(Me))
End Sub

'@Description("Returns a new instance prepoulated with fail values")
Public Function Deb() As ResultLong


    Guard UsePredeclaredIdInstance, Types.IsNotSameInstance(Me, ResultLong), mcLocation & "Class_Initialize", TypeName(Me)
    
    If Me Is ResultLong Then
    
        With New ResultLong
        
            Set Deb = .ReadyToUseInstance()
        
        End With
    
    End If
    
End Function

Public Function ReadyToUseInstance() As ResultLong

    Guard UsePredeclaredIdInstance, Types.IsNotSameInstance(Me, ResultLong), mcLocation & "Class_Initialize", TypeName(Me)
    
    Set ReadyToUseInstance = Me
    
End Function
    
    
' All properties are read/write
' This is because
'  set opresult = ResultLong.deb
'  Set tryget<thing> = opResult
' assigns to tryGet<thing> the reference of the object created
' by resultLong.deb and not a reference to opResult.
'
' This if we then do
' set opresult = resultlong.SetSuccess(value)
' then opresult and TryGet<thing> refer to different objects
' TryGet<thing> will not refer to the object created by resultlong.SetSuccess
' but will keep the reference to resultlong.deb

'@DefaultMember
Public Property Get Status() As Boolean
    Status = p.Status
End Property

Public Property Let Status(ByVal ipStatus As Boolean)
    p.Status = ipStatus
End Property


Public Property Get Reason() As GuardClause
    Reason = p.Reason
End Property

Public Property Let Reason(ByVal ipReason As GuardClause)
    p.Reason = ipReason
End Property


Public Property Get ReasonMsg() As String
    ReasonMsg = p.ReasonMsg
End Property

Public Property Let ReasonMsg(ByVal ipReasonMsg As String)
    p.ReasonMsg = ipReasonMsg
End Property


Public Property Get Value() As Long
    Value = p.Value
End Property

Public Property Let Value(ByVal ipValue As Long)
    p.Value = CLng(ipValue)
End Property


Public Property Get ReasonArgs() As Variant
    ReasonArgs = p.ReasonArgs
End Property

Public Property Let ReasonArgs(ByVal ipReasonArgs As Variant)
    p.ReasonArgs = ipReasonArgs
End Property


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
    If Types.Iterable.LacksItems(p.ReasonArgs) Then
        
        myMessage = p.ReasonMsg
        
    Else
        
        myMessage = Fmt.TxtArr(p.ReasonMsg, p.ReasonArgs)
        
    End If
    
    Guard p.Reason, True, p.Location, myMessage
    
End Sub

Public Function SetReasonArgs(ParamArray ipargs() As Variant) As ResultLong
    
    p.ReasonArgs = ipargs
    Set SetReasonArgs = Me
    
End Function

Public Function SetNotOkay _
( _
    Optional ByVal ipReason As GuardClause = NotOkay, _
    Optional ByVal ipLocation As String = vbNullString, _
    Optional ByVal ipReasonMsg As String = vbNullString, _
    Optional ByVal ipReasonArgs As Variant _
) As ResultLong

    p.Reason = ipReason
    p.Status = False
    p.Value = -1
    
    If VBA.Len(ipLocation) > 0 Then p.Location = ipLocation
        
    p.ReasonMsg = IIf(VBA.Len(ipReasonMsg) = 0, Enums.GuardClauses.ToString(ipReason), ipReasonMsg)
      
    Dim myargs As Variant
    myargs = IIf(Not VBA.IsMissing(ipReasonArgs), ipReasonArgs, p.ReasonArgs)
    
    p.ReasonMsg = Fmt.TxtArr(p.ReasonMsg, ipReasonArgs)
        
    Set SetNotOkay = Me
        
End Function

Public Sub SetOkay(Optional ByVal ipValue As Long = -1)
    
    p.Status = True
    p.Reason = IsOkay
    p.ReasonMsg = "Okay"
    p.Value = ipValue
        
End Sub

Public Function Clone() As ResultLong
    
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