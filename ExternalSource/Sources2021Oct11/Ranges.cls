VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ranges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed

Private Const mcLocation                        As String = "VBALib.Ranges."
Public Enum ValidateRangeReason
    
    Status As Boolean
    StartIndex As Long
    Run As Long
    
End Enum

'@Ignore AssignedByValParameter
Public Function GetSeries(ByVal ipstart As Variant, ByVal ipCount As Long, Optional ByVal ipStep As Variant = 1) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set GetSeries = myLyst
    
    If Types.Group.IsNotNumber(ipstart, ipCount, ipStep) Then Exit Function
    
    Dim myItem As Variant
    For myItem = 0 To ipCount - 1
    
        myLyst.Add ipstart
        ipstart = ipstart + ipStep
        
    Next
    
End Function


Public Function GetConstSeries(ByVal ipCount As Long, ByVal ipValue As Variant) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set GetConstSeries = myLyst
    
    Dim myItem As Long
    For myItem = 0 To ipCount - 1
    
        myLyst.Add ipValue
        
    Next
    
End Function


'@Description("Returns a ValidationResult of Status, StartIndex and Run. The startindex and endindex are offsets using a zerobased index. Negative values count from higher Index to lower Index. ")
Friend Function ValidateRange _
( _ 
    ByVal ipItem As Variant, _
    ByVal ipStartIndex As Long, _
    Optional ByVal ipRun As Variant, _
    Optional ipEndIndex As Variant, _
    Optional ByVal ipRank As Long = 1 _
) As ValidationResult

    ' To emulate other programming languages, values
    ' for the start end or run can be negative to indicate counting down from a high to low index.
    ' e.g. in terms of this library an ipstartindex of -1 would correspond to the LastIndex property.
    '
    ' This method takes start, run and end values and calculatess a start and run
    ' based on ipStartIndex being adjusted to be a positive offset from FirstIndex
    ' and run increasing the positive offset from startindex
    
    Dim myValidationResult As ValidationResult
    myValidationResult.Status = False
    
    ' Get the extend of ipItem
    ' handle string seperately to Iterables
    Dim myItemFirstIndex As Long
    Dim myItemLastIndex As Long
    Dim myItemCount As Long
    
    If Types.Group.IsNotIterable(ipItem) And Types.IsNotString(ipItem) Then Exit Function
    If Types.IsString(ipItem) Then
         
        myItemFirstIndex = 1
        myItemLastIndex = VBA.Len(ipItem)
        myItemCount = myItemLastIndex
            
        
    Else
        If Not Types.Iterable.TryFirstIndex(ipItem, myItemFirstIndex) Then Exit Function
        
        myItemCount = Types.Iterable.Count(ipItem)
        
        
    End If
    
    
    If Types.IsMissing(ipRun, ipEndIndex) Then
        
        myValidationResult.Run = myItemLastIndex - myValidationResult.StartIndex
        
    ElseIf Types.IsNotMissing(ipRun) Then
        
        
    Else
        
 
    End If
    
    
End Function

Friend Function HelperGetPositiveStartRunFromAnyStartRunEnd _
    ( _
        Optional ByVal ipStartIndex As Long = 1, _
        Optional ByVal ipRun As Long = 0, _
        Optional ByVal ipEnd As Long = 0 _
    ) As ValidationResult
    
        
        Dim myStart As Long
        myStart = ipStartIndex
        If myStart = 0 Then myStart = 1
        
        If ipEnd = 0 And ipRun = 0 Then
        
            Set HelperGetPositiveStartRunFromAnyStartRunEnd = HelperGetPositiveStartRunToEndOfString(myStart)
            Exit Function
            
        End If
        
        If Abs(ipRun) > 0 Then
        
            Set HelperGetPositiveStartRunFromAnyStartRunEnd = HelperGetPositiveStartRunFromStartRun(myStart, ipRun)
            Exit Function
            
        End If
        
        Set HelperGetPositiveStartRunFromAnyStartRunEnd = HelperGetPositiveStartRunFromStartEnd(myStart, ipEnd)
        
    End Function
        
    
    Public Function HelperGetPositiveStartRunToEndOfString(ByVal ipStart As Long) As StartRun
    ' Used by HelperGetPositiveStartRunFromAnyStartRunEnd
        Dim myStartRun As StartRun
        Set myStartRun = New StartRun
        
        If ipStart < 0 Then
        
            myStartRun.Start = 1
            myStartRun.Run = VBA.Len(p.Value) + ipStart + 1
            
        Else
        
            myStartRun.Start = ipStart
            myStartRun.Run = VBA.Len(p.Value) - ipStart + 1
            
        End If
        
        Set HelperGetPositiveStartRunToEndOfString = myStartRun
        
    End Function
    
    
    Public Function HelperGetPositiveStartRunFromStartRun(ByVal ipStart As Long, ByVal ipRun As Long) As StartRun
    'used by HelperGetPositiveStartRunFromAnyStartRunEnd
        Dim localIpStart As Long
        localIpStart = ipStart
        Dim myStartRun As StartRun
        Set myStartRun = New StartRun
        
        If localIpStart < 0 Then
        
            localIpStart = VBA.Len(p.Value) + localIpStart + 1
            
        End If
          
        If ipRun < 0 Then
        
            myStartRun.Start = localIpStart + ipRun + 1
            myStartRun.Run = Abs(ipRun)
            
        Else
        
            myStartRun.Start = localIpStart
            myStartRun.Run = ipRun
            
        End If
        
        Set HelperGetPositiveStartRunFromStartRun = myStartRun
    
    End Function
    
    
    Friend Function GetPositiveStartRunFromStartEnd(ByVal ipStartIndex As Long, ByVal ipEnd As Long) As StartRun
    'used by HelperGetPositiveStartRunFromAnyStartRunEnd
        Dim localIpEnd As Long
        localIpEnd = ipEnd
        Dim localIpStart As Long
        localIpStart = ipStart
        Dim myStartRun As StartRun
        Set myStartRun = New StartRun
        
        If localIpStart < 0 Then
        
            localIpStart = VBA.Len(p.Value) + localIpStart + 1
            
        End If
        
        If localIpEnd < 0 Then
        
            localIpEnd = VBA.Len(p.Value) + localIpEnd + 1
            
        End If
        
        If localIpStart < localIpEnd Then
        
            myStartRun.Start = localIpStart
            myStartRun.Run = localIpEnd - localIpStart + 1
        
        Else
        
            myStartRun.Start = localIpEnd - localIpStart + 1
            myStartRun.Run = localIpEnd
        
        End If
        
        Set HelperGetPositiveStartRunFromStartEnd = myStartRun
        
    End Function