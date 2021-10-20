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
'@Folder("VBALib")

Private Type InternalConstants

    Location                        As String
    

End Type

Private ic                          As InternalConstants

Public Enum ValidateRangeReason
    
    Status
    StartIndex
    Run
    
    
End Enum

Private Sub Class_Initialize()
    ic.Location = "VBALib.Ranges."
End Sub

'@Ignore AssignedByValParameter
Public Function GetSeries(ByVal ipStart As Variant, ByVal ipCount As Long, Optional ByVal ipStep As Variant = 1) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    Set GetSeries = myLyst
    
    If Types.Group.IsNotNumber(ipStart, ipCount, ipStep) Then Exit Function
    
    Dim myItem As Variant
    For myItem = 0 To ipCount - 1
    
        myLyst.Add ipStart
        ipStart = ipStart + ipStep
        
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


'@Description("Returns a ValidationResult of Status, StartIndex and Run. StartIndex and EndIndex use the indexing of the Item. Negative values count from higher Index to lower Index. ")
Friend Function ParseStartRunEndToStartRun _
( _
    ByVal ipItem As Variant, _
    ByVal ipStartIndex As Long, _
    Optional ByVal ipRun As Variant, _
    Optional ipEndIndex As Variant, _
    Optional ByVal ipRank As Long = 1 _
) As Result
Attribute ParseStartRunEndToStartRun.VB_Description = "Returns a ValidationResult of Status, StartIndex and Run. StartIndex and EndIndex use the indexing of the Item. Negative values count from higher Index to lower Index. "

    ' To emulate other programming languages, values
    ' for the start end or run can be negative to indicate counting down from a high to low index.
    ' e.g. in terms of this library an ipstartindex of -1 would correspond to the LastIndex property.
    '
    ' This method takes start, run and end values and calculatess a start and run
    ' based on ipStartIndex being adjusted to be a positive offset from FirstIndex
    ' and run increasing the positive offset from startindex
    
    Dim myStartRun As Result
    Set myStartRun = Result.Deb
    Set ParseStartRunEndToStartRun = myStartRun
    
    Dim myItemExtent As Result
    Set myItemExtent = GetItemExtent(ipItem, ipStartIndex, ipRun, ipEndIndex, ipRank)
    
    ' If the startIndex is outside of the item span then there
    ' is no legitimate starting point so exit function
    If ipStartIndex < myItemExtent.Item(Extent.IsSpanFirstIndex) Then Exit Function
    If ipStartIndex > myItemExtent.Item(Extent.IsLastIndex) Then Exit Function
        
    ' normalise the startindex
    myStartRun.Item(StartRun.StartIndex) = GetNormalisedStartIndex(myItemExtent, ipStartIndex)
    
    If Types.IsMissing(ipRun, ipEndIndex) Then
        ' by default the run is from startindex to the end of the range
        myStartRun.Item(StartRun.IsRun) = myItemExtent.Item(Extent.IsLastIndex) - myStartRun.Item(StartRun.StartIndex) + 1
        
        
    ElseIf Types.IsNotMissing(ipRun) Or Types.IsNotMissing(ipRun, ipEndIndex) Then
        'If a Run Is provided then any value For endindex is ignored
        ' If run is negative adjust the startindex accordingly
        If ipRun < 0 Then
            
            myStartRun.Item(StartRun.StartIndex) = myStartRun.Item(StartRun.StartIndex) - VBA.Abs(ipRun) + 1
            

        End If
        
        ' make sure that the run is always positive after adjusting startindex
        myStartRun.Item(StartRun.IsRun) = VBA.Abs(ipRun)
        
        ' Check that adjusting startindex for a negative run hasn't put it outside of
        ' first to last index if it is then adjust startindex to first index
        
       ' If myStartRunEndResult.Item(StartRunEndResult.IsStartIndex) > myItemExtent.Item(ExtentResult.IsLastIndex) Then Exit Function
            
        If myStartRun.Item(StartRun.StartIndex) < myItemExtent.Item(Extent.IsFirstIndex) Then
            
            myStartRun.Item(StartRun.StartIndex) = myItemExtent.Item(Extent.IsFirstIndex)
            
            
        End If
        
        myStartRun.Status = True
        
    Else ' we are left with ipstartindex and ipendindex
        
        myStartRun.Status = True
        
        ' deal with end index being outside the extent span
        Dim myEndIndex As Long
        myEndIndex = ipEndIndex
        If ipEndIndex < myItemExtent.Item(Extent.IsSpanFirstIndex) Then myEndIndex = myItemExtent.Item(Extent.IsSpanFirstIndex)
        If ipEndIndex > myItemExtent.Item(Extent.IsLastIndex) Then myEndIndex = myItemExtent.Item(Extent.IsLastIndex)
            
   
        'If endindex is less then the original startindex we need to normalise
        If myEndIndex < ipStartIndex Then
            
            myStartRun.Item(StartRun.IsRun) = VBA.Abs(myEndIndex - ipStartIndex) + 1
            myStartRun.Item(StartRun.StartIndex) = myEndIndex

        Else
            
            myStartRun.Item(StartRun.IsRun) = ipStartIndex - myEndIndex + 1
            
        End If
        
        
    End If
    
    
End Function

   

Friend Function GetNormalisedStartIndex(ByVal ipItemExtent As Result, ipStartIndex As Long) As Long
   
    ' If the startindex is less then first index adjust
    ' to be a value counted from first index
    If ipStartIndex < ipItemExtent.Item(Extent.IsFirstIndex) Then
        
        GetNormalisedStartIndex = ipStartIndex - ipItemExtent.Item(Extent.IsFirstIndex)
        
    Else
        
        GetNormalisedStartIndex = ipStartIndex
        
    End If
    
End Function

Friend Function GetItemExtent _
    ( _
        ByVal ipItem As Variant, _
        ByVal ipStartIndex As Long, _
        Optional ByVal ipRun As Variant, _
        Optional ipEndIndex As Variant, _
        Optional ByVal ipRank As Long = 1 _
    ) As Result
    
    Dim myExtentResult As Result
    Set myExtentResult = Result.Deb
    Set GetItemExtent = myExtentResult
    
    ' handle string seperately to Iterables
   ' Dim myItemFirstIndex As Long
   ' Dim myItemLastIndex As Long
   ' Dim myItemCount As Long
    'the span variable gives the most negative range based on counting down
    ' from firstindex.  This if firstindex = -5 and count is 10 then spanfirstindex will be -15
    ' because 10 below -5 is -6 to -15
    'Dim myItemSpanFirstIndex As Long
    
    If Types.Group.IsNotIterable(ipItem) And Types.IsNotString(ipItem) Then Exit Function
    If Types.IsString(ipItem) Then
         
        With myExtentResult
        
            .Item(Extent.IsFirstIndex) = 1
            .Item(Extent.IsLastIndex) = VBA.Len(ipItem)
            .Item(Extent.IsCount) = VBA.Len(ipItem)
            
        
        End With
        
    Else
    
        If Types.Iterable.TryExtent(ipItem, myExtentResult, ipRank).Failed Then Exit Function
            
        
    End If
    
    With myExtentResult
    
        .Item(Extent.IsSpanFirstIndex) = .Item(Extent.IsFirstIndex) - .Item(Extent.IsCount) + 1
        
        
    End With
    
End Function

'Friend Function HelperGetPositiveStartRunFromAnyStartRunEnd _
'    ( _
'        Optional ByVal ipStartIndex As Long = 1, _
'        Optional ByVal ipRun As Long = 0, _
'        Optional ByVal ipEnd As Long = 0 _
'    ) As ValidationResult
'
'
'        Dim myStart As Long
'        myStart = ipStartIndex
'        If myStart = 0 Then myStart = 1
'
'        If ipEnd = 0 And ipRun = 0 Then
'
'            Set HelperGetPositiveStartRunFromAnyStartRunEnd = HelperGetPositiveStartRunToEndOfString(myStart)
'            Exit Function
'
'        End If
'
'        If Abs(ipRun) > 0 Then
'
'            Set HelperGetPositiveStartRunFromAnyStartRunEnd = HelperGetPositiveStartRunFromStartRun(myStart, ipRun)
'            Exit Function
'
'        End If
'
'        Set HelperGetPositiveStartRunFromAnyStartRunEnd = HelperGetPositiveStartRunFromStartEnd(myStart, ipEnd)
'
'    End Function
'
'
'    Public Function HelperGetPositiveStartRunToEndOfString(ByVal ipStart As Long) As StartRun
'    ' Used by HelperGetPositiveStartRunFromAnyStartRunEnd
'        Dim myStartRun As StartRun
'        Set myStartRun = New StartRun
'
'        If ipStart < 0 Then
'
'            myStartRun.Start = 1
'            myStartRun.Run = VBA.Len(p.Value) + ipStart + 1
'
'        Else
'
'            myStartRun.Start = ipStart
'            myStartRun.Run = VBA.Len(p.Value) - ipStart + 1
'
'        End If
'
'        Set HelperGetPositiveStartRunToEndOfString = myStartRun
'
'    End Function
'
'
'    Public Function HelperGetPositiveStartRunFromStartRun(ByVal ipStart As Long, ByVal ipRun As Long) As StartRun
'    'used by HelperGetPositiveStartRunFromAnyStartRunEnd
'        Dim localIpStart As Long
'        localIpStart = ipStart
'        Dim myStartRun As StartRun
'        Set myStartRun = New StartRun
'
'        If localIpStart < 0 Then
'
'            localIpStart = VBA.Len(p.Value) + localIpStart + 1
'
'        End If
'
'        If ipRun < 0 Then
'
'            myStartRun.Start = localIpStart + ipRun + 1
'            myStartRun.Run = Abs(ipRun)
'
'        Else
'
'            myStartRun.Start = localIpStart
'            myStartRun.Run = ipRun
'
'        End If
'
'        Set HelperGetPositiveStartRunFromStartRun = myStartRun
'
'    End Function
'
'
'    Friend Function GetPositiveStartRunFromStartEnd(ByVal ipStartIndex As Long, ByVal ipEnd As Long) As StartRun
'    'used by HelperGetPositiveStartRunFromAnyStartRunEnd
'        Dim localIpEnd As Long
'        localIpEnd = ipEnd
'        Dim localIpStart As Long
'        localIpStart = ipStart
'        Dim myStartRun As StartRun
'        Set myStartRun = New StartRun
'
'        If localIpStart < 0 Then
'
'            localIpStart = VBA.Len(p.Value) + localIpStart + 1
'
'        End If
'
'        If localIpEnd < 0 Then
'
'            localIpEnd = VBA.Len(p.Value) + localIpEnd + 1
'
'        End If
'
'        If localIpStart < localIpEnd Then
'
'            myStartRun.Start = localIpStart
'            myStartRun.Run = localIpEnd - localIpStart + 1
'
'        Else
'
'            myStartRun.Start = localIpEnd - localIpStart + 1
'            myStartRun.Run = localIpEnd
'
'        End If
'
'        Set HelperGetPositiveStartRunFromStartEnd = myStartRun
'
'    End Function

