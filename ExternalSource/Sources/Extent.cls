VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Extent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = TRue
Attribute VB_Exposed = True
Option Explicit
'@Folder("VBALib")
'@PredeclaredId
'@Exposed
'



Public Function IsFirstIndex() As Long
    FirstIndex = 0
End Function


Public Function IsLBound() As Long
    IsLBound = 0
End Function

Public Function IsLastIndex() As Long
    IsLastIndex = 1
End Function


Public Function IsUBound() As Long
    IsUBound = 1
End Function

Public Function IsCount() As Long
    IsCount = 2
End Function


Public Function IsSpanFirstIndex() As Long
    IsSpanFirstIndex = 3
End Enum


'@Description("Returns a ValidationResult of Status, StartIndex and Run. StartIndex and EndIndex use the indexing of the Item. Negative values count from higher Index to lower Index. ")
Friend Function ParseStartRunEndToStartRun _
( _
    ByVal ipItem As Variant, _
    ByVal ipStartIndex As Long, _
    Optional ByVal ipRun As Variant, _
    Optional ipEndIndex As Variant, _
    Optional ByVal ipRank As Long = 1 _
) As Result

    ' To emulate other programming languages, values
    ' for the start end or run can be negative to indicate counting down from a high to low index.
    ' e.g. in terms of this library an ipstartindex of -1 would correspond to the LastIndex property.
    '
    ' This method takes start, run and end values and calculatess a start and run
    ' based on ipStartIndex being adjusted to be a positive offset from FirstIndex
    ' and run increasing the positive offset from startindex
    
    Dim myStartRun As Result
    myStartRun = Result.Deb
    Set ParseStartRunEndToStartRun = myStartRun
    
    Dim myItemExtent As Result
    myItemExtent = GetItemExtent(ipItem, ipStartIndex, ipRun, ipEndIndex, ipRank)
    
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
            myStartRun.Item (StartRun.StartIndex = myEndIndex)

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
    Dim myItemFirstIndex As Long
    Dim myItemLastIndex As Long
    Dim myItemCount As Long
    'the span variable gives the most negative range based on counting down
    ' from firstindex.  This if firstindex = -5 and count is 10 then spanfirstindex will be -15
    ' because 10 below -5 is -6 to -15
    Dim myItemSpanFirstIndex As Long
    
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

