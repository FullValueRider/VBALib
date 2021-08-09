VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypesArrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Helpers")
'@PredeclaredId
'@Exposed
Option Explicit

'12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Public Enum ArrayDirections

    FourWay
    Eightway
    
End Enum

Public Enum KeyOrientation

    ByRow
    ByColumn
    
End Enum

'ToDo: Move perimiter and associated functions to its own Area class
'@Description("Returns an array of valid coordinated based on the provided x,y coordinate and the available area.")
Public Function GetPerimeterCoordinates _
( _
    ByVal ipCurRow As Long, _
    ByVal ipRowMin As Long, _
    ByVal ipRowMax As Long, _
    ByVal ipCurCol As Long, _
    ByVal ipColMin As Long, _
    ByVal ipColMax As Long, _
    ByVal ipType As ArrayDirections _
) As Variant


    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    If ipCurRow - 1 >= ipRowMin Then myLyst.Add Array(ipCurRow - 1, ipCurCol)
    If ipCurRow + 1 <= ipRowMax Then myLyst.Add Array(ipCurRow + 1, ipCurCol)
    
    If ipCurCol - 1 >= ipColMin Then myLyst.Add Array(ipCurRow, ipCurCol - 1)
    If ipCurCol + 1 <= ipColMax Then myLyst.Add Array(ipCurRow, ipCurCol + 1)
  
    If ipType = ArrayDirections.Eightway Then
    
        If ipCurRow - 1 >= ipRowMin And ipCurCol - 1 >= ipColMin Then myLyst.Add Array(ipCurRow - 1, ipCurCol - 1)
        If ipCurRow - 1 >= ipRowMin And ipCurCol + 1 <= ipColMax Then myLyst.Add Array(ipCurRow - 1, ipCurCol + 1)
        
        If ipCurRow + 1 <= ipRowMax And ipCurCol - 1 >= ipColMin Then myLyst.Add Array(ipCurRow + 1, ipCurCol - 1)
        If ipCurRow + 1 <= ipRowMax And ipCurCol + 1 <= ipColMax Then myLyst.Add Array(ipCurRow + 1, ipCurCol + 1)
        
    End If
    
    GetPerimeterCoordinates = myLyst.ToArrayList.ToArray
    
End Function


Public Function GetMoveCoordinates(ByVal ipType As ArrayDirections) As Scripting.Dictionary

    Dim myDirections As Scripting.Dictionary
    Set myDirections = New Scripting.Dictionary
    
    With myDirections
            
        .Add "N", Array(1, 0)
        .Add "E", Array(0, 1)
        .Add "S", Array(-1, 0)
        .Add "W", Array(0, -1)
            
        If ipType = ArrayDirections.Eightway Then
            
            .Add "NE", Array(1, 1)
            .Add "NW", Array(1, -1)
            .Add "SW", Array(-1, -1)
            .Add "SE", Array(-1, 1)
                
        End If
        
    End With
    
    Set GetMoveCoordinates = myDirections

End Function


Public Function Ranks(ByVal iparray As Variant, Optional ByRef iopResult As ResultLong) As ResultLong
    
    If iopResult Is Nothing Then Set iopResult = ResultLong.Deb.SetFailure
    Set Ranks = iopResult
    
    If Not VBA.IsArray(iparray) Then
    
        iopResult.SetFailure rcNotArray, "Parameter must be an Array"
        Exit Function
        
    End If
    
    'Visual Basic for Applications arrays can have up to 60000 dimensions; this allows for that
    Dim myDims As Long
    For myDims = 1 To 60000
    
        '@Ignore VariableNotUsed
        Dim myUbound As Long
        On Error Resume Next
        myUbound = UBound(iparray, myDims)
        If Err.Number <> 0 Then
            
            On Error GoTo 0
            iopResult.SetSuccess myDims - 1
            Exit Function
            
        End If
        
    Next
    
    On Error GoTo 0
    
End Function


Public Function HasRank(ByRef iparray As Variant, Optional ByVal ipRank As Long = 1) As Boolean

    ' It is legal for an upper bound to be -1
    ' so if the ubound test returns -1 we also need to
    ' check that Lbound is <= -1
    HasRank = False
    If Types.IsNotArray(iparray) Then Exit Function
    
    On Error Resume Next
    If LBound(iparray, ipRank) <= UBound(iparray, ipRank) Then
    
        If Err.Number = 0 Then
        
            On Error GoTo 0
            HasRank = True
            Exit Function
        
        End If
    
    End If
    
    On Error GoTo 0
    
End Function


Public Function LacksRank(ByRef iparray As Variant, Optional ByVal ipRank As Long = 1) As Boolean
    LacksRank = Not HasRank(iparray, ipRank)
End Function


Public Function IsListArray(ParamArray ipArgs() As Variant) As Boolean
    IsListArray = pvIsListArrayByCvar(CVar(ipArgs))
End Function

Public Function IsNotListArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotListArray = Not pvIsListArrayByCvar(CVar(ipArgs))
End Function

Public Function pvIsListArrayByCvar(ByVal ipArgs As Variant) As Boolean

    pvIsListArrayByCvar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
        
        If Types.IsNotArray(myArg) Then Exit Function
        If Types.Arrays.Ranks(myArg).value <> 1 Then Exit Function
            
    Next
    
    pvIsListArrayByCvar = True
    
End Function


Public Function IsTableArray(ParamArray ipArgs() As Variant) As Boolean
    IsTableArray = IsTableArrayByCVar(CVar(ipArgs))
End Function

Public Function IsNotTableArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotTableArray = Not IsTableArrayByCVar(CVar(ipArgs))
End Function

Public Function IsTableArrayByCVar(ByVal ipArgs As Variant) As Boolean

    IsTableArrayByCVar = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Not VBA.IsArray(myArg) Then Exit Function
        If Types.Arrays.Ranks(myArg).value <> 2 Then Exit Function
    
    Next
    
    IsTableArrayByCVar = True
    
End Function


Public Function IsMDArray(ParamArray ipArgs() As Variant) As Boolean

    IsMDArray = False
    
    Dim myArg As Variant
    For Each myArg In ipArgs
    
        If Not VBA.IsArray(myArg) Then Exit Function
        If Types.Arrays.Ranks(myArg).value < 3 Then Exit Function
    
    Next
    
    IsMDArray = True
    
End Function

'@Description("Returns -1 for an unitialised array otherwise the number of items in the array")
Public Function GetSize(ByRef iparray As Variant, Optional ByVal ipRank As Long = 1) As Long
    GetSize = TryGetSize(iparray, ipRank).value
End Function


Public Function IsEmpty(ByRef iparray As Variant, Optional ByVal ipRank As Long = 1) As Boolean
    IsEmpty = GetSize(iparray, ipRank) = 0
End Function


'@Description("Returns false if the array is uninitialised.  If an array is uninitialised the size value is 0 to be consistent with empty collection objects.")
Public Function TryGetSize(ByRef iparray As Variant, Optional ByVal ipRank As Long = 1, Optional ByRef iopResult As ResultLong) As ResultLong

    If iopResult Is Nothing Then Set iopResult = ResultLong.Deb.SetFailure
    Set TryGetSize = iopResult
    
    If Not VBA.IsArray(iparray) Then
    
        iopResult.SetFailure rcNotArray
        Exit Function
        
    End If
    
    If Types.Arrays.LacksRank(iparray, ipRank) Then
    
        iopResult.SetFailure rcInvalidRank, Fmt.Txt(ResultReason.AsString(rcInvalidRank), ipRank, Types.Arrays.Ranks(iparray).value)
        Exit Function
        
    End If

    Dim myUbound As Long
    On Error Resume Next
    myUbound = UBound(iparray, ipRank)
    If Err.Number <> 0 Or myUbound = -1 Then
    
        iopResult.SetFailure rcArrayNotInitialised
        On Error GoTo 0
        Exit Function
        
    End If
    On Error GoTo 0
    
    iopResult.SetSuccess myUbound - LBound(iparray, ipRank) + 1
    
End Function

Public Function GetLBound(ByRef iparray As Variant, Optional ByVal ipRank As Long = 1) As Long
    GetLBound = TryGetLbound(iparray, ipRank:=ipRank).value
End Function

Public Function TryGetLbound(ByRef iparray As Variant, ByVal ipRank As Long, Optional ByRef iopResult As ResultLong) As ResultLong

    If iopResult Is Nothing Then Set iopResult = ResultLong.Deb.SetFailure
    Set TryGetLbound = iopResult
    
    If Types.IsNotArray(iparray) Then
    
        iopResult.SetFailure rcNotArray, "Parameter is not an Array"
        Exit Function
        
    End If
    
    If Types.Arrays.LacksRank(iparray, ipRank) Then
    
       iopResult.SetFailure rcInvalidRank, "The Rank is not valid"
       Exit Function
       
    End If
    
    
    ' If Ubound is -1 or gives error the array is not initialised
    Dim myUbound As Long
    On Error Resume Next
    myUbound = UBound(iparray)
    If Err.Number = 0 Then
    
        If UBound(iparray, ipRank) > 0 Then
        
            iopResult.SetSuccess LBound(iparray, ipRank)
            
        End If
        
    Else
    
        iopResult.SetFailure rcArrayNotInitialised, "Dynamic array not initialised"
    
    End If
    
    On Error GoTo 0
        
End Function


Public Function GetUBound(ByRef iparray As Variant, Optional ByVal ipRank As Long = 1) As Long
    GetUBound = TryGetUbound(iparray, ipRank:=ipRank).value
End Function

Public Function TryGetUbound(ByRef iparray As Variant, ByVal ipRank As Long, Optional ByRef iopResult As ResultLong) As ResultLong

    If iopResult Is Nothing Then Set iopResult = ResultLong.Deb
    Set TryGetUbound = iopResult
    
    If Not Types.IsArray(iparray) Then
    
        iopResult.SetFailure rcNotArray, "Parameter is not an Array"
        Exit Function
        
    End If
    
    If Types.Arrays.LacksRank(iparray, ipRank) Then
    
       iopResult.SetFailure rcInvalidRank, "The Rank is not valid"
       Exit Function
       
    End If
    
    Dim myUbound As Long
    On Error Resume Next
    myUbound = UBound(iparray, ipRank)
    If Err.Number <> 0 Or myUbound = -1 Then
    
        iopResult.SetFailure rcArrayNotInitialised, "Dynamic array not initialised"
        Exit Function
    
    End If
    On Error GoTo 0
    
    iopResult.SetSuccess UBound(iparray, ipRank)
    
End Function


'@Description("Returns a Lyst of Lyst where the inner lyst is a lyst the TableArray row items")
Public Function TryToLystOfLyst(ByVal ipTableArray As Variant, ByRef iopResult As ResultLyst, Optional ByVal ipOrientation As KeyOrientation = KeyOrientation.ByRow) As ResultLyst

    ' nomeclature for Table arrays is Var(rows,columns) as this matches
    ' data taken from excel ranges.
    If iopResult Is Nothing Then Set iopResult = ResultLyst.Deb.SetFailure
    Set TryToLystOfLyst = iopResult

    If Types.Arrays.IsNotTableArray(ipTableArray) Then
    
        iopResult.SetFailure rcNotTableArray
        Exit Function
        
    End If
    
    Dim myFirstDim1 As Long
    myFirstDim1 = LBound(ipTableArray, 1)
    
    Dim myLastDim1 As Long
    myLastDim1 = UBound(ipTableArray, 1)
    
    Dim myFirstDim2 As Long
    myFirstDim2 = LBound(ipTableArray, 2)
    
    Dim myLastDIm2 As Long
    myLastDIm2 = UBound(ipTableArray, 2)
    
    
    
    Dim myLystOfLyst As Lyst
    Set myLystOfLyst = Lyst.Deb
    Dim myDim1 As Long
    Dim myDim2 As Long
    
    If ipOrientation = ByRow Then
        
        For myDim1 = myFirstDim1 To myLastDim1
        
            myLystOfLyst.Add Lyst.Deb
            
            For myDim2 = myFirstDim2 To myLastDIm2
            
                myLystOfLyst.Item(myDim1 - 1).Add ipTableArray(myDim1, myDim2)
                
            Next
        
        Next
        
    Else
    
        For myDim2 = myFirstDim2 To myLastDIm2
        
            myLystOfLyst.Add Lyst.Deb
            
            For myDim1 = myFirstDim1 To myLastDim1
            
                myLystOfLyst.Item(myDim2 - 1).Add ipTableArray(myDim1, myDim2)
                
            Next
        
        Next

    End If
    
    iopResult.SetSuccess myLystOfLyst
    
End Function


'@Description("Returns a Lyst of Lyst where item(0) is a list of the first items of the rank and item(1) is the lyst of the rank values
Public Function TryToLystOfLystSplitFirst _
( _
    ByVal ip2DArray As Variant, _
    ByRef iopResult As ResultLyst, _
    Optional ByVal ipOrientation As KeyOrientation = KeyOrientation.ByRow _
) As ResultLyst

    If iopResult Is Nothing Then Set iopResult = ResultLyst.Deb.SetFailure
    Set TryToLystOfLystSplitFirst = iopResult
    
    If TryToLystOfLyst(ip2DArray, iopResult, ipOrientation).Failed Then Exit Function
    
    Dim myFirstValues  As Lyst
    Set myFirstValues = Lyst.Deb
    
    Dim myRankValues As Lyst
    Set myRankValues = Lyst.Deb
    
    Dim myLystOfLyst As Lyst
    Set myLystOfLyst = iopResult.value
    
    Dim myInnerLyst As Lyst
    For Each myInnerLyst In myLystOfLyst
    
        myFirstValues.Add myInnerLyst.Item(0)
        myRankValues.Add myInnerLyst.RemoveAt(0).Clone
            
     Next
    
     Dim myResult As Lyst
     Set myResult = Lyst.Deb.Add(myFirstValues).Add(myRankValues)
    iopResult.SetSuccess myResult
    
End Function

'@Description("Returns a Lyst of Lyst wheere the inner lyst is two items, the first value of the row, thesecond is the remainder of the row items as a Lyst")
Public Function TryToLystOfLystCopyFirst _
( _
    ByVal ip2DArray As Variant, _
    ByRef iopResult As ResultLyst, _
    Optional ByVal ipOrientation As KeyOrientation = KeyOrientation.ByRow _
) As ResultLyst

    If iopResult Is Nothing Then Set iopResult = ResultLyst.Deb.SetFailure
        Set TryToLystOfLystCopyFirst = iopResult
        
        If TryToLystOfLyst(ip2DArray, iopResult, ipOrientation).Failed Then Exit Function
        
        Dim myFirstValues  As Lyst
        Set myFirstValues = Lyst.Deb
        
        Dim myRankValues As Lyst
        Set myRankValues = Lyst.Deb
        
        Dim myLystOfLyst As Lyst
        Set myLystOfLyst = iopResult.value
        
        Dim myInnerLyst As Lyst
        For Each myInnerLyst In myLystOfLyst
        
            myFirstValues.Add myInnerLyst.Item(0)
            myRankValues.Add myInnerLyst.Clone
                
         Next
        
         Dim myResult As Lyst
         Set myResult = Lyst.Deb.Add(myFirstValues).Add(myRankValues)
        iopResult.SetSuccess myResult
    
End Function

'@Description("transposes rows and colums of a 2d array")
Public Function TryTranspose(ByVal ipTableArray As Variant, Optional ByRef iopResult As ResultVariant) As ResultVariant


    If iopResult Is Nothing Then Set iopResult = ResultVariant.Deb.SetFailure
    Set TryTranspose = iopResult
    On Error GoTo exitfailed
    
    If Types.Arrays.IsNotTableArray(ipTableArray) Then
    
        iopResult.SetFailure rcNotTableArray
        Exit Function
        
    End If
    
    Dim myFirstDim1 As Long
    myFirstDim1 = LBound(ipTableArray, 1)
    
    Dim myLastDim1 As Long
    myLastDim1 = UBound(ipTableArray, 1)

    Dim myFirstDim2 As Long
    myFirstDim2 = LBound(ipTableArray, 2)
    
    Dim myLastDIm2 As Long
    myLastDIm2 = UBound(ipTableArray, 2)

    '@Ignore VariableNotAssigned
    Dim myTableArray As Variant
    ReDim myTableArray(myFirstDim2 To myLastDIm2, myFirstDim1 To myLastDim1)
    
    Dim myDim1 As Long
    For myDim1 = myFirstDim1 To myLastDim1
    
        Dim myDim2 As Long
        For myDim2 = myFirstDim2 To myLastDIm2
        
            myTableArray(myDim2, myDim1) = ipTableArray(myDim1, myDim2)
            
        Next
    
    Next
    
    '@Ignore UnassignedVariableUsage
    iopResult.SetSuccess myTableArray
    
exitfailed:
End Function

