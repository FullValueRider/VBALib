VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arrays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'@Folder("VBALib")
'@PredeclaredId
'@Exposed

'12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Public Enum ArraysConst

'     TryLBoundResult = 0
'     TryUboundResult = 0
    
' End Enum

Private Type InternalConstants
    
    Location                                  As String
    ArrayNotCountable                         As Long

End Type

Private i                                     As InternalConstants


Public Enum ArrayDirections

    FourWay
    Eightway
    
End Enum

Public Enum RankOrientation

    ByRow
    ByColumn
    
End Enum

' Private Type State
    
'     Results As Stack
    
' End Type


Public Sub Class_Initialize()
    
    If VBA.Len(i.Location) = 0 Then
        
        i.Location = "VBALib.Arrays."
        i.ArrayNotCountable = -1
        
        
    End If
    
    Guard InvalidNew, Not (Me Is Arrays), i.Location & "Class_Initialise"

End Sub


' Public Function GetConst(ByVal ipConst As ArraysConst) As ArraysConst
'     GetConst = ipConst
' End Function



'ToDo: Move perimiter and associated functions to its own Area class
' ToDO use Lbound and Ubound to get min and max values
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
Attribute GetPerimeterCoordinates.VB_Description = "Returns an array of valid coordinated based on the provided x,y coordinate and the available area."

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

Public Function IsEmptyParamArray(ByVal ipParamarray As Variant) As Boolean
    IsEmptyParamArray = UBound(ipParamarray) = -1
End Function

'In VBAlib it is common practise to parse a paramarray based on the number of items in the array
' this method returns a Lyst where the input paramarray has been processed into an appropriate set of
' lysts to allow for each actions on the one or more of the lysts.  This becomes very repetitive boiler plate code in many methods.
'
' The general pattern for parsing a paramarray is to determine if a paramarray consits of a sinlge item
' and if it is to take different actions depending on if it is an iterable object or not.
' if the param array contains more than one item then the list is just treated as an iterable array.
'
' To be successful this method must not allow for infinite recursions,
' i.e it must alway return a value without either directly or indirectly calling itself.
'
' Public Function ParseParamArrayToLystIsFail(ByVal ipParamArray As Variant, ByRef opResult As Lyst) As Boolean
'     ParseParamArrayToLystIsFail = Not ParseParamArrayToLystIsOkay(ipParamArray, opResult)
' End Function




Public Function IsSameSize _
( _
    ByVal ipTestArray As Variant, _
    ByVal ipRefArray As Variant, _
    Optional ByVal ipTestRank As Long = ArrayFirstRank, _
    Optional ByVal ipRefRank As Long = ArrayFirstRank _
) As Boolean
    
    IsSameSize = False

    If (ipTestRank = 0) And (ipRefRank = 0) Then
        
        IsSameSize = pvIsArraySameSize(ipTestArray, ipRefArray)
    Else
        
        IsSameSize = pvIsRankSameSize(ipTestArray, ipRefArray, ipTestRank, ipRefRank)
        
    End If

End Function


Public Function pvIsRankSameSize _
( _
    ByVal ipTestArray As Variant, _
    ByVal ipRefArray As Variant, _
    Optional ByVal ipTestRank As Long = ArrayFirstRank, _
    Optional ByVal ipRefRank As Long = ArrayFirstRank _
) As Boolean
    
    pvIsRankSameSize = False
    
    If LacksRank(ipTestArray, ipTestRank) Then Exit Function
    If LacksRank(ipRefArray, ipRefRank) Then Exit Function
    
    pvIsRankSameSize = (UBound(ipTestArray) - LBound(ipTestArray)) = (UBound(ipRefArray) = LBound(ipRefArray))
    
End Function

Public Function pvIsArraySameSize _
( _
    ByVal ipTestArray As Variant, _
    ByVal ipRefArray As Variant _
) As Boolean

    pvIsArraySameSize = False
    
    Dim myTestRanks As Long
    myTestRanks = Ranks(ipTestArray)
    If myTestRanks = 0 Then Exit Function
            
    Dim myRefRanks As Long
    myRefRanks = Ranks(ipRefArray)
    If myRefRanks = 0 Then Exit Function
        
    If myRefRanks <> myTestRanks Then Exit Function
        
    Dim mySize As Long
    mySize = MinLong
    
    Dim myRank As Long
    For myRank = ArrayFirstRank To myRefRanks
    
        Dim myRefSize As Long
        myRefSize = Count(ipRefArray, myRank)
        
        If myRefSize = Sys.Failed Then Exit Function
        If myRefRanks = MinLong Then myRefRanks = myRefSize
        If mySize <> myRefSize Then Exit Function
            
        Dim myTestSize As Long
        myTestSize = Count(ipTestArray(myRank))
        If mySize <> myTestSize Then Exit Function
            
    Next
    
    pvIsArraySameSize = True
    
End Function

' '@Description("Returns true if the Array has been initialised. The term 'HasItems' is used t align with 'HasItems' for collection objects")
' Public Function HasItems(ByVal ipArray As Variant, Optional ipRank As Long = ArrayFirstRank) As Boolean
    
'     HasItems = False
'     If LacksItems(ipArray) Then Exit Function
'     HasItems = True
    
' End Function

' Public Function LacksItems(ByVal ipArray As Variant, Optional ByVal ipRank As Long = 1) As Boolean
'     LacksItems = Not HasItems(ipArray, ipRank)
' End Function

'@Description("Returns the number of dimensions of an array.  A return value of -1 means the input could not be treated as an array")
Public Function Ranks(ByVal iparray As Variant) As Long
Attribute Ranks.VB_Description = "Returns the number of dimensions of an array.  A return value of -1 means the input could not be treated as an array"
    
    Ranks = 0
    'If LacksItems(iparray) Then Exit Function
    
    'Visual Basic for Applications arrays can have up to 60000 dimensions; this allows for that
    Dim myDims As Long
    For myDims = 1 To 60000
    
        '@Ignore VariableNotUsed
        Dim myUbound As Long
        On Error Resume Next
        myUbound = UBound(iparray, myDims)
        If Err.Number <> 0 Then
            
            On Error GoTo 0
            Ranks = myDims - 1
            Exit Function
            
        End If
        
    Next
    
    On Error GoTo 0
    
End Function

'@Description("Returns True is an array has the requested dimension")
Public Function HasRank(ByVal iparray As Variant, Optional ByVal ipRank As Long = 1) As Boolean
Attribute HasRank.VB_Description = "Returns True is an array has the requested dimension"
    
    HasRank = False
    'If LacksItems(iparray) Then Exit Function
    If ipRank < 1 Then Exit Function
    If ipRank > Ranks(iparray) Then Exit Function
    HasRank = True
    
End Function

Public Function LacksRank(ByVal iparray As Variant, Optional ByVal ipRank As Long = 1) As Boolean
    LacksRank = Not HasRank(iparray, ipRank)
End Function



Public Function IsSingleItem(ByVal iparray As Variant) As Boolean
    IsSingleItem = (UBound(iparray) - LBound(iparray) + 1) = 1
End Function

Public Function IsNotSingleItem(ByVal iparray As Variant) As Boolean
    IsNotSingleItem = Not IsSingleItem(iparray)
End Function

Public Function IsListArray(ParamArray ipArgs() As Variant) As Boolean
    IsListArray = HasRanks(CVar(ipArgs), Comparer.Action.Equal, 1)
End Function

Public Function IsNotListArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotListArray = Not HasRanks(CVar(ipArgs), Comparer.Action.Equal, 1)
End Function


Public Function IsTableArray(ParamArray ipArgs() As Variant) As Boolean
    IsTableArray = HasRanks(CVar(ipArgs), Comparer.Action.Equal, 2)
End Function

Public Function IsNotTableArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotTableArray = Not HasRanks(CVar(ipArgs), Action.Equal, 2)
End Function


Public Function IsMDArray(ParamArray ipArgs() As Variant) As Boolean

    IsMDArray = HasRanks(ipArgs, Comparer.Action.MoreThan, 2)
    
End Function
    
Public Function IsNotMDArray(ParamArray ipArgs() As Variant) As Boolean
    IsNotMDArray = Not IsMDArray(ipArgs, Comparer.Action.MoreThan, 2)
End Function

Friend Function HasRanks(ByVal ipArgs As Variant, ByVal ipCompare As Action, ByVal ipRanks As Long) As Boolean

    HasRanks = False
    If HasNoItems(ipArgs) Then Exit Function
    
    Dim myArg As Variant
    For Each myArg In ipArgs
        
        If Not Comparer.Compare(ipCompare, Ranks(myArg), ipRanks) Then Exit Function
            
    Next
    
    HasRanks = True
    
End Function

'@Description("Returns -1 for an unitialised array/not array otherwise the number of items in the rank of the array")
Public Function Count(ByRef iparray As Variant, Optional ByVal ipRank As Long = ArrayFirstRank) As Long
Attribute Count.VB_Description = "Returns -1 for an unitialised array/not array otherwise the number of items in the rank of the array"
    
    If HasNoItems(iparray, ipRank) Then
        
        Count = i.ArrayNotCountable
    
    Else
    
        Count = UBound(iparray, ipRank) - LBound(iparray, ipRank) + 1
        
    End If
        
End Function

'@Description("Returns -1 for an unitialised array/not array otherwise the total number of items in the array")
Public Function CountAll(ByRef iparray As Variant) As Long
Attribute CountAll.VB_Description = "Returns -1 for an unitialised array/not array otherwise the total number of items in the array"

    CountAll = i.ArrayNotCountable
    
    If HasNoItems(iparray) Then Exit Function
    
    Dim myRank As Long
    Dim myRanks As Long
    myRanks = Ranks(iparray)
    Dim myCount As Long
    myCount = 1
    For myRank = 1 To myRanks
        
        myCount = myCount * Count(iparray, myRank)
        
        
    Next
        
    CountAll = myCount
        
End Function

'@Description("Returns true if a the array held by a variant has a Ubound of -1.  By definition, ParamArrays contain at least an empty array, and consequently cannot be tested with IsEmpty")
Public Function HasItems(ByVal iparray As Variant, Optional ByVal ipRank As Long = ArrayFirstRank) As Boolean
Attribute HasItems.VB_Description = "Returns true if a the array held by a variant has a Ubound of -1.  By definition, ParamArrays contain at least an empty array, and consequently cannot be tested with IsEmpty"
    HasItems = TryExtent(iparray, ipRank:=ipRank).Status
End Function

Public Function HasNoItems(ByVal iparray As Variant, Optional ByVal ipRank As Long = ArrayFirstRank) As Boolean
    HasNoItems = Not HasItems(iparray, ipRank)
End Function

Public Function TryExtent(ByRef iparray As Variant, Optional ByRef opResult As Result, Optional ByVal ipRank As Long = ArrayFirstRank) As Result
    
    If opResult Is Nothing Then Set opResult = Result.Deb
    Set TryExtent = opResult
    
    Dim myErrorNo As Long
    On Error Resume Next
    opResult.Item(Extent.IsUbound) = UBound(iparray, ipRank)
    myErrorNo = Err.Number
    
    Err.Clear
    On Error GoTo 0
    
    If myErrorNo <> 0 Then Exit Function
    'AN uninitialised variant array has Ubound of -1 and Lbound of 0
    opResult.Item(Extent.IsLbound) = LBound(iparray, ipRank)
    If opResult.Item(Extent.IsUbound) < opResult.Item(Extent.IsLbound) Then Exit Function
        
    opResult.Status = True
    opResult.Item(Extent.IsCount) = opResult.Item(Extent.IsUbound) - opResult.Item(Extent.IsLbound) + 1
    
End Function
''@Description("Returns False for an unitialised variant array. If True, opbound cantains the value of Lbound for the rank.")
'Public Function TryLBound(ByRef iparray As Variant, Optional ByRef opResult As Result, Optional ByVal ipRank As Long = ArrayFirstRank) As Result
'
'    Set opResult = Result.Deb
'    Set TryLBound = opResult
'    Dim myErrorNo As Long
'    On Error Resume Next
'    opResult.Item(0) = LBound(iparray, ipRank) ' item 0 is lbound
'    myErrorNo = Err.Number
'
'    Err.Clear
'    On Error GoTo 0
'
'    If myErrorNo <> 0 Then Exit Function
'
'    ' for a variant array we need to check that Ubound>Lbound
'    ' to avoid the 0,-1 for unintnitialsed variant arrays
'    If UBound(iparray, ipRank) < opResult.Item(0) Then Exit Function
'    opResult.Status = True
'
'End Function
'
'
'Public Function TryUBound(ByRef iparray As Variant, Optional ByRef opUBound As Result, Optional ByVal ipRank As Long = ArrayFirstRank) As Result
'
'    Set opUBound = Result.Deb
'    Set TryUBound = opUBound
'
'    Dim myErrorNo As Long
'    On Error Resume Next
'    opUBound.Item(0) = UBound(iparray, ipRank)
'    myErrorNo = Err.Number
'
'    Err.Clear
'    On Error GoTo 0
'
'    If myErrorNo <> 0 Then Exit Function
'    'AN uninitialised variant array has Ubound of -1 and Lbound of 0
'    If opUBound.Item(0) < LBound(iparray, ipRank) Then Exit Function
'
'    opUBound.Status = True
'
'End Function



'@Description("Returns a Lyst of Lyst where the inner lyst is a lyst the selected ranks items")
Public Function ToLystOfRankLysts _
( _
    ByVal ipTablearray As Variant, _
    Optional ByVal ipAction As TableToLystAction = TableToLystAction.RankIsRowFirstItemActionIsNoAction _
) As Lyst
Attribute ToLystOfRankLysts.VB_Description = "Returns a Lyst of Lyst where the inner lyst is a lyst the selected ranks items"

    ' nomeclature for Table arrays is Var(rows,columns) as this matches
    ' the organisation of tables taken from Excel.
    Guard NotTableArray, Me.IsNotTableArray(ipTablearray), i.Location & "ToLystOfRankLysts"
    'Debug.Print ipAction, Enums.TableToLystActions.LacksMember(ipAction)
    Guard NotEnumerationMember, Enums.TableToLystActions.LacksMember(ipAction), i.Location & "TryToLystOfLysts"
    
    Set ToLystOfRankLysts = Enums.TableToLystActions.Exec(ipAction, ipTablearray)
    
End Function


'@Description("transposes rows and colums of a 2d array")
Public Function Transpose(ByVal ipTablearray As Variant) As Variant
Attribute Transpose.VB_Description = "transposes rows and colums of a 2d array"

    Guard NotTableArray, Me.IsNotTableArray(ipTablearray), i.Location & "Transpose"
    
    Dim myFirstRow As Long
    myFirstRow = LBound(ipTablearray, 1)
    
    Dim myLastRow As Long
    myLastRow = UBound(ipTablearray, 1)

    Dim myFirstCol As Long
    myFirstCol = LBound(ipTablearray, 2)
    
    Dim myLastCol As Long
    myLastCol = UBound(ipTablearray, 2)

    '@Ignore VariableNotAssigned
    Dim myTableArray As Variant
    ReDim myTableArray(myFirstCol To myLastCol, myFirstRow To myLastRow)
    
    Dim myRow As Long
    For myRow = myFirstRow To myLastRow
    
        Dim myCol As Long
        For myCol = myFirstCol To myLastCol
        
            myTableArray(myCol, myRow) = ipTablearray(myRow, myCol)
            
        Next
    
    Next
    
    '@Ignore UnassignedVariableUsage
    Transpose = myTableArray
    
End Function
    
