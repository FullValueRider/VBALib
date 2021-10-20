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
Option Explicit

'12345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Const mcLocation                    As String = "VBALib.Arrays."
Const mcArrayNotCountable                As String = -1

Public Enum ArrayDirections

    FourWay
    Eightway
    
End Enum

Public Enum RankOrientation

    ByRow
    ByColumn
    
End Enum

Private Type State
    
    Results As stack
    
End Type


Public Sub Class_Initialize()
    
    Guard InvalidNew, Not (Me Is Arrays), mcLocation & "Class_Initialise"
    ' Set s.Res = Result.Deb

End Sub

' Public Property Get Result() As Result
'     Set Result = s.Res
' End Property


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

Public Function IsEmptyParamArray(ByVal ipParamArray As Variant) As Boolean
    IsEmptyParamArray = UBound(ipParamArray) = -1
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
    ipTestArray As Variant, _
    ipRefArray As Variant, _
    Optional ipTestRank As Long = ArrayFirstRank, _
    Optional ipRefRank As Long = ArrayFirstRank _
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
    ipTestArray As Variant, _
    ipRefArray As Variant, _
    Optional ipTestRank As Long = ArrayFirstRank, _
    Optional ipRefRank As Long = ArrayFirstRank _
) As Boolean
    
    pvIsRankSameSize = False
    
    If LacksRank(ipTestArray, ipTestRank) Then Exit Function
    If LacksRank(ipRefArray, ipRefRank) Then Exit Function
    
    pvIsRankSameSize = (UBound(ipTestArray) - LBound(ipTestArray)) = (UBound(ipRefArray) = LBound(ipRefArray))
    
End Function

Public Function pvIsArraySameSize _
( _
    ipTestArray As Variant, _
    ipRefArray As Variant _
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
        myTestSize = Count(ipTestArray(, myRank))
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
Public Function HasRank(ByRef iparray As Variant, Optional ByVal ipRank As Long = 1) As Boolean
    
    HasRank = False
    'If LacksItems(iparray) Then Exit Function
    If ipRank < 1 Then Exit Function
    If ipRank > Ranks(iparray) Then Exit Function
    HasRank = True
    
End Function

Public Function LacksRank(ByRef iparray As Variant, Optional ByVal ipRank As Long = 1) As Boolean
    LacksRank = Not HasRank(iparray, ipRank)
End Function



Public Function IsSingleItem(ByVal ipArray As Variant) As Boolean
    IsSingleItem = (UBound(ipArray) - LBound(ipArray) + 1) = 1
End Function

Public Function IsNotSingleItem(ByVal ipArray As Variant) As Boolean
    IsNotSingleItem = Not IsSingleItem(ipArray)
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
    IsNotTableArray = Not HasRanks(CVar(ipArgs), Comparer.Action.Equal, 2)
End Function


Public Function IsMDArray(ParamArray ipArgs() As Variant) As Boolean

    IsMDArray = HasRanks(ipArgs, Comparer.Action.MoreThan, 2)
    
End Function
    
Public Function IsNotMDArray(ParamArray ipargs() As Variant) As Boolean
    IsNotMDArray = Not IsMDArray(ipargs, Comparer.Action.MoreThan, 2)
End Function

Friend Function HasRanks(ByVal ipArgs As Variant, ByVal ipCompare As Comparer.Action, ByVal ipRanks As Long) As Boolean

    HasRanks = False
    If HasNoItems(ipArgs) Then Exit Function
    
    Dim myArg As Variant
    For Each myArg In ipArgs
        
        If Not Comparer.Compare(ipCompare, Ranks(myArg), ipRanks) Then Exit Function
            
    Next
    
    HasRanks = True
    
End Function

'@Description("Returns -1 for an unitialised array/not array otherwise the number of items in the rank of the array")
Public Function Count(ByRef ipArray As Variant, Optional ByVal ipRank As Long = ArrayFirstRank) As Long
    
    If HasNoItems(ipArray, ipRank) Then
        
        Count = mcArrayNotCountable
    
    Else
    
        Count = UBound(ipArray, ipRank) - LBound(ipArray, ipRank) + 1
        
    End If
        
End Function

'@Description("Returns -1 for an unitialised array/not array otherwise the total number of items in the array")
Public Function CountAll(ByRef ipArray As Variant) As Long

    CountAll = mcArrayNotCountable
    
    If HasNoItems(ipArray) Then Exit Function
    
    Dim myRank As Long
    Dim myRanks As Long
    myRanks = Ranks(ipArray)
    Dim myCount As Long
    myCount = 1
    For myRank = 1 To myRanks
        
        myCount = myCount * Count(ipArray, myRank)
        
        
    Next
        
    CountAll = myCount
        
End Function

'@Description("Returns true if a the array held by a variant has a Ubound of -1.  By definition, ParamArrays contain at least an empty array, and consequently cannot be tested with IsEmpty")
Public Function HasItems(ByVal ipArray As Variant, Optional ipRank As Long = ArrayFirstRank) As Boolean
    HasItems = TryLBound(ipArray, iprank:=ipRank)
End Function

Public Function HasNoItems(ByVal ipArray As Variant, Optional ipRank As Long = ArrayFirstRank) As Boolean
    HasNoItems = Not HasItems(ipArray, ipRank)
End Function

'@Description("Returns False for an unitialised variant array. If True, opbound cantains the value of Lbound for the rank.")
Public Function TryLBound(ByRef iparray As Variant, Optional ByRef opResult As Result, Optional ByVal ipRank As Long = ArrayFirstRank) As Result
    
    
    TryLBound = False
    Dim myErrorNo As Long
    On Error Resume Next
    opLbound = LBound(iparray, ipRank)
    myErrorNo = Err.Number
    
    Err.Clear
    On Error GoTo 0
    
    If myErrorNo <> 0 Then Exit Function
        
    ' for a variant array we need to check that Ubound>Lbound 
    ' to avoid the 0,-1 for unintnitialsed variant arrays
    If UBound(iparray, ipRank) < opLbound Then Exit Function
    TryLBound = True
        
End Function


Public Function TryUBound(ByRef iparray As Variant, Optional opUbound As Long, Optional ByVal ipRank As Long = ArrayFirstRank) As Boolean
    
    TryUBound = False
    Dim myErrorNo As Long
    On Error Resume Next
    opUbound = UBound(iparray, ipRank)
    myErrorNo = Err.Number
    
    Err.Clear
    On Error GoTo 0
    
    If myErrorNo <> 0 Then Exit Function
    'AN uninitialised variant array has Ubound of -1 and Lbound of 0
    If opUbound < LBound(iparray, ipRank) Then Exit Function
        
    TryUBound = True
    
End Function



'@Description("Returns a Lyst of Lyst where the inner lyst is a lyst the selected ranks items")
Public Function ToLystOfRankLysts _
( _
    ByVal ipTableArray As Variant, _
    Optional ByVal ipAction As TableToLystAction = TableToLystAction.RankIsRowFirstItemActionIsNoAction _
) As Lyst

    ' nomeclature for Table arrays is Var(rows,columns) as this matches
    ' the organisation of tables taken from Excel.
    Guard NotTableArray, Arrays.IsNotTableArray(ipTableArray), mcLocation & "ToLystOfRankLysts"
    'Debug.Print ipAction, Enums.TableToLystActions.LacksMember(ipAction)
    Guard NotEnumerationMember, Enums.TableToLystActions.LacksMember(ipAction), mcLocation & "TryToLystOfLysts"
    
    Set ToLystOfRankLysts = Enums.TableToLystActions.Exec(ipAction, ipTableArray)
    
End Function


'@Description("transposes rows and colums of a 2d array")
Public Function Transpose(ByVal ipTableArray As Variant) As Variant

    Guard NotTableArray, Arrays.IsNotTableArray(ipTableArray), mcLocation & "Transpose"
    
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
        
            myTableArray (myDim2, myDim1) = ipTableArray(myDim1, myDim2)
            
        Next
    
    Next
    
    Transpose = myTableArray
    
End Function
    
