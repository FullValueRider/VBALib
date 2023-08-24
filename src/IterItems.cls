VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IterItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A class to iterate any single dimension containers, strings and single items"
'@PredeclaredId
'@Exposed
'@Folder("Iterators")
'@ModuleDescription("A class to iterate any single dimension containers, strings and single items")
Option Explicit
' This class allows iterating over any one dimensional contains, string or single value
' The iteration can be forwards or backward for any given item
' Items may be reference with an offset  e.g. Item(0) is the current item, .Item(-2) is two items before hte current item
' The class also maintains two other moveable values
' Key = the native index of the  host item
' Index = the offset from the first index of the host item
' the indexes provided to FTS must be based on normalised indexes StartIndexing at 1

Private Const MY_TYPENAME               As String = "IterItems"

Private Type State

    Host                                As Variant ' holds ipItems
    Keys                                As Variant ' Keys if ipItems is a dictionary
    Items                               As Variant ' Items if ipItems is a ToArray object
    
    FirstIndex                          As Long
    LastIndex                           As Long
    GlobalOffset                        As Long
    Size                                As Long
    StartIndex                          As Long
    EndIndex                            As Long
    Step                                As Long
    WrapAround                          As Boolean
    
    
End Type


Private s                               As State


'@Defaultmember
Public Function Deb(ByRef ipItems As Variant) As IterItems
Attribute Deb.VB_UserMemId = 0
    With New IterItems
        Set Deb = .ConstructInstance(ipItems)
    End With
End Function


Friend Function ConstructInstance(ByRef ipItems As Variant) As IterItems

    ' If VBA.IsArray(ipItems) Then
    '     If ArrayOp.Ranks(ipItems) > 1 Then
    '         Err.Raise 17 + vbObjectError, _
    '             "Iteritems: Invalid Array", _
    '             "Arrays must be 1 dimension only"
    '     End If
    ' End If
    
    Select Case True
        
        Case GroupInfo.IsDictionary(ipItems)
           
            s.Keys = ipItems.Keys
    
            Set s.Host = ipItems
            s.FirstIndex = LBound(s.Keys)
            s.LastIndex = UBound(s.Keys)
            
        Case GroupInfo.IsItemByToArray(ipItems)
        
            s.Items = ipItems.ToArray
            Set s.Host = ipItems
            s.FirstIndex = LBound(s.Items)
            s.LastIndex = UBound(s.Items)
            
        Case GroupInfo.IsString(ipItems)
        
            s.Host = ipItems
            s.FirstIndex = 1
            s.LastIndex = VBA.Len(ipItems)
            
        Case GroupInfo.IsIndexFrom1(ipItems)
        
            Set s.Host = ipItems
            ' cannot use firstindex lastindex as not supported by collection
            s.FirstIndex = 1
            s.LastIndex = ipItems.Count
            
        Case GroupInfo.IsIndexFrom0(ipItems)
        
            Set s.Host = ipItems
            s.FirstIndex = 0
            s.LastIndex = ipItems.Count - 1
            
        Case GroupInfo.IsArray(ipItems)
            If ArrayOp.Ranks(ipItems) > 1 Then
                Err.Raise 17 + vbObjectError, _
                "IterItems", _
                "IterItems does not currently support Arrays with more than one dimension"
            End If
            s.Host = ipItems
            s.FirstIndex = LBound(ipItems)
            s.LastIndex = UBound(ipItems)
            
        
        Case Else
        
            If VBA.IsObject(ipItems) Then
                Set s.Host = ipItems
            Else
                s.Host = ipItems
            End If
            s.FirstIndex = 1
            s.LastIndex = 1
            
    End Select
    
    s.StartIndex = s.FirstIndex
    s.EndIndex = s.LastIndex
    s.Step = 1
    s.GlobalOffset = 0
    s.Size = s.LastIndex - s.FirstIndex + 1
    
'    If s.Size = 0 Then
'        AlertNoItems
'    End If
    
    Set ConstructInstance = Me
    
End Function


Public Function Size() As Long
    Size = s.Size
End Function


Public Function HoldsItems() As Boolean
    HoldsItems = s.Size > 0
End Function


Public Function LacksItems() As Boolean
    LacksItems = s.Size < 1
End Function


'@Description("Sets the subrange and stepping to be used when using Iteritems. From,To parameters must use 1 based indexing. Step in the number of position to move in response to MoveNext/MovePrev. Mapping to the ative index is done by IterItems")
Public Function FTS(Optional ByRef ipStartIndex As Variant = Empty, Optional ByRef ipEndIndex As Variant = Empty, Optional ByRef ipStep As Long = 1) As IterItems
Attribute FTS.VB_Description = "Sets the subrange and stepping to be used when using Iteritems. From,To parameters must use 1 based indexing. Step in the number of position to move in response to MoveNext/MovePrev. Mapping to the ative index is done by IterItems"
    ' Start and End indexes are variants because at some stage we will also be allowing dictionary keys
    
    If Not VBA.IsEmpty(ipStartIndex) Then
        s.StartIndex = ipStartIndex
    End If
    
    If Not VBA.IsEmpty(ipEndIndex) Then
        s.EndIndex = ipEndIndex
    End If
    
    ' s.StartIndex = EnsureNumber(ipStartIndex, s.FirstIndex)
    ' s.EndIndex = EnsureNumber(ipEndIndex, s.EndIndex)
    If s.Step < 1 Then
        AlertNegativeStep ipStep
    Else
        s.Step = ipStep
    End If
    
    'check for valid numbers in the cntext of 1 based indexing
    
    If s.StartIndex < 1 Or s.StartIndex > s.Size Then
        AlertInvalidFTs "StartIndex", s.StartIndex
    End If
    
    If s.EndIndex < 1 Or s.EndIndex > s.Size Then
        AlertInvalidFTs "EndIndex", s.EndIndex
    End If
    
    If s.StartIndex > s.EndIndex Then
        AlertInvalidStartEnd s.StartIndex, s.EndIndex
    End If
    
    ' finally,  convert start and end indexes to native indexes
    s.StartIndex = s.FirstIndex + s.StartIndex - 1
    s.EndIndex = s.FirstIndex + s.EndIndex - 1
    ' global offset is offset from firstIndex
    s.GlobalOffset = s.StartIndex - s.FirstIndex
    
    Set FTS = Me
    
End Function


' Private Function EnsureNumber(ByRef ipNumber As Variant, ByRef ipDefault As Long) As Long

'     If VBA.IsEmpty(ipNumber) Then
'         EnsureNumber = ipDefault
'     Else
'         If GroupInfo.IsNumber(ipNumber) Then
'             EnsureNumber = ipNumber
'         Else
'             AlertNotNumber ipNumber
'         End If
'     End If
    
' End Function


'@Description("MoveToEndIndex moves to the nearest step position before or equal to EndIndex)
Public Function MoveToEndIndex() As IterItems
    s.GlobalOffset = ((s.EndIndex - s.FirstIndex) \ s.Step) * s.Step
    Set MoveToEndIndex = Me
End Function


'@Description("Moves to the defined Strt position
Public Function MoveToStartIndex() As IterItems
    s.GlobalOffset = s.StartIndex - s.FirstIndex '+ 1
    Set MoveToStartIndex = Me
End Function


'@Description("Moves to firstindex, resets StartIndex stop EndIndex")
Public Function MoveToFirst() As IterItems
Attribute MoveToFirst.VB_Description = "Moves to firstindex, resets StartIndex stop EndIndex"
    s.GlobalOffset = 0
    s.StartIndex = s.FirstIndex
    s.EndIndex = s.LastIndex
    s.Step = 1
    Set MoveToFirst = Me
End Function


'@Description("Moves to firstindex, resets StartIndex stop EndIndex")
Public Function MoveToLast() As IterItems
Attribute MoveToLast.VB_Description = "Moves to firstindex, resets StartIndex stop EndIndex"
    s.GlobalOffset = s.Size - 1
    s.StartIndex = s.FirstIndex
    s.EndIndex = s.LastIndex
    s.Step = 1
    Set MoveToLast = Me
End Function


Public Function HasNext(Optional ByRef ipLocalOffset As Long = 0) As Boolean

    HasNext = False
    Dim myLoc As Long: myLoc = s.GlobalOffset + ipLocalOffset * s.Step + s.FirstIndex '- 1
    If myLoc > s.EndIndex - s.Step Then
        Exit Function
    End If
    HasNext = True
    
End Function


Public Function HasNoNext(Optional ByRef ipLocalOffset As Long = 0) As Boolean
    HasNoNext = Not HasNext(ipLocalOffset)
End Function


Public Function HasPrev(Optional ByRef ipLocalOffset As Long = 0) As Boolean

    HasPrev = False
    Dim myLoc As Long: myLoc = s.GlobalOffset + ipLocalOffset * s.Step + s.FirstIndex '- 1
   
    If myLoc < s.StartIndex + s.Step Then
        Exit Function
    End If

    HasPrev = True
    
End Function


Public Function HasNoPrev(Optional ByRef ipLocalOffset As Long = 0) As Boolean
    HasNoPrev = Not HasPrev(ipLocalOffset)
End Function


Public Function MoveNext() As Boolean
    If HasNext Then
        MoveNext = True
        s.GlobalOffset = s.GlobalOffset + s.Step
    Else
        MoveNext = False
    End If
End Function


Public Function MovePrev() As Boolean
    If HasPrev Then
        MovePrev = True
        s.GlobalOffset = s.GlobalOffset - s.Step
    Else
        MovePrev = False
    End If
End Function


' '@Description("Returns the native index of the specified item")
' Private Function TryGetIndex(ByRef ipLocalOffset As Long, ByRef ipIndex As Long) As Boolean
' Attribute TryGetIndex.VB_Description = "Returns the native index of the specified item"
    
'     Dim myNewIndex As Long
'     myNewIndex = s.FirstIndex + s.GlobalOffset + (ipLocalOffset * s.Step)
    
'     If myNewIndex >= s.FirstIndex And myNewIndex <= s.LastIndex Then
'         ipIndex = myNewIndex
'         TryGetIndex = True
'     Else
'         TryGetIndex = False
'     End If
    
' End Function


'@Description("Get the items at the global + local offset.  Default local offset = 0. offset = number of steps")
Public Property Get CurItem(ByVal ipLocalOffset As Long) As Variant
Attribute CurItem.VB_Description = "Get the items at the global + local offset.  Default local offset = 0. offset = number of steps"

    Dim myLoc As Long: myLoc = s.GlobalOffset + ipLocalOffset * s.Step + s.FirstIndex
    If myLoc < s.StartIndex Or myLoc > s.EndIndex Then
        CurItem = Null
        Exit Property
    End If
    
    Dim myResult As Variant
    Select Case True
    
        Case GroupInfo.IsArray(s.Host):
            myResult = Array(s.Host(myLoc))
        Case GroupInfo.IsString(s.Host):
            myResult = Array(VBA.Mid$(s.Host, myLoc, 1))
        Case GroupInfo.IsItemByToArray(s.Host):
            myResult = Array(s.Items(myLoc))
        Case GroupInfo.IsDictionary(s.Host):
            Dim myKey As Variant: myKey = Array(s.Keys(myLoc))
            myResult = Array(s.Host.Item(myKey(0)))
        Case GroupInfo.IsIndexFrom1(s.Host):
            myResult = Array(s.Host.Item(myLoc))
        Case GroupInfo.IsIndexFrom0(s.Host):
            myResult = Array(s.Host.Item(myLoc))
        Case Else:
            myResult = Array(s.Host)
        
    End Select
    
    
    If VBA.IsObject(myResult(0)) Then
        Set CurItem = myResult(0)
    Else
        CurItem = myResult(0)
    End If
     
End Property


''@Description("Assigns a value to the curretn non object Item")
Public Property Let CurItem(ByVal ipLocalOffset As Long, ByVal ipItem As Variant)

    Dim myLoc As Long: myLoc = s.GlobalOffset + ipLocalOffset * s.Step + s.FirstIndex
    If myLoc < s.StartIndex Or myLoc > s.EndIndex Then
        ' is it preferable to error here
        Exit Property
    End If
    
    Select Case True
    
        Case GroupInfo.IsArray(s.Host):             s.Host(s.FirstIndex + myLoc) = ipItem
        Case GroupInfo.IsString(s.Host):            Mid$(s.Host, myLoc, 1) = ipItem
        Case GroupInfo.IsItemByToArray(s.Host):     s.Items(myLoc) = ipItem
        Case GroupInfo.IsDictionary(s.Host):        s.Host.Item(myLoc) = ipItem
        Case GroupInfo.IsIndexFrom0(s.Host):        s.Host.Item(myLoc) = ipItem
        Case GroupInfo.IsIndexFrom1(s.Host):        s.Host.Item(myLoc) = ipItem
        Case Else:                                  s.Host = ipItem
    End Select
    
End Property


'@("Sets the value of the current object Item")
Public Property Set CurItem(ByVal ipLocalOffset As Long, ByVal ipItem As Variant)

    Dim myLoc As Long: myLoc = s.GlobalOffset + ipLocalOffset * s.Step + s.FirstIndex
    If myLoc < s.StartIndex Or myLoc > s.EndIndex Then
        ' is it preferable to error here
        Exit Property
    End If
    
    Select Case True
    
        Case GroupInfo.IsArray(s.Host):             Set s.Host(s.FirstIndex + myLoc) = ipItem
            'Case GroupInfo.IsString(s.Host):            Mid$(s.Host, myLoc, 1) = ipItem
        Case GroupInfo.IsItemByToArray(s.Host):     Set s.Items(myLoc) = ipItem
        Case GroupInfo.IsDictionary(s.Host):        Set s.Host.Item(myLoc) = ipItem
        Case GroupInfo.IsIndexFrom0(s.Host):        Set s.Host.Item(myLoc) = ipItem
        Case GroupInfo.IsIndexFrom1(s.Host):        Set s.Host.Item(myLoc) = ipItem
        Case Else:                                  Set s.Host = ipItem
    End Select
    
End Property


'@("Returns the host index of the current item")
Public Property Get CurKey(ByVal ipLocalOffset As Long) As Variant

    Dim myLoc As Long: myLoc = s.GlobalOffset + ipLocalOffset * s.Step + s.FirstIndex
    If myLoc < s.StartIndex Or myLoc > s.EndIndex Then
        CurKey = Null
        Exit Property
    End If
    
    Dim myResult As Variant
    Select Case True
    
        Case GroupInfo.IsArray(s.Host):                 myResult = Array(myLoc)
        Case GroupInfo.IsString(s.Host):                myResult = Array(myLoc)
        Case GroupInfo.IsItemByToArray(s.Host):         myResult = Array(myLoc)
        Case GroupInfo.IsDictionary(s.Host):            myResult = Array(s.Keys(myLoc))
        Case GroupInfo.IsIndexFrom0(s.Host):            myResult = Array(myLoc)
        Case GroupInfo.IsIndexFrom1(s.Host):            myResult = Array(myLoc)
        Case Else:                                      myResult = Array(1)
            
    End Select
    
    If VBA.IsObject(myResult(0)) Then
        Set CurKey = myResult(0)
    Else
        CurKey = myResult(0)
    End If

End Property


'@Description("Returns the 1 based index of the current item")
Public Property Get CurOffset(ByVal ipLocalOffset As Long) As Variant
Attribute CurOffset.VB_Description = "Returns the 1 based index of the current item"

    Dim myOffset As Long: myOffset = s.GlobalOffset + ipLocalOffset * s.Step
    If myOffset < 0 Or myOffset > s.Size - 1 Then
        CurOffset = Null
    Else
        CurOffset = myOffset
    End If
    

End Property

Public Property Get TypeName() As String
    TypeName = MY_TYPENAME
End Property

Public Property Get FirstIndex() As Long
    FirstIndex = s.FirstIndex
End Property

Public Property Get StartIndex() As Long
    StartIndex = s.StartIndex
End Property

Public Property Get LastIndex() As Long
    LastIndex = s.LastIndex
End Property

Public Property Get EndIndex() As Long
    EndIndex = s.EndIndex
End Property

Private Sub AlertInvalidFTs(ByRef ipType As String, ByRef ipValue As Long)

    Err.Raise 17 + vbObjectError, _
    Fmt.Text("VBALib: Invalid FTS Paramater: {0}", ipType), _
    Fmt.Text("Expecting {0} to be in the normalised range {1} to {2}. Got {3}", ipType, s.GlobalOffset, s.Size, ipValue)
        
End Sub


Private Sub AlertNotNumber(ByRef ipItem As Variant)

    Err.Raise 17 + vbObjectError, _
    "IterItems: Invalid Number", _
    Fmt.Text("Expecting an number. Got Type: {0}", VBA.TypeName(ipItem))
        
End Sub


Private Sub AlertNegativeStep(ByRef ipStep As Variant)

    Err.Raise 17 + vbObjectError, _
    "IterItems: Invalid Negative Step", _
    Fmt.Text("Negative steps ({0}) are deprecated.  Use MovePrev instead", ipStep)
        
End Sub


Private Sub AlertNoItems()
    Err.Raise 17 + vbObjectError, _
    "Iteritems: No Items", _
    "Cannot iterate an Item count of 0"
End Sub


Private Sub AlertInvalidStartEnd(ByRef ipStart As Long, ByRef ipEnd As Long)
    Err.Raise 17 + vbObjectError, _
    "IterItems.FTS|", _
    Fmt.Text("Expcting start index before End index. Got Start of {0}, End of {1}", ipStart, ipEnd)
End Sub
