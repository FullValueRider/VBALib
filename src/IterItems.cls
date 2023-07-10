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

Private Type State

    Host                                As Variant   ' holds ipItems
    Keys                                As Variant   ' Keys if ipItems is a dictionary
    Items                               As Variant   ' Items if ipItems is a ToArray object
    
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

    If VBA.IsArray(ipItems) Then
        If ArrayInfo.Ranks(ipItems) > 1 Then
            Err.Raise 17 + vbObjectError, _
                "VBALib.Iteritems: Invalid Array", _
                "Arrays must be 1 dimension only"
        End If
    End If
    
    Select Case True
    
        Case GroupInfo.IsDictionary(ipItems)
            ' this is a horrible hack.
            ' Probably best to make kvp produce arrays rather than seq for Keys and Items
            If VBA.TypeName(ipItems) = "Dictionary" Then
                s.Keys = ipItems.Keys
                ReDim Preserve s.Keys(1 To UBound(s.Keys) + 1)
            Else
                s.Keys = ipItems.Keys.ToArray
            End If
    
            Set s.Host = ipItems
            s.FirstIndex = 1
            s.LastIndex = ipItems.Count
            
        Case GroupInfo.IsItemByToArray(ipItems)
        
            s.Items = ipItems.ToArray
            Set s.Host = ipItems
            s.FirstIndex = 0
            s.LastIndex = ipItems.Count - 1
            
        Case GroupInfo.IsString(ipItems)
        
            s.Host = ipItems
            s.FirstIndex = 1
            s.LastIndex = VBA.Len(ipItems)
            
        Case GroupInfo.IsIndexFrom1(ipItems)
        
            Set s.Host = ipItems
            s.FirstIndex = 1
            s.LastIndex = ipItems.Count
            
        Case GroupInfo.IsIndexFrom0(ipItems)
        
            Set s.Host = ipItems
            s.FirstIndex = 0
            s.LastIndex = ipItems.Count - 1
            
        Case GroupInfo.IsArray(ipItems)
        
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
    
    s.StartIndex = 1
    s.EndIndex = s.LastIndex - s.FirstIndex + 1
    s.Step = 1
    s.GlobalOffset = 0
    s.Size = s.EndIndex
    
    If s.Size = 0 Then
        AlertNoItems
    End If
    
    Set ConstructInstance = Me
    
End Function

'@Description("Sets the subrange and stepping to be used when using Iteritems")
Public Function FTS(Optional ByRef ipStartIndex As Variant = Empty, Optional ByRef ipEndIndex As Variant = Empty, Optional ByRef ipStep As Variant = Empty) As IterItems
Attribute FTS.VB_Description = "Sets the subrange and stepping to be used when using Iteritems"

    s.StartIndex = EnsureNumber(ipStartIndex, s.FirstIndex)
    s.EndIndex = EnsureNumber(ipEndIndex, s.EndIndex)
    s.Step = EnsureNumber(ipStep, s.Step)
    
    If ipStep < 0 Then
        AlertNegativeStep ipStep
    End If
    
    ' now check for valid numbers
    
    If s.StartIndex < 1 Or s.StartIndex > s.Size Then
        AlertInvalidFTs "StartIndex", s.StartIndex
    End If
    
    If s.EndIndex < 1 Or s.StartIndex > s.Size Or s.EndIndex < s.StartIndex Then
        AlertInvalidFTs "EndIndex", s.EndIndex
    End If
    
    s.GlobalOffset = s.StartIndex
    
    Set FTS = Me
    
End Function


Private Function EnsureNumber(ByRef ipNumber As Variant, ByRef ipDefault As Long) As Long

    If VBA.IsEmpty(ipNumber) Then
        EnsureNumber = ipDefault
    Else
        If GroupInfo.IsNumber(ipNumber) Then
            EnsureNumber = ipNumber
        Else
            AlertNotNumber ipNumber
        End If
    End If
    
End Function


'@Description("MoveToEndIndex moves to the nearest step position before or equal to EndIndex)
Public Function MoveToEndIndex() As IterItems
    s.GlobalOffset = ((s.EndIndex - s.StartIndex + 1) Mod s.Step * s.Step)
    Set MoveToEndIndex = Me
End Function


'@Description("Moves to the defined Strt position
Public Function MoveToStartIndex() As IterItems
    s.GlobalOffset = s.StartIndex
    Set MoveToStartIndex = Me
End Function


'@Description("Moves to firstindex, resets StartIndex stop EndIndex")
Public Function MoveToFirst() As IterItems
Attribute MoveToFirst.VB_Description = "Moves to firstindex, resets StartIndex stop EndIndex"
    s.GlobalOffset = 1
    s.StartIndex = 1
    s.EndIndex = s.Size
    s.Step = 1
    Set MoveToFirst = Me
End Function


'@Description("Moves to firstindex, resets StartIndex stop EndIndex")
Public Function MoveToLast() As IterItems
Attribute MoveToLast.VB_Description = "Moves to firstindex, resets StartIndex stop EndIndex"
    s.GlobalOffset = s.Size
    s.StartIndex = 1
    s.EndIndex = s.Size
    s.Step = 1
    Set MoveToLast = Me
End Function


Public Function HasNext(Optional ByRef ipLocalOffset As Long = 0) As Boolean
    HasNext = (s.GlobalOffset + ipLocalOffset + 1 >= s.StartIndex) And (s.GlobalOffset + ipLocalOffset + 1 <= s.EndIndex - s.Step)
End Function

Public Function HasNoNext(Optional ByRef ipLocalOffset As Long = 0) As Boolean
    HasNoNext = Not HasNext(ipLocalOffset)
End Function


Public Function HasPrev(Optional ByRef ipLocalOffset As Long = 0) As Boolean
    HasPrev = (s.GlobalOffset - ipLocalOffset + 1 >= s.StartIndex + s.Step) And (s.GlobalOffset - ipLocalOffset + 1 <= s.EndIndex)
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

'@Description("Returns the native index of the specified item")
Private Function TryGetIndex(ByRef ipLocalOffset As Long, ByRef ipIndex As Long) As Boolean
Attribute TryGetIndex.VB_Description = "Returns the native index of the specified item"
    
    Dim myNewIndex As Long
    myNewIndex = s.FirstIndex + s.GlobalOffset + (ipLocalOffset * s.Step)
    
    If myNewIndex >= s.FirstIndex And myNewIndex <= s.LastIndex Then
        ipIndex = myNewIndex
        TryGetIndex = True
    Else
        TryGetIndex = False
    End If
    
End Function


'@Description("Get the items at the current 1 based offset.  Optionally specify an offset in terms of the number of steps")
Public Property Get CurItem(ByVal ipLocalOffset As Long) As Variant
Attribute CurItem.VB_Description = "Get the items at the current 1 based offset.  Optionally specify an offset in terms of the number of steps"

    Dim myIndex As Long
    If Not TryGetIndex(ipLocalOffset, myIndex) Then
        CurItem = Null
        Exit Property
    End If
    
    Dim myresult As Variant
    Select Case True
    
        Case GroupInfo.IsArray(s.Host):             myresult = Array(s.Host(myIndex))
        Case GroupInfo.IsString(s.Host):            myresult = Array(VBA.Mid$(s.Host, myIndex, 1))
        Case GroupInfo.IsItemByToArray(s.Host):     myresult = Array(s.Items(s.FirstIndex + myIndex - 1))
        Case GroupInfo.IsDictionary(s.Host):        myresult = Array(s.Host.Item(s.Keys(myIndex)))
        Case GroupInfo.IsIndexFrom0(s.Host):        myresult = Array(s.Host.Item(myIndex))
        Case GroupInfo.IsIndexFrom1(s.Host):        myresult = Array(s.Host.Item(myIndex))
        Case Else:                                  myresult = Array(s.Host)
        
    End Select
    
    
    If VBA.IsObject(myresult(0)) Then
        Set CurItem = myresult(0)
    Else
        CurItem = myresult(0)
    End If
     
End Property


''@Description("Assigns a value to the curretn non object Item")
Public Property Let CurItem(ByVal ipLocalOffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long
    If Not TryGetIndex(ipLocalOffset, myIndex) Then
        Exit Property
    End If
    
    Select Case True
    
        Case GroupInfo.IsArray(s.Host):             s.Host(s.FirstIndex + myIndex - 1) = ipItem
        Case GroupInfo.IsString(s.Host):            Mid$(s.Host, myIndex, 1) = ipItem
        Case GroupInfo.IsItemByToArray(s.Host):     s.Items(s.FirstIndex + myIndex - 1) = ipItem
        Case GroupInfo.IsDictionary(s.Host):        s.Host.Item(s.Keys(myIndex)) = ipItem
        Case GroupInfo.IsIndexFrom0(s.Host):        s.Host.Item(myIndex - 1) = ipItem
        Case GroupInfo.IsIndexFrom1(s.Host):        s.Host.Item(myIndex) = ipItem
        Case Else:                                  s.Host = ipItem
        
    End Select
    
End Property

'@("Sets the value of the current object Item")
Public Property Set CurItem(ByVal ipLocalOffset As Long, ByVal ipItem As Variant)

    Dim myIndex As Long
    If Not TryGetIndex(ipLocalOffset, myIndex) Then
        Exit Property
    End If
    
    Select Case True
    
        Case GroupInfo.IsArray(s.Host):                 Set s.Host(s.FirstIndex + myIndex - 1) = ipItem
        'set does not apply for a string
        Case GroupInfo.IsItemByToArray(s.Host):         Set s.Items(s.FirstIndex + myIndex - 1) = ipItem
        Case GroupInfo.IsDictionary(s.Host):            Set s.Host.Item(s.Keys(myIndex)) = ipItem
        Case GroupInfo.IsIndexFrom0(s.Host):            Set s.Host.Item(myIndex - 1) = ipItem
        Case GroupInfo.IsIndexFrom1(s.Host):            Set s.Host.Item(myIndex) = ipItem
        Case Else:                                      Set s.Host = ipItem
            
    End Select
    
End Property

'@("Returns the host index of the current item")
Public Property Get CurKey(ByVal ipLocalOffset As Long) As Variant

    Dim myIndex As Long
    If Not TryGetIndex(ipLocalOffset, myIndex) Then
        CurKey = Null
        Exit Property
    End If
    
    Dim myresult As Variant
    Select Case True
    
        Case GroupInfo.IsArray(s.Host):                 myresult = Array(myIndex)
        Case GroupInfo.IsString(s.Host):                myresult = Array(myIndex)
        Case GroupInfo.IsItemByToArray(s.Host):         myresult = Array(s.FirstIndex + myIndex)
        Case GroupInfo.IsDictionary(s.Host):            myresult = Array(s.Keys(myIndex))
        Case GroupInfo.IsIndexFrom0(s.Host):            myresult = Array(myIndex)
        Case GroupInfo.IsIndexFrom1(s.Host):            myresult = Array(myIndex)
        Case Else:                                      myresult = Array(1)
            
    End Select
    
    If VBA.IsObject(myresult(0)) Then
        Set CurKey = myresult(0)
    Else
        CurKey = myresult(0)
    End If

End Property

'@Description("Returns the 1 based index of the current item")
Public Property Get CurOffset(ByVal ipLocalOffset As Long) As Variant
Attribute CurOffset.VB_Description = "Returns the 1 based index of the current item"

    Dim myIndex As Long
    myIndex = s.GlobalOffset + (s.Step * ipLocalOffset) + 1
    If myIndex >= 1 And myIndex <= s.Size Then
        CurOffset = myIndex
    Else
        CurOffset = Null
    End If
    
End Property





Private Sub AlertInvalidFTs(ByRef ipType As String, ByRef ipValue As Long)

    Err.Raise 17 + vbObjectError, _
        Fmt.Text("VBALib: Invalid FTS Paramater: {0}", ipType), _
        Fmt.Text("Expecting {0} to be in the normalised range {1} to {2}. Got {3}", ipType, s.GlobalOffset, s.Size, ipValue)
        
End Sub


Private Sub AlertNotNumber(ByRef ipItem As Variant)

    Err.Raise 17 + vbObjectError, _
        "VBALib.IterItems: Invalid Number", _
        Fmt.Text("Expecting an number. Got Type: {0}", VBA.TypeName(ipItem))
        
End Sub


Private Sub AlertNegativeStep(ByRef ipStep As Variant)

    Err.Raise 17 + vbObjectError, _
        "VBALib.IterItems: Invalid Negative Step", _
        Fmt.Text("Negative steps ({0}) are deprecated.  Use MovePrev instead", ipStep)
        
End Sub


Private Sub AlertNoItems()
    Err.Raise 17 + vbObjectError, _
        "VBALib.Iteritems: No Items", _
        "Cannot iterate an Item count of 0"
End Sub
