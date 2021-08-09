Attribute VB_Name = "Module1"
Option Explicit

'@IgnoreModule

Sub ttest()

    Dim myarray As Variant
    ReDim myarray(1 To 5)
    Debug.Print TypeName(myarray)
    Dim myList As ArrayList
    Set myList = New ArrayList
    myList.AddRange myarray
    
End Sub
Public Sub ArrayTypeInfo()

    Dim myUbound As Long
    Dim myUboundMsg As String
    On Error Resume Next
    Debug.Print , , , "TypeName", "VarType", "IsArray", "IsNull", "IsEmpty", "Ubound"
    
    Dim myLongs() As Long
    myUbound = UBound(myLongs)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "Dim myLongs() As Long", , TypeName(myLongs), VarType(myLongs), IsArray(myLongs), IsNull(myLongs), IsEmpty(myLongs), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
    myLongs = Array(1&, 2&, 3&, 4&, 5&)
    myUbound = UBound(myLongs)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "myLongs = Array(1&, 2&, 3&, 4&, 5&)", TypeName(myLongs), VarType(myLongs), IsArray(myLongs), IsNull(myLongs), IsEmpty(myLongs), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
    ReDim myLongs(1 To 5)
    myUbound = UBound(myLongs)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "ReDim myLongs(1 To 5)", , TypeName(myLongs), VarType(myLongs), IsArray(myLongs), IsNull(myLongs), IsEmpty(myLongs), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
    Debug.Print
    Dim myarray As Variant
    myUbound = UBound(myarray)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "Dim myArray As Variant", , TypeName(myarray), VarType(myarray), IsArray(myarray), IsNull(myarray), IsEmpty(myarray), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
    myarray = Array()
    myUbound = UBound(myarray)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "myArray = Array()", , TypeName(myarray), VarType(myarray), IsArray(myarray), IsNull(myarray), IsEmpty(myarray), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
    myarray = Array(1, 2, 3, 4, 5)
    myUbound = UBound(myarray)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "myArray = Array(1, 2, 3, 4, 5)", TypeName(myarray), VarType(myarray), IsArray(myarray), IsNull(myarray), IsEmpty(myarray), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
    myarray = myLongs
    myUbound = UBound(myarray)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "myArray = myLongs", , TypeName(myarray), VarType(myarray), IsArray(myarray), IsNull(myarray), IsEmpty(myarray), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
    Debug.Print
    Dim myArrayOfVar() As Variant
    myUbound = UBound(myArrayOfVar)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "Dim myArrayOfVar() As Variant", TypeName(myArrayOfVar), VarType(myArrayOfVar), IsArray(myArrayOfVar), IsNull(myArrayOfVar), IsEmpty(myArrayOfVar), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
    myArrayOfVar = Array()
    myUbound = UBound(myArrayOfVar)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "myArrayOfVar = Array()", , TypeName(myArrayOfVar), VarType(myArrayOfVar), IsArray(myArrayOfVar), IsNull(myArrayOfVar), IsEmpty(myArrayOfVar), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
    myArrayOfVar = Array(1&, 2&, 3&, 4&, 5&)
    myUbound = UBound(myArrayOfVar)
    myUboundMsg = IIf(Err.Number = 0, CStr(myUbound), "Ubound error")
    Debug.Print "myArrayOfVar = Array(1&, 2&, 3&, 4&, 5&)", TypeName(myArrayOfVar), VarType(myArrayOfVar), IsArray(myArrayOfVar), IsNull(myArrayOfVar), IsEmpty(myArrayOfVar), myUboundMsg
    On Error GoTo 0
    On Error Resume Next
    
End Sub
Sub TestObjectsBeforeAndAfterNewing()

    Dim myColl As Collection
    Dim myList As ArrayList
    Dim myQueue As Queue
    Dim myStack As Stack
    Dim myDict As Scripting.Dictionary
    
    Debug.Print "Collection", TypeName(myColl)
    Debug.Print "Arraylist", TypeName(myList)
    Debug.Print "QUeue", TypeName(myQueue)
    Debug.Print "Stack", TypeName(myStack)
    
    Set myColl = New Collection
    Set myList = New ArrayList
    Set myQueue = New Queue
    Set myDict = New Scripting.Dictionary
    
    Debug.Print "Collection", TypeName(myColl)
    Debug.Print "Arraylist", TypeName(myList)
    Debug.Print "QUeue", TypeName(myQueue)
    Debug.Print "Stack", TypeName(myStack)
    
End Sub

Public Sub NoParamArray(ParamArray ipArgs() As Variant)

    Debug.Print UBound(ipArgs)
End Sub

Public Sub TestNoParamArray()
    
    NoParamArray 1, 2, 3, 4, 5
    NoParamArray
End Sub


Public Sub ForeachStack()

    Dim myStack As Stack
    Set myStack = New Stack
    
    myStack.Push 10
    myStack.Push 20
    myStack.Push 30
    myStack.Push 40
    
    Dim myItem As Variant
    For Each myItem In myStack
    
        Debug.Print myItem
    Next
    
End Sub

Public Sub ForeachQueUE()

    Dim myQueue As Queue
    Set myQueue = New Queue
    
    myQueue.enqueue 10
    myQueue.enqueue 20
    myQueue.enqueue 30
    myQueue.enqueue 40
    
    Dim myItem As Variant
    For Each myItem In myQueue
    
        Debug.Print myItem
    Next

End Sub
