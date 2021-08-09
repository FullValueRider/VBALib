VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stringifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Static class: Repository for non standard conversions methods"
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder("Helpers")
'@ModuleDescription "Static class: Repository for non standard conversions methods"
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Private Const miMethods   As String = "ToString,ToJSon,ToJson,JSon,Json,Value,value"

Public Enum MarkupEnum

    ARRAY_SEPARATOR = 0
    ARRAY_LEFT = 1
    ARRAY_RIGHT = 2

End Enum

Private Type State

    ItemSeparator           As String
    ArrayOpen               As String
    ArrayClose              As String
    ArrayItemSeperator      As String
    ObjectOpen              As String
    ObjectClose             As String
    ObjectItemSeperator     As String
    
End Type

Private s                   As State
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Class admin
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Sub Class_Initialize()

    If Me Is Stringifier Then

        With s
        
            .ItemSeparator = Char.comma
            .ArrayOpen = Char.LBracket
            .ArrayClose = Char.RBracket
            .ArrayItemSeperator = Char.comma
            .ObjectOpen = Char.LCurly
            .ObjectClose = Char.RCurly
            .ObjectItemSeperator = Char.comma
    
        End With
        
    End If

End Sub


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Implementation
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Sub SetArrayMarkup _
( _
    Optional ByVal ipOpen As String = Char.LBracket, _
    Optional ByVal ipClose As String = Char.RBracket, _
    Optional ByVal ipSeparator As String = Char.comma _
)
    With s
    
        .ArrayOpen = ipOpen
        .ArrayClose = ipClose
        .ArrayItemSeperator = ipSeparator
        
    End With
    
End Sub


Public Sub SetObjectMarkup _
( _
    Optional ByVal ipOpen As String = Char.LCurly, _
    Optional ByVal ipClose As String = Char.RCurly, _
    Optional ByVal ipSeparator As String = Char.comma _
)
    With s
    
        .ObjectOpen = ipOpen
        .ObjectClose = ipClose
        .ObjectItemSeperator = ipSeparator
        
    End With
    
End Sub


Public Sub SetItemSeparator(Optional ByVal ipSeparator As String = Char.comma)

    With s
    
        .ItemSeparator = ipSeparator
        
    End With
    
End Sub

'@Description("Converts the parameters a single string.")
Public Function ToString(ParamArray ipArgs() As Variant) As String
Attribute ToString.VB_Description = "Converts the parameters a single string."

    If UBound(ipArgs) = -1 Then
    
        ToString = vbNullString
        Exit Function
            
    End If
        
    ' Convert each item in ipParamArray to a string representation
    Dim myArg As Variant
    Dim MyParamsAsString As Lyst
    Set MyParamsAsString = Lyst.Deb
    For Each myArg In ipArgs
            
        Select Case True
        
            Case Types.Group.IsPrimitive(myArg)
            
                MyParamsAsString.Add pvStringifyPrimitive(myArg)
        
            Case Types.Group.IsIterable(myArg)
            
                MyParamsAsString.Add pvStringifyIterable(myArg)
            
            Case Types.Group.IsAdmin(myArg)
             
                MyParamsAsString.Add pvStringifyAdmin(myArg)
                
            Case VBA.IsObject(myArg)
            
                MyParamsAsString.Add pvStringifyObject(myArg)
        
            Case Else
            
                Debug.Print "Missing error message in ToString"
        
        End Select
        
    Next
    
    ToString = VBA.Join(MyParamsAsString.ToArray, s.ItemSeparator)
    
    SetArrayMarkup
    SetObjectMarkup
    
End Function


 Public Function pvStringifyPrimitive(ByVal ipPrimitive As Variant) As String
    
    pvStringifyPrimitive = CStr(ipPrimitive)
    
End Function


Public Function pvStringifyAdmin(ByVal ipAdmin As Variant) As String

    Select Case True
    
        
            
        Case Types.IsNothing(ipAdmin)
        
            pvStringifyAdmin = s.ObjectOpen & Char.Quotation & "Object" & Char.Quotation & " Nothing" & s.ObjectClose
            
        Case VBA.IsNull(ipAdmin)
    
            pvStringifyAdmin = TypeName(Null)
        
        Case VBA.IsError(ipAdmin)
        
            Dim myLyst As Lyst
            Set myLyst = Lyst.Deb
            
            myLyst.Add Char.Quotation & "Error" & Char.Quotation
            myLyst.Add Err.Source
            myLyst.Add CStr(Err.Number)
            myLyst.Add Err.Description
            myLyst.Add Err.HelpFile
            
            pvStringifyAdmin = s.ObjectOpen & VBA.Join(myLyst.ToArray, s.ItemSeparator) & s.ObjectClose
            
        Case VBA.IsEmpty(ipAdmin)
        
            pvStringifyAdmin = "Empty"
            
        Case Types.IsArray(ipAdmin)
    
            pvStringifyAdmin = pvStringifyIterable(ipAdmin)
        
    End Select
    
End Function


Public Function pvStringifyIterable(ByRef ipIterable As Variant) As String

    Dim myLyst As Lyst
    '@Ignore AssignmentNotUsed
    Set myLyst = Lyst.Deb
    
    Dim myOpen As String
    Dim myClose As String
    If VBA.IsArray(ipIterable) Then
    
        myOpen = s.ArrayOpen
        myClose = s.ArrayClose
    Else
    
        myOpen = s.ObjectOpen
        myClose = s.ObjectClose
        
    End If

    Select Case True
    
        Case Types.Group.IsIterableItemsByEnum(ipIterable)
    
            Set myLyst = pvGetIterableItemsByEnumAsLystOfStrings(ipIterable)
        
        Case Types.Group.IsIterableKeysByEnum(ipIterable)
        
            Set myLyst = pvGetIterableKeysByEnumAsLystOfStrings(ipIterable)
            
        Case Else
        
            VBA.Err.Raise _
                17 + vbObjectError, _
                "VBALIB.Stringifier.pvStringifyIterable", _
                "The Iterable Type '" & TypeName(ipIterable) & "' is not known"
            
    End Select
    
    pvStringifyIterable = myOpen & VBA.Join(myLyst.ToArray, s.ItemSeparator) & myClose
    
End Function


Public Function pvGetIterableItemsByEnumAsLystOfStrings(ByVal ipIterable As Variant) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    Dim myItem As Variant
    For Each myItem In ipIterable
    
        myLyst.Add ToString(myItem)
        
    Next
    
    Set pvGetIterableItemsByEnumAsLystOfStrings = myLyst
        
End Function


Public Function pvGetIterableKeysByEnumAsLystOfStrings(ByVal ipIterable As Variant) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    Dim myKey As Variant
    For Each myKey In ipIterable.Keys
    
        myLyst.Add Char.Quotation & ToString(myKey) & Char.Quotation & Char.Space & ToString(ipIterable.Item(myKey))
        
    Next

    Set pvGetIterableKeysByEnumAsLystOfStrings = myLyst

End Function


Public Function pvGetLystOfIterableItemsByIndexStrings(ByVal ipIterable As Variant) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    Dim myIndex As Long
    Dim myStartIndex As Long
    myStartIndex = Types.Iterable.GetFirstIndex(ipIterable)
    Dim myLastIndex As Long
    myLastIndex = Types.Iterable.GetLastIndex(ipIterable)
    For myIndex = myStartIndex To myLastIndex
    
        myLyst.Add ToString(ipIterable(myIndex))
        
    Next
    
    Set pvGetLystOfIterableItemsByIndexStrings = myLyst
    
End Function


'@Description("Trys to get string description of a non iterable object. Defaullt result is '{Object "Name"}'"
Public Function pvStringifyObject(ByVal ipObject As Object) As String

    Dim myMethods As Variant
    myMethods = VBA.Split(miMethods, Char.comma)
    
    Dim myString As String
    
    Dim myMethod As Variant
    For Each myMethod In myMethods
    
        On Error Resume Next
        myString = CallByName(ipObject, myMethod, VbGet)
        If Err.Number = 0 Then
        
             pvStringifyObject = s.ObjectOpen & myString & s.ObjectClose
            On Error GoTo 0
            Exit Function
            
        End If
        
        On Error GoTo 0
        On Error Resume Next
        myString = CallByName(ipObject, myMethod, VbMethod)
        If Err.Number = 0 Then
        
             pvStringifyObject = s.ObjectOpen & myString & s.ObjectClose
            On Error GoTo 0
            Exit Function
            
        End If
    
        On Error GoTo 0
        
    Next
    
    ' final try using default member
    On Error Resume Next
    
    'try default method
    '@Ignore ImplicitUnboundDefaultMemberAccess
    myString = ipObject
    If Err.Number = 0 Then
    
        On Error GoTo 0
         pvStringifyObject = s.ObjectOpen & myString & s.ObjectClose
        Exit Function
        
    End If
    
    On Error GoTo 0
     pvStringifyObject = s.ObjectOpen & "Object " & Char.Quotation & TypeName(ipObject) & Char.Quotation & s.ObjectClose
        
End Function
