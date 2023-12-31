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
'@IgnoreModule FunctionReturnValueAlwaysDiscarded
Option Explicit
'@Exposed
'@PredeclaredId
'@Folder("Helpers")
'@ModuleDescription "Static class: Repository for non standard conversions methods"
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
' Code line limit should be 120 characters.
' Comment line limit should be 80 characters
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C

Private Const mcMethods                 As String = "ToString,ToJSon,ToJson,JSon,Json,Value,value"
Private Const mcLocation                As String = "VBALib.Stringifier"

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

        ResetMarkup
        
    End If

End Sub


'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' Implementation
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Friend Sub ResetMarkup()
    
    Me _
        .ResetArrayMarkup _
        .ResetObjectMarkup _
        .ResetItemSeparator
        
    
End Sub


Public Function ResetArrayMarkup _
( _
    Optional ByVal ipOpen As String = Char.LBracket, _
    Optional ByVal ipClose As String = Char.RBracket, _
    Optional ByVal ipSeparator As String = Char.comma _
) As Stringifier
    With s
    
        .ArrayOpen = ipOpen
        .ArrayClose = ipClose
        .ArrayItemSeperator = ipSeparator
        
    End With
    
    Set ResetArrayMarkup = Me
    
End Function


Public Function ResetObjectMarkup _
( _
    Optional ByVal ipOpen As String = Char.LCurly, _
    Optional ByVal ipClose As String = Char.RCurly, _
    Optional ByVal ipSeparator As String = Char.comma _
) As Stringifier
    With s
    
        .ObjectOpen = ipOpen
        .ObjectClose = ipClose
        .ObjectItemSeperator = ipSeparator
        
    End With
    
    Set ResetObjectMarkup = Me
    
End Function


Public Function ResetItemSeparator(Optional ByVal ipSeparator As String = Char.comma) As Stringifier

    With s
    
        .ItemSeparator = ipSeparator
        
    End With
    
    Set ResetItemSeparator = Me
    
End Function


'@Description("Converts the parameters a single string.")
Public Function ToString(ParamArray ipArgs() As Variant) As String
Attribute ToString.VB_Description = "Converts the parameters a single string."
    
    ToString = vbNullString

    If Arrays.IsEmptyParamArray(ipArgs) Then Exit Function
    
    ' Convert each item in ipParamArray to a string representation
    Dim myArg As Variant
    Dim MyParamsAsString As Lyst
    Set MyParamsAsString = Lyst.Deb
      
    For Each myArg In ipArgs
            
        Select Case True
        
            Case Types.Group.IsPrimitive(myArg)
            
                MyParamsAsString.AddValidatedItem StringifyPrimitive(myArg)
                
                
            Case Types.Group.IsAdmin(myArg)
             
                MyParamsAsString.AddValidatedItem StringifyAdmin(myArg)
        
                
            Case Types.Group.IsNotIterable(myArg)
            
                MyParamsAsString.AddValidatedItem StringifyNonIterableObject(myArg)
    
                
            Case Types.Group.IsIterable(myArg)
            
                MyParamsAsString.AddValidatedItem StringifyIterable(myArg)
            
                
            Case Else
            
            'todo: Add error message
                Debug.Print "Missing error message in ToString"
        
                
        End Select
        
    Next
    
    ToString = VBA.Join(MyParamsAsString.ToArray, s.ItemSeparator)
    
    ResetMarkup
    
End Function


 Friend Function StringifyPrimitive(ByVal ipPrimitive As Variant) As String
    
    StringifyPrimitive = CStr(ipPrimitive)
    
End Function


Friend Function StringifyAdmin(ByVal ipAdmin As Variant) As String

    Select Case True
            
        Case VBA.Information.IsEmpty(ipAdmin)
        
            StringifyAdmin = Types.Name.OfEmpty
        
            
        Case Types.IsNothing(ipAdmin)
        
            StringifyAdmin = s.ObjectOpen & "Nothing" & s.ObjectClose

            
        Case Types.IsNull(ipAdmin)
    
            StringifyAdmin = Types.Name.OfNull
        
            
        Case Types.IsUnknown(ipAdmin)
        
            StringifyAdmin = Types.Name.OfUnknown
            
            
        Case VBA.Information.IsError(ipAdmin)
            
            StringifyAdmin = _
                s.ObjectOpen _
                & VBA.Join _
                ( _
                    Array _
                    ( _
                        Char.Quotation & Types.Name.OfError & Char.Quotation, _
                        Err.Source, _
                        CStr(Err.Number), _
                        Err.Description, _
                        Err.HelpFile _
                    ), _
                    s.ItemSeparator _
                ) _
                & s.ObjectClose

            
    End Select
    
End Function


Friend Function StringifyIterable(ByRef ipIterable As Variant) As String

    Guard InvalidIterable, Types.Group.IsNotIterable(ipIterable), mcLocation & "pvStringifyIterable"
    
    StringifyIterable = vbNullString
    Dim myLyst As Lyst
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
    
            Set myLyst = GetIterableItemsByEnumAsLystOfStrings(ipIterable)
        
            
        Case Types.Group.IsIterableKeysByEnum(ipIterable)
        
            Set myLyst = GetIterableKeysByEnumAsLystOfStrings(ipIterable)
           
            
        Case Else
        
            VBA.Err.Raise _
                17 + vbObjectError, _
                "VBALIB.Stringifier.pvStringifyIterable", _
                "The Iterable Type '" & TypeName(ipIterable) & "' is not known"
            
                
    End Select
    
    StringifyIterable = myOpen & VBA.Join(myLyst.ToArray, s.ItemSeparator) & myClose
    
End Function

Friend Function GetIterableItemsByEnumAsLystOfStrings(ByVal ipIterable As Variant) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    Dim myItem As Variant
    For Each myItem In ipIterable
    
        myLyst.AddValidatedItem ToString(myItem)
        
    Next
    
    Set GetIterableItemsByEnumAsLystOfStrings = myLyst
        
End Function


Friend Function GetIterableKeysByEnumAsLystOfStrings(ByVal ipIterable As Variant) As Lyst

    Dim myLyst As Lyst
    Set myLyst = Lyst.Deb
    
    Dim myKey As Variant
    For Each myKey In ipIterable.Keys
    
        Dim myKeyStr As String
        myKeyStr = ToString(myKey)
        
        Dim myItemStr As String
        myItemStr = ToString(ipIterable.Item(myKey))
        
        myLyst.AddValidatedItem Char.Quotation & myKeyStr & Char.Quotation & Char.Colon & Char.Space & myItemStr
        
    Next

    Set GetIterableKeysByEnumAsLystOfStrings = myLyst

End Function


' Friend Function GetLystOfIterableItemsByIndexStrings(ByVal ipIterable As Variant) As Lyst

'     Dim myLyst As Lyst
'     Set myLyst = Lyst.Deb
    
'     Dim myIndex As Long
'     Dim myResult As Result
'     If Types.Iterable.TryExtent(ipIterable, myResult).Failed Then Exit Function
        
    
'     For myIndex = myResult.Item(Types.Iterable.GetConst(TryFirstIndexResult)) To myResult.Item(ExtentResult.IsFirstIndex) 'item 0 is startindex
    
'         myLyst.AddValidatedItem ToString(ipIterable(myIndex))
        
'     Next
    
'     Set GetLystOfIterableItemsByIndexStrings = myLyst
    
' End Function


'@Description("Trys to get string description of a non iterable object. Defaullt result is '{Object "Name"}'"
Friend Function StringifyNonIterableObject(ByVal ipObject As Object) As String

    Dim myMethods As Variant
    myMethods = VBA.Split(mcMethods, Char.comma)
    
    Dim myString As String
    
    Dim myMethod As Variant
    For Each myMethod In myMethods
    
        On Error Resume Next
        myString = CallByName(ipObject, myMethod, vbGet)
        If Err.Number = 0 Then
        
             StringifyNonIterableObject = s.ObjectOpen & myString & s.ObjectClose
            On Error GoTo 0
            Exit Function
            
        End If
        
        On Error GoTo 0
        On Error Resume Next
        myString = CallByName(ipObject, myMethod, vbMethod)
        If Err.Number = 0 Then
        
             StringifyNonIterableObject = s.ObjectOpen & myString & s.ObjectClose
            On Error GoTo 0
            Exit Function
            
        End If
    
        On Error GoTo 0
        
    Next
    
    ' MyStringAssignment is for the err.number side effect
    On Error Resume Next
    
    'try default method
    '@Ignore ImplicitUnboundDefaultMemberAccess, AssignmentNotUsed
    myString = ipObject
    If Err.Number = 0 Then
    
        On Error GoTo 0
         StringifyNonIterableObject = s.ObjectOpen & myString & s.ObjectClose
        Exit Function
        
    End If
    
    On Error GoTo 0
     StringifyNonIterableObject = s.ObjectOpen & Types.Name.OfObject & Char.Space & Char.Quotation & TypeName(ipObject) & Char.Quotation & s.ObjectClose
        
End Function
