VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypesNamesArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit
'@PredeclaredId
'@Exposed
'@Folder("VBALib")

    

' The odd one out as we can't get the name 'variant' from typename

Public Property Get ofVariant() As String
    ofVariant = "Variant()"
End Property
Public Property Get OfBoolean() As String
    OfBoolean = "Boolean()"
End Property

Public Property Get OfByte() As String
    OfByte = "Byte()"
End Property

Public Property Get OfCurrency() As String
    OfCurrency = "Currency()"
End Property

Public Property Get OfDate() As String
    OfDate = "Date()"
End Property

Public Property Get OfDecimal() As String
    OfDecimal = "Decimal()"   ' remember that decimal can only be represented as a variant
End Property

Public Property Get OfDouble() As String
    OfDouble = "Double()"
End Property

Public Property Get OfInteger() As String
    OfInteger = "Integer()"
End Property

Public Property Get OfLong() As String
    OfLong = "Long()"
End Property

Public Property Get OfLongLong() As String
    OfLongLong = "LongLong()"
End Property

Public Property Get OfLongPtr() As String
    OfLongPtr = "LongPtr()"
End Property

Public Property Get OfSingle() As String
    OfSingle = "Single()"
End Property

Public Property Get OfString() As String
    OfString = "String()"
End Property

Public Property Get OfCollection() As String
    OfCollection = "Collection()"
End Property

Public Property Get OfObject() As String
    OfObject = "Object()"
End Property


Public Property Get OfEmpty() As String
    OfEmpty = "Empty()"
End Property

Public Property Get OfError() As String
    OfError = "Error()"
End Property

Public Property Get OfNothing() As String
    OfNothing = "Nothing()"
End Property

Public Property Get OfNull() As String
    OfNull = "Null()"
End Property

Public Property Get OfUnknown() As String
    OfUnknown = "Unknown()"
End Property
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
'
' External types
'
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========C
Public Property Get OfScriptingDictionary() As String
    OfScriptingDictionary = "Dictionary()"
End Property

Public Property Get OfQueue() As String
    OfQueue = "Queue()"
End Property

Public Property Get OfStack() As String
    OfStack = "Stack()"
End Property

Public Property Get OfArrayList() As String
    OfArrayList = "ArrayList()"
End Property

Public Property Get OfLyst() As String ' Not a typo too many other list types in Office space
     OfLyst = "Lyst()"
End Property

Public Property Get OfKvp() As String
    OfKvp = "Kvp()"
End Property

Public Property Get OfKVPair() As String
     OfKVPair = "KVPair()"
End Property

Public Property Get OfArrayMarkup() As String
    OfArrayMarkup = "ArrayMarkup()"
End Property

Public Property Get OfAutoKeyByNumber() As String
    OfAutoKeyByNumber = "AutoKeybyNumber()"
End Property

Public Property Get OfAutoKeybyString() As String
    OfAutoKeybyString = "AutoKeyByString()"
End Property

Public Property Get OfAutoKeyByArray() As String
    OfAutoKeyByArray = "AotoKeyByArray()"
End Property


