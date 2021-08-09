Attribute VB_Name = "TypesIsGroupType"
'@IgnoreModule UnassignedVariableUsage
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Rubberduck.AssertClass
'Private Fakes As Rubberduck.FakesProvider


'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = New Rubberduck.AssertClass
'    Set Fakes = New Rubberduck.FakesProvider
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
'    Set Fakes = Nothing
End Sub


''@TestInitialize
'Private Sub TestInitialize()
'    'This method runs before every test in the module..
'End Sub
'
'
''@TestCleanup
'Private Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub

'@TestMethod("IsNumber")
Private Sub Test01_ByteIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Byte
    myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsNumber")
Private Sub Test02_CurrencyIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    Dim myValue As Currency
    myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsNumber")
Private Sub Test03_DateIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Date
    myValue = "28-Jun-2021"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsNumber")
Private Sub Test04_DecimalIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    ' VBA does not support declaring a variable as Decimal
    Dim myValue As Variant
    myValue = VBA.CDec(42)
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsNumber")
Private Sub Test05_DoubleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Double
    myValue = VBA.CDec(42)
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsNumber")
Private Sub Test06_IntegerIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Integer
    myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsNumber")
Private Sub Test07_LongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Long
    myValue = 42&
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsNumber")
Private Sub Test08_LongLongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As LongLong
    myValue = 42^
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsNumber")
Private Sub Test09_SingleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Single
    myValue = 42!
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsNumber")
Private Sub Test10_LongPtrIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As LongPtr
    myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsNumber")
Private Sub Test11_StringIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myValue As String
    myValue = "42"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("AnObject")
Private Sub Test12_IsAnObjectIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.IsObject(New Collection, New ArrayList, New Stack, New Queue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an Object: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("AnObject")
Private Sub Test13_IsAnObjectIsFalse()
    On Error GoTo TestFail


    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean

    'Act:
    myResult = Types.IsObject(New Collection, New ArrayList, New Stack, 100#)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an Object: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("AnObject")
Private Sub Test14_IsArrayOfObjectIsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myResult As Boolean
    Dim myFirstColl(1 To 5) As Collection
    Dim mySecondColl(1 To 4) As Collection
    'Act:
    myResult = Types.IsObjectArray(myFirstColl, mySecondColl)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an Object: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("AnObject")
Private Sub Test15_IsArrayOfObjectIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Collection
    Dim mySecondArray(1 To 4) As Double
    'Act:
    myResult = Types.IsObjectArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an Object: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsShort")
Private Sub Test16_BooleanIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Boolean
    myValue = True
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsShort")
Private Sub Test17_ByteIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Byte
    myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsShort")
Private Sub Test18_CurrencyIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    Dim myValue As Currency
    myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsShort")
Private Sub Test19_DateIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Date
    myValue = "28-Jun-2021"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsShort")
Private Sub Test20_DecimalIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    ' VBA does not support declaring a variable as Decimal
    Dim myValue As Variant
    myValue = VBA.CDec(42)
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsShort")
Private Sub Test21_DoubleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Double
    myValue = VBA.CDec(42)
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsShort")
Private Sub Test22_IntegerIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Integer
    myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsShort")
Private Sub Test23_LongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Long
    myValue = 42&
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsShort")
Private Sub Test24_LongLongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As LongLong
    myValue = 42^
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsShort")
Private Sub Test25_SingleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Single
    myValue = 42!
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsShort")
Private Sub Test26_LongPtrIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As LongPtr
    myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsShort")
Private Sub Test27_StringIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myValue As String
    myValue = "42"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsPrimitive")
Private Sub Test28_BooleanIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Boolean
    myValue = True
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsPrimitive")
Private Sub Test29_ByteIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Byte
    myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsPrimitive")
Private Sub Test30_CurrencyIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    Dim myValue As Currency
    myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsPrimitive")
Private Sub Test31_DateIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Date
    myValue = "28-Jun-2021"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsPrimitive")
Private Sub Test32_DecimalIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    ' VBA does not support declaring a variable as Decimal
    Dim myValue As Variant
    myValue = VBA.CDec(42)
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsPrimitive")
Private Sub Test33_DoubleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Double
    myValue = VBA.CDec(42)
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsPrimitive")
Private Sub Test334_IntegerIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Integer
    myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsPrimitive")
Private Sub Test35_LongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Long
    myValue = 42&
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsPrimitive")
Private Sub Test36_LongLongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As LongLong
    myValue = 42^
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsPrimitive")
Private Sub Test37_SingleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Single
    myValue = 42!
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsPrimitive")
Private Sub Test38_LongPtrIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As LongPtr
    myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsPrimitive")
Private Sub Test39_StringIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As String
    myValue = "42"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsPrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub



'@TestMethod("IsIterableNumber")
Private Sub Test41_ByteIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Byte
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterableNumber")
Private Sub Test32_CurrencyIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    Dim myValue(1 To 5) As Currency
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterableNumber")
Private Sub Test43_DateIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Date
    'myValue = "28-Jun-2021"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterableNumber")
Private Sub Test44_DecimalArrayCannotBeTested()

' Cannot test for decimal array because we can't decalre a type as ecimal
' only convert a variant to decimal
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = True
'
'    ' VBA does not support declaring a variable as Decimal
'    Dim myValue(1 To 5) As Variant
'    myValue(1) = VBA.CDec(42)
'    myValue(2) = VBA.CDec(42)
'    myValue(3) = VBA.CDec(42)
'    myValue(4) = VBA.CDec(42)
'    myValue(5) = VBA.CDec(42)
'
'    Dim myResult As Boolean
'    Debug.Print TypeName(myValue)
'    'Act:
'    myResult = Types.Group.IsIterableNumber(myValue)
'    'Assert:
    Assert.IsTrue True
'
'TestExit:
'    Exit Sub
'
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
End Sub


'@TestMethod("IsIterableNumber")
Private Sub Test45_DoubleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Double
    'myValue = 42#
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableNumber")
Private Sub Test46_IntegerIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Integer
    'myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableNumber")
Private Sub Test47_LongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Long
    'myValue = 42&
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableNumber")
Private Sub Test48_LongLongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As LongLong
    'myValue = 42^
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableNumber")
Private Sub Test49_SingleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Single
    'myValue = 42!
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableNumber")
Private Sub Test50_LongPtrIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As LongPtr
    'myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterableNumber")
Private Sub Test51_StringIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myValue(1 To 5) As String
    
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableNumber(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableShort")
Private Sub Test52_BooeanIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Boolean
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableShort")
Private Sub Test53_ByteIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Byte
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterableShort")
Private Sub Test54_CurrencyIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    Dim myValue(1 To 5) As Currency
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterableShort")
Private Sub Test55_DateIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Date
    'myValue = "28-Jun-2021"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterableShort")
Private Sub Test56_DecimalArrayCannotBeTested()

' Cannot test for decimal array because we can't decalre a type as ecimal
' only convert a variant to decimal
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = True
'
'    ' VBA does not support declaring a variable as Decimal
'    Dim myValue(1 To 5) As Variant
'    myValue(1) = VBA.CDec(42)
'    myValue(2) = VBA.CDec(42)
'    myValue(3) = VBA.CDec(42)
'    myValue(4) = VBA.CDec(42)
'    myValue(5) = VBA.CDec(42)
'
'    Dim myResult As Boolean
'    Debug.Print TypeName(myValue)
'    'Act:
'    myResult = Types.Group.IsIterableShort(myValue)
'    'Assert:
    Assert.IsTrue True
'
'TestExit:
'    Exit Sub
'
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
End Sub


'@TestMethod("IsIterableShort")
Private Sub Test57_DoubleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Double
    'myValue = 42#
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableShort")
Private Sub Test58_IntegerIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Integer
    'myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableShort")
Private Sub Test59_LongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Long
    'myValue = 42&
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableShort")
Private Sub Test60_LongLongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As LongLong
    'myValue = 42^
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableShort")
Private Sub Test61_SingleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Single
    'myValue = 42!
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableShort")
Private Sub Test62_LongPtrIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As LongPtr
    'myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterableShort")
Private Sub Test63_StringIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myValue(1 To 5) As String
    
    
    Dim myResult As Boolean
    Debug.Print TypeName(myValue)
    'Act:
    myResult = Types.Group.IsIterableShort(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterablePrimitive")
Private Sub Test64_BooeanIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Boolean
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterablePrimitive")
Private Sub Test65_ByteIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Byte
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterablePrimitive")
Private Sub Test66_CurrencyIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    Dim myValue(1 To 5) As Currency
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterablePrimitive")
Private Sub Test67_DateIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Date
    'myValue = "28-Jun-2021"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterablePrimitive")
Private Sub Test68_DecimalArrayCannotBeTested()

' Cannot test for decimal array because we can't decalre a type as ecimal
' only convert a variant to decimal
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = True
'
'    ' VBA does not support declaring a variable as Decimal
'    Dim myValue(1 To 5) As Variant
'    myValue(1) = VBA.CDec(42)
'    myValue(2) = VBA.CDec(42)
'    myValue(3) = VBA.CDec(42)
'    myValue(4) = VBA.CDec(42)
'    myValue(5) = VBA.CDec(42)
'
'    Dim myResult As Boolean
'    Debug.Print TypeName(myValue)
'    'Act:
'    myResult = Types.Group.IsIterablePrimitive(myValue)
'    'Assert:
    Assert.IsTrue True
'
'TestExit:
'    Exit Sub
'
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
End Sub


'@TestMethod("IsIterablePrimitive")
Private Sub Test69_DoubleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Double
    'myValue = 42#
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterablePrimitive")
Private Sub Test70_IntegerIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Integer
    'myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterablePrimitive")
Private Sub Test719_LongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Long
    'myValue = 42&
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterablePrimitive")
Private Sub Test72_LongLongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As LongLong
    'myValue = 42^
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterablePrimitive")
Private Sub Test73_SingleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Single
    'myValue = 42!
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterablePrimitive")
Private Sub Test74_LongPtrIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As LongPtr
    'myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterablePrimitive")
Private Sub Test75_StringIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As String
    
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterablePrimitive(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterable")
Private Sub Test78_BooeanIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Boolean
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterable")
Private Sub Test79_ByteIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Byte
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterable")
Private Sub Test80_CurrencyIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    Dim myValue(1 To 5) As Currency
    'myValue = 42
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterable")
Private Sub Test81_DateIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Date
    'myValue = "28-Jun-2021"
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterable")
Private Sub Test82_DecimalArrayCannotBeTested()

' Cannot test for decimal array because we can't decalre a type as ecimal
' only convert a variant to decimal
'
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = True
'
'    ' VBA does not support declaring a variable as Decimal
'    Dim myValue(1 To 5) As Variant
'    myValue(1) = VBA.CDec(42)
'    myValue(2) = VBA.CDec(42)
'    myValue(3) = VBA.CDec(42)
'    myValue(4) = VBA.CDec(42)
'    myValue(5) = VBA.CDec(42)
'
'    Dim myResult As Boolean
'    Debug.Print TypeName(myValue)
'    'Act:
'    myResult = Types.Group.IsIterable(myValue)
'    'Assert:
    Assert.IsTrue True
'
'TestExit:
'    Exit Sub
'
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
End Sub


'@TestMethod("IsIterable")
Private Sub Test83_DoubleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Double
    'myValue = 42#
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterable")
Private Sub Test84_IntegerIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Integer
    'myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterable")
Private Sub Test85_LongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Long
    'myValue = 42&
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterable")
Private Sub Test86_LongLongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As LongLong
    'myValue = 42^
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterable")
Private Sub Test87_SingleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As Single
    'myValue = 42!
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterable")
Private Sub Test88_LongPtrIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As LongPtr
    'myValue = 42
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterable")
Private Sub Test89_StringIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue(1 To 5) As String
    
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterable")
Private Sub Test90_CollectionIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    ' Collection needs to be initialised otherwise it will be seen as just nothing
    ' but shows as variant/object in Locals window
    Dim myValue As Collection
    Set myValue = New Collection
    
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterable")
Private Sub Test91_ArrayListIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    ' ArrayList needs to be initialised otherwise it will be seen as just nothing
    ' but shows as variant/object in Locals window
    Dim myValue As ArrayList
    Set myValue = New ArrayList
    
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterable")
Private Sub Test92_VariantHoldingArrayIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Variant
    Dim myIntegers(1 To 5) As Integer
    'Set myValue = New Collection
    myValue = myIntegers
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterable")
Private Sub Test93_IntegerIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myValue As Integer
    myValue = 1
    
    Dim myResult As Boolean

    'Act:
    '@Ignore UnassignedVariableUsage
    myResult = Types.Group.IsIterable(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''@TestMethod("IsIterableByKeys")
'Private Sub Test94_KvpIsTrue()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = True
'
'    Dim myValue As kvp
'    Set myValue = kvp.deb
'
'    Dim myResult As Boolean
'
'    'Act:
'    myResult = Types.Group.IsIterableByKeys(myValue)
'    'Assert:
'    Assert.AreEqual myExpected, myResult
'
'TestExit:
'    Exit Sub
'
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

'@TestMethod("IsIterableByKeys")
Private Sub Test95_ScriptingDictionaryIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myValue As Scripting.Dictionary
    Set myValue = New Scripting.Dictionary
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterableKeysByEnum(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("IsIterableByKeys")
Private Sub Test96_CollectionIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    Dim myValue As Collection
    Set myValue = New Collection

    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterableKeysByEnum(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("IsIterableByKeys")
Private Sub Test96_VariantEmptyIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    Dim myValue As Variant
    myValue = 1

    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.Group.IsIterableKeysByEnum(myValue)
    'Assert:
    Assert.AreEqual myExpected, myResult

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

