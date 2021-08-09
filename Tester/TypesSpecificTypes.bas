Attribute VB_Name = "TypesSpecificTypes"
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

'@TestMethod("Boolean")
Private Sub Test01_IsBoooleanIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.IsBoolean(True, False, False, True)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Boolean")
Private Sub Test02_IsBoooleanIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.IsBoolean(True, "True", False, "False")
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Boolean")
Private Sub Test03_IsBoooleanArrayIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Boolean
    Dim mySecondArray(1 To 4) As Boolean
    'Act:
    myResult = Types.IsBooleanArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Boolean")
Private Sub Test04_IsBoooleanArrayIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Boolean
    Dim mySecondArray(1 To 4) As Integer
    'Act:
    myResult = Types.IsBooleanArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Byte")
Private Sub Test05_IsByteIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Byte
    myExpected = True
    
    Dim myResult As Byte
   
    'Act:
    myResult = Types.IsByte(CByte(10), CByte(20), CByte(30), CByte(40))
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Byte")
Private Sub Test06_IsByteIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Byte
    myExpected = False
    
    Dim myResult As Byte

    'Act:
    myResult = Types.IsByte(CByte(10), CByte(20), CByte(30), 40)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Byte")
Private Sub Test07_IsByteArrayIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Byte
    myExpected = True
    
    Dim myResult As Byte
    Dim myFirstArray(1 To 5) As Byte
    Dim mySecondArray(1 To 4) As Byte
    'Act:
    myResult = Types.IsByteArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Byte")
Private Sub Test08_IsByteArrayIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Byte
    myExpected = False
    
    Dim myResult As Byte
    Dim myFirstArray(1 To 5) As Byte
    Dim mySecondArray(1 To 4) As Integer
    'Act:
    myResult = Types.IsByteArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Currency")
Private Sub Test09_IsCurrencyIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.IsCurrency(CCur(10), CCur(20), CCur(30), CCur(40))
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Currency")
Private Sub Test10_IsCurrencyIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.IsCurrency(CCur(10), CCur(20), CCur(30), 40)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Currency")
Private Sub Test11_IsCurrencyArrayIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Currency
    Dim mySecondArray(1 To 4) As Currency
    'Act:
    myResult = Types.IsCurrencyArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Currency")
Private Sub Test12_IsCurrencyArrayIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Currency
    Dim mySecondArray(1 To 4) As Integer
    'Act:
    myResult = Types.IsCurrencyArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Date")
Private Sub Test13_IsDateIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.IsDate(CDate("27-jun-21"), CDate("28-jun-21"), CDate("29-jun-21"), CDate("30-jun-21"))
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Date")
Private Sub Test14_IsDateIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.IsDate(CDate("27-jun-21"), CDate("28-jun-21"), CDate("29-jun-21"), 40)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Date")
Private Sub Test15_IsDateArrayIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Date
    Dim mySecondArray(1 To 4) As Date
    'Act:
    myResult = Types.IsDateArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Date")
Private Sub Test16_IsDateArrayIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Date
    Dim mySecondArray(1 To 4) As Integer
    'Act:
    myResult = Types.IsDateArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Decimal")
Private Sub Test17_IsDecimalIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.IsDecimal(CDec(20), CDec(20), CDec(20), CDec(40))
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Decimal")
Private Sub Test18_IsDecimalIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.IsDecimal(CDec(20), CDec(30), CDec(40), 40)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

' can't create a decimal array in VBA, only a variant array containing decimals
''@TestMethod("IsDecimalArray")
'Private Sub Test19_IsTrue()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = True
'
'    Dim myResult As Boolean
''    Dim myFirstArray(1 To 5) As Variant
''    Dim mySecondArray(1 To 4) As Variant
'    'Act:
'    Debug.Print TypeName(Array(CDec(20), CDec(30), CDec(40)))
'    myResult = Types.IsDecimalArray(Array(CDec(20), CDec(30), CDec(40)), Array(CDec(20), CDec(20), CDec(20), CDec(40)))
'    'Assert:
'    Assert.AreEqual myResult, myExpected
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub
'
''@TestMethod("IsDecimalArray")
'Private Sub Test20_IsFalse()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = False
'
'    Dim myResult As Boolean
''    Dim myFirstArray(1 To 5) As Decimal
''    Dim mySecondArray(1 To 4) As Integer
'    'Act:
'    myResult = Types.IsDecimalArray(Array(CDec(20), CDec(30), CDec(40)), Array(20, 30, 40))
'    'Assert:
'    Assert.AreEqual myResult, myExpected
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

'@TestMethod("Double")
Private Sub Test21_IsDoubleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.IsDouble(27.2, 98.3, 0.9, 100#)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Double")
Private Sub Test22_IsDoubleIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean

    'Act:
    myResult = Types.IsDouble(27.2, 98.3, 0.9, 100)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Double")
Private Sub Test23_IsDoubleArrayIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Double
    Dim mySecondArray(1 To 4) As Double
    'Act:
    myResult = Types.IsDoubleArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Double")
Private Sub Test24_IsDoubleArrayIsFalse()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False
    
    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Double
    Dim mySecondArray(1 To 4) As Integer
    'Act:
    myResult = Types.IsDoubleArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Integer")
Private Sub Test25_IsIntegerIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.IsInteger(5, 6, 7, 8)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Integer: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Integer")
Private Sub Test26_IsIntegerIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean

    'Act:
    myResult = Types.IsInteger(5, 6, 7, 100#)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Integer: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Integer")
Private Sub Test27_IsIntegerArrayIsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Integer
    Dim mySecondArray(1 To 4) As Integer
    'Act:
    myResult = Types.IsIntegerArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Integer: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Integer")
Private Sub Test28_IsIntegerArrayIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Integer
    Dim mySecondArray(1 To 4) As Double
    'Act:
    myResult = Types.IsIntegerArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Long: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Long")
Private Sub Test29_IsLongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.IsLong(5&, 6&, 7&, 8&)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Long: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Long")
Private Sub Test30_IsLongIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean

    'Act:
    myResult = Types.IsLong(5&, 6&, 7&, 100#)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Long: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Long")
Private Sub Test31_IsLongArrayIsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Long
    Dim mySecondArray(1 To 4) As Long
    'Act:
    myResult = Types.IsLongArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Long: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Long")
Private Sub Test32_IsLongArrayIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Long
    Dim mySecondArray(1 To 4) As Double
    'Act:
    myResult = Types.IsLongArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Long: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("LongLong")
Private Sub Test33_IsLongLongIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
   
    'Act:
    myResult = Types.IsLongLong(5^, 6^, 7^, 8^)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an LongLong: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("LongLong")
Private Sub Test34_IsLongLongIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean

    'Act:
    myResult = Types.IsLongLong(5^, 6^, 7^, 100#)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an LongLong: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LongLong")
Private Sub Test35_IsLongLongArrayIsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As LongLong
    Dim mySecondArray(1 To 4) As LongLong
    'Act:
    myResult = Types.IsLongLongArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an LongLong: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("LongLong")
Private Sub Test36_IsLongLongArrayIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As LongLong
    Dim mySecondArray(1 To 4) As Double
    'Act:
    myResult = Types.IsLongLongArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an LongLong: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


''@TestMethod("ObjectObject")
'Private Sub Test37_IsObjectObjectIsTrue()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = True
'
'    Dim myResult As Boolean
'    '@Ignore VariableNotAssigned
'    Dim myObject As Object
'    'Set myObject = New object
'    'Set myObject = New object
'    'Act:
'
'
'    '@Ignore UnassignedVariableUsage
'    myResult = Types.IsObjectObject(myObject, myObject, myObject, myObject)
'
'    'Assert:
'    Assert.AreEqual myResult, myExpected
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub


''@TestMethod("ObjectObject")
'Private Sub Test38_IsObjectObjectIsFalse()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = False
'
'    Dim myResult As Boolean
'    '@Ignore VariableNotAssigned
'    Dim myObject As Object
'    'Act:
'
'
'    '@Ignore UnassignedVariableUsage
'    myResult = Types.IsObjectObject(myObject, myObject, myObject, 100#)
'
'    'Assert:
'    Assert.AreEqual myResult, myExpected
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

''@TestMethod("ObjectObject")
'Private Sub Test39_IsObjectObjectArrayIsTrue()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = True
'
'    Dim myResult As Boolean
'    Dim myFirstArray(1 To 5) As Object
'    Dim mySecondArray(1 To 4) As Object
'    'Act:
'    myResult = Types.IsObjectObjectArray(myFirstArray, mySecondArray)
'    'Assert:
'    Assert.AreEqual myResult, myExpected
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

''@TestMethod("ObjectObject")
'Private Sub Test40_IsObjectObjectArrayIsFalse()
'    On Error GoTo TestFail
'
'    'Arrange:
'    Dim myExpected  As Boolean
'    myExpected = False
'
'    Dim myResult As Boolean
'    Dim myFirstArray(1 To 5) As Object
'    Dim mySecondArray(1 To 4) As Double
'    'Act:
'    myResult = Types.IsObjectObjectArray(myFirstArray, mySecondArray)
'    'Assert:
'    Assert.AreEqual myResult, myExpected
'
'TestExit:
'    Exit Sub
'TestFail:
'    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
'    Resume TestExit
'End Sub

'@TestMethod("Single")
Private Sub Test41_IsSingleIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    
    'Act:
    myResult = Types.IsSingle(1!, 2!, 3!, 4!, 5!)
    
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("Single")
Private Sub Test42_IsSingleIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean
    
    'Act:
    myResult = Types.IsSingle(1!, 2!, 3!, 4!, 100#)
    
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Single")
Private Sub Test43_IsSinglerrayIsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Single
    Dim mySecondArray(1 To 4) As Single
    'Act:
    myResult = Types.IsSingleArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("Single")
Private Sub Test45_IsSinglerrayIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As Single
    Dim mySecondArray(1 To 4) As Double
    'Act:
    myResult = Types.IsSingleArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
    
End Sub

'@TestMethod("String")
Private Sub Test45_IsStringIsTrue()
    On Error GoTo TestFail
    
    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True
    
    Dim myResult As Boolean
    
    
    'Act:
    myResult = Types.IsString("Hello", "There", "World", "Nice", "Day")
    
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub


'@TestMethod("String")
Private Sub Test46_IsStringIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean
    
    'Act:
    myResult = Types.IsString("Hello", "There", "World", "Nice", 100#)
    
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("String")
Private Sub Test47_IsStringArrayIsTrue()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = True

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As String
    Dim mySecondArray(1 To 4) As String
    'Act:
    myResult = Types.IsStringArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
        Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("String")
Private Sub Test48_IsStringArrayIsFalse()
    On Error GoTo TestFail

    'Arrange:
    Dim myExpected  As Boolean
    myExpected = False

    Dim myResult As Boolean
    Dim myFirstArray(1 To 5) As String
    Dim mySecondArray(1 To 4) As Double
    'Act:
    myResult = Types.IsStringArray(myFirstArray, mySecondArray)
    'Assert:
    Assert.AreEqual myResult, myExpected

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an Error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

