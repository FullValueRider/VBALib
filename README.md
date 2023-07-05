# VBALib:  Advent Of Code Library in VBA

 I developed this library in response to trying to solve Advent of Code problems in VBA.  
 
 VBA is an OK Oop language but needs a lot of boilerplate code to allow it to be more usable and expressive.

The library offers
- a degree of enhanced reflection,
- an enhanced Collection (SeqC)
- an en enhanced DIctionary (KvpC)
- string interpolation for formatting characters and variables
- Map, Reduce,Filter,and Count functionality on the items in a SeqC or KvpC and a set of functions to use in such operations
- IterItems class for enhanced for each functionality
- Maths functions not in VBA
- Useful global constants such Maximim and Minimum valiues for some integer types

## Example Code: Advent of Code 2015 Day 06
```VBA
'@PredeclaredId
'@Exposed
Option Explicit

Private Const TODAY                             As String = "\Day06.txt"

Private Type State

    Display                                     As BulbDisplay
    Instructions                                As SeqC
    Lights                                      As Variant

End Type

Private s                                       As State


Public Sub Execute()
    Part01
    Part02
End Sub


Private Sub Part01()

    Initialise
    
    Dim myArea As Variant
    myArea = Array(0, 0, 999, 999)
    ReDim Preserve myArea(1 To 4)
    'BulbDisplay is a separate class not presented here
    Set s.Display = BulbDisplay(myArea)
    s.Display.SwitchOff myArea
    
    Dim myItem As Variant
    For Each myItem In s.Instructions
    
        Dim myS As SeqC
        Set myS = myItem
    
        Select Case myS.First

            Case "on": s.Display.SwitchOn myS.Tail.Toarray

            Case "off":  s.Display.SwitchOff myS.Tail.Toarray

            Case "toggle": s.Display.Toggle myS.Tail.Toarray

        End Select

    Next
    
    Dim myResult As Long: myResult = s.Display.LitBulbs(myArea)
    
    fmt.dbg "The answer for Day {0} Part 1 is 543903. Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult

End Sub


Private Sub Part02()

    Initialise
    
    Dim myArea As Variant
    myArea = Array(0, 0, 999, 999)
    ReDim Preserve myArea(1 To 4)
    
    Set s.Display = BulbDisplay(myArea)
    s.Display.SwitchOff myArea
    s.Display.UseBrightness = True
    
    Dim myItem As Variant
    For Each myItem In s.Instructions
        
        Dim myS As SeqC
        Set myS = myItem
    
        Select Case myS.First

            Case "on": s.Display.SwitchOn myS.Tail.Toarray

            Case "off":  s.Display.SwitchOff myS.Tail.Toarray

            Case "toggle": s.Display.Toggle myS.Tail.Toarray

        End Select

    Next
    
    Dim myResult As Long: myResult = s.Display.Brightness(myArea)
  
    fmt.dbg "The answer for Day {0} Part 02 is 14687245. Found is {1}", VBA.Mid$(TODAY, 5, 2), myResult


End Sub

Private Sub Initialise()
    ' Converts a file of strings of the the form
    ' toggle 461,550 through 564,900
    ' turn off 812,389 through 865,874
    ' turn on 599,989 through 806,993
    ' to a SeqC of SeqC where the inner sequences are
    '  String, Long,Long,Long,Long
    ' e.g.
    ' "off", 812,389,564,900
   
    Set s.Instructions = SeqC(Filer.GetFileAsArrayOfStrings(AoCRawData & Year & TODAY)) _
        .mapit(mpMultireplace(SeqC(Array("through ", vbNullString), Array("turn ", vbNullString), Array(chars.twcomma, chars.twSpace), Array("  ", " ")))) _
        .mapit(mpInner(mpSPlit(chars.twSpace))) _
        .mapit(mpInner(mpConvert(e_Convertto.m_Long)))
   
End Sub
```

## Modules

### Chars
Module chars provides text definitions of control characters and some collections of characters as constants.

### Globals
Module Globals provides some useful constants/functions that provide Type values missing from VBA, e.g. MaxLong, MinLong etc

### ComparerHelpers
A module contining some functions that help with comparing values

### Helpers
A module where I put things for which I don't have an obvious home.

## Classes
All classes in this library are created with a PredeclaredIds using the Rubberduck annotations.  This is to enforce the use of the class name qualifier but also to minimise name clashes.

### Class ArrayInfo

#### Function IsAllocated(ByRef ipArray As Variant) As Boolean
Returns True if an array contains elements.  Returns False otherwise

#### Function IsNotAllocated(ByRef ipArray As Variant) As Boolean

#### Function IsArray(ByRef ipArray As Variant, Optional ByRef ipArrayType As e_ArrayType = e_ArrayType.m_AnyArrayType) As Boolean
Returns true if ipArray is an array and is allocated and is of the specified ArrayType.  The default actions is for an array type of Any.

#### Function IsNotArray(ByRef ipArray As Variant, Optional ByRef ipArrayType As e_ArrayType = m_AnyArrayType) As Boolean

#### Function Count(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 0) As Long
Default is to return the total number of elements in an array.  If a range is specified it returns the Ubound - Lbound+ 1 of the rank.
If the array is not allocated returns -1

#### Function Ranks(ByVal ipArray As Variant) As Long
Returns the number of dimensions as specified in a DIm or Redim statement.
If the array is not allocted returns -1

#### Function HasRank(ByRef ipArray As Variant, ByVal ipRank As Long) As Boolean
Returns true if the array dimensions include the specified rank

#### Function HoldsItems(ByRef ipArray As Variant) As Boolean
An alternative for IsAllocated

#### LacksItems(ByRef ipArray As Variant) As Boolean
An alternative for IsNotAllocated. 

#### Public Function FirstIndex(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 1) As long
Alternative to Lbound. Returns 'Empty' if ipRank is not in the array

#### Function LastIndex(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 1) As Variant
ALternative to Ubound.  Returns 'Empty' if ipRank is not in the array

## Comparer classes
Comparer classes are used to provide a compare function for the SeqC and KvpC FilterIt, ReduceIt and CountIt classes
A class is provided for each comparision , equals, not equals, more than, morethan or equal, lessthen and lessthan or equals.
The classes are primed with a reference value at creation or default to a reference of zero.  The cmp classes may rely of a COmpereHelper function for a comparision.
cmp classes implement the IComparer inerface.

### cmpEq

#### Function Deb(ByVal ipReference As Variant) As cmpEQ
DefaultMember

#### Function IComparer_ExecCmp(ByRef ipHostItem As Variant) As Boolean
IpHostItem is compared against ipReference

#### Property Get IComparer_TypeName() As String
Returns the typename as a string.  Useful when debugging

### cmpNEQ
### cmpLT
### cmpLTEQ
### cmpMT
### cmpMTEQ

### Example code
```VBA
DIm myNubmers as SeqC
Set myNumbers = SeqC(-5,-4,-3,-2,-1,0,1,2,3,4,-10)
Dim myNegatives as seqC
set myNegatives = myNumbers.FilterIt(cmpLT(0))
debug.print vba.Join(myNegatives.ToArray,chars.Comma)
-5,-4,-3,-2,-1,-10
```
### Class Filer
Provides file related actions using the Scripting.FileSystemObject

#### Function GetFileAsArrayOfStrings(ByVal ipFilePath As String, Optional ByRef ipSplitAtToken As String = vbCrLf) As Variant

#### Function GetFileAsString(ByVal ipPath As String) As String

