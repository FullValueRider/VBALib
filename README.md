# VBALib:  Advent Of Code Library in VBA

## ##DANGER WILL ROBINSON DANGER##
I'm not a professional programmer.  Code in this Library represents my best efforts following several decades of using VBA on an ad-hoc basis to assist with massaging technical reports written in Word.
More recently I became interested in solving Advent of Code problems and used VBA as this was the language with which I was most familiar.

Hoever I quickly became frustrated by the immense amount of boiler plate needed in VBA to make useful things happen.  Eventually, and after reading and rereading the Rubberduck blogs, I began to get a better grip on Oop and saw patterns in the boilerplate code I was creating that were amenable encapsulation in classes.
And thus started the treadmill because reducing one pattern to a class highlighted andther set of patterns and so on.

This library represent my current ' State of the art' for pure VBA.  I have a parallel project in twinBasic which is in the refactoring doldrums at present whilst I try to eliminate the use of variants.

Run the tests from TestAll.  Do not use Rubberduck unit testing. There are currently over 400 tests and Rubberduck just grinds to a halt.

## Overview of library contents

The library offers
- a degree of enhanced reflection,  e.g. GroupInfo.IsNumber, GroupInfo.IsContainer, GroupInfo.IsList etc, Similarly for Arrays, ArrayInfo.Ranks, ArrayInfo.Count, ArrayInfo.HoldsItems etc
- enhanced Collections (SeqA, SeqC, SeqL with SeqAL, SeqH planned.)
- an en enhanced DIctionaries (KvpA,KvpC, KvpL with KvpAL and KvpH planned)
- string interpolation for formatting strings (formatting characters and positional variables)
- Map, Reduce,Filter,and Count functionality on seq and Kvp classes and a set of functions to use in such operations
- IterItems class for enhanced 'for each' functionality. Enumerate any type, move forwards or backwards at will, methods for current item, current offset from first inde and current index, Specify from to step increments, forward the iteritem =instance as a parameter.
- Maths functions not in VBA
- Useful global constants such Maximim and Minimum valiues for some integer types

## ToDo insert some example code here highlighting the library functionality

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
All classes in this library are created with PredeclaredIds using the Rubberduck annotations.  This is to enforce the use of the class name qualifier but also to minimise name clashes.

### Class ArrayInfo

#### Function Holdsitemsd(ByRef ipArray As Variant) As Boolean
Returns True if an array contains elements.  Returns False otherwise. An Array() will return false

#### Function LacksItesd(ByRef ipArray As Variant) As Boolean

#### Function IsArray(ByRef ipArray As Variant, Optional ByRef ipArrayType As e_ArrayType = e_ArrayType.m_AnyArrayType) As Boolean
Returns true if ipArray is an array and is allocated and is of the specified ArrayType.  The default actions is for an array type of Any. An Array() will return false

#### Function IsNotArray(ByRef ipArray As Variant, Optional ByRef ipArrayType As e_ArrayType = m_AnyArrayType) As Boolean

#### Function Count(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 0) As Long
Default is to return the total number of elements in an array.  If a rank is specified it returns the Ubound - Lbound+ 1 of the rank.
If the array LacksItems returns -1

#### Function Ranks(ByVal ipArray As Variant) As Long
Returns the number of dimensions as specified in a Dim or Redim statement.
If the array has no ranks returns -1

#### Function HasRank(ByRef ipArray As Variant, ByVal ipRank As Long) As Boolean
Returns true if the array dimensions include the specified rank

#### Public Function FirstIndex(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 1) As long
Alternative to Lbound. Returns 'Empty' if ipRank is not in the array

#### Function LastIndex(ByRef ipArray As Variant, Optional ByVal ipRank As Long = 1) As Variant
ALternative to Ubound.  Returns 'Empty' if ipRank is not in the array

## Comparer classes
Comparer classes are used to provide a compare function for the Seq and Kvp FilterIt, ReduceIt and CountIt classes
A class is provided for the comparisions of equals, not equals, more than, more than or equal, less than and less than or equals.
The classes are primed with a reference value at creation or default to a reference of zero.  The cmp classes may rely on a CompereHelper function for their comparision.
cmp classes implement the IComparer interface.

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
Provides file related actions specifically related to parsing Advent of Code  dataset.  Uses the Scripting.FileSystemObject

#### Function GetFileAsArrayOfStrings(ByVal ipFilePath As String, Optional ByRef ipSplitAtToken As String = vbCrLf) As Variant

#### Function GetFileAsString(ByVal ipPath As String) As String

TO BE CONTINUED
