# VBALib:  Advent Of Code Library in VBA

## ##DANGER WILL ROBINSON DANGER##
I'm not a professional programmer.  Code in this Library represents my best efforts following several decades of using VBA on an ad-hoc basis to assist with massaging technical reports written in Word.
More recently I became interested in solving Advent of Code problems and used VBA as this was the language with which I was most familiar.

Hoever I quickly became frustrated by the immense amount of boiler plate needed in VBA to make useful things happen.  Eventually, and after reading and rereading the Rubberduck blogs, I began to get a better grip on Oop and saw patterns in the boilerplate code I was creating that were amenable encapsulation in classes.
And thus started the treadmill because reducing one pattern to a class highlighted andther set of patterns and so on.

This library represent my current ' State of the art' for pure VBA.  I have a parallel project in twinBasic which is in the refactoring doldrums at present whilst I try to eliminate the use of variants.

Run the tests from TestAll.  Rubberduck unit testing is now working but is very slow (for the current 700+ tests it takes about 12 minutes).  Recent updates to code have slowed the manual tests quite a bit from around 20 seconds to nearly 180 !! But I also notice quite a lot of variability in the manual timings.

I'm currently using the MIT license **BUT** if you find yourslf using any of the library code for commercial reasons please make a contribution proportional to the utility you gained.

__VBALib requires the installation of vbWatchdog.  The library has been tested with the community version.__
### A Note on naming
In the library I use the following as I have previously found them useful for indicating the role of a variable and to virtually eliminate the bane of clashing names.
#### Method parameter prefixes
- ip:  input only parameter
- op:  output only parameter, the input value is not used
- iop: input output parameter, the input parameter is mutated and provided back.
#### Prefix for variables local to Methods
- my
#### Prefixes at the Module or class level
I use two private User Defined Types to organise module level varibles
- State: for variables used internally by the class/module.
- Properties: for variables exposed via properties of the Class/Method.

State and properties are represented by the single letter variables s and p respectively

Please remember that I am a human being and consequently, as I don't have a built in compiler to enfore the above rules, there may be some usage that is not consistent with the above rules.

## Overview of library contents

The library offers
- a degree of enhanced reflection,  e.g. GroupInfo.IsNumber, GroupInfo.IsContainer, GroupInfo.IsList etc, Similarly for Arrays, ArrayOp.Ranks, ArrayOp.Count, ArrayOp.HoldsItems etc
- enhanced Collections (SeqA, SeqC, SeqL etc.)
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


### Class Filer
Provides file related actions specifically related to parsing Advent of Code  dataset.  Uses the Scripting.FileSystemObject

#### Function GetFileAsArrayOfStrings(ByVal ipFilePath As String, Optional ByRef ipSplitAtToken As String = vbCrLf) As Variant

#### Function GetFileAsString(ByVal ipPath As String) As String

## Class IterItems
IterItems is a class which has knowledge of the different types available (via class GroupInfo) and can threrefore enumerate according to the Type. This it is possible to use IterItems to Iterate an array, collection,dictionary, string or single item using the same code.
Currently Iteritems only accepts 1 dimensional arrays.  This just means I haven't yet got around to adding code that wil allow multidimensional arrays.

Iteritems supports the following operations
- Enumerate forward or backwards
- Move in any direction during the iteration loop using .MoveNext and .MovePrev
- Enumerate over any valid subrange using the FTS method to set the F)rom, T)o and S)tep parameters
- Iteritems is a class so multiple iterators can be used in parallel
- Test for if it is ok to movenext or moveprev
- Has methods for the current Item, Key()*, Or Offset as .curItem(0),.curKey(0).curOffset(0)
- Item.Key and Offset methods also support relative addreessing e.g. .Item(5) means the Item 5 places after the current item
- Can arbritrarily move to the first, Last, Start,End position in the enumeration loop.

*Key Key is used in the wider sense of a value used to define the location of a value in a container of items.  Thus for a Dictionary, Key will return the Key value, but for other container items it returns the Native index of the class.  e.g. If a single dimension array (Array(-10 to 5)) is passed to Iteritems then at the start of iteration
- .Offset(0) = 0 , .Offset(5) =5, and Offset(-5) = Null (cant go outside the range specified by Start End (equivalent to First/Last if FTS is not specified)
- .Key(0) = -10 ,Key(5) = -5, and Key(-5) = null  
After one .MoveNext
- .Offset(0) = 1, offset(5) =6, and offset(-1) = 0
- .Key =(0)=-9 , .Key(5) = -4,  and .Key(-1) = -10

**WARNING** I am current considering the naming of a number of methods of this class, and also whether or not to keep Deb as the default member in preference to CurItem.

### Function Deb(ByRef iopItems As Variant) As IterItems  (DefaultMember)
Deb is the factory method of the PredecalredId that returns a new instance of Iteritems.  The name 'Deb' is an in joke refering to Debutate (in the Old fashined sense the concept of being presented to society as available for marriage)
Deb (and all other Iteritems methods that allow it) return the instance of me so that the Iteritems class can be used as a fluid interface.  Thus Iteritem can give the impression of a 'True' constructor as Deb does not need to be specified as Deb **Requires** a parameter.
```VBA
Dim myItems as Iteritems: Set myItems = Iteritems.Deb(myVar)
' but, as deb is the default member the above can be written as
Dim myItems as Iteritems: Set myItems = Iteritems(myVar)
```
The internal pointer for the current index is set the the first position of myVar.  To enumerate in reverse use
```VBA
Dim myItems as Iteritems: Set myItems = Iteritems(myVar).MoveToEndIndex
```
The generic template for a loop using iteritems is
```VBA
dim myItems as Iteritems: set myItems = IterItems(Variable)
do
  <code>
loop while myItems.Movenext ' movenext/moveprev return True is the move was successful
```
### Function MoveNext() As Boolean
Moves the internal current item pointer to the next available slot.  Returns True if the move was successful, False if Not.
### Function MovePrev() As Boolean
Moves the internal current item pointer to the previous available slot.  Returns True if the move was successful, False if Not.
### Function HasNext(Optional ByRef ipLocalOffset As Long = 0) As Boolean
Returns True if the internal  current item pointer can be moved to the specified location
### Function HasNoNext(Optional ByRef ipLocalOffset As Long = 0) As Boolean
Returns True if the internal current item pointer cannot be moved to the specified location
###Function HasPrev(Optional ByRef ipLocalOffset As Long = 0) As Boolean
Returns True if the internal  current item pointer can be moved to the specified location
### Function HasNoPrev(Optional ByRef ipLocalOffset As Long = 0)
Returns True if the internal current item pointer cannot be moved to the specified location
### MoveToFirst() As IterItems
Moves the internal pointer to the First location of the host container object.  Resets StartIndex and EndIndex to FirstIndex and LastIndex respectively.
### Function MoveToLast() As IterItems
Moves the internal pointer to the lastLocation of the host container object.  Resets StartIndex and EndIndex to FirstIndex and LastIndex respectively.
### MoveToEndIndex
Moves the internal pointer to the start location of the host container object as sppecified by the startindex parameter for the FTS method
### Function MoveToEndIndex() As IterItems
Moves the internal pointer to the end location of the host container object as sppecified by the end index parameter for the FTS method
### Function FTS(Optional ByRef ipStartIndex As Variant = Empty, Optional ByRef ipEndIndex As Variant = Empty, Optional ByRef ipStep As Variant = Empty) As IterItems
Allows a start and end position to de defined for the enumeration.  Also allows the number of positions to move in response to a moveNext or MovePrev call.  
By defulat, the startindex, endIndex and Step are set to FirstIndex,LastIndex and 1 respectively.
FTS places a restriction on the values that can be used for First and Last, is that they must represent positions using 1 based indexing.  Iteritems does the translation between 1 based indexing and the Native indexing.  Using 1 based indexing allows positions to be specified irrspectiv of the indexing of he host container.  Currently, for dictionaries, the native indexing is based on the position in the Arrays returned by the Keys and Items methds.  For Dictionaries, a future update will allow start and end positions to be specified as actual keys.
### Property Get CurItem(ByVal ipLocalOffset As Long) As Variant
### Property Let CurItem(ByVal ipLocalOffset As Long, ByVal ipItem As Variant)
### Property Set CurItem(ByVal ipLocalOffset As Long, ByVal ipItem As Variant)
Returns/Sets the value at the position pointed to by the internal current item pointer. 
This currently doesn't work for arrays because I haven't yet has the time to port the code that facilitates arrays byreference.  
An offset relative to the current position may be specified. If the offset results in a location outside the start to End location it is ignored (for Let/Set) or return the value of 'Null' (for Get).
### Property Get CurKey(ByVal ipLocalOffset As Long) As Variant
Returns the item used to locate a value in a container.  For Dictionaries this is the Key value.  For other containers it returns the native index of the host container.
### Property Get CurOffset(ByVal ipLocalOffset As Long) As Variant
Return the offset of the current position from the first location of the host container.
### Function Size() As Long
Returns the number of elemnts in the host container
### Function HoldsItems() As Boolean
True if the host container is populated with items
### Function LacksItems() As Boolean
True if the host container has a count of zero or is an unallocated Array
###How it works
Iteritems depends on the class GroupInfo to classify Items in terms of a group to which they are assigned (see Id method) or a property that they have (e.g. IsDictionary, IsItemObject, IsNumber, IsContainer).
Iteritems manages the mapping of th native indexing of container objects to a 1 based indexing system. 
For dictionaries, and other objects that do not allow a numeric indexed access, the following is used
- Stack, Queue (from mscorlin) native indexing of the arrays produced by the .ToArray Method
- Dictionaries - the native indexing of the arrays produced by the .Items and .Keys methods
**Iteritems does not support keys for collection objects.**

