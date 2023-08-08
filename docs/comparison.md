## Comparison
Comparison of two entities can be done as usual in VBA.
In addition, VBALib offers a set of comparers which use more restrictive conditions, but enable the contents of Containers to be compared.

The Comparer module provides
- EQ (formarly Equals)
- NEQ
- LT
- LTEQ
- MT
- MTEQ

These methods are called by the equivalent comparer functors
- cmpEQ ->   Comparer.Eq
- cmpNEQ ->  Comparer.NEQ
- cmpLT ->   Comparer.LT
- cmpLTEQ -> Comparer.LTEQ
- cmpMT ->   Comparer.MT
- cmpMTEQ->  Comparer.MTEQ

The rules used by the Comparer methods are
1. Items must be of the same group otherwise the comparison is false (i.e. "5"=5 = false (string vs number)  
   The GroupInfo class is used to ensure that the items being compared belong to the same group
   - Boolean
   - Number
   - String
   - ItemObject (Objects that are not contained in the admin or container groups)
   - Admin (empty, null, nothing, error)
   - Array
   - List
   - Dictionary
2. Strings, Arrays, Lists and Dictionaries are first compared on the basis of size.  
   -Array(1,2,3) is lessthan Array(1,1,1,1): Comparison of content is done only if Items are the same size.  
   -Array (1,2,3) is greataer than Array(1,1,1000):  Second item 2 is greater than 1)
4. String comparisons use VBA rules based on Option Compare  
5. Booleans and Admin types can only be compared for EQ and NEQ (but True>False would be straightforward to implement if found desirable)  

### Comparer Functors (Is there a better word than Functor?)

Comparer functors are defined as Predecalred Id's and have the factory method (Deb) set as the default method. 
Comparer functors are provided as parameters to the CountIt, FilterIt and ReduceIt methods of the Seq, Kvp and ArrayOp classes
The general template for use is  
```
  Seq.FilterIt(cmpNEQ(42))
```
The above example returns a new seq with items that are not equal to 42.

### Example code
```VBA
Dim myNumbers as SeqC
Set myNumbers = SeqC(-5,-4,-3,-2,-1,0,1,2,3,4,-10)
Dim myNegatives as SeqC
set myNegatives = myNumbers.FilterIt(cmpLT(0))
Fmt.Dbg "{0}", myNegatives
Output is {-5,-4,-3,-2,-1,-10}

'or

Dim myNegatives as Variant
myNegatives = myNumbers.FilterIt(cmpNEQ(42)).ToArray
Fmt.Dbg "{0}", myNegatives
Output is [-5,-4,-3,-2,-1,-10]

```
