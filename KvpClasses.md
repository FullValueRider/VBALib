# Kvp Classes

KvpX Classes (K)ey V)alue P)air) are Dictionaries with enhanced functionality.  The X denotes the underlying objects used to manage the Keys and Items. 
- KvpA:  The Keys and Items are held in two SeqA objects. Searching is linear. Items may be retieved by Key or Index
- KvpC:  The Keys and Items are held in two SeqC objects. Searching is linear. Items may be retieved by key or index
- KvpH:  The Keys and Items are held in a linked list node.  The list nodes are linked by order of addition/insertion.  The list nodes are also indexed in VBA array according to the Hash of the key. Searching for keys is done via the hash value of the Key. Searching for Items is linear via the linked list.  The Linked list nodes automatically manage the index in the linked list. Items may be retieved by Key or Index.
- KvpL:  The Keys and Items are held in two SeqL objects. Searching is linear. Items may be retieved by key or Index
  
To be implemented    
  
- KvpAL: The keys and Items are held in two ArrayList objects.  Searching is linear.  Items may be retieved by key or index
- KvpT:   The keys and items are held in a Key/Value treap.  Searching is quicker than linear.  Items may be retieved by Key or Index

Kvp classes are restricted to unique keys by default.  
The use of duplicate keys is permitted but is an opt in option which must be applied after the Kvp is created.
String keys use binary comparisons so "Hello" <> "hello".   
Arrays and container objects may be used as keys.  Comparisons are done using the strings produced by the Text method of the Fmt class.
Non contianer objects are compared based on the default member value or, if no suitable default member is available,  the objptr value of the instance.

## Constructors
Kvp classes use a PredeclaredId which allows a constructor method (Deb) to be used rather than the New KvpX syntax.  .  Constructor like behaviour is achieved by declaring the constructor method to be the default member of the class.  When not returning a specific value, Kvp objects return the instance of 'Me'.  This allows fluid interface use of Kvp classes.
```
Dim myK as KvpA
set myK = KvpA.Deb  '**Do not use New KvpA **

' fluid interface example

Set myK = KvpA.Deb.AddPairs(container1,container2).Mapit(mpInc(5)).FiltertIt(cmpLT(42))
```
