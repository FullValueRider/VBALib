Attribute VB_Name = "modHashD"
Option Explicit
 
Public Type SAFEARRAY1D
  cDims As Integer
  fFeatures As Integer
  cbElements As Long
  cLocks As Long
  pvData As LongPtr
  cElements1D As Long
  lLbound1D As Long
End Type

 
Public Declare PtrSafe Sub BindArray Lib "kernel32" Alias "RtlMoveMemory" (PArr() As Any, pSrc As LongPtr, Optional ByVal CB As Long = LenB(myPtrLen))
Public Declare PtrSafe Function VariantCopy Lib "oleaut32" (Dst As Any, Src As Any) As Long
Public Declare PtrSafe Function VariantCopyInd Lib "oleaut32" (Dst As Any, Src As Any) As Long
Private Declare PtrSafe Function CharLowerBuffW Lib "user32" (lpsz As Any, ByVal cchLength As Long) As Long

' we need to define a variable as VBA doesn't support lenB(of Type)
Private myPtrLen As LongPtr
Public LWC(-32768 To 32767) As Integer


Public Sub InitLWC()
  Dim i As Long
  For i = -32768 To 32767: LWC(i) = i: Next 'init the Lookup-Array to the full WChar-range
  CharLowerBuffW LWC(-32768), 65536 '<-- and convert its whole content to LowerCase-WChars
End Sub


