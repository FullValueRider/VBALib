VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Maths"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
Option Explicit

Public Function IsEven(ByRef ipItem As Variant) As Boolean
    IsEven = (ipItem Mod 2) = 0
End Function

Public Function IsOdd(ByRef ipItem As Variant) As Boolean
    IsOdd = Not IsEven(ipItem)
End Function

Public Function Pi() As Double
    Pi = 4 * VBA.Atn(1)
End Function

'@Description("Returns 2 * Pi")
Public Function Tau() As Double
Attribute Tau.VB_Description = "Returns 2 * Pi"
    Tau = 2 * Pi
End Function

'@Description(" a factor to convert a value in radians to a value in Degrees")
Public Function RadiansToDegrees() As Double
Attribute RadiansToDegrees.VB_Description = " a factor to convert a value in radians to a value in Degrees"
    RadiansToDegrees = 180 / Pi
End Function

Public Function ATan2(ByVal ipY As Double, ByVal ipX As Double) As Double
        
    Dim myTheta As Double

    If (Abs(ipX) < 0.0000001) Then
        
        If (Abs(ipY) < 0.0000001) Then
            
            myTheta = 0#
                
        ElseIf (ipY > 0#) Then
            
            myTheta = 1.5707963267949
        Else
            
            myTheta = -1.5707963267949
                
        End If
            
    Else
        
        myTheta = VBA.Atn(ipY / ipX)
    
        If (ipX < 0) Then
            
            If (ipY >= 0#) Then
                
                myTheta = 3.14159265358979 + myTheta
                    
            Else
                
                myTheta = myTheta - 3.14159265358979
                    
            End If
                
        End If
            
    End If
        
    ATan2 = myTheta
        
End Function

Public Function Ceiling(ByVal ipX As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Ceiling = (Int(ipX / Factor) - (ipX / Factor - Int(ipX / Factor) > 0)) * Factor
End Function

Public Function Floor(ByVal ipX As Double, Optional ByVal Factor As Double = 1) As Double
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Floor = Int(ipX / Factor) * Factor
End Function

Public Function Factorial(ByVal ipNumber As Variant) As Variant

    Dim myFact As Variant
    myFact = CDec(ipNumber)
    Dim myCounter As Variant
    myCounter = CDec(ipNumber - 1)
    Do
        
        myFact = myFact * myCounter
        myCounter = myCounter - 1
        
    Loop Until myCounter = 1
        
    Factorial = myFact
        
End Function

Public Function NextPrime(ByVal ipNumber As Long) As Long

    Dim myNumber As Long
    myNumber = ipNumber
    Do
        myNumber = myNumber + 1
    Loop Until IsPrime(myNumber)
        
    NextPrime = myNumber
        
End Function

Public Function IsPrime(ByVal ipNumber As LongLong) As Boolean
        
    If ipNumber = 1 Then
        
        IsPrime = False
        Exit Function
            
    ElseIf ipNumber < 4 Then
        
        IsPrime = True
        Exit Function
            
    ElseIf ipNumber Mod 2 = 0 Then
        
        IsPrime = False
        Exit Function
            
    ElseIf ipNumber < 9 Then
        
        IsPrime = True                           'we have already excluded 4,6 and 8.
        Exit Function
            
    ElseIf ipNumber Mod 3 = 0 Then
        
        IsPrime = False
        Exit Function
            
    Else
        
        Dim myR As Double
        myR = Floor(Sqr(ipNumber))               ' n rounded to the greatest integer r so that r*r<=n
            
        Dim myF As Double
        myF = 5
            
        Do While myF <= myR
            
            If ipNumber Mod myF = 0 Then
                
                IsPrime = False                  'return false (and step out of the function)
                Exit Function
                    
            End If
                
            If ipNumber Mod (myF + 2) = 0 Then
                
                IsPrime = False
                Exit Function                    'return false (and step out of the function)
                    
            End If
            
            myF = myF + 6
        Loop

        IsPrime = True
            
    End If
        
End Function
'@Description("Ln. Log to base e to provide a more distictive function than of VBA which used Log for log to base e")
Public Function Ln(ByVal ipX As Double) As Double
Attribute Ln.VB_Description = "Ln. Log to base e to provide a more distictive function than of VBA which used Log for log to base e"
Ln = VBA.Log(ipX)
End Function

'@Description("Log to base 10.  nb native VBA uses Log to give Log to base e.  See Maths.Ln")
Public Function Log(ByVal ipX As Double) As Double
Attribute Log.VB_Description = "Log to base 10.  nb native VBA uses Log to give Log to base e.  See Maths.Ln"
    Log = Log(ipX) / Log(10#)
End Function

' Public Function TriangularNumber(ByVal ipLast As LongLong, Optional ByVal ipfirst As LongLong = 1) As LongLong
Public Function TriangularNumber(ByVal ipLast As LongLong) As LongLong
    TriangularNumber = ipLast * (ipLast + 1) / 2
End Function

Public Function Max(ByVal ipReference As Variant, ByVal ipTest As Variant) As Variant
    Max = IIf(ipTest > ipReference, ipTest, ipReference)
End Function

Public Function Min(ByVal ipReference As Variant, ByVal ipTest As Variant) As Variant
    Min = IIf(ipTest < ipReference, ipTest, ipReference)
End Function

Public Function GetRndLong() As Long
    GetRndLong = VBA.Int((VBA.Rnd - 0.5) * 2 ^ 32)
End Function


' Generate permutations of the input seq using Heap's Algorithm
' https://www.mrexcel.com/board/threads/permutations-heaps-algorithm.959919/
' https://iq.opengenus.org/heaps-algorithm-for-generating-permutations/
Public Sub Permutations(ByRef ipSeq As SeqC, ByRef ipSwapIndex As Long, ByRef opPermutations As SeqC)
    
        If VBA.IsObject(ipSeq.First) Then
            Err.Raise 17 + vbObjectError, _
                "VBALib.Maths.Permutations", _
                "Cannot generate permutations of objects"
        End If
        
        If ipSwapIndex = 1 Then
        
            opPermutations.Add ipSeq.Clone
            Exit Sub
            
        Else
            
            Dim i As Long
            For i = ipSeq.FirstIndex To ipSwapIndex
                Permutations ipSeq, ipSwapIndex - 1, opPermutations
    
                ' if SwapIndex is odd, swap 0th i.e (first) and
                ' (SwapIndex-1)th i.e (last) element
                'Dim temp As Variant
                If Maths.IsEven(ipSwapIndex) Then
                
                    ipSeq.Swap i, ipSwapIndex
                    
                ' If SwapIndex is even, swap ith
                ' and (SwapIndex-1)th i.e last element
                Else
                
                    ipSeq.Swap 1, ipSwapIndex
                  
                End If
                
            Next
            
       End If
      
    End Sub
