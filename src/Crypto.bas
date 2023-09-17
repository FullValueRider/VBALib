Attribute VB_Name = "Crypto"
Option Explicit

'DefObj A-Z

#Const HasPtrSafe = (VBA7 <> 0) Or (TWINBASIC <> 0)
#Const HasOperators = (TWINBASIC <> 0)

#If HasPtrSafe Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
#End If

Private LNG_POW2(0 To 31)   As Long
Private s(0 To 15)          As Long
Private K(0 To 63)          As Long

#If Not HasOperators Then
Private Function ROTL32(ByVal lX As Long, ByVal lN As Long) As Long
    '--- ROTL32 = LShift(X, n) Or RShift(X, 32 - n)
    Debug.Assert lN <> 0
    ROTL32 = ((lX And (LNG_POW2(31 - lN) - 1)) * LNG_POW2(lN) Or -((lX And LNG_POW2(31 - lN)) <> 0) * LNG_POW2(31)) Or _
        ((lX And (LNG_POW2(31) Xor -1)) \ LNG_POW2(32 - lN) Or -(lX < 0) * LNG_POW2(lN - 1))
End Function

Private Function UAdd(ByVal lX As Long, ByVal lY As Long) As Long
    If (lX Xor lY) > 0 Then
        UAdd = ((lX Xor &H80000000) + lY) Xor &H80000000
    Else
        UAdd = lX + lY
    End If
End Function
#End If

#If HasOperators Then
[ IntegerOverflowChecks (False) ]
#End If
'@Ignore AssignedByValParameter
Public Sub CryptoMd5(ByRef baOutput() As Byte, ByRef baInput() As Byte, Optional ByVal Pos As Long, Optional ByVal Size As Long = -1)
    Dim vElem           As Variant
    Dim lIdx            As Long
    Dim lA              As Long
    Dim lB              As Long
    Dim lC              As Long
    Dim lD              As Long
    Dim lA2             As Long
    Dim lB2             As Long
    Dim lC2             As Long
    Dim lD2             As Long
    Dim lR              As Long
    Dim lE              As Long
    Dim lShift          As Long
    Dim lTemp           As Long
    Dim aBuffer()       As Long
    Dim lBufPos         As Long
    Dim lBufIdx         As Long
    
    If LNG_POW2(0) = 0 Then
        LNG_POW2(0) = 1
        For lIdx = 1 To 30
            LNG_POW2(lIdx) = LNG_POW2(lIdx - 1) * 2
        Next
        LNG_POW2(31) = &H80000000
        lIdx = 0
        For Each vElem In Split("7 12 17 22 5 9 14 20 4 11 16 23 6 10 15 21")
            s(lIdx) = vElem
            lIdx = lIdx + 1
        Next
        For lIdx = 0 To 63
            vElem = VBA.Abs(VBA.Sin(lIdx + 1)) * 4294967296#
            K(lIdx) = Int(IIf(vElem > 2147483648#, vElem - 4294967296#, vElem))
        Next
    End If
    If Size < 0 Then
        Size = UBound(baInput) + 1 - Pos
    End If
    '--- pad input buffer to 64 bytes
    lIdx = 64 - (Size Mod 64)
    If lIdx < 9 Then
        lIdx = lIdx + 64
    End If
    ReDim aBuffer(0 To (Size + lIdx) \ 4 - 1) As Long
    If Size > 0 Then
        CopyMemory aBuffer(0), baInput(Pos), Size
    End If
    CopyMemory ByVal VarPtr(aBuffer(0)) + Size, &H80, 1
    aBuffer(UBound(aBuffer) - 1) = Size * 8
    '--- md5 step
    lA = &H67452301: lB = &HEFCDAB89: lC = &H98BADCFE: lD = &H10325476
    Do While lBufPos < UBound(aBuffer)
        lA2 = lA: lB2 = lB: lC2 = lC: lD2 = lD
        For lIdx = 0 To 63
            lR = lIdx \ 16
            Select Case lR
            Case 0
                lE = (lB2 And lC2) Or (Not lB2 And lD2)
                lBufIdx = lIdx
            Case 1
                lE = (lB2 And lD2) Or (lC2 And Not lD2)
                lBufIdx = (lIdx * 5 + 1) And 15
            Case 2
                lE = lB2 Xor lC2 Xor lD2
                lBufIdx = (lIdx * 3 + 5) And 15
            Case 3
                lE = lC2 Xor (lB2 Or Not lD2)
                lBufIdx = (lIdx * 7) And 15
            End Select
            lShift = s((lR * 4) Or (lIdx And 3))
            #If HasOperators Then
                lE = lE + lA2 + K(lIdx) + aBuffer(lBufPos + lBufIdx)
            #Else
                lE = UAdd(UAdd(UAdd(lE, lA2), K(lIdx)), aBuffer(lBufPos + lBufIdx))
            #End If
            lTemp = lD2
            lD2 = lC2
            lC2 = lB2
            #If HasOperators Then
                lB2 = lB2 +(lE << lShift) Or (lE >> (32 - lShift))
            #Else
                lB2 = UAdd(lB2, ROTL32(lE, lShift))
            #End If
            lA2 = lTemp
        Next
        #If HasOperators Then
            lA = lA +lA2: lB += lB2: lC += lC2: lD += lD2
        #Else
            lA = UAdd(lA, lA2): lB = UAdd(lB, lB2): lC = UAdd(lC, lC2): lD = UAdd(lD, lD2)
        #End If
        lBufPos = lBufPos + 16
    Loop
    '--- complete output
    aBuffer(0) = lA: aBuffer(1) = lB: aBuffer(2) = lC: aBuffer(3) = lD
    ReDim baOutput(0 To 15) As Byte
    CopyMemory baOutput(0), aBuffer(0), 16
End Sub

Public Function CryptoMd5Text(ByVal sText As String) As String
    Const CP_UTF8       As Long = 65001
    Dim lSize           As Long
    Dim baInput()       As Byte
    Dim baHash()        As Byte
    Dim aRetVal(0 To 15) As String
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baInput(0 To lSize - 1) As Byte
        WideCharToMultiByte CP_UTF8, 0, StrPtr(sText), Len(sText), baInput(0), lSize, 0, 0
    Else
        baInput = vbNullString
    End If
    CryptoMd5 baHash, baInput
    For lSize = 0 To 15
        aRetVal(lSize) = Right$("0" & Hex$(baHash(lSize)), 2)
    Next
    CryptoMd5Text = LCase$(Join(aRetVal, vbNullString))
End Function


