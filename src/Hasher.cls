VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hasher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
'@Exposed
'@Folder("Helpers")
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========
' Based on cHashD from https://www.vbforums.com/showthread.php?834515-Simple-and-fast-lightweight-HashList-Class-(no-APIs)&s=f049d882b4fc4d27f4bc2f9bc383a53e
' cHashD, Olaf Schmidt in August 2016 (enhanced about mixed, variable KeyTypes and Remove-functions in 2020)
'========1=========2=========3=========4=========5=========6=========7=========8=========9=========A=========B=========


Public Enum e_HashSize
    m_First = 0
    m_PreviousOrDefault = m_First
    m_Min
    m_Max
    m_Default
    m_last = m_Default
End Enum


Option Explicit

Const HMul                                  As Long = 3727      '  Hash multiplikator' as per olaf schmidt
Const MIN_HASH_TABLE_SIZE                   As Long = 128
Const MAX_HASH_TABLE_SIZE                   As Long = 524288    '  2^19
Const DEFAULT_HASH_TABLE_SIZE               As Long = 16348
'@Ignore ConstantNotUsed
Const USE_DEFINED_SIZE_OR_DEFAULT           As Long = -1

Private Type State
    ' The hashing algorith of Olaf Schmidt uses either a stream of 16 bit
    ' character values, or the data region of a variant.
    ' the data areas of the string or variant are pointed to by usuroung the
    ' data area pointer of W() using the Saw UDT.

    W()                                     As Integer
    sAW                                     As SAFEARRAY1D
    ' HashBitMask is a value of 2^x-1 (e.g. righthand bits are all 1)
    ' It is initialised at 16 and then increased by repeatedly multiplying by 2
    ' until it is greater than ipCapacity
    ' subtracting 1 then sets all the RH bits to 1
    ' This value is then used to AND intermediate calculations
    'during hashing to ensure we don't exceed the maximum allows hashvalue
    MaxHash                                 As Long
    sizeMask                                As Long
    CompareMode                             As Long
    Size                                    As Long
End Type

Private s                                   As State

'@DefaultMember
Public Function Deb(Optional ByVal ipHashTableSize As Variant) As Hasher
Attribute Deb.VB_UserMemId = 0
    With New Hasher
        Set Deb = .ConstructInstance(ipHashTableSize)
    End With
End Function

'@Description("HashRanges (0 to x) are set using the x^2 bigger or equal to ipHashTableSize. The default is 16384 (0 to (2^14)-1)")
Friend Function ConstructInstance(ByRef ipHashTableSize As Variant) As Hasher
Attribute ConstructInstance.VB_Description = "HashRanges (0 to x) are set using the x^2 bigger or equal to ipHashTableSize. The default is 16384 (0 to (2^14)-1)"

    Select Case True
    
        '@Ignore IsMissingOnInappropriateArgument
        Case VBA.IsMissing(ipHashTableSize)
        
            If s.Size = 0 Then
                ipHashTableSize = DEFAULT_HASH_TABLE_SIZE
            Else
                ipHashTableSize = s.Size
            End If
            
        Case GroupInfo.IsNotNumber(ipHashTableSize)
        
            Err.Raise 17 + vbObjectError, _
                "VBALib.Hasher", _
                Fmt.Text("The hash table size must be a number between {0} and {1}. Default is {2}", MIN_HASH_TABLE_SIZE, MAX_HASH_TABLE_SIZE, DEFAULT_HASH_TABLE_SIZE)
                
     End Select
     
    Reinit ipHashTableSize
    Set ConstructInstance = Me
    
End Function

Public Sub Reinit(ByVal ipHashTableSize As Long)

    ' ipHashtablesize will have an initial value >=128
    ' the maximum hash value could be an arbitrary value but in practice
    ' the most efficient use of the hash space is for the maximum hash value
    ' to be a power of 2 minus 1
    ' thus sizemask is initially calculated as a power of 2
    ' and hence its binary representation will be 1 followed by zeros.
    ' when the power of 2 size is fixed we substract 1 to give a value
    ' whose binary representation is 0 followed by all 1's
    ' this mask is used to AND the value in the hash calculation to ensure that
    ' the hash value never exceeds the maximum hash value
    
    'find the first power of 2 that is bigger or equal to ipHashTableSize
    s.sizeMask = 128
    Do Until s.sizeMask >= ipHashTableSize
        s.sizeMask = s.sizeMask * 2
    Loop
    
    If s.sizeMask > MAX_HASH_TABLE_SIZE Then
        s.sizeMask = MAX_HASH_TABLE_SIZE
    End If
    
    ' The size mask is 1 larger than it needs to  be
    ' so we subtract 1 to give a binary representation of all 1's
    s.Size = s.sizeMask
    
    s.MaxHash = s.sizeMask - 1
    s.sizeMask = s.sizeMask - 1
    
    ' setup for using W as an integer array
    s.sAW.cDims = 1
    s.sAW.cbElements = 2
    s.sAW.cLocks = 1
    s.sAW.fFeatures = &H11 'FADF_AUTO=&H1 || FADF_FIXEDSIZE=&H10
    BindArray s.W, VarPtr(s.sAW)
 
    If LWC(97) = 0 Then
        InitLWC
    End If
    
End Sub

Public Function MaxHash() As Long
    MaxHash = s.MaxHash
End Function


Public Function Size() As Long
    Size = s.Size
End Function
'@Description("Hash alogorith based on the algorith in cHAshD by Olaf Schmidt")
Public Function OSHash(ByRef ipHashItem As Variant) As Long  'return -1, when no Key can be found
Attribute OSHash.VB_Description = "Hash alogorith based on the algorith in cHAshD by Olaf Schmidt"

    OSHash = -1  ' a hash of -1 means we couldn't calculate a hash value
    Dim i As Long
    Dim D As Double
    Dim L As LongPtr
    Dim F As Single
    Dim VT As VbVarType
     
    Dim C As Variant: C = VBA.CDec(C)
    
    'Dim s.sizemask As Long: s.sizemask = s.sizemask

    Dim H As Long: H = s.sizeMask 'init the HashValue (all bits to 1)
    
    Dim myHashItem As Variant
    If GroupInfo.IsContainer(ipHashItem) Or GroupInfo.IsItemObject(ipHashItem) Or GroupInfo.IsAdmin(ipHashItem) Then
        myHashItem = Fmt.Text("{0}", ipHashItem)
    Else
        myHashItem = ipHashItem
    End If
  
    VT = VarType(myHashItem)
    Select Case VT
        Case vbString
            
            s.sAW.cElements1D = VBA.Len(myHashItem)
            s.sAW.pvData = StrPtr(myHashItem)
            
            If s.CompareMode = 0 Then
                myHashItem = VBA.UCase$(myHashItem)
            End If
            
            For i = 0 To s.sAW.cElements1D - 1
                H = (H + s.W(i)) * HMul And s.sizeMask
            Next
         
' the case of vb Object is now captured by GroupInfo.ItemObject
'       Case vbObject
'            L = ObjPtr(Key): sAW.cElements1D = 2: sAW.pvData = VarPtr(L)
'           H = (H + W(0)) * HMul And s.sizemask 'loop-unrolling (we have only 2 16Bit integers in the array)
'           H = (H + W(1)) * HMul And s.sizemask
         
       Case vbCurrency
            C = myHashItem
            s.sAW.cElements1D = 4
            s.sAW.pvData = VarPtr(C)
            H = (H + s.W(0)) * HMul And s.sizeMask 'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And s.sizeMask
            H = (H + s.W(2)) * HMul And s.sizeMask
            H = (H + s.W(3)) * HMul And s.sizeMask
         
       Case vbLong, vbInteger, vbByte
            L = myHashItem
            s.sAW.cElements1D = 2
            s.sAW.pvData = VarPtr(L)
            H = (H + s.W(0)) * HMul And s.sizeMask 'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And s.sizeMask
         
    
       Case vbDouble
            D = myHashItem
            s.sAW.cElements1D = 4
            s.sAW.pvData = VarPtr(D)
            H = (H + s.W(0)) * HMul And s.sizeMask 'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And s.sizeMask
            H = (H + s.W(2)) * HMul And s.sizeMask
            H = (H + s.W(3)) * HMul And s.sizeMask
               
  
       Case vbDate
            D = myHashItem
            s.sAW.cElements1D = 4
            s.sAW.pvData = VarPtr(D)
            H = (H + s.W(0)) * HMul And s.sizeMask 'loop-unrolling (we have only 4 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And s.sizeMask
            H = (H + s.W(2)) * HMul And s.sizeMask
            H = (H + s.W(3)) * HMul And s.sizeMask
            
       Case vbSingle
            F = myHashItem:
            s.sAW.cElements1D = 2:
            s.sAW.pvData = VarPtr(F)
            H = (H + s.W(0)) * HMul And s.sizeMask 'loop-unrolling (we have only 2 16Bit integers in the array)
            H = (H + s.W(1)) * HMul And s.sizeMask
            
     End Select
  
     OSHash = H
    
End Function
