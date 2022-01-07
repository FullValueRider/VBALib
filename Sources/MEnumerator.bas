Attribute VB_Name = "MEnumerator"
' Modified by cxw from code by http://www.vbforums.com/member.php?255623-DEXWERX
' posted at http://www.vbforums.com/showthread.php?854963-VB6-IEnumVARIANT-For-Each-support-without-a-typelib&p=5229095&viewfull=1#post5229095
' License: "Use it how you see fit." - http://www.vbforums.com/showthread.php?854963-VB6-IEnumVARIANT-For-Each-support-without-a-typelib&p=5232689&viewfull=1#post5232689
' Explanation at https://stackoverflow.com/a/52261687/2877364

'
' MEnumerator.bas
'
' Implementation of IEnumVARIANT to support For Each in VB6
' CHanged some long values to LongPtr whereerrors were indicated
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Private Type TENUMERATOR
    VTablePtr   As LongPtr
    References  As Long
    Enumerable  As IValueProvider
    Index       As Long
End Type

Private Enum API
    NULL_ = 0
    S_OK = 0
    S_FALSE = 1
    E_NOTIMPL = &H80004001
    E_NOINTERFACE = &H80004002
    E_POINTER = &H80004003
#If False Then
    Dim NULL_, S_OK, S_FALSE, E_NOTIMPL, E_NOINTERFACE, E_POINTER
#End If
End Enum

Private Declare PtrSafe Function FncPtr Lib "msvbvm60" Alias "VarPtr" (ByVal Address As LongPtr) As Long
Private Declare PtrSafe Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
Private Declare PtrSafe Function CopyBytesZero Lib "msvbvm60" Alias "__vbaCopyBytesZero" (ByVal Length As Long, Dst As Any, Src As Any) As Long
Private Declare PtrSafe Function CoTaskMemAlloc Lib "ole32" (ByVal cb As Long) As Long
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByVal lpiid As LongPtr) As Long
Private Declare PtrSafe Function SysAllocStringByteLen Lib "oleaut32" (ByVal psz As Long, ByVal cblen As Long) As Long
Private Declare PtrSafe Function VariantCopyToPtr Lib "oleaut32" Alias "VariantCopy" (ByVal pvargDest As Long, ByRef pvargSrc As Variant) As Long
Private Declare PtrSafe Function InterlockedIncrement Lib "kernel32" (ByRef Addend As Long) As Long
Private Declare PtrSafe Function InterlockedDecrement Lib "kernel32" (ByRef Addend As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function NewEnumerator(ByRef Enumerable As IValueProvider) As IEnumVARIANT
' Class Factory
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Static VTable(6) As LongLong
    If VTable(0) = NULL_ Then
        ' Setup the COM object's virtual table
        VTable(0) = FncPtr(AddressOf IUnknown_QueryInterface)
        VTable(1) = FncPtr(AddressOf IUnknown_AddRef)
        VTable(2) = FncPtr(AddressOf IUnknown_Release)
        VTable(3) = FncPtr(AddressOf IEnumVARIANT_Next)
        VTable(4) = FncPtr(AddressOf IEnumVARIANT_Skip)
        VTable(5) = FncPtr(AddressOf IEnumVARIANT_Reset)
        VTable(6) = FncPtr(AddressOf IEnumVARIANT_Clone)
    End If

    Dim this As TENUMERATOR
    With this
        ' Setup the COM object
        .VTablePtr = VarPtr(VTable(0))
        .References = 1
        Set .Enumerable = Enumerable
    End With

    ' Allocate a spot for it on the heap
    Dim pThis As Long
    pThis = CoTaskMemAlloc(LenB(this))
    If pThis Then
        ' CopyBytesZero is used to zero out the original
        ' .Enumerable reference, so that VB doesn't mess up the
        ' reference count, and free our enumerator out from under us
        CopyBytesZero LenB(this), ByVal pThis, this
        DeRef(VarPtr(NewEnumerator)) = pThis
    End If
End Function

Private Function RefToIID$(ByVal riid As Long)
    ' copies an IID referenced into a binary string
    Const IID_CB As Long = 16&  ' GUID/IID size in bytes
    DeRef(VarPtr(RefToIID)) = SysAllocStringByteLen(riid, IID_CB)
End Function

Private Function StrToIID$(ByRef iid As String)
    ' converts a string to an IID
    StrToIID = RefToIID$(NULL_)
    IIDFromString StrPtr(iid), StrPtr(StrToIID)
End Function

Private Function IID_IUnknown() As String
    Static iid As String
    If StrPtr(iid) = NULL_ Then _
        iid = StrToIID$("{00000000-0000-0000-C000-000000000046}")
    IID_IUnknown = iid
End Function

Private Function IID_IEnumVARIANT() As String
    Static iid As String
    If StrPtr(iid) = NULL_ Then _
        iid = StrToIID$("{00020404-0000-0000-C000-000000000046}")
    IID_IEnumVARIANT = iid
End Function

Private Function IUnknown_QueryInterface(ByRef this As TENUMERATOR, _
                                         ByVal riid As Long, _
                                         ByVal ppvObject As Long _
                                         ) As Long
    If ppvObject = NULL_ Then
        IUnknown_QueryInterface = E_POINTER
        Exit Function
    End If

    Select Case RefToIID$(riid)
        Case IID_IUnknown, IID_IEnumVARIANT
            DeRef(ppvObject) = VarPtr(this)
            IUnknown_AddRef this
            IUnknown_QueryInterface = S_OK
        Case Else
            IUnknown_QueryInterface = E_NOINTERFACE
    End Select
End Function

Private Function IUnknown_AddRef(ByRef this As TENUMERATOR) As Long
    IUnknown_AddRef = InterlockedIncrement(this.References)
End Function

Private Function IUnknown_Release(ByRef this As TENUMERATOR) As Long
    IUnknown_Release = InterlockedDecrement(this.References)
    If IUnknown_Release = 0& Then
        Set this.Enumerable = Nothing
        CoTaskMemFree VarPtr(this)
    End If
End Function

Private Function IEnumVARIANT_Next(ByRef this As TENUMERATOR, _
                                   ByVal celt As Long, _
                                   ByVal rgVar As Long, _
                                   ByRef pceltFetched As Long _
                                   ) As Long

    Const VARIANT_CB As Long = 16 ' VARIANT size in bytes

    If rgVar = NULL_ Then
        IEnumVARIANT_Next = E_POINTER
        Exit Function
    End If

    Dim Fetched As Long
    Fetched = 0
    Dim element As Variant

    With this
        Do While this.Enumerable.HasMore
            element = .Enumerable.GetNext
            VariantCopyToPtr rgVar, element
            Fetched = Fetched + 1&
            If Fetched = celt Then Exit Do
            rgVar = PtrAdd(rgVar, VARIANT_CB)
        Loop
    End With

    If VarPtr(pceltFetched) Then pceltFetched = Fetched
    If Fetched < celt Then IEnumVARIANT_Next = S_FALSE
End Function

Private Function IEnumVARIANT_Skip(ByRef this As TENUMERATOR, ByVal celt As Long) As Long
    IEnumVARIANT_Skip = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Reset(ByRef this As TENUMERATOR) As Long
    IEnumVARIANT_Reset = E_NOTIMPL
End Function

Private Function IEnumVARIANT_Clone(ByRef this As TENUMERATOR, ByVal ppEnum As Long) As Long
    IEnumVARIANT_Clone = E_NOTIMPL
End Function

Private Function PtrAdd(ByVal Pointer As Long, ByVal Offset As Long) As Long
    Const SIGN_BIT As Long = &H80000000
    PtrAdd = (Pointer Xor SIGN_BIT) + Offset Xor SIGN_BIT
End Function

Private Property Let DeRef(ByVal Address As LongPtr, ByVal Value As LongPtr)
    GetMem4 Value, ByVal Address
End Property
