VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' *****************************************************************************
' *  VCOM library: vbWatchdog  http://www.everythingaccess.com/vbwatchdog.htm *
' *  3.9.2 (Ultimate Edition TRI built May 26 2021                            *
' *---------------------------------------------------------------------------*
' *  This is a demonstration version of the library for TESTING purposes only.*
' *                                                                           *
' *  This software is provided "as-is," without any express or implied        *
' *  warranty.  In no event shall the author be held liable for any           *
' *  consequential, incidental, direct, indirect, special, punitive, or other *
' *  damages arising from the use of this software (including without         *
' *  limitation, damages for loss of business profits, business interruption, *
' *  loss of information or other pecuniary loss).                            *
' *****************************************************************************

Option Explicit

Private Const OPTION_BASE                  As Long = 0
Private Const OPTION_FLAGS                 As Long = 2
Private Const OPTION_INCLUDE_REFERENCEDOCS As Long = 0
Private Const OPTION_DISABLEDCLASSES       As String = ""
Private Const PAGE_EXECUTE_RW              As Long = &H40
Private Const MEM_RESERVE_AND_COMMIT       As Long = &H3000
Private Const MEM_RELEASE                  As Long = &H8000
Private Const ERR_OUT_OF_MEMORY            As Long = &H7
Private Const ROOTOBJECT_SIZE              As Long = &H4ED4C

Private m_Loader As VCOMInitializerStruct
Private m_VCOMObject As Object

#If VBA7 = False Then
    Private Declare Function VirtualAlloc Lib "kernel32" (ByVal Address As Long, ByVal Size As Long, ByVal AllocationType As Long, ByVal Protect As Long) As Long
    Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal ProcName As String) As Long
    Private Declare Function GetProcAddress Lib "kernel32" (ByVal Module As Long, ByVal ProcName As String) As Long
    Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal Size As Long, ByVal dwFreeType As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Size As Long)
    Private Const VBA_VERSION              As Long = 6

    Private Type VCOMInitializerStruct
        vtbl_QueryInterface As Long
        vtbl_AddRef As Long
        vtbl_Release As Long
        vtbl_GetTypeInfoCount As Long
        vtbl_GetTypeInfo As Long
        vtbl_GetIDsOfNames As Long
        vtbl_Invoke As Long
        RootObjectMem As Long
        HelperObject As Object
        SysFreeString As Long
        WideCharToMultiByte As Long
        GetProcAddress As Long
        NativeCode As String
        LoaderMem As Long
        IgnoreFlag As Boolean
        VTablePtr As Long
        Kernel32Handle As Long
        RootObject As Object
        ClassFactory As Object
    End Type
#Else
    Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal Address As LongPtr, ByVal Size As LongPtr, ByVal AllocationType As Long, ByVal Protect As Long) As LongPtr
    Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" (ByVal ProcName As String) As LongPtr
    Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal Module As LongPtr, ByVal ProcName As String) As LongPtr
    Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal Size As LongPtr, ByVal dwFreeType As Long) As Long
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Dest As Any, ByRef Source As Any, ByVal Size As LongPtr)
    Private Const VBA_VERSION              As Long = 7

    Private Type VCOMInitializerStruct
        vtbl_QueryInterface As LongPtr
        vtbl_AddRef As LongPtr
        vtbl_Release As LongPtr
        vtbl_GetTypeInfoCount As LongPtr
        vtbl_GetTypeInfo As LongPtr
        vtbl_GetIDsOfNames As LongPtr
        vtbl_Invoke As LongPtr
        RootObjectMem As LongPtr
        HelperObject As Object
        SysFreeString As LongPtr
        WideCharToMultiByte As LongPtr
        GetProcAddress As LongPtr
        NativeCode As String
        LoaderMem As LongPtr
        IgnoreFlag As Boolean
        VTablePtr As LongPtr
        Kernel32Handle As LongPtr
        RootObject As Object
        ClassFactory As Object
    End Type
#End If

' ----------------------------------------------
' Virtual-COM initialization routine:
' ----------------------------------------------
Private Sub Class_Initialize()

    With m_Loader
        .NativeCode = "%EEEE%::::PPPPPPPPPH+D$ XXXtNXXXXXXVSPPPPj PPPPPPPP4T)D$04P)D$,4'4 )D$($ PZ3D$@+D$ YQ3H +L$ XP3Q +T$0XPf55ntvf)B|+T$0+T$0+T$0R[YQ^VXP2CP<0tF1D$$kD$$@!L$$2CQ1D$$kD$$@!L$$2CR1D$$kD$$@!L$$2CS+\$,3BP1BP1rP3rP+T$(  XXXXXXXXXXXXX[^tJAYAZQ4tPPPPH)D$@4pH)D$84'4 H)D$0$ PH+L$ H3AtH+D$ L3PtL+T$HXPf55{L+T$HL+T$HtqfA)B8ARA[YQXPA2CD<0tR1D$0kD$0@L!L$0A2CE1D$0kD$0@L!L$0A2CF1D$0kD$0@L!L$0A2CGL+\$@A3BDA1BDA1JDA3JDL+T$8  XXXXXYXXXXqBLHOJA@n[??n[=ezoieZZprkhs^ljbZljbZ=bNZ_Q_>HirF[Q^Z[IrzRM wGDDoeTtKTfdGVduCVduCGhiCGhygGhygCmzXGcH[D_J^DV VfF VX<TI@<_veu]flqomliCuelQxpdudatE@hrwIkzSMzvOizw_Mzw_MssLJssLZBCLZ@A]^@A]^TNa^oFmn^nIv@aSsbT?WeWnSg_DCgKjKWCgHe[wJGe;?@fj;Ifyr@cfMAmTN_rNKNzxilIhMnADMgDV@cm;<jihu?aE=]rdY\puMUpgDuAa;UqSWBSPSUG=LUFNNESSOPGVYEbGXQWROj__GHKjOj_MIHKj^x?IRh=XVh=XVKH<VYKlJWLbAEtOIg@nIDT^HJVOD[KGudwGDEeFT[reTWJ@\ht>a;r>cruLna<Mniy?eKL_]zy?\pznXpznXANNXIL_\IL_\xSc\iMIUzQIdEoomgyo=XAyzJCDBXN>=QKmvHmtvO]HXO]J\O]J\m]hV?]mXmQvgl=tdpaS RUqPBV \PRocNMQflywB>;gFluaO?jKF@UIO ai_vUJ[apwFqeFGfACZVu>[0"

        .LoaderMem = VirtualAlloc(0, Len(.NativeCode), MEM_RESERVE_AND_COMMIT, PAGE_EXECUTE_RW)
        If .LoaderMem = 0 Then Err.Raise ERR_OUT_OF_MEMORY

        .RootObjectMem = VirtualAlloc(0, ROOTOBJECT_SIZE, MEM_RESERVE_AND_COMMIT, PAGE_EXECUTE_RW)
        If .RootObjectMem = 0 Then Err.Raise ERR_OUT_OF_MEMORY

        .vtbl_QueryInterface = .LoaderMem
        .VTablePtr = VarPtr(m_Loader)
        .Kernel32Handle = GetModuleHandleA("KERNEL32")
        .GetProcAddress = GetProcAddress(.Kernel32Handle, "GetProcAddress")
        .SysFreeString = GetProcAddress(GetModuleHandleA("OLEAUT32"), "SysFreeString")
        .WideCharToMultiByte = GetProcAddress(GetModuleHandleA("KERNEL32"), "WideCharToMultiByte")
        Set .HelperObject = New ErrEx_Helper
        Call CopyMemory(ByVal .LoaderMem, ByVal .NativeCode, Len(.NativeCode))
        Call CopyMemory(.RootObject, VarPtr(.VTablePtr), LenB(.VTablePtr))
        .IgnoreFlag = TypeOf .RootObject Is VBA.Collection
        Set .ClassFactory = (.RootObject)
        Set .RootObject = Nothing
        VirtualFree .LoaderMem, 0, MEM_RELEASE
        Call .ClassFactory.Init(.Kernel32Handle, .GetProcAddress, OPTION_BASE + OPTION_FLAGS, VBA_VERSION, .HelperObject)
        Set m_VCOMObject = .ClassFactory.GetErrEx()
    End With

End Sub

Sub Catch(ByVal ErrorNumber1 As Long, Optional ByVal ErrorNumber2 As Long, Optional ByVal ErrorNumber3 As Long, Optional ByVal ErrorNumber4 As Long)
    Call m_VCOMObject.Catch(ErrorNumber1, ErrorNumber2, ErrorNumber3, ErrorNumber4)
End Sub

Sub Finally(Optional ByVal Ignore1 As Long = &HFFFFFFFF, Optional ByVal Ignore2 As Long = &HFFFFFFFF, Optional ByVal Ignore3 As Long = &HFFFFFFFF, Optional ByVal Ignore4 As Long = &HFFFFFFFF)
    Call m_VCOMObject.Finally(Ignore1, Ignore2, Ignore3, Ignore4)
End Sub

Sub CatchAll(Optional ByVal Ignore1 As Long = &HFFFFFFFE, Optional ByVal Ignore2 As Long = &HFFFFFFFE, Optional ByVal Ignore3 As Long = &HFFFFFFFE, Optional ByVal Ignore4 As Long = &HFFFFFFFE)
    Call m_VCOMObject.CatchAll(Ignore1, Ignore2, Ignore3, Ignore4)
End Sub

Sub DoFinally()
    Call m_VCOMObject.DoFinally
End Sub

Property Get IsEnabled() As Boolean
    IsEnabled = m_VCOMObject.IsEnabled
End Property

Property Get SourceProject() As String
    SourceProject = m_VCOMObject.SourceProject
End Property

Property Get SourceProjectFilename() As String
    SourceProjectFilename = m_VCOMObject.SourceProjectFilename
End Property

Property Get SourceModule() As String
    SourceModule = m_VCOMObject.SourceModule
End Property

Property Get SourceProcedure() As String
    SourceProcedure = m_VCOMObject.SourceProcedure
End Property

Property Get SourceProcedureIndex() As Long
    SourceProcedureIndex = m_VCOMObject.SourceProcedureIndex
End Property

Property Get SourceLineNumber() As Long
    SourceLineNumber = m_VCOMObject.SourceLineNumber
End Property

Property Get SourceLineCode() As String
    SourceLineCode = m_VCOMObject.SourceLineCode
End Property

Property Get SourceVBEProject() As Object
    Set SourceVBEProject = m_VCOMObject.SourceVBEProject
End Property

Property Get SourceProjectIsCompiled() As Boolean
    SourceProjectIsCompiled = m_VCOMObject.SourceProjectIsCompiled
End Property

Property Get SourceProjectIsSaved() As Boolean
    SourceProjectIsSaved = m_VCOMObject.SourceProjectIsSaved
End Property

Property Get SourceProjectConditionalCompilationArgs() As String
    SourceProjectConditionalCompilationArgs = m_VCOMObject.SourceProjectConditionalCompilationArgs
End Property

Property Get Callstack() As ErrExCallstack
    Set Callstack = m_VCOMObject.AssignObj(New ErrExCallstack, m_VCOMObject.Callstack)
End Property

Property Get VariablesInspector() As ErrExVariables
    Set VariablesInspector = m_VCOMObject.AssignObj(New ErrExVariables, m_VCOMObject.VariablesInspector)
End Property

Property Get DialogOptions() As ErrExDialogOptions
    Set DialogOptions = m_VCOMObject.AssignObj(New ErrExDialogOptions, m_VCOMObject.DialogOptions)
End Property

Property Get VariablesDialogOptions() As ErrExDialogOptions
    Set VariablesDialogOptions = m_VCOMObject.AssignObj(New ErrExDialogOptions, m_VCOMObject.VariablesDialogOptions)
End Property

Property Get State() As OnErrorStatus
    State = m_VCOMObject.State
End Property

Property Let State(ByVal Value As OnErrorStatus)
    m_VCOMObject.State = Value
End Property

Property Get StateAsStr() As String
    StateAsStr = m_VCOMObject.StateAsStr
End Property

Property Get Description() As String
    Description = m_VCOMObject.Description
End Property

Property Let Description(ByVal Value As String)
    m_VCOMObject.Description = Value
End Property

Property Get HelpFile() As String
    HelpFile = m_VCOMObject.HelpFile
End Property

Property Let HelpFile(ByVal Value As String)
    m_VCOMObject.HelpFile = Value
End Property

Property Get Source() As String
    Source = m_VCOMObject.Source
End Property

Property Let Source(ByVal Value As String)
    m_VCOMObject.Source = Value
End Property

Property Get HelpContext() As Long
    HelpContext = m_VCOMObject.HelpContext
End Property

Property Let HelpContext(ByVal Value As Long)
    m_VCOMObject.HelpContext = Value
End Property

Property Get LastDLLError() As Long
    LastDLLError = m_VCOMObject.LastDLLError
End Property

Property Get Number() As Long
    Number = m_VCOMObject.Number
End Property

Property Let Number(ByVal Value As Long)
    m_VCOMObject.Number = Value
End Property

Property Get VerifyOnErrorProcName() As Boolean
    VerifyOnErrorProcName = m_VCOMObject.VerifyOnErrorProcName
End Property

Property Let VerifyOnErrorProcName(ByVal Value As Boolean)
    m_VCOMObject.VerifyOnErrorProcName = Value
End Property

Sub Enable(ByVal OnErrorProcName As String)
    Call m_VCOMObject.Enable(OnErrorProcName)
End Sub

Property Get EvalObject() As Object
    Set EvalObject = m_VCOMObject.EvalObject
End Property

Property Set EvalObject(ByVal EventsObject As Object)
    m_VCOMObject.EvalObject = EventsObject
End Property

Sub Disable(Optional ByVal Flags As Long)
    Call m_VCOMObject.Disable(Flags)
End Sub

Sub CallGlobalErrorHandler(Optional ByVal UserData As Variant)
    Call m_VCOMObject.CallGlobalErrorHandler(UserData)
End Sub

Function LiveCallstack() As ErrExCallstack
    Set LiveCallstack = m_VCOMObject.AssignObj(New ErrExCallstack, m_VCOMObject.LiveCallstack)
End Function

Property Get ProjectFilterList() As String
    ProjectFilterList = m_VCOMObject.ProjectFilterList
End Property

Property Let ProjectFilterList(ByVal Value As String)
    m_VCOMObject.ProjectFilterList = Value
End Property

Sub ShowHelp()
    Call m_VCOMObject.ShowHelp
End Sub

Function ShowErrorDialog() As OnErrorStatus
    ShowErrorDialog = m_VCOMObject.ShowErrorDialog
End Function

Property Get UserData() As Variant
    Call m_VCOMObject.AssignVar(UserData, m_VCOMObject.UserData)
End Property

Property Let UserData(ByVal Value As Variant)
    m_VCOMObject.UserData = Value
End Property

Property Get IsDebugable() As Boolean
    IsDebugable = m_VCOMObject.IsDebugable
End Property

Property Get VBEVersion() As String
    VBEVersion = m_VCOMObject.VBEVersion
End Property

Property Get Version() As String
    Version = m_VCOMObject.Version
End Property

Property Get VariablesInspectorEnabled() As Boolean
    VariablesInspectorEnabled = m_VCOMObject.VariablesInspectorEnabled
End Property

Property Let VariablesInspectorEnabled(ByVal Value As Boolean)
    m_VCOMObject.VariablesInspectorEnabled = Value
End Property

Property Get TypeInfoInspectorEnabled() As Boolean
    TypeInfoInspectorEnabled = m_VCOMObject.TypeInfoInspectorEnabled
End Property

Property Let TypeInfoInspectorEnabled(ByVal Value As Boolean)
    m_VCOMObject.TypeInfoInspectorEnabled = Value
End Property

Property Get CustomVars(ByVal Name As String) As String
    CustomVars = m_VCOMObject.CustomVars(Name)
End Property

Property Let CustomVars(ByVal Name As String, ByVal Value As String)
    m_VCOMObject.CustomVars(Name) = Value
End Property

Property Get PropagateUnhandledClassErrors() As Boolean
    PropagateUnhandledClassErrors = m_VCOMObject.PropagateUnhandledClassErrors
End Property

Property Let PropagateUnhandledClassErrors(ByVal Value As Boolean)
    m_VCOMObject.PropagateUnhandledClassErrors = Value
End Property

Property Get AlwaysAllowPropagation() As Boolean
    AlwaysAllowPropagation = m_VCOMObject.AlwaysAllowPropagation
End Property

Property Let AlwaysAllowPropagation(ByVal Value As Boolean)
    m_VCOMObject.AlwaysAllowPropagation = Value
End Property

Property Let DebugLogPath(ByVal Path As String)
    m_VCOMObject.DebugLogPath = Path
End Property

Sub ReThrow(Optional ByVal OverrideNumber As Long, Optional ByVal OverrideSource As String, Optional ByVal OverrideDescription As String)
    Call m_VCOMObject.ReThrow(OverrideNumber, OverrideSource, OverrideDescription)
End Sub

Sub Helper_SetClipboardText(ByVal Text As String)
    Call m_VCOMObject.Helper_SetClipboardText(Text)
End Sub

Function Helper_ChooseColor(Optional ByVal ParentWindowHandle As Long, Optional ByVal InitialColorRGB As Long) As Long
    Helper_ChooseColor = m_VCOMObject.Helper_ChooseColor(ParentWindowHandle, InitialColorRGB)
End Function

Function Helper_ChooseFile(Optional ByVal Filter As String) As String
    Helper_ChooseFile = m_VCOMObject.Helper_ChooseFile(Filter)
End Function

Sub Helper_ShellExecute(Optional ByVal WindowHandle As Long, Optional ByVal Operation As String, Optional ByVal File As String, Optional ByVal Params As String, Optional ByVal Directory As String, Optional ByVal ShowCmd As Long = &H1, Optional ByVal Wait As Boolean = True)
    Call m_VCOMObject.Helper_ShellExecute(WindowHandle, Operation, File, Params, Directory, ShowCmd, Wait)
End Sub

Property Get LoadedDLLs(Optional ByVal ShowFullPaths As Boolean, Optional ByVal ShowVersionInfo As Boolean = True) As String
    LoadedDLLs = m_VCOMObject.LoadedDLLs(ShowFullPaths, ShowVersionInfo)
End Property

Property Get SortVariables() As Boolean
    SortVariables = m_VCOMObject.SortVariables
End Property

Property Let SortVariables(ByVal Value As Boolean)
    m_VCOMObject.SortVariables = Value
End Property

Property Get IgnoreHandledErrors() As Boolean
    IgnoreHandledErrors = m_VCOMObject.IgnoreHandledErrors
End Property

Property Let IgnoreHandledErrors(ByVal Value As Boolean)
    m_VCOMObject.IgnoreHandledErrors = Value
End Property

Property Get HostAllowsBreakNow() As Boolean
    HostAllowsBreakNow = m_VCOMObject.HostAllowsBreakNow
End Property

Property Get DisableMemorySafetyChecks() As Boolean
    DisableMemorySafetyChecks = m_VCOMObject.DisableMemorySafetyChecks
End Property

Property Let DisableMemorySafetyChecks(ByVal Value As Boolean)
    m_VCOMObject.DisableMemorySafetyChecks = Value
End Property

