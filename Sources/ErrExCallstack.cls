VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrExCallstack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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

Enum CallstackProcKind
    CALLSTACK_PROCKIND_SUB = &H1
    CALLSTACK_PROCKIND_FUNCTION = &H2
    CALLSTACK_PROCKIND_GET = &H3
    CALLSTACK_PROCKIND_LET = &H4
    CALLSTACK_PROCKIND_SET = &H5
End Enum

Enum OnErrorStatus
    OnErrorGoto0 = &H1
    OnErrorResumeNext = &H2
    OnErrorGotoLabel = &H3
    OnErrorEnd = &H4
    OnErrorDebug = &H5
    CalledByLocalHandler = &H6
    OnErrorRetry = &H7
    OnErrorPropagate = &H8
    OnErrorExitProcedure = &H9
    OnErrorCatch = &HA
    OnErrorCatchAll = &HB
    OnErrorInsideCatch = &HC
    OnErrorInsideCatchAll = &HD
    OnErrorInsideFinally = &HE
    OnErrorPropagateCatch = &HF
    OnErrorPropagateCatchAll = &H10
    OnErrorUnwind = &H11
    OnErrorUnwindNoFinally = &H12
End Enum

Public VCOMObject As Object      ' DO NOT USE!!! THIS IS FOR INTERNAL USE ONLY!!!
Attribute VCOMObject.VB_VarUserMemId = -4

Property Get ProjectName() As String
    ProjectName = VCOMObject.ProjectName
End Property

Property Get ProjectFilename() As String
    ProjectFilename = VCOMObject.ProjectFilename
End Property

Property Get ModuleName() As String
    ModuleName = VCOMObject.ModuleName
End Property

Property Get ProcedureName() As String
    ProcedureName = VCOMObject.ProcedureName
End Property

Property Get ProcedureIndex() As Long
    ProcedureIndex = VCOMObject.ProcedureIndex
End Property

Property Get LineNumber() As Long
    LineNumber = VCOMObject.LineNumber
End Property

Property Get SourceLineSub() As Long
    SourceLineSub = VCOMObject.SourceLineSub
End Property

Property Get LineCode() As String
    LineCode = VCOMObject.LineCode
End Property

Property Get VBEProject() As Object
    Set VBEProject = VCOMObject.VBEProject
End Property

Property Get ProjectIsCompiled() As Boolean
    ProjectIsCompiled = VCOMObject.ProjectIsCompiled
End Property

Property Get ProjectIsSaved() As Boolean
    ProjectIsSaved = VCOMObject.ProjectIsSaved
End Property

Property Get ProjectConditionalCompilationArgs() As String
    ProjectConditionalCompilationArgs = VCOMObject.ProjectConditionalCompilationArgs
End Property

Sub FirstLevel()
    Call VCOMObject.FirstLevel
End Sub

Function NextLevel() As Boolean
    NextLevel = VCOMObject.NextLevel
End Function

Property Get VariablesInspector() As ErrExVariables
    Set VariablesInspector = VCOMObject.AssignObj(New ErrExVariables, VCOMObject.VariablesInspector)
End Property

Property Get HasActiveErrorHandler() As Boolean
    HasActiveErrorHandler = VCOMObject.HasActiveErrorHandler
End Property

Property Get HasResumePoint() As Boolean
    HasResumePoint = VCOMObject.HasResumePoint
End Property

Property Get ProcedureKind() As CallstackProcKind
    ProcedureKind = VCOMObject.ProcedureKind
End Property

Function ProjectCreateClassInstance(ByVal ClassName As String) As Object
    Set ProjectCreateClassInstance = VCOMObject.ProjectCreateClassInstance(ClassName)
End Function

