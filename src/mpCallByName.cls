VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "mpCallByName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Implements call by name on the object provided by ipHostItem"
'@PredeclaredId
'@Exposed
'Folder("Mappers")
'@ModuleDescription("Implements call by name on the object provided by ipHostItem")
    Option Explicit

 Implements IMapper
 
 Private Const MY_TYPENAME           As String = "mpCallByName"
 
 Private Type State
         
     MethodName                       As String
     CallType                         As VbCallType
         
 End Type

 Private s                               As State

 
Public Function IMapper_ExecMapper(ByRef ipHostItem As Variant) As Variant
    
    Dim myresult() As Variant
    If TryCallByName(ipHostItem, s.MethodName, s.CallType, myresult) Then
        IMapper_ExecMapper = myresult
    Else
        IMapper_ExecMapper = Array("Call by name failed")
    End If

End Function
 
 Public Function TryCallByName(ByRef ipHostItem As Variant, ByVal ipMethodName As String, ByVal ipcalltype As VbCallType, ByRef opResult() As Variant) As Boolean
     
     On Error Resume Next
     opResult = Array(VBA.CallByName(ipHostItem, ipMethodName, ipcalltype))
     
     On Error GoTo 0
     TryCallByName = Err.Number = 0
     
 End Function

 '@DefaultMember
 Public Function Deb(ByVal ipMethodName As String, ByVal ipcalltype As VbCallType) As mpCallByName
Attribute Deb.VB_UserMemId = 0
     With New mpCallByName
         Set Deb = .ConstructInstance(ipMethodName, ipcalltype)
     End With
 End Function


 Friend Function ConstructInstance(ByVal ipMethodName As String, ByVal ipcalltype As VbCallType) As mpCallByName
   
    s.MethodName = ipMethodName
    s.CallType = ipcalltype
    Set ConstructInstance = Me
    
 End Function


Private Property Get IMapper_TypeName() As String
     IMapper_TypeName = MY_TYPENAME
 End Property
