VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Filer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'@PredeclaredId
'@Exposed
Const ForReading                            As Integer = 1  'Scripting.IOMode.ForReading

Public Function GetFileAsArrayOfStrings(ByVal ipFilePath As String, Optional ByRef ipSplitAtToken As String = vbCrLf) As Variant

    Dim myFso  As Scripting.FileSystemObject
    Set myFso = New Scripting.FileSystemObject
        
    GetFileAsArrayOfStrings = VBA.Split(myFso.OpenTextFile(ipFilePath, ForReading).ReadAll, ipSplitAtToken)
   
        
End Function

Public Function GetFileAsString(ByVal ipPath As String) As String

    Dim myFso  As Scripting.FileSystemObject
    Set myFso = New Scripting.FileSystemObject
    GetFileAsString = myFso.OpenTextFile(ipPath, ForReading).ReadAll
    
End Function
