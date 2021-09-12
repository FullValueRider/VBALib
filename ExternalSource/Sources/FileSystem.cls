VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Helpers")
Option Explicit
'@PredeclaredId
'@Exposed


Public Function GetFileAsStringThenSplitToArrayOfStrings(ByVal ipFilePath As String, Optional ByRef ipSplitAtToken As String = vbCrLf) As Variant

    Dim myFso  As Scripting.FileSystemObject
    Set myFso = New Scripting.FileSystemObject
    
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile(ipFilePath, Scripting.IOMode.ForReading)
        
    GetFileAsStringThenSplitToArrayOfStrings = VBA.Split(myfile.ReadAll, ipSplitAtToken)
    
End Function


Public Function GetFileAsString(ByVal ipPath As String) As String

    Dim myFso  As Scripting.FileSystemObject: Set myFso = New Scripting.FileSystemObject
    Dim myfile As TextStream
    Set myfile = myFso.OpenTextFile(ipPath, Scripting.IOMode.ForReading)
    
    Dim myFileAsString As String
    myFileAsString = myfile.ReadAll

    myfile.Close
    GetFileAsString = myFileAsString
    
End Function
