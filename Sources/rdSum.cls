VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "rdSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
 Option Explicit
 '@PredeclaredId
 '@Exposed
 '@Folder("VBALib.Filters")
Implements IReduce

Private Type State

    Comparer                        As Variant 'IComparer


End Type

Private s                           As State


Public Function Deb(Optional ByVal ipComparer As IComparer = Nothing) As rdSum

    With New rdSum

        Set Deb = .ReadyToUseInstance(ipComparer)

    End With

End Function

Friend Function ReadyToUseInstance(ByVal ipComparer As IComparer) As rdSum

    If ipComparer Is Nothing Then
    
        s.Comparer = Empty ' nothing causes a problem later on
        
    Else
    
        Set s.Comparer = ipComparer
        
    End If
    
    Set ReadyToUseInstance = Me

End Function

'@Description("Adds the value in ipArgsArray(0) to each item in the Lyst.  If the host item is not a number the string NAN is returned")
Private Function IReduce_ExecRed(ByVal ipHostItems As Lyst) As Variant
Attribute IReduce_ExecRed.VB_Description = "Adds the value in ipArgsArray(0) to each item in the Lyst.  If the host item is not a number the string NAN is returned"
    IReduce_ExecRed = ExecRed(ipHostItems)
End Function

Public Function ExecRed(ByVal ipHostItems As Lyst) As Variant

    Dim mySum As LongLong
    mySum = 0
    Dim myItem As Variant
    
    If VBA.IsEmpty(s.Comparer) Then
    
        For Each myItem In ipHostItems
    
            mySum = mySum + myItem
         
        Next
           
    Else
        
        For Each myItem In ipHostItems
           
           If s.Comparer.ExecCmp(myItem) Then mySum = mySum + myItem
           
        Next
             
    End If

    ExecRed = mySum

End Function
