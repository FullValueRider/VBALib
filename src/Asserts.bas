Attribute VB_Name = "Asserts"
'@IgnoreModule
'@Folder("Tests")
Option Explicit

        
Public Sub AssertStrictAreEqual(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
#If twinbasic Then
    Assert.Strict.AreEqual ipExpected, ipResult, ipWhere
#Else
    Assert.AreEqual ipExpected, ipResult, ipWhere
#End If

End Sub


Public Sub AssertStrictAreNotEqual(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
    #If twinbasic Then
        Assert.Strict.AreNotEqual ipExpected, ipResult, ipWhere
    #Else
        Assert.AreNotEqual ipExpected, ipResult, ipWhere
    #End If
    
    End Sub


Public Sub AssertStrictSequenceEquals(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
#If twinbasic Then
     Assert.Strict.SequenceEquals ipExpected, ipResult, ipWhere
#Else
    Assert.SequenceEquals ipExpected, ipResult, ipWhere
#End If

End Sub


Public Sub AssertPermissiveSequenceEquals(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
    #If twinbasic Then
        Assert.Permissive.SequenceEquals ipExpected, ipResult, ipWhere
    #Else
        Assert.SequenceEquals ipExpected, ipResult, ipWhere
    #End If
    
End Sub


Public Sub AssertPermissiveAreEqual(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
    #If twinbasic Then
        Assert.Permissive.AreEqual ipExpected, ipResult, ipWhere
    #Else
        Assert.AreEqual ipExpected, ipResult, ipWhere
    #End If
    
End Sub


Public Sub AssertExactAreEqual(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
    #If twinbasic Then
        Assert.Exact.AreEqual ipExpected, ipResult, ipWhere
    #Else
        Assert.AreEqual ipExpected, ipResult, ipWhere
    #End If
    
End Sub


Public Sub AssertStrictAreSame(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
    #If twinbasic Then
        Assert.Strict.AreSame ipExpected, ipResult, ipWhere
    #Else
        Assert.AreSame ipExpected, ipResult, ipWhere
    #End If
    
End Sub

Public Sub AssertFail(ByRef ipComponent As String, ipProcedure As String, ByRef ipMessage As String)
    Fmt.Dbg "{0}, {1}, {2}", ipComponent, ipProcedure, ipMessage
End Sub
