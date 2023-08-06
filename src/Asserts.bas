Attribute VB_Name = "Asserts"
'@IgnoreModule
'@Folder("Tests")
Option Explicit

        
Public Sub AssertExactAreEqual(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
#If twinbasic Then
    Assert.Strict.AreEqual ipExpected, ipResult, ipWhere
#Else
    If VBATesting Then
        Dim myExpected As String: myExpected = Fmt.Text("{0}: {1}", VBA.TypeName(ipExpected), ipExpected)
        Dim myResult As String: myResult = Fmt.Text("{0}: {1}", VBA.TypeName(ipResult), ipResult)
        
        If myExpected <> myResult Then
        Debug.Print
            Fmt.Dbg "{0}: Exact AreEqual assertion failed: {nl}{1}{nl}{2}", ipWhere, myExpected, myResult
        End If
    Else
        Assert.AreEqual ipExpected, ipResult, ipWhere
    End If
#End If

End Sub


Public Sub AssertExactAreNotEqual(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
#If twinbasic Then
    Assert.Strict.AreNotEqual ipExpected, ipResult, ipWhere
#Else
    If VBATesting Then
        Dim myExpected As String: myExpected = Fmt.Text("{0}: {1}", VBA.TypeName(ipExpected), ipExpected)
        Dim myResult As String: myResult = Fmt.Text("{0}: {1}", VBA.TypeName(ipResult), ipResult)
        
        
        If myExpected = myResult Then
            Debug.Print
            Fmt.Dbg "{0}: Exact AreEqual assertion failed: {nt}{1}{nt}{2}", ipWhere, myExpected, myResult
        End If
    Else
        Assert.AreNotEqual ipExpected, ipResult, ipWhere
    End If
#End If

End Sub



Public Sub AssertStrictSequenceEquals(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
#If twinbasic Then
    Assert.Strict.SequenceEquals ipExpected, ipResult, ipWhere
#Else
    If VBATesting Then
        Dim myExpected As String: myExpected = Fmt.Text("{0}", Array(ipExpected))
        Dim myResult As String: myResult = Fmt.Text("{0}", Array(ipResult))
        
        If myExpected <> myResult Then
            Debug.Print
            Fmt.Dbg "{0}: Strict SequenceEquals assertion failed: {nt}{1}{nt}{2}", ipWhere, myExpected, myResult
        End If
    Else
        Assert.SequenceEquals ipExpected, ipResult, ipWhere
    End If
#End If

End Sub

Public Sub AssertStrictSequenceNotEquals(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
#If twinbasic Then
    Assert.Strict.SequenceNotEquals ipExpected, ipResult, ipWhere
#Else
    If VBATesting Then
        Dim myExpected As String: myExpected = Fmt.Text("{0}", ipExpected)
        Dim myResult As String: myResult = Fmt.Text("{0}", ipResult)
        
        If myExpected = myResult Then
            Debug.Print
            Fmt.Dbg "SequenceEquals assertion failed: {0}{nt}{1}{nt}{2}", ipWhere, myExpected, myResult
        End If
    Else
        Assert.SequenceEquals ipExpected, ipResult, ipWhere
    End If
#End If

End Sub

Public Sub AssertExactSequenceEquals(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
#If twinbasic Then
    Assert.exact.SequenceEquals ipExpected, ipResult, ipWhere
#Else
    If VBATesting Then
    
        Dim myEItems As IterItems: Set myEItems = IterItems(ipExpected)
        Dim myRItems As IterItems: Set myRItems = IterItems(ipResult)
        
        If myEItems.LacksItems Or myRItems.LacksItems Then
            Fmt.Dbg "{0}: SequenceEquals assertion failed - no items: ", ipWhere
            Exit Sub
        End If
        
        Dim myE As Variant: ReDim myE(myEItems.StartIndex To myEItems.EndIndex)
        Dim myR As Variant: ReDim myR(myRItems.StartIndex To myRItems.EndIndex)
        
        
        Do
        
            Dim myIndex As Long: myIndex = myEItems.CurKey(0)
            
            myE(myIndex) = Fmt.Text("({0}): {1}: {2}", myIndex, VBA.TypeName(myEItems.CurItem(0)), myEItems.CurItem(0))
            myR(myIndex) = Fmt.Text("({0}): {1}: {2}", myIndex, VBA.TypeName(myRItems.CurItem(0)), myRItems.CurItem(0))
            
            
            If myE(myIndex) <> myR(myIndex) Then
                myE(myIndex) = "** " & myE(myIndex) & " **"
                myR(myIndex) = "** " & myR(myIndex) & " **"
            End If
            
        Loop While myEItems.MoveNext And myRItems.MoveNext
        
                
        Dim myExpected As String: myExpected = VBA.Join(myE, ",")
        Dim myResult As String: myResult = VBA.Join(myR, ",")
        
        If myExpected <> myResult Then
            Debug.Print
            Fmt.Dbg "{0}: SequenceEquals assertion failed: {0}{nt}{1}{nt}{2}", ipWhere, myExpected, myResult
        End If
    Else
        Assert.SequenceEquals ipExpected, ipResult, ipWhere
    End If
#End If

End Sub


'Public Sub AssertPermissiveSequenceEquals(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
'
'    #If twinbasic Then
'        Assert.Permissive.SequenceEquals ipExpected, ipResult, ipWhere
'    #Else
'        Assert.SequenceEquals ipExpected, ipResult, ipWhere
'    #End If
'
'End Sub
'
'
'Public Sub AssertPermissiveAreEqual(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
'
'    #If twinbasic Then
'        Assert.Permissive.AreEqual ipExpected, ipResult, ipWhere
'    #Else
'        Assert.AreEqual ipExpected, ipResult, ipWhere
'    #End If
'
'End Sub


'Public Sub AssertExactAreEqual(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
'
'    #If twinbasic Then
'        Assert.Exact.AreEqual ipExpected, ipResult, ipWhere
'    #Else
'        Assert.AreEqual ipExpected, ipResult, ipWhere
'    #End If
'
'End Sub


Public Sub AssertExactAreSame(ByRef ipExpected As Variant, ipResult As Variant, ipWhere As String)
    
    #If twinbasic Then
        Assert.Strict.AreSame ipExpected, ipResult, ipWhere
    #Else
        Assert.AreSame ipExpected, ipResult, ipWhere
    #End If
    
End Sub

Public Sub AssertFail(ByRef ipComponent As String, ipProcedure As String, ByRef ipMessage As String)
    Fmt.Dbg "{0}{nl}{1}{nl}{2}", ipComponent, ipProcedure, ipMessage
End Sub
