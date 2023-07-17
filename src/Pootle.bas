Attribute VB_Name = "Pootle"
Option Explicit
'@IgnoreModule


Sub TestSeqLAdd()
    Dim myS As SeqL
    Set myS = SeqL.Deb
    Debug.Print myS.Add(42)
    Debug.Print myS.Add(43)
    Debug.Print myS.Add(44)
End Sub
