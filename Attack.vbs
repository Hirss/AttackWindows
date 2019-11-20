Sub DoMyProc50Times()
   Dim x
   For x = 1 To 100000000
     Set fso = WScript.CreateObject("Scripting.FileSystemObject")
     fso.CreateFolder("Espace sans echanges "&x)
   Next
End Sub
DoMyProc50Times()