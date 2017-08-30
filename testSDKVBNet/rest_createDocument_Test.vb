Imports click2mailVBNET

Module rest_createDocument_Test
    Private WithEvents r As New click2mailVBNET.Restc2mAPI("username", "password", Restc2mAPI.liveMode.Stage)
    Sub Main()
        Console.Write("DocumentID IS:" & r.createDocumentSimple("C:\c2m\test.pdf")) 'This will show full output
        Console.ReadLine()
    End Sub
End Module
