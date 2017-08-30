Imports click2mailVBNET

Module rest_checkStatusOfJob
    Private WithEvents r As New click2mailVBNET.Restc2mAPI("username", "password", Restc2mAPI.liveMode.Stage)
    Sub Main()
        r.jobId = 12345
        Console.Write(r.checkJobStatus()) 'This will show full output
        Console.ReadLine()

    End Sub
    Private Sub jobStatusCheck(ByVal id As String, status As String, description As String) Handles r.jobStatusCheck
        Console.WriteLine("jobId is:" & id)
        Console.WriteLine("job Status is:" & status)
        Console.WriteLine("job Description is:" & description)
    End Sub
End Module
