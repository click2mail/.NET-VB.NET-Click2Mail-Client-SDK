Imports click2mailVBNET

Module rest_SimpleSend_Test
    Private WithEvents r As New click2mailVBNET.Restc2mAPI("username", "password", Restc2mAPI.liveMode.Stage)
    Sub Main()
        r.addressList.Clear()
        Dim x As click2mailVBNET.Restc2mAPI.addressItem
        x = New click2mailVBNET.Restc2mAPI.addressItem("John", "Smith", "TestOrg", "1234 Test Street", "Ste 335", "", "Oak Brook", "IL", "60523", "")
        r.addressList.Add(x)
        x = New click2mailVBNET.Restc2mAPI.addressItem("John", "Smith2", "TestOrg", "1234 Test Street", "Ste 335", "", "Oak Brook", "IL", "60523", "")
        r.addressList.Add(x)
        r.runComplete("C:\c2m\test.pdf", r.createXMLFromAddressList(), "Letter 8.5 x 11", "Address on Separate Page", "Next Day", "#10 Double Window", "Black and White", "White 24#", "Printing Both sides")
        Console.ReadLine()
    End Sub
    Private Sub statuschanged(ByVal message As String) Handles r.statusChanged
        Console.WriteLine(message)
    End Sub
    Private Sub jobStatusCheck(ByVal id As String, status As String, description As String) Handles r.jobStatusCheck
        Console.WriteLine("jobId is:" & id)
        Console.WriteLine("job Status is:" & status)
        Console.WriteLine("job Description is:" & description)
    End Sub
End Module
