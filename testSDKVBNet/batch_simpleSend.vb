Imports click2mailVBNET
Module batch_simpleSend_Test
    Private WithEvents r As New click2mailVBNET.Batchc2mAPI("username", "password", Restc2mAPI.liveMode.Stage)
    Sub Main()
        Dim batchJobs As New List(Of click2mailVBNET.Batchc2mAPI.batchJob)
        Dim addList As New List(Of click2mailVBNET.Batchc2mAPI.addressItem)
        Dim job As click2mailVBNET.Batchc2mAPI.batchJob

        Dim address1 As New click2mailVBNET.Batchc2mAPI.addressItem("John1", "Smith", "MyOrg", "1234 Test Street", "Ste 335", "Oak Brook", "IL", "60523", "United States")
        Dim address2 As New click2mailVBNET.Batchc2mAPI.addressItem("John2", "Smith", "MyOrg", "1234 Test Street", "Ste 335", "Oak Brook", "IL", "60523", "")

        addList.Add(address1) 'Address1 this job goes to
        addList.Add(address2) 'Address2 this job goes to
        job = New click2mailVBNET.Batchc2mAPI.batchJob(1, 6, "Letter 8.5 x 11", "Address on First Page", "Next Day", "#10 Double Window", "Full Color", "White 24#", "Printing One side", "First Class", addList)
        'Addes job to batch
        batchJobs.Add(job)

        'Second Job in batch
        addList = New List(Of click2mailVBNET.Batchc2mAPI.addressItem)
        Dim address3 As New click2mailVBNET.Batchc2mAPI.addressItem("John3", "Smith", "MyOrg", "1234 Test Street", "Ste 335", "Oak Brook", "IL", "60523", "United States")
        addList.Add(address3)
        job = New click2mailVBNET.Batchc2mAPI.batchJob(7, 10, "Letter 8.5 x 11", "Address on First Page", "Next Day", "#10 Double Window", "Full Color", "White 24#", "Printing One side", "First Class", addList)
        batchJobs.Add(job)

        r.runComplete("C:\c2m\test.pdf", batchJobs)
        Console.ReadLine()
    End Sub
    Private Sub statuschanged(ByVal message As String) Handles r.statusChanged
        Console.WriteLine(message)
    End Sub
End Module
