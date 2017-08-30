Imports click2mailVBNET
Module rest_SimpleSendCustomAddress_Test
    Private WithEvents r As New click2mailVBNET.Restc2mAPI("username", "password", Restc2mAPI.liveMode.Stage)
    Sub Main()
        Dim addressList As New List(Of List(Of KeyValuePair(Of String, String)))
        Dim customAddressItem As List(Of KeyValuePair(Of String, String)) = New List(Of KeyValuePair(Of String, String))
        customAddressItem.Add(New KeyValuePair(Of String, String)("name", "John"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("Address1", "1234 Test Street"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("Address2", "Ste 335"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("city", "Oak Brook"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("StAtE", "IL"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("Zip", "60523"))
        addressList.Add(customAddressItem)
        customAddressItem = New List(Of KeyValuePair(Of String, String))
        customAddressItem.Add(New KeyValuePair(Of String, String)("name", "John2"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("Address1", "1234 Test Street"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("Address2", "Ste 335"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("city", "Oak Brook"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("StAtE", "IL"))
        customAddressItem.Add(New KeyValuePair(Of String, String)("Zip", "60523"))
        addressList.Add(customAddressItem)
        r.runComplete("C:\c2m\test.pdf", r.createXMLFromCustomList(addressList, 12345), "Letter 8.5 x 11", "Address on Separate Page", "Next Day", "#10 Double Window", "Black and White", "White 24#", "Printing Both sides")

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
