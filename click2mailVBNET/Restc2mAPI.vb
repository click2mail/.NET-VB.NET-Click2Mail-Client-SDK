Imports System.Text
Imports System.Xml
Imports System.Net
Imports System.IO

Public Class Restc2mAPI
    Dim _username As String = ""
    Dim _password As String = ""
    Dim _pdfFile As String = ""
    Dim _batchPDFFile As String = ""

    Private Const _lRestmainurl As String = "https://rest.click2mail.com"
    Private Const _sRestmainurl As String = "https://stage-rest.click2mail.com"
    Private _authinfo As String = String.Empty
    Private _addressListName = ""
    Private _al As New List(Of addressItem)
    Public Property addressListId As Integer
    Public Property documentId As Integer
    Public Property jobId As Integer
    Public Property mode As liveMode

    Public Event statusChanged(ByVal Reason As String)
    Public Event jobStatusCheck(ByVal id As String, ByVal status As String, ByVal description As String)

    Public Sub New(username As String, pw As String, mode As liveMode)
        Me.mode = mode
        _username = username
        _password = pw
        _authinfo = _username & ":" & _password
    End Sub
    'TOOLS
    Private Function parseReturnxml(strxml As String, lookfor As String) As String

        Dim s As String = 0

        ' Create an XmlReader
        Using reader As XmlReader = XmlReader.Create(New StringReader(strxml))

            '            reader.ReadToFollowing(lookfor)
            'reader.MoveToFirstAttribute()
            'Dim genre As String = reader.Value
            'output.AppendLine("The genre value: " + genre)

            reader.ReadToFollowing(lookfor)
            s = reader.ReadElementContentAsString()
            reader.Close()
        End Using
        Return s
    End Function
    Public Function runComplete(ByVal PDF As String, addressList As String, docClass As String, layout As String, productionTime As String, envelope As String, color As String, papertype As String, printOption As String) As String
        createDocumentSimple(PDF)
        RaiseEvent statusChanged("DocumentID:" & documentId)
        createAddressListSimple(addressList)
        RaiseEvent statusChanged("AddressID:" & addressListId)
        waitForCompletedAddressList()
        createJobSimple(docClass, layout, productionTime, envelope, color, papertype, printOption)
        RaiseEvent statusChanged("JobID:" & jobId)
        submitJobSimple()
        checkJobStatus()

        'RaiseEvent statusChanged(checkJobStatus())
        RaiseEvent statusChanged("Completed")
    End Function
    Public Function checkJobStatus() As String
        Dim results As String
        Dim y As System.Collections.Specialized.NameValueCollection = New Specialized.NameValueCollection
        y.Clear()
        results = createJobPost(getRestURL() & "/molpro/jobs/" & jobId, y, "GET")
        RaiseEvent jobStatusCheck(parseReturnxml(results, "id"), parseReturnxml(results, "status"), parseReturnxml(results, "description"))
        Return results
    End Function
    Public Function submitJobSimple() As String
        Dim results As String
        Dim y As System.Collections.Specialized.NameValueCollection = New Specialized.NameValueCollection
        y.Add("billingType", "User Credit")
        results = createJobPost(getRestURL() & "/molpro/jobs/" & jobId & "/submit", y, "POST")
        Return results
    End Function
    Public Function createJobSimple(docClass As String, layout As String, productionTime As String, envelope As String, color As String, papertype As String, printOption As String) As String
        Dim y As System.Collections.Specialized.NameValueCollection = New Specialized.NameValueCollection
        y.Add("documentClass", docClass)
        y.Add("layout", layout)
        y.Add("productionTime", productionTime)
        y.Add("envelope", envelope)
        y.Add("color", color)
        y.Add("paperType", papertype)
        ''y.Add("printOption", "Printing One side")
        y.Add("printOption", printOption)

        y.Add("documentId", documentId)
        y.Add("addressId", addressListId)
        Dim results As String
        results = createJobPost(getRestURL() & "/molpro/jobs", y, "POST")
        jobId = parseReturnxml(results, "id")
        Return jobId
    End Function
    Public Sub waitForCompletedAddressList()
        Dim y As System.Collections.Specialized.NameValueCollection = New Specialized.NameValueCollection

        Dim status As String = "0"
        Dim results As String = "0"
        results = createJobPost(getRestURL() & "/molpro/addressLists/" & addressListId, y, "GET")
        status = parseReturnxml(results, "status")
        If (status <> "3") Then

            While (status <> "3")
                RaiseEvent statusChanged("Waiting Address List to processes.  Current Status is: " & status)
                results = createJobPost(getRestURL() & "/molpro/addressLists/" & addressListId, y, "GET")
                status = parseReturnxml(results, "status")
                System.Threading.Thread.Sleep(5000)
            End While
        End If
        RaiseEvent statusChanged("The status received is 3, which means we can proceed")
    End Sub
    Public Function createDocumentSimple(pdf As String) As Integer
        Dim x As System.Collections.Specialized.NameValueCollection = New Specialized.NameValueCollection
        x.Add("documentName", "sample Letter")
        x.Add("documentClass", "Letter 8.5 x 11")
        x.Add("documentFormat", "PDF")
        Dim results As String = createDocument(getRestURL() & "/molpro/documents/", pdf, "file", "application/pdf", x)
        documentId = parseReturnxml(results, "id")
        Return documentId
    End Function

    Public Function createAddressListSimple(Xml As String)
        Dim results As String = createAddressList(getRestURL() & "/molpro/addressLists/", Xml)
        addressListId = parseReturnxml(results, "id")
        Return addressListId
    End Function

    Public Property addressList As List(Of addressItem)
        Get

            Return _al
        End Get
        Set(value As List(Of addressItem))
            _al = value
        End Set
    End Property
    Public Class addressItem
        Public _First_name As String = ""
        Public _Last_name As String = ""
        Public _Organization As String = ""
        Public _Address1 As String = ""
        Public _Address2 As String = ""
        Public _Address3 As String = ""
        Public _City As String = ""
        Public _State As String = ""
        Public _Zip As String = ""
        Public _Country_nonUS As String = ""
        Public Sub New(ByVal fname As String, lname As String, org As String, address1 As String, address2 As String, address3 As String, city As String, state As String, zip As String, country As String)
            _First_name = fname
            _Last_name = lname
            _Organization = org
            _Address1 = address1
            _Address2 = address2
            _Address3 = address3
            _City = city
            _State = state
            _Zip = zip
            _Country_nonUS = country
        End Sub
    End Class
    Public Function GenerateStreamFromString(s As String) As Stream
        Dim stream As New MemoryStream()
        Dim writer As New StreamWriter(stream)
        writer.Write(s)
        writer.Flush()
        stream.Position = 0
        Return stream
    End Function
    Private Function getRestURL() As String
        If _mode = True Then
            Return _lRestmainurl
        Else
            Return _sRestmainurl
        End If
    End Function
    Private Function createAddressList( _
   ByVal uri As String, _
   xml As String) As String
        Dim responseText As String = ""
        Dim request As Net.HttpWebRequest = Net.WebRequest.Create(uri)

        request.ContentType = "application/xml"
        request.Method = "POST"
        request.KeepAlive = True
        'request.Credentials = Net.CredentialCache.DefaultCredentials

        Dim authinfo As String
        authinfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authinfo))
        request.Headers("Authorization") = "Basic " + authinfo

        Using requestStream As IO.Stream = request.GetRequestStream()

            Using fileStream As Stream = GenerateStreamFromString(xml)

                Dim buffer(4096) As Byte
                Dim bytesRead As Int32 = fileStream.Read(buffer, 0, buffer.Length)

                Do While (bytesRead > 0)

                    requestStream.Write(buffer, 0, bytesRead)
                    bytesRead = fileStream.Read(buffer, 0, buffer.Length)

                Loop

            End Using



        End Using

        Dim response As Net.WebResponse = Nothing

        Try

            response = request.GetResponse()

            Using responseStream As IO.Stream = response.GetResponseStream()

                Using responseReader As New IO.StreamReader(responseStream)

                    responseText = responseReader.ReadToEnd()


                End Using

            End Using

        Catch exception As Net.WebException

            response = exception.Response

            If (response IsNot Nothing) Then

                Using reader As New IO.StreamReader(response.GetResponseStream())

                    responseText = reader.ReadToEnd()


                End Using

                response.Close()

            End If

        Finally

            request = Nothing

        End Try
        Return responseText

    End Function

    Public Enum liveMode
        Live = True
        Stage = False
    End Enum
    Public Function RemoveTroublesomeCharacters(inString As String) As String
        If inString Is Nothing Then
            Return Nothing
        End If

        Dim newString As New StringBuilder()
        Dim ch As Char

        For i As Integer = 0 To inString.Length - 1

            ch = inString(i)
            ' remove any characters outside the valid UTF-8 range as well as all control characters
            ' except tabs and new lines
            'if ((ch < 0x00FD && ch > 0x001F) || ch == '\t' || ch == '\n' || ch == '\r')
            'if using .NET version prior to 4, use above logic
            If XmlConvert.IsXmlChar(ch) Then
                'this method is new in .NET 4
                newString.Append(ch)
            End If
        Next
        Return newString.ToString()

    End Function
    Public Function createJobPost(url As String, nameValueCollection As Specialized.NameValueCollection, Method As String) As String
        ' Here we convert the nameValueCollection to POST data.
        ' This will only work if nameValueCollection contains some items.
        Dim parameters = New StringBuilder()
        Dim responseText As String = ""
        For Each key As String In nameValueCollection.Keys
            parameters.AppendFormat("{0}={1}&", WebUtility.UrlEncode(key), WebUtility.UrlEncode(nameValueCollection(key)))
        Next
        If (parameters.Length > 0) Then
            parameters.Length -= 1
        End If

        ' Here we create the request and write the POST data to it.
        Dim request = DirectCast(HttpWebRequest.Create(url), HttpWebRequest)
        request.Method = Method
        request.KeepAlive = True
        'request.Credentials = Net.CredentialCache.DefaultCredentials

        Dim authinfo As String
        authinfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authinfo))
        request.Headers("Authorization") = "Basic " + authinfo

        If (parameters.Length > 0) Then
            Using writer = New StreamWriter(request.GetRequestStream())
                writer.Write(parameters.ToString())
            End Using
        End If

        Dim response As Net.WebResponse = Nothing

        Try

            response = request.GetResponse()

            Using responseStream As IO.Stream = response.GetResponseStream()

                Using responseReader As New IO.StreamReader(responseStream)

                    responseText = responseReader.ReadToEnd()


                End Using

            End Using

        Catch exception As Net.WebException

            response = exception.Response

            If (response IsNot Nothing) Then

                Using reader As New IO.StreamReader(response.GetResponseStream())

                    responseText = reader.ReadToEnd()


                End Using

                response.Close()

            End If

        Finally

            request = Nothing

        End Try
        Return responseText
    End Function

    Public Function createXMLFromAddressList() As String
        Dim doc As New Xml.XmlDocument
        'create nodes
        Dim root As Xml.XmlElement = doc.CreateElement("addressList")


        Dim addressListName As Xml.XmlElement = doc.CreateElement("addressListName")
        _addressListName = Guid.NewGuid().ToString("N")
        addressListName.InnerXml = _addressListName

        root.AppendChild(addressListName)

        Dim addressMappingId As Xml.XmlElement = doc.CreateElement("addressMappingId")
        addressMappingId.InnerXml = 2
        root.AppendChild(addressMappingId)

        Dim addresses As Xml.XmlElement = doc.CreateElement("addresses")
        root.AppendChild(addresses)

        For Each a As addressItem In addressList
            Dim address As Xml.XmlElement = doc.CreateElement("address")
            Dim fname As Xml.XmlElement = doc.CreateElement("First_name")
            fname.InnerXml = a._First_name
            address.AppendChild(fname)
            Dim lname As Xml.XmlElement = doc.CreateElement("Last_name")
            lname.InnerXml = a._Last_name
            address.AppendChild(lname)
            Dim Organization As Xml.XmlElement = doc.CreateElement("Organization")
            Organization.InnerXml = a._Organization
            address.AppendChild(Organization)
            Dim Address1 As Xml.XmlElement = doc.CreateElement("Address1")
            Address1.InnerXml = a._Address1
            address.AppendChild(Address1)
            Dim Address2 As Xml.XmlElement = doc.CreateElement("Address2")
            Address2.InnerXml = a._Address2
            address.AppendChild(Address2)
            Dim Address3 As Xml.XmlElement = doc.CreateElement("Address3")
            Address3.InnerXml = a._Address3
            address.AppendChild(Address3)
            Dim City As Xml.XmlElement = doc.CreateElement("City")
            City.InnerXml = a._City
            address.AppendChild(City)
            Dim State As Xml.XmlElement = doc.CreateElement("State")
            State.InnerXml = a._State
            address.AppendChild(State)
            Dim zip As Xml.XmlElement = doc.CreateElement("zip")
            zip.InnerXml = a._Zip
            address.AppendChild(zip)
            Dim country As Xml.XmlElement = doc.CreateElement("Country_non-US")
            country.InnerXml = a._Country_nonUS
            address.AppendChild(country)
            addresses.AppendChild(address)
        Next

        doc.AppendChild(root)
        Dim xmlString As String
        Using stringWriter = New StringWriter()
            Using xmlTextWriter = XmlWriter.Create(stringWriter)
                doc.WriteTo(xmlTextWriter)
                xmlTextWriter.Flush()

                xmlString = stringWriter.GetStringBuilder().ToString()
            End Using
        End Using
        Return xmlString

    End Function
    Public Function createXMLFromCustomList(ByVal myList As List(Of List(Of KeyValuePair(Of String, String))), ByVal AddressListId As Integer) As String
        Dim doc As New Xml.XmlDocument
        'create nodes
        Dim root As Xml.XmlElement = doc.CreateElement("addressList")


        Dim addressListName As Xml.XmlElement = doc.CreateElement("addressListName")
        _addressListName = Guid.NewGuid().ToString("N")
        addressListName.InnerXml = _addressListName

        root.AppendChild(addressListName)

        Dim addressMappingId As Xml.XmlElement = doc.CreateElement("addressMappingId")
        addressMappingId.InnerXml = AddressListId

        root.AppendChild(addressMappingId)

        Dim addresses As Xml.XmlElement = doc.CreateElement("addresses")
        root.AppendChild(addresses)
        Dim address As Xml.XmlElement = Nothing
        For Each a In myList
            For Each aa As KeyValuePair(Of String, String) In a
                address = doc.CreateElement("address")
                Dim i As Xml.XmlElement = doc.CreateElement(aa.Key)
                i.InnerXml = aa.Value
                address.AppendChild(i)
            Next
            addresses.AppendChild(address)
        Next

        doc.AppendChild(root)
        Dim xmlString As String
        Using stringWriter = New StringWriter()
            Using xmlTextWriter = XmlWriter.Create(stringWriter)
                doc.WriteTo(xmlTextWriter)
                xmlTextWriter.Flush()

                xmlString = stringWriter.GetStringBuilder().ToString()
            End Using
        End Using
        Return xmlString

    End Function
    Private Function createDocument( _
     ByVal uri As String, _
     ByVal filePath As String, _
     ByVal fileParameterName As String, _
     ByVal contentType As String, _
     ByVal otherParameters As Specialized.NameValueCollection) As String
        Dim responseText As String = ""
        Dim boundary As String = "---------------------------" & DateTime.Now.Ticks.ToString("x")
        Dim newLine As String = System.Environment.NewLine
        Dim boundaryBytes As Byte() = Text.Encoding.ASCII.GetBytes(newLine & "--" & boundary & newLine)
        Dim request As Net.HttpWebRequest = Net.WebRequest.Create(uri)

        request.ContentType = "multipart/form-data; boundary=" & boundary
        request.Method = "POST"
        request.KeepAlive = True
        'request.Credentials = Net.CredentialCache.DefaultCredentials

        Dim authinfo As String
        authinfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authinfo))
        request.Headers("Authorization") = "Basic " + authinfo

        Using requestStream As IO.Stream = request.GetRequestStream()

            Dim formDataTemplate As String = "Content-Disposition: form-data; name=""{0}""{1}Content-Type: text/plain; charset=UTF-8{1}{1}{2}"

            For Each key As String In otherParameters.Keys

                requestStream.Write(boundaryBytes, 0, boundaryBytes.Length)
                Dim formItem As String = String.Format(formDataTemplate, key, newLine, otherParameters(key))
                Dim formItemBytes As Byte() = Text.Encoding.UTF8.GetBytes(formItem)
                requestStream.Write(formItemBytes, 0, formItemBytes.Length)

            Next key

            requestStream.Write(boundaryBytes, 0, boundaryBytes.Length)

            Dim headerTemplate As String = "Content-Disposition: form-data; name=""{0}""; filename=""{1}""{2}Content-Type: {3}{2}Content-Transfer-Encoding: binary{2}{2}"
            Dim header As String = String.Format(headerTemplate, fileParameterName, filePath, newLine, contentType)
            Dim headerBytes As Byte() = Text.Encoding.UTF8.GetBytes(header)
            requestStream.Write(headerBytes, 0, headerBytes.Length)

            Using fileStream As New IO.FileStream(filePath, IO.FileMode.Open, IO.FileAccess.Read)

                Dim buffer(4096) As Byte
                Dim bytesRead As Int32 = fileStream.Read(buffer, 0, buffer.Length)

                Do While (bytesRead > 0)

                    requestStream.Write(buffer, 0, bytesRead)
                    bytesRead = fileStream.Read(buffer, 0, buffer.Length)

                Loop

            End Using

            Dim trailer As Byte() = Text.Encoding.ASCII.GetBytes(newLine & "--" + boundary + "--" & newLine)
            requestStream.Write(trailer, 0, trailer.Length)

        End Using

        Dim response As Net.WebResponse = Nothing

        Try

            response = request.GetResponse()

            Using responseStream As IO.Stream = response.GetResponseStream()

                Using responseReader As New IO.StreamReader(responseStream)

                    responseText = responseReader.ReadToEnd()


                End Using

            End Using

        Catch exception As Net.WebException

            response = exception.Response

            If (response IsNot Nothing) Then

                Using reader As New IO.StreamReader(response.GetResponseStream())

                    responseText = reader.ReadToEnd()


                End Using

                response.Close()

            End If

        Finally

            request = Nothing

        End Try
        Return responseText

    End Function
End Class
