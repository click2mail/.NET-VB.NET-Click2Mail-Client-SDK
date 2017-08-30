Imports System.Text
Imports System.Xml
Imports System.Net
Imports System.IO

Public Class Batchc2mAPI
    Dim _username As String = ""
    Dim _password As String = ""
    Dim _pdfFile As String = ""
    Dim _batchPDFFile As String = ""

    Private Const _Smainurl As String = "https://stage-batch.click2mail.com"
    Private Const _lmainurl As String = "https://batch.click2mail.com"
    
    Private _authinfo As String = String.Empty
    Private _addressListName = ""

    Public Property pdf As String
    Public Property batchId As Integer
    Public Property mode As liveMode
    Public Property jobList As List(Of batchJob)

    Public Event statusChanged(ByVal Reason As String)


    Public Sub New(username As String, pw As String, mode As liveMode)
        Me.mode = mode
        _username = username
        _password = pw
        _authinfo = _username & ":" & _password
    End Sub
    Public Function createBatchSimple() As Integer
        Dim results As String = createbatch()
        batchId = parseReturnxml(results, "id")

        Return batchId
    End Function
    'TOOLS

    Private Function createbatch() As String
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader
        Dim address As Uri
        Dim data As StringBuilder = Nothing
        Dim byteData() As Byte = Nothing
        Dim postStream As Stream = Nothing

        address = New Uri(getBatchURL() & "/v1/batches")

        Dim authinfo As String
        authinfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authinfo))

        ' Create the web request  
        request = DirectCast(WebRequest.Create(address), HttpWebRequest)
        request.Headers("Authorization") = "Basic " + authinfo
        request.Method = "POST"
        request.ContentType = "text/plain"
        Try
            response = request.GetResponse()
        Catch wex As WebException
            ' This exception will be raised if the server didn't return 200 - OK  
            ' Try to retrieve more information about the network error  
            If Not wex.Response Is Nothing Then
                Dim errorResponse As HttpWebResponse = Nothing
                Try
                    errorResponse = DirectCast(wex.Response, HttpWebResponse)
                    Console.WriteLine( _
                      "The server returned '{0}' with the status code {1} ({2:d}).", _
                      errorResponse.StatusDescription, errorResponse.StatusCode, _
                      errorResponse.StatusCode)
                Finally
                    If Not errorResponse Is Nothing Then errorResponse.Close()
                End Try
            End If
        Finally
            If Not postStream Is Nothing Then postStream.Close()
        End Try
        '
        Try

            reader = New StreamReader(response.GetResponseStream())

            ' Console application output  
            Dim s As String = reader.ReadToEnd()
            reader.Close()
            ' Console.Write(s)
            Return s
            '    c2m.StatusPick.jobStatus = parsexml(s, "status")
            'MsgBox(s)

        Finally
            ' If c2m.jobid = 0 Then
            '            c2m.StatusPick.jobStatus = 99
            'End If
            'If Not response Is Nothing Then response.Close()
        End Try
    End Function
    Public Function createXMLBatchPost() As String
        Dim doc As New Xml.XmlDocument
        doc.AppendChild(doc.CreateXmlDeclaration("1.0", "utf-8", Nothing))
        'create nodes
        Dim root As Xml.XmlElement = doc.CreateElement("batch")

        '"    <username>" & username & "</username>" &
        '"    <password>" & password & "</password>" &
        '"    <filename>" & fileName & "</filename>" &
        '"    <appSignature>MyTest App</appSignature>" &
        Dim attr As Xml.XmlElement = doc.CreateElement("username")
        attr.InnerXml = Me._username
        root.AppendChild(attr)

        attr = doc.CreateElement("password")
        attr.InnerXml = Me._password
        root.AppendChild(attr)

        attr = doc.CreateElement("filename")
        attr.InnerXml = Me.pdf
        root.AppendChild(attr)


        attr = doc.CreateElement("appSignature")
        attr.InnerXml = ".NET SDK API"
        root.AppendChild(attr)
        For Each b As batchJob In jobList
            Dim job = doc.CreateElement("job")

            attr = doc.CreateElement("startingPage")
            attr.InnerXml = b.startingPage
            job.AppendChild(attr)

            attr = doc.CreateElement("endingPage")
            attr.InnerXml = b.endingPage
            job.AppendChild(attr)

            Dim printProductIOptions = doc.CreateElement("printProductionOptions")
            job.AppendChild(printProductIOptions)
            '<documentClass>Letter 8.5 x 11</documentClass>" &
            '"            <layout>Address on First Page</layout>" &
            '"            <productionTime>Next Day</productionTime>" &
            '"            <envelope>#10 Double Window</envelope>" &
            '"            <color>Full Color</color>" &
            '"            <paperType>White 24#</paperType>" &
            '"            <printOption>Printing One side</printOption>" &
            '"            <mailClass>First Class</mailClass>" &
            attr = doc.CreateElement("documentClass")
            attr.InnerXml = b.documentClass
            printProductIOptions.AppendChild(attr)

            attr = doc.CreateElement("layout")
            attr.InnerXml = b.layout
            printProductIOptions.AppendChild(attr)

            attr = doc.CreateElement("productionTime")
            attr.InnerXml = b.productionTime
            printProductIOptions.AppendChild(attr)
            attr = doc.CreateElement("envelope")
            attr.InnerXml = b.envelope
            printProductIOptions.AppendChild(attr)
            attr = doc.CreateElement("color")
            attr.InnerXml = b.color
            printProductIOptions.AppendChild(attr)

            attr = doc.CreateElement("paperType")
            attr.InnerXml = b.paperType
            printProductIOptions.AppendChild(attr)
            attr = doc.CreateElement("printOption")
            attr.InnerXml = b.printOption
            printProductIOptions.AppendChild(attr)
            attr = doc.CreateElement("mailClass")
            attr.InnerXml = b.mailClass
            printProductIOptions.AppendChild(attr)

            Dim addressList As XmlElement = doc.CreateElement("recipients")
            job.AppendChild(addressList)
            For Each ai In b.addressList
                Dim address As XmlElement = doc.CreateElement("address")
                addressList.AppendChild(address)
                attr = doc.CreateElement("name")
                If (ai._First_name.Length > 0 And ai._Last_name.Length > 0) Then
                    attr.InnerText = ai._Last_name & ", " & ai._First_name
                Else
                    attr.InnerText = Trim(ai._First_name & " " & ai._Last_name)
                End If
                address.AppendChild(attr)

                attr = doc.CreateElement("organization")
                attr.InnerText = Trim(ai._Organization)
                address.AppendChild(attr)

                attr = doc.CreateElement("address1")
                attr.InnerText = Trim(ai._Address1)
                address.AppendChild(attr)

                attr = doc.CreateElement("address2")
                attr.InnerText = Trim(ai._Address2)
                address.AppendChild(attr)

                attr = doc.CreateElement("address3")
                attr.InnerText = ""
                address.AppendChild(attr)

                attr = doc.CreateElement("city")
                attr.InnerText = Trim(ai._City)
                address.AppendChild(attr)

                attr = doc.CreateElement("state")
                attr.InnerText = Trim(ai._State)
                address.AppendChild(attr)

                attr = doc.CreateElement("postalCode")
                attr.InnerText = Trim(ai._Zip)
                address.AppendChild(attr)

                attr = doc.CreateElement("country")
                attr.InnerText = Trim(ai._Country_nonUS)
                address.AppendChild(attr)
                
   
            Next
            root.AppendChild(job)

        Next

        doc.AppendChild(root)

        'doc.Declaration = New XDeclaration("1.0", "utf-8", Nothing)
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
    Public Class batchJob
        Public Property startingPage As Integer
        Public Property endingPage As Integer
        Public Property documentClass As String
        Public Property layout As String
        Public Property productionTime As String
        Public Property envelope As String
        Public Property color As String
        Public Property paperType As String
        Public Property printOption As String
        Public Property mailClass As String
        Public Property addressList As List(Of addressItem)

        Public Sub New(startingPage As Integer, endingPage As Integer, documentClass As String, layout As String, productionTime As String, envelope As String, color As String, paperType As String, printOption As String, mailClass As String, addressList As List(Of addressItem))
            Me.startingPage = startingPage
            Me.endingPage = endingPage
            Me.documentClass = documentClass
            Me.layout = layout
            Me.productionTime = productionTime
            Me.envelope = envelope
            Me.color = color
            Me.paperType = paperType
            Me.printOption = printOption
            Me.mailClass = mailClass
            Me.addressList = addressList
        End Sub
    End Class
    Public Class addressItem
        Public _First_name As String = ""
        Public _Last_name As String = ""
        Public _Organization As String = ""
        Public _Address1 As String = ""
        Public _Address2 As String = ""

        Public _City As String = ""
        Public _State As String = ""
        Public _Zip As String = ""
        Public _Country_nonUS As String = ""
        Public Sub New(ByVal fname As String, lname As String, org As String, address1 As String, address2 As String, city As String, state As String, zip As String, country As String)
            _First_name = fname
            _Last_name = lname
            _Organization = org
            _Address1 = address1
            _Address2 = address2
            _City = city
            _State = state
            _Zip = zip
            _Country_nonUS = country
        End Sub
    End Class
    Private Sub uploadBatchxml()
        Dim _XMLDOC As New XmlDocument
        'Console.Write(createXMLBatchPost())
        'Return
        _XMLDOC.LoadXml(createXMLBatchPost())
        Dim strURI As String = String.Empty
        strURI = getBatchURL() & "/v1/batches/" & batchId
        PutObject(strURI, _XMLDOC)
        Return

        Dim request As HttpWebRequest = CType(WebRequest.Create(strURI), HttpWebRequest)
        Dim authinfo As String
        authinfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authinfo))
        request.Headers("Authorization") = "Basic " + authinfo
        request.Accept = "text/xml"
        request.Method = "PUT"
        Using ms As New MemoryStream()
            _XMLDOC.Save(ms)
            request.ContentLength = ms.Length
            ms.WriteTo(request.GetRequestStream())
        End Using
        Dim result As String

        Using response As WebResponse = request.GetResponse()
            Using reader As New StreamReader(response.GetResponseStream())
                result = reader.ReadToEnd()
            End Using
        End Using

        Return


        'Console.WriteLine(result)
    End Sub

    Public Function PutObject(postUrl As String, xmlDoc As XmlDocument) As [String]
        Dim myCreds As New NetworkCredential(_username, _password)

        Dim xmlStream As New MemoryStream()
        xmlDoc.Save(xmlStream)

        Dim result As String = ""
        xmlStream.Flush()
        'Adjust this if you want read your data 
        xmlStream.Position = 0

        Using client = New System.Net.WebClient()
            client.Credentials = myCreds
            client.Headers.Add("Content-Type", "application/xml")
            Dim b As Byte() = client.UploadData(postUrl, "PUT", xmlStream.ToArray())
            'Dim b As Byte() = client.UploadFile(postUrl, "PUT", "C:\test\test.xml")

            result = client.Encoding.GetString(b)
        End Using

        Return result
    End Function
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
    Public Sub uploadBatchPDF()

        Dim client As New WebClient

        Dim strURI As String = String.Empty
        strURI = getBatchURL() & "/v1/batches/" & batchId
        Dim authinfo As String
        authinfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authinfo))
        client.Headers("Authorization") = "Basic " + authinfo
        client.Headers.Add("Content-Type", "application/pdf")
        'Dim sentXml As Byte() = System.Text.Encoding.ASCII.GetBytes(_XMLDOC.OuterXml)

        Dim fInfo As New FileInfo(pdf)

        Dim numBytes As Long = fInfo.Length

        Dim fStream As New FileStream(pdf, FileMode.Open, FileAccess.Read)

        Dim br As New BinaryReader(fStream)

        Dim data As Byte() = br.ReadBytes(CInt(numBytes))

        ' Show the number of bytes in the array.


        br.Close()

        fStream.Close()




        Dim response As Byte() = client.UploadData(strURI, "PUT", Data)

        Console.WriteLine(System.Text.Encoding.Default.GetString(response))


        'Console.WriteLine(response.ToString())
    End Sub
    Public Function getbatchstatus() As String
        Dim strURI As String = String.Empty
        strURI = getBatchURL() & "/v1/batches/" & batchId
        Console.WriteLine(strURI)
        Dim request As HttpWebRequest = CType(WebRequest.Create(strURI), HttpWebRequest)
        Dim authinfo As String
        authinfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authinfo))
        request.Headers("Authorization") = "Basic " + authinfo
        request.Method = System.Net.WebRequestMethods.Http.Get
        Dim result As String
        Try
            Using response As WebResponse = request.GetResponse()
                Using reader As New StreamReader(response.GetResponseStream())
                    result = reader.ReadToEnd()

                End Using
            End Using
            Return result
            'Console.Write(result)

            Return parseReturnxml(result, "hasErrors")
        Catch ex As Exception
            Return ex.Message
            'Console.Write(ex.Message)
        End Try
        Return ""
    End Function
    Private Sub submitbatch()
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim reader As StreamReader
        Dim address As Uri

        Dim postStream As Stream = Nothing

        address = New Uri(getBatchURL() & "/v1/batches/" & batchId)

        Dim authinfo As String
        authinfo = Convert.ToBase64String(Encoding.Default.GetBytes(_authinfo))

        ' Create the web request  
        request = DirectCast(WebRequest.Create(address), HttpWebRequest)
        request.Headers("Authorization") = "Basic " + authinfo
        request.Method = "POST"

        Try
            response = request.GetResponse()
        Catch wex As WebException
            ' This exception will be raised if the server didn't return 200 - OK  
            ' Try to retrieve more information about the network error  
            If Not wex.Response Is Nothing Then
                Dim errorResponse As HttpWebResponse = Nothing
                Try
                    errorResponse = DirectCast(wex.Response, HttpWebResponse)
                    Console.WriteLine( _
                      "The server returned '{0}' with the status code {1} ({2:d}).", _
                      errorResponse.StatusDescription, errorResponse.StatusCode, _
                      errorResponse.StatusCode)
                Finally
                    If Not errorResponse Is Nothing Then errorResponse.Close()
                End Try
            End If
        Finally
            If Not postStream Is Nothing Then postStream.Close()
        End Try
        '
        Try

            reader = New StreamReader(response.GetResponseStream())
        Finally
        End Try

    End Sub
    Public Function runComplete(ByVal PDF As String, jobList As List(Of batchJob)) As String
        batchId = createBatchSimple()
        Me.pdf = PDF
        Me.jobList = jobList
        RaiseEvent statusChanged("BatchID Created:" & batchId)
        'RaiseEvent statusChanged(createXMLBatchPost())
        uploadBatchxml()
        RaiseEvent statusChanged("XML UPLOAD Completed")
        uploadBatchPDF()
        RaiseEvent statusChanged("PDF UPLOAD Completed")
        submitbatch()
        RaiseEvent statusChanged("Batch UPLOAD Completed")

        RaiseEvent statusChanged(getbatchstatus())
        Return ""
    End Function

    Private Function getBatchURL() As String
        If _mode = True Then
            Return _lmainurl
        Else
            Return _Smainurl
        End If
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

End Class
