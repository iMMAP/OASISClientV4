
<%

Dim FileDoc, bytecount, records, connect, file, fs
Dim sQueryString: sQueryString = BinToText(Request.BinaryRead(Request.TotalBytes), Request.TotalBytes)
Dim OASISStringConv
Set OASISStringConv = CreateObject("OASISStringCompression.OASISCompression")

bytecount = Request.TotalBytes

if bytecount > 10 then

	Set connect = CreateObject("ADODB.Connection")
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set file = fs.OpenTextFile(Server.mappath("..\dns.txt"), 1)  ' 1 = ForReading
    sConn = file.ReadLine
  	file.Close()
  	connect.Open sConn

	Set FileDoc = CreateObject("MSXML2.DomDocument")
	Set records = CreateObject("ADODB.Recordset")


  	sQueryString = OASISStringConv.DecompressStringToString(sQueryString)
	FileDoc.LoadXML sQueryString 'Request.BinaryRead(bytecount)



	'Response.Write FileDoc.xml
	records.CursorLocation = 3
	records.LockType = 4
	records.Open FileDoc
	records.ActiveConnection = connect
	records.UpdateBatch
	records.Close

	If Err.Number <> 0 Then
		ResponseWriteCompressed Err.Description & Err.Number
	Else
		ResponseWriteCompressed "Data Updated"
	End If

end if

Function BinToText(varBinData, intDataSizeInBytes)    ' as String
        dim objRS
        Const adFldLong = &H00000080
        Const adVarChar = 200
        Set objRS = Server.CreateObject("ADODB.Recordset")

        objRS.Fields.Append "txt", adVarChar, intDataSizeInBytes, adFldLong
        objRS.Open

        objRS.AddNew
        objRS.Fields("txt").AppendChunk varBinData
        BinToText = objRS("txt").Value

        objRS.Close
        Set objRS = Nothing
    End Function

 Function ResponseWriteCompressed(sStringPassed)

 	'On Error Resume Next

 	Response.ContentType = "text/plain"
 	'Response.ContentType = "application/x-gzip-compressed"
 	'Response.ContentType = "application/octet-stream"

 	Response.Charset = "windows-1252"
 	'Response.Charset = "iso-8859-1"
 	'Response.Charset = "utf-8"
 	'Response.ContentEncoding = Encoding.GetEncoding(1252)

 	If Err.Number = 0 AND Not len(sStringPassed) = 0 then
 		Response.Flush
 		Response.Buffer = True
 		'Response.BinaryWrite OASISStringConv.CompressStringToByteArray(len(sStringPassed))
 		Response.BinaryWrite OASISStringConv.CompressStringToByteArray(sStringPassed)
 		Response.Flush
 		Response.End
 	Else
 		Response.Flush
 		Response.Buffer = True
 		Response.BinaryWrite OASISStringConv.CompressStringToByteArray("Error: [" & err.description & "] ..... sStringPassed: [" & sStringPassed & "]")
 		Response.Flush
 		Response.End
 	End If

 End Function

%>


