<%@ Language=VBScript %>

<% Option Explicit %>
<%

Const ServVers = "4.00"
Const ServUpdDate = "04 July 2012"

Dim FileDoc, connect, SQLString, sConn, records
Dim rs, fs, file, cat, i, sUser, sPwd, rsprefix
Dim sBatchDownloadNum, OASISStringConv, sQueryStringArray

Dim sQueryString: sQueryString = BinToText(Request.BinaryRead(Request.TotalBytes), Request.TotalBytes)
Response.Charset = "utf-8"
Set OASISStringConv = CreateObject("OASISStringCompression.OASISCompression")
Set fs = CreateObject("Scripting.FileSystemObject")

sQueryString = OASISStringConv.DecompressStringToString(sQueryString)
sQueryStringArray = Split(sQueryString, "&")
Set file = fs.OpenTextFile(Server.mappath("..\dns.txt"), 1)   
sConn = file.ReadLine
file.Close()

If (fs.FileExists(Server.mappath("batchdownloadnum.txt"))) = True Then
    Set file = fs.OpenTextFile(Server.mappath("batchdownloadnum.txt"), 1)   
    sBatchDownloadNum = file.ReadLine ' first line is not needed
    sBatchDownloadNum = file.ReadLine
    file.Close()
Else
    sBatchDownloadNum = "200"
End If

Set fs = Nothing
Set file = Nothing
Set connect = CreateObject("ADODB.Connection")
Set records = CreateObject("ADODB.Recordset")
Set cat = CreateObject("ADOX.Catalog")
Set FileDoc = CreateObject("MSXML2.DomDocument")
connect.CommandTimeout = 0
connect.Open sConn
connect.CommandTimeout = 0
Set cat.ActiveConnection = connect
  
Select Case True
  
    Case ReqExists("getservertime") 
  
        ResponseWriteCompressed (RFC3339DateTime())
		
	Case ReqExists("batchcounttodownload") 
  
        ResponseWriteCompressed sBatchDownloadNum
  
    Case ReqExists("getDDtables") 

        ReturnXMLRecordset "select * from information_schema.tables where table_name LIKE '" & Req("getDDtables") & "%' AND TABLE_TYPE='BASE TABLE' order by table_name"

    Case ReqExists("gettables")
  
        ReturnXMLRecordset "select * from information_schema.tables order by table_name"
  
    Case ReqExists("getDDviews") 
  
        ReturnXMLRecordset "select * from information_schema.VIEWS WHERE [TABLE_NAME] LIKE '" & Req("getDDviews") & "%'"
    
    Case ReqExists("getddcolumnsdesc")

        ReturnXMLRecordset "SELECT OBJECT_NAME(c.object_id) as [TABLE_NAME], [Column Name] = c.name, [Description] = CAST(ex.value AS varchar(255)) FROM sys.columns c LEFT OUTER JOIN sys.extended_properties ex ON ex.major_id = c.object_id AND ex.minor_id = c.column_id AND ex.name = 'MS_Description' WHERE OBJECTPROPERTY(c.object_id, 'IsMsShipped')=0 AND OBJECT_NAME(c.object_id) like '" & Req("getddcolumnsdesc") & "%' ORDER BY OBJECT_NAME(c.object_id), c.column_id"

    Case ReqExists("getDDcolumnsNEW") 
  
        'ReturnXMLRecordset "select * from information_schema.columns where [table_name] LIKE '" & req("getDDcolumnsNEW") & "%' order by [table_name]"
        ReturnXMLRecordset "select * from information_schema.columns order by table_name"
  
    Case ReqExists("ID") 
       
        ReturnXMLRecordset Req("ID")
  
    Case ReqExists("user") 

        sUser = Split(Req("user"), "|||")(0)
        sPwd = Split(Req("user"), "|||")(1)
        Set rsprefix = Server.CreateObject("ADODB.Recordset")
        SQLString = "SELECT SettingTablePrefix FROM UserGroups WHERE ID IN (SELECT UserGroupID FROM Users WHERE [user] = '" & sUser & "' AND [pwd] = '" & sPwd & "')"
        rsprefix.Open SQLString, connect, 1, 4
        ReturnXMLRecordset "SELECT * FROM " & rsprefix.Fields.item("SettingTablePrefix").Value & "AppSettings"
        rsprefix.Close
        Set rsprefix = Nothing
      
    Case ReqExists("getservers") 
  
        SQLString = "immapsrv.arvixededicated.com/georgia;"
        SQLString = SQLString & "colombia.oasiswebservice.org;"
        SQLString = SQLString & "afghanistan.oasiswebservice.org;"
        SQLString = SQLString & "nomad.oasiswebservice.org;"
        SQLString = SQLString & "srfpakistan.pk/oasis;"
        SQLString = SQLString & "iraq.oasiswebservice.org;"
        SQLString = SQLString & "atlantis.oasiswebservice.org;"
        SQLString = SQLString & "yemen.oasiswebservice.org;"
        SQLString = SQLString & "somalia.oasiswebservice.org;"
        SQLString = SQLString & "syria.oasiswebservice.org"
        ResponseWriteCompressed SQLString
  
	Case ReqExists("changepwd") 
  
        sPassUser = Split("changepwd", "|||")(0)
        sPassOld = Split("changepwd", "|||")(1)
        sPassNew = Split("changepwd", "|||")(2)
        
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM [Users] WHERE [user] = '" & sPassUser & "' AND [pwd] = '" & sPassOld & "'", connect

        If Not rs.EOF Then
            connect.Execute "UPDATE [Users] SET [pwd] = '" & sPassNew & "' WHERE [user] = '" & sPassUser & "'"

            If Err.Number <> 0 Then
                ResponseWriteCompressed Err.Description
            Else
                ResponseWriteCompressed "done"
            End If

        Else
            ResponseWriteCompressed "Incorrect user/password!"
        End If
		
	case else
	
        records.CursorLocation = 3
        records.open "SELECT [name] FROM sys.sysdatabases  where [name] like 'OasisDb%' order by [name]", connect, 1, 4  
        records.Save FileDoc, 1
        ResponseWriteUnCompressed FileDoc.xml
End Select
  
Function ReturnXMLRecordset(sSqlStringSent)

    records.CursorLocation = 3
    records.Open sSqlStringSent, connect, 1, 4

    If Err.Number <> 0 Then
        ResponseWriteCompressed "-1"
    Else
        records.Save FileDoc, 1
        ResponseWriteCompressed FileDoc.xml
    End If

    records.Close

End Function
    
Function Req(item)

    Dim i
    Req = ""

    Do Until i > UBound(sQueryStringArray)

        If LCase(Left(sQueryStringArray(i), Len(item))) = LCase(item) Then
            Req = Right(sQueryStringArray(i), Len(sQueryStringArray(i)) - Len(item) - 1)
        End If

        i = i + 1
    Loop

End Function

Function ReqExists(item)

    Dim i
    ReqExists = false

    Do Until i > UBound(sQueryStringArray)

        If LCase(Left(sQueryStringArray(i), Len(item))) = LCase(item) Then
            ReqExists = true
        End If

        i = i + 1
    Loop

End Function

Function GetGUID()

    Dim NEWGUID
    Set NEWGUID = CreateObject("Scriptlet.TypeLib")
    GetGUID = Left(NEWGUID.Guid, 38)

End Function

Function RFC3339DateTime()
    '  Get the current timedate and format it as RFC 3339

    Dim g_CurrentDateTime  'As Date
    Dim iYear 'As Integer
    Dim sMonth 's String
    Dim sDay ' As String
    Dim shour 'As String
    Dim sMinute 'As String
    Dim sSec 'As String

    iYear = Year(Date)
    sMonth = Month(Date)
    sDay = Day(Date)
    shour = Hour(Now())
    sMinute = Minute(Now())
    sSec = Second(Now())

    If (iYear < 70) Then
        iYear = iYear + 2000
    ElseIf (iYear < 1900) Then
        iYear = iYear + 1900
    End If

    'var g_Month = g_CurrentDateTime.getMonth() + 1;

    If (CInt(sMonth) <= 9) Then
        sMonth = "0" & sMonth
    End If

    If (CInt(sDay) <= 9) Then
        sDay = "0" & sDay
    End If

    If (CInt(shour) <= 9) Then
        shour = "0" & shour
    End If

    If (CInt(sMinute) <= 9) Then
        sMinute = "0" & sMinute
    End If

    If (CInt(sSec) <= 9) Then
        sSec = "0" & sSec
    End If

    RFC3339DateTime = iYear & "-" & sMonth & "-" & sDay & "T" & shour & ":" & sMinute & ":" & sSec & "Z"

End Function

Function BinToText(varBinData, intDataSizeInBytes)    ' as String

if intDataSizeInBytes > 0 THEN

    Dim objRS
    Const adFldLong = &H80
    Const adVarChar = 200
    Set objRS = Server.CreateObject("ADODB.Recordset")

    objRS.Fields.Append "txt", adVarChar, intDataSizeInBytes, adFldLong
    objRS.Open

    objRS.AddNew
    objRS.Fields("txt").AppendChunk varBinData
    BinToText = objRS("txt").Value

    objRS.Close
    Set objRS = Nothing
	
End If
	
End Function

Function ResponseWriteCompressed(sStringPassed)

    Response.ContentType = "text/plain"
    'Response.ContentType = "application/x-gzip-compressed"
    'Response.ContentType = "application/octet-stream"
    Response.Charset = "windows-1252"
    'Response.Charset = "iso-8859-1"
    'Response.Charset = "utf-8"
    'Response.ContentEncoding = Encoding.GetEncoding(1252)

    If Err.Number = 0 And Not Len(sStringPassed) = 0 Then
        Response.Flush
        Response.Buffer = True
        Response.BinaryWrite OASISStringConv.CompressStringToByteArray(sStringPassed)
        Response.Flush
        Response.End
    Else
        Response.Flush
        Response.Buffer = True
        Response.BinaryWrite OASISStringConv.CompressStringToByteArray("Error: [" & Err.Description & "] ..... sStringPassed: [" & sStringPassed & "]")
        Response.Flush
        Response.End
    End If

End Function

Function ResponseWriteUnCompressed(sStringPassed)
    Response.ContentType = "text/plain"
    Response.Charset = "windows-1252"
    Response.Write sStringPassed
End Function

%>

