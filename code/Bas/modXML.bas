Attribute VB_Name = "modXML"
Option Explicit

'
'Option Explicit
'
'
''=============================================================================================================
''
'' modXML Module
'' -------------
''
'' Created By  : Kevin Wilson
''               http://www.TheVBZone.com   ( The VB Zone )
''               http://www.TheVBZone.net   ( The VB Zone .net )
''
'' Created On  : June 15, 2001
'' Last Update : February 19, 2003
''
'' VB Versions : 5.0 / 6.0
''
'' Requires    : XML 2.0 (or better) support (This support comes with IE 4.x or better, or by installing the XML components)
''               ADO 2.5 (or better) support (This support comes with the Microsoft Data Access Components (MDAC) installation)
''
'' Description : This module gives you the ability to easily load and save XML in various formats as well as
''               convert XML back and forth between XML and ADO (Recordset).
''
'' See Also    : XML Reference
''               http://msdn.microsoft.com/library/default.asp?URL=/library/psdk/xmlsdk/xml_9yg5.htm
''
''               XML (Extensible Markup Language)
''               http://msdn.microsoft.com/library/default.asp?URL=/library/psdk/xmlsdk/xml_9yg5.htm
''
''               Universal Data Access Web Site
''               http://www.microsoft.com/data/ado/
''
''               Universal Data Access - Free Downloads
''               http://www.microsoft.com/data/download.htm
''
''               ADO Version 2.6
''               http://msdn.microsoft.com/library/default.asp?URL=/library/psdk/dasdk/ados4piv.htm
''
'
'
'Private strFiltersSub()    As String
'Private intFiltersSubCount As Integer
'
'
'
''=============================================================================================================
'' xml_LoadXML
''
'' Purpose:
'' --------
'' Loads the specified XML file to a string and returns that string.  This function uses straight ASCII TEXT
'' instead of objects so it is less error prone and more flexible as compared to "xml_LoadXML_Ex".
''
'' NOTE : If BOTH the "strXmlSource" and "strLoadPath" parameters are specified, the "strLoadPath"
'' parameter is used.  If neither are specified, the function fails.
''
'' Param                Use
'' ------------------------------------
'' Return_XML           Returns the XML document as a string
'' strLoadPath          Specifies the XML file to retrieve the XML from
'' blnIncludeCRLF       Optional. If set to TRUE will include the vbCrLf characters at the end of each line.
''                      Otherwise will leave them off to conserve size of XML.
''
'' Return:
'' -------
'' Returns TRUE if the function executed successfully
'' Returns FALSE if the function failed
''
''=============================================================================================================
'Public Function xml_LoadXML(ByRef Return_XML As String, _
'                            ByVal strLoadPath As String, _
'                            Optional ByVal blnIncludeCRLF As Boolean = False) As Boolean
'On Error Resume Next
'
'  Dim intFileNum     As Integer
'  Dim strCurrentLine As String
'
'  ' Erase the return variables
'  Return_XML = ""
'
'  ' Validate parameters
'  strLoadPath = Trim(strLoadPath)
'
'  ' Make sure the file specified exists
'  If FileExists(strLoadPath) = False Then Exit Function
'
'  ' Go to the file and import the XML
'  intFileNum = FreeFile
'  Open strLoadPath For Input As #intFileNum
'    Do While EOF(intFileNum) = False
'      Line Input #intFileNum, strCurrentLine
'      If blnIncludeCRLF = True Then
'        Return_XML = Return_XML & strCurrentLine & vbCrLf
'      Else
'        Return_XML = Return_XML & strCurrentLine
'      End If
'    Loop
'  Close #intFileNum
'
'  ' Function executed successfully
'  xml_LoadXML = True
'
'End Function
'
'
''=============================================================================================================
'' xml_LoadXML_Ex
''
'' Purpose:
'' --------
'' This function takes the XML source or XML loaded from the specified file and converts it to an
'' MSXML.DOMDocument object and returns a reference to it.  This function uses the MSXML object to do the XML
'' work so is more complex and more error prone compared to the "xml_LoadXML" function.
''
'' NOTE : If BOTH the "strXmlSource" and "strLoadPath" parameters are specified, the "strLoadPath"
'' parameter is used.  If neither are specified, the function fails.
''
'' Param                Use
'' ------------------------------------
'' Return_XML           Returns the XML as an MSXML.DOMDocument object
'' strXmlSource         Optional. Specifies the XML source to load to the MSXML object (string format).
''                      If this parameter is not specified, the "strLoadPath" parameter MUST be specified.
'' strLoadPath          Optional. Specifies the location of the file to load to an MSXML object.  If this
''                      is not specified, the "strXmlSource" parameter MUST be specified.
'' blnShowErrorMsgs     Optional. If set to TRUE and an error occurs, an error message will be shown to the user.
'' Return_ErrNum        Optional. Returns the error number to any error that occured.
'' Return_ErrDesc       Optional. Returns the error description to any error that occured.
''
'' Return:
'' -------
'' Returns TRUE if the function executed successfully
'' Returns FALSE if the function failed
''
''=============================================================================================================
'Public Function xml_LoadXML_Ex(ByRef Return_XML As MSXML.DOMDocument, _
'                               Optional ByVal strXmlSource As String, _
'                               Optional ByVal strLoadPath As String, _
'                               Optional ByVal blnShowErrorMsgs As Boolean = False, _
'                               Optional ByRef Return_ErrNum As Long, _
'                               Optional ByRef Return_ErrDesc As String) As Boolean
'On Error Resume Next
'
'  ' Erase the return variables
'  Set Return_XML = Nothing
'  Return_ErrNum = 0
'  Return_ErrDesc = ""
'
'  ' Validate parameters
'  strXmlSource = Trim(strXmlSource)
'  strLoadPath = Trim(strLoadPath)
'
'  ' Make sure at least one valid source was specified
'  If (strXmlSource = "") And (strLoadPath = "") Then Exit Function
'
'  ' If the user specified a file path, use it
'  Set Return_XML = New MSXML.DOMDocument
'  GoSub CheckError
'  If FileExists(strLoadPath) = True Then
'    If Return_XML.Load(strLoadPath) = False Then
'      Return_ErrNum = Err.Number
'      Return_ErrDesc = Err.Description
'      If Return_ErrNum = 0 Then Return_ErrNum = -1
'      If Return_ErrDesc = "" Then Return_ErrDesc = "Load(strLoadPath) Failed"
'      If blnShowErrorMsgs = True Then _
'        MsgBox "xml_LoadXML_Ex() caused the following error:" & Chr(13) & Chr(13) & _
'               "Error Number = " & CStr(Return_ErrNum) & Chr(13) & _
'               "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Error " & CStr(Return_ErrNum)
'      Err.Clear
'      Exit Function
'    End If
'
'  ' If the user specified
'  Else
'    If Return_XML.loadXML(strXmlSource) = False Then
'      Return_ErrNum = Err.Number
'      Return_ErrDesc = Err.Description
'      If Return_ErrNum = 0 Then Return_ErrNum = -1
'      If Return_ErrDesc = "" Then Return_ErrDesc = "loadXML(strXmlSource) Failed"
'      If blnShowErrorMsgs = True Then _
'        MsgBox "xml_LoadXML_Ex() caused the following error:" & Chr(13) & Chr(13) & _
'               "Error Number = " & CStr(Return_ErrNum) & Chr(13) & _
'               "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Error " & CStr(Return_ErrNum)
'      Err.Clear
'      Exit Function
'    End If
'  End If
'
'  xml_LoadXML_Ex = True
'
'ExitOut:
'
'  Exit Function
'
'CheckError:
'
'  ' Check if an error occured
'  Return_ErrNum = Err.Number
'  Return_ErrDesc = Err.Description
'  If Return_ErrNum = 0 Then Return
'  Set Return_XML = Nothing
'
'  ' An error did occured
'  If blnShowErrorMsgs = True Then _
'    MsgBox "xml_LoadXML_Ex() caused the following error:" & Chr(13) & Chr(13) & _
'           "Error Number = " & CStr(Return_ErrNum) & Chr(13) & _
'           "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Error " & CStr(Return_ErrNum)
'  Err.Clear
'  Resume ExitOut
'
'End Function
'
'
''=============================================================================================================
'' xml_RS_to_XML
''
'' Purpose:
'' --------
'' This function takes an ADODB.Recordset object, or the database connection information to retrieve an
'' ADODB.Recordset object and converts it to XML.. filtering out unwanted nodes and sub-nodes.
''
'' NOTE:
'' -----
'' If *BOTH* the recordset and the DB connection information are passed, the recordset is taken over
'' the DB connection information (saves resources).  If *NIETHER* are passed, the function fails.
''
'' Param                Use
'' ------------------------------------
'' Return_XML           Returns the XML document in String form
'' rsRecordSet          Optional. Specifies the ADODB.Recordset object to retrieve the data from.  If this
''                      is not specified, bot the "strAdoConnString" and "strSqlStatement" must be specified.
'' strAdoConnString     Optional. Specifies the ADO connection string to use to connect to the database to
''                      retrieve data from.  If this is not specified, the "rsRecordSet" must be specified.
'' strSqlStatement      Optional. Specifies the SQL statement to execute against the ADO DB connection
''                      created by the "strAdoConnString" parameter. If this is not specified, you must specify
''                      the "rsRecordSet" parameter.
'' blnIsStoredProc      Optional. If the "rsRecordSet" is not used and this parameter is set to TRUE, the SQL
''                      statement passed to the "strSqlStatement" is treated as a Stored Procedure.
'' strFilterNodes       Optional. The recordset is converted to an XML document by the MSXML object. Once this
''                      occurs, you can specifies which nodes are to be retrieved from that XML document by
''                      specifying the Node names in this parameter, seperated by commas (,).
''                      * NOTE - Usually there are only 2 nodes created when a recordset is converted to XML...
''                        "s:Schema" which contains the XML document's schema, and "rs:data" which contains the
''                        recordset's records and the records' fields as node attributes.
'' strFilterSubNodes    Optional. The recordset is converted to an XML document by the MSXML object. Once this
''                      occurs, you can specify which nodes to retrieve from that XML document via the
''                      "strFilterNodes".  Using this parameter ("strFilterSubNodes") you can specify which
''                      recordset fields (node attributes) to return.
'' blnIncludeRoot       Optional. If set to TRUE, the root of the XML document is returned.  Otherwise it is
''                      left off.
'' blnMakeReadable      Optional. If set to TRUE, the XML code returned includes vbCrLf and vbTab characters
''                      to make the code more readable.  Otherwise they are left out to conserve file size.
'' blnShowErrorMsgs     Optional. If set to TRUE and an error occurs, an error message is displayed to the user.
'' Return_ErrNum        Optional. Returns the error number of any error that occured in this function.
'' Return_ErrDesc       Optional. Returns the error description of any error that occured in this function.
''
'' Return:
'' -------
'' Returns TRUE if the function executed successfully
'' Returns FALSE if the function failed
''
''=============================================================================================================
'Public Function xml_RS_to_XML(ByRef Return_XML As String, _
'                              Optional ByRef rsRecordSet As ADODB.Recordset, _
'                              Optional ByVal strAdoConnString As String, _
'                              Optional ByVal strSqlStatement As String, _
'                              Optional ByVal blnIsStoredProc As Boolean = False, _
'                              Optional ByVal strFilterNodes As String, _
'                              Optional ByVal strFilterSubNodes As String, _
'                              Optional ByVal blnIncludeRoot As Boolean = True, _
'                              Optional ByVal blnMakeReadable As Boolean = False, _
'                              Optional ByVal blnShowErrorMsgs As Boolean = False, _
'                              Optional ByRef Return_ErrNum As Long, _
'                              Optional ByRef Return_ErrDesc As String) As Boolean
'On Error Resume Next
'
'  Dim rsResults       As ADODB.Recordset
'  Dim conConnection   As ADODB.Connection
'  Dim xmlResults      As MSXML.DOMDocument
'  Dim lnodList        As MSXML.IXMLDOMNodeList
'  Dim rnodRoot        As MSXML.IXMLDOMElement
'  Dim cnodChild       As MSXML.IXMLDOMNode
'  Dim cnodSubChild    As MSXML.IXMLDOMNode
'  Dim attNodeAttrib   As MSXML.IXMLDOMAttribute
'  Dim lngOptions      As Long
'  Dim blnDestroyRS    As Boolean
'  Dim strFilters()    As String
'  Dim intFiltersCount As Integer
'  Dim strLineEnd      As String
'  Dim strTabChar      As String
'  Dim intCounter      As Integer
'
'  ' Clear the return variables
'  Return_XML = ""
'  Return_ErrNum = 0
'  Return_ErrDesc = ""
'  Erase strFiltersSub
'  intFiltersSubCount = 0
'
'  ' Add the processing instruction to the top of the XML document : <?xml version="1.0" ?>
'  ' (Note: This gets returned reguardless of whether the function succeeds or fails to make even a blank return valid XML)
'  Return_XML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " ?>"
'
'  ' Validate parameters passed
'  blnDestroyRS = True
'  strAdoConnString = Trim(strAdoConnString)
'  strSqlStatement = Trim(strSqlStatement)
'  strFilterNodes = Trim(strFilterNodes)
'  strFilterSubNodes = Trim(strFilterSubNodes)
'  If blnMakeReadable = True Then
'    strLineEnd = vbCrLf
'    strTabChar = vbTab
'  Else
'    strLineEnd = ""
'    strTabChar = " "
'  End If
'
'  ' If the user passed a recordset that is valid, use it.  If not get it based on the connection information passed.
'  ' Otherwise exit the function.
'  If rsRecordSet Is Nothing Then
'    If strAdoConnString = "" Then
'      Exit Function
'    Else
'      Set conConnection = New ADODB.Connection ' Setup the connection object
'      GoSub CheckError
'      conConnection.Open strAdoConnString ' Connect to the database
'      GoSub CheckError
'      lngOptions = adCmdText
'      If blnIsStoredProc = True Then lngOptions = adCmdStoredProc
'      Set rsResults = conConnection.Execute(strSqlStatement, , lngOptions) ' Retrieve the recordset from the DB
'      GoSub CheckError
'    End If
'  Else
'    blnDestroyRS = False
'    Set rsResults = rsRecordSet
'  End If
'
'  ' Check if the recordset is EMPTY
'  If rsResults.BOF = True And rsResults.EOF = True Then GoTo ExitOut
'  If rsResults.State = adStateClosed Then GoTo ExitOut
'  rsResults.MoveFirst
'
'  ' Create the XML object to recieve the recordset
'  Set xmlResults = New MSXML.DOMDocument
'  GoSub CheckError
'
'  ' Put the returned Recordset into XML to parse
'  rsResults.Save xmlResults, adPersistXML
'  rsResults.MoveFirst
'  GoSub CheckError
'
'  ' Get the filters passed
'  If ParseFilters(strFilters, intFiltersCount, strFilterNodes) = False Then GoTo ExitOut
'
'  ' Get the filters for sub-nodes
'  If ParseFilters(strFiltersSub, intFiltersSubCount, strFilterSubNodes) = False Then GoTo ExitOut
'
'  ' Get the XML root and loop through the root's attributes
'  Set rnodRoot = xmlResults.documentElement
'  GoSub CheckError
'  If blnIncludeRoot = True Then
'    Return_XML = Return_XML & strLineEnd & "<" & rnodRoot.nodeName & " "
'    For Each attNodeAttrib In rnodRoot.Attributes
'      Return_XML = Return_XML & strTabChar & attNodeAttrib.Name & "='" & attNodeAttrib.Text & "'" & strLineEnd
'    Next
'    Return_XML = Left(Return_XML, Len(Return_XML) - Len(strLineEnd)) ' Strip off the last cariage return
'    Return_XML = Return_XML & ">" & strLineEnd
'  End If
'
'  ' Loop through the root's children nodes
'  For Each cnodChild In rnodRoot.childNodes
'
'    ' Parse only nodes specified in filter(s)
'    If intFiltersCount > 0 Then
'      For intCounter = 1 To intFiltersCount
'        If UCase(strFilters(intCounter)) = UCase(cnodChild.nodeName) Then
'          Return_XML = Return_XML & strTabChar & "<" & cnodChild.nodeName
'          For Each attNodeAttrib In cnodChild.Attributes
'            Return_XML = Return_XML & " " & attNodeAttrib.Name & "='" & attNodeAttrib.Text & "'"
'          Next
'          Return_XML = Return_XML & ">" & strLineEnd
'          If cnodChild.hasChildNodes = True Then Return_XML = Return_XML & NodeLoop(cnodChild, strTabChar)
'          Return_XML = Return_XML & strTabChar & "</" & cnodChild.nodeName & ">" & strLineEnd
'        End If
'      Next
'
'    ' No filters, parse all nodes
'    Else
'      Return_XML = Return_XML & strTabChar & "<" & cnodChild.nodeName
'      For Each attNodeAttrib In cnodChild.Attributes
'        Return_XML = Return_XML & " " & attNodeAttrib.Name & "='" & attNodeAttrib.Text & "'"
'      Next
'      Return_XML = Return_XML & ">" & strLineEnd
'      If cnodChild.hasChildNodes = True Then Return_XML = Return_XML & NodeLoop(cnodChild, strTabChar)
'      Return_XML = Return_XML & strTabChar & "</" & cnodChild.nodeName & ">" & strLineEnd
'    End If
'  Next
'
'  ' End the XML root
'  If blnIncludeRoot = True Then Return_XML = Return_XML & "</" & rnodRoot.nodeName & ">"
'
'  ' Function executed successfully
'  xml_RS_to_XML = True
'
'ExitOut:
'
'  Erase strFilters
'  Erase strFiltersSub
'  intFiltersSubCount = 0
'  If blnDestroyRS = True Then
'    If Not rsResults Is Nothing Then
'      If rsResults.State <> adStateClosed Then rsResults.Close
'      Set rsResults = Nothing
'    End If
'  End If
'  If Not conConnection Is Nothing Then
'    If conConnection.State <> adStateClosed Then conConnection.Close
'    Set conConnection = Nothing
'  End If
'  Set xmlResults = Nothing
'  Set lnodList = Nothing
'  Set rnodRoot = Nothing
'  Set cnodChild = Nothing
'  Set cnodSubChild = Nothing
'  Set attNodeAttrib = Nothing
'  Exit Function
'
'CheckError:
'
'  ' Check if an error occured
'  Return_ErrNum = Err.Number
'  Return_ErrDesc = Err.Description
'  If Return_ErrNum = 0 Then Return
'
'  ' An error did occur
'  If blnShowErrorMsgs = True Then _
'    MsgBox "xml_RS_to_XML() caused the following error:" & Chr(13) & Chr(13) & _
'           "Error Number = " & CStr(Return_ErrNum) & Chr(13) & _
'           "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Error " & CStr(Return_ErrNum)
'  Err.Clear
'  Resume ExitOut
'
'End Function
'
'
''=============================================================================================================
'' xml_SaveXML
''
'' Purpose:
'' --------
'' This function takes the specified XML source and saves it to the specified save path.  This function works
'' with straight ASCII TEXT rather than MSXML objects so it is less error prone and more flexible.
''
'' Param                Use
'' ------------------------------------
'' strXmlSource         Specifies the XML source to save to the file (in string format)
'' strSavePath          Specifies the path to the file to save as.
'' blnOverwriteExisting Optional. If TRUE and the file specified exists, the file will automatically
''                      be overwritten.
'' blnPromptToOverwrite Optional. If the "blnOverwriteExisting" parameter is set to FALSE and this parameter
''                      is set to TRUE, the user will be prompted to overwrite the file.  Otherwise the file
''                      is NOT overwritten and the function fails.
''
'' Return:
'' -------
'' Returns TRUE if the function executed successfully
'' Returns FALSE if the function failed
''
''=============================================================================================================
'Public Function xml_SaveXML(ByVal strXmlSource As String, _
'                            ByVal strSavePath As String, _
'                            Optional ByVal blnOverwriteExisting As Boolean = True, _
'                            Optional ByVal blnPromptToOverwrite As Boolean = False) As Boolean
'On Error Resume Next
'
'  Dim MyAnswer    As VbMsgBoxResult
'  Dim intFileNum  As Integer
'
'  ' Validate parameters
'  strXmlSource = Trim(strXmlSource)
'  If strXmlSource = "" Then Exit Function
'  strSavePath = Trim(strSavePath)
'  If strSavePath = "" Then Exit Function
'
'  ' Check if the file exists already, if it does delete it
'  If FileExists(strSavePath) = True Then
'    If blnOverwriteExisting = False Then
'      If blnPromptToOverwrite = False Then
'        Exit Function
'      Else
'        MyAnswer = MsgBox(strSavePath & Chr(13) & "This file already exists." & Chr(13) & Chr(13) & "Overwrite existing file?", vbYesNo + vbExclamation, "  Confirm File Overwrite")
'        If MyAnswer <> vbYes Then Exit Function
'      End If
'    End If
'  End If
'
'  ' Delete any existing file
'  Kill strSavePath
'
'  ' Save the file out as a text
'  intFileNum = FreeFile
'  Open strSavePath For Output As #intFileNum
'    Print #intFileNum, strXmlSource
'  Close #intFileNum
'
'  ' Function executed successfully
'  xml_SaveXML = True
'
'End Function
'
'
''=============================================================================================================
'' xml_SaveXML_Ex
''
'' Purpose:
'' --------
'' This function takes the specified XML source and
''
'' NOTE : If BOTH the "strXmlSource" and "xmlSource" parameters are passed, the "xmlSource" parameter is used
'' If NEITHER parameter is passed, the function fails.
''
'' Param                Use
'' ------------------------------------
''
'' Return:
'' -------
'' Returns TRUE if the function executed successfully
'' Returns FALSE if the function failed
''
''=============================================================================================================
'Public Function xml_SaveXML_Ex(ByVal strSavePath As String, _
'                               Optional ByVal strXmlSource As String, _
'                               Optional ByRef xmlSource As MSXML.DOMDocument, _
'                               Optional ByVal blnOverwriteExisting As Boolean = True, _
'                               Optional ByVal blnPromptToOverwrite As Boolean = False, _
'                               Optional ByVal blnShowErrorMsgs As Boolean = False, _
'                               Optional ByRef Return_ErrNum As Long, _
'                               Optional ByRef Return_ErrDesc As String) As Boolean
'On Error Resume Next
'
'  Dim xmlSave       As MSXML.DOMDocument
'  Dim MyAnswer      As VbMsgBoxResult
'
'  ' Erase the return variables
'  Return_ErrNum = 0
'  Return_ErrDesc = ""
'
'  ' Validate parameters
'  strXmlSource = Trim(strXmlSource)
'  strSavePath = Trim(strSavePath)
'  If strSavePath = "" Then Exit Function
'  If (strXmlSource = "") And (xmlSource Is Nothing) Then Exit Function
'
'  ' Check if the file exists already, if it does delete it
'  If FileExists(strSavePath) = True Then
'    If blnOverwriteExisting = False Then
'      If blnPromptToOverwrite = False Then
'        Exit Function
'      Else
'        MyAnswer = MsgBox(strSavePath & Chr(13) & "This file already exists." & Chr(13) & Chr(13) & "Overwrite existing file?", vbYesNo + vbExclamation, "  Confirm File Overwrite")
'        If MyAnswer <> vbYes Then Exit Function
'      End If
'    End If
'  End If
'
'  ' The user passed the XML as an object
'  If Not xmlSource Is Nothing Then
'    xmlSource.Save strSavePath
'    GoSub CheckError
'
'  ' The user passed the XML as a string
'  Else
'
'    ' Create an XML object to save the file
'    Set xmlSave = New MSXML.DOMDocument
'    GoSub CheckError
'
'    ' Load the specified XML source into the XML object
'    If xmlSave.loadXML(strXmlSource) = False Then
'      Return_ErrNum = Err.Number
'      Return_ErrDesc = Err.Description
'      If Return_ErrNum = 0 Then Return_ErrNum = -1
'      If Return_ErrDesc = "" Then Return_ErrDesc = "loadXML(strXmlSource) Failed"
'      If blnShowErrorMsgs = True Then _
'        MsgBox "xml_SaveXML_Ex() caused the following error:" & Chr(13) & Chr(13) & _
'               "Error Number = " & CStr(Return_ErrNum) & Chr(13) & _
'               "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Error " & CStr(Return_ErrNum)
'      Err.Clear
'      If Not xmlSave Is Nothing Then Set xmlSave = Nothing
'      Exit Function
'    End If
'
'    ' Delete any existing file
'    Kill strSavePath
'
'    ' Use XML to save the file out (XML adds some formatting)
'    xmlSave.Save strSavePath
'
'    If Not xmlSave Is Nothing Then Set xmlSave = Nothing
'  End If
'
'  ' Function executed successfully
'  xml_SaveXML_Ex = True
'
'ExitOut:
'
'  Exit Function
'
'CheckError:
'
'  ' Check if an error occured
'  Return_ErrNum = Err.Number
'  Return_ErrDesc = Err.Description
'  If Return_ErrNum = 0 Then Return
'
'  ' An error did occured
'  If blnShowErrorMsgs = True Then _
'    MsgBox "xml_SaveXML_Ex() caused the following error:" & Chr(13) & Chr(13) & _
'           "Error Number = " & CStr(Return_ErrNum) & Chr(13) & _
'           "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Error " & CStr(Return_ErrNum)
'  Err.Clear
'  If Not xmlSave Is Nothing Then Set xmlSave = Nothing
'  Resume ExitOut
'
'End Function
'
'
''=============================================================================================================
'' xml_XML_to_RS
''
'' Purpose:
'' --------
'' This function takes an XML document and converts the specified node collection to a recordset and returns
'' an ADODB.Recordset object.
''
'' NOTE : The way this module works is it takes the node name specified by the "strNodeName" parameter and
''        treats that like a recordset.  This function will loop through the entire XML document looking
''        for occurances of that node name, and if one is found will add another record to the recordset.
''        The sub-nodes below the nodes (to the first level sub-nodes only) found with node names that match
''        one of the fields specified in the "strFieldNames" parameter (commad delimited) will be added to
''        the recordset.  If the "blnSearchAttrib" parameter is set to TRUE, the attributes of the
''        recordset nodes will be searched for the fields matching the ones specified in the "strFieldNames"
''        parameter instead of sub-nodes (sub-nodes won't be searched in this case).
''
'' NOTE : Specifying "BOOK" as the strNodeName parameter with fields specified in the strFieldNames
''        parameter will return a recordset representing all nodes in the specified XML document with
''        the name "BOOK".  The sub-nodes or attributes (depending on the blnSearchAttrib parameter's
''        setting) of those nodes which match the field names specified in the strFieldNames parameter
''        will be returned as the field values of the records returned in the recordset.  Likewise,
''        specifying "BOOK/AUTHOR" as the strNodeName parameter will return all of the AUTHOR child
''        nodes with a BOOK parent.
''
'' NOTE : If BOTH the "strXmlSource" and "strLoadPath" parameters are specified, the "strLoadPath"
''        parameter is used.  If NEITHER are specified, the function fails.
''
'' NOTE : If there are no child nodes matching the one specified in the "strNodeName" parameter, the
''        function fails.
''
'' NOTE : If nodes are found but none of those node's children have names that match any of the fields
''        specified in the "strFieldNames" parameter, a blank ADODB.Recordset object is returned
''        (no records in the recordset)
''
'' Param                Use
'' ------------------------------------
'' Return_Recordset     Returns an ADODB.Recordset object that represents the data found in the XML that
''                      matched the specified parameters.
'' Return_RecordCount   Returns the number of records returned in the "Return_Recordset" parameter
'' strNodeName          Optional. Specifies the name of the node to search for in the XML document and make
''                      a recordset from.  YOu can use a "file path" approach with this parameter by passing
''                      "PARENT/CHILD/SUBCHILD" to access specific nodes that you know the parent(s) of.
'' strFieldNames        Optional. This is a comma delimited list of node names to look for under the node
''                      specified in the "strNodeName" parameter.  This list will turn into the fields (columns)
''                      of the recordset object.
'' strXmlSource         Optional. Specifies the XML document in String form to be converted.  If this
''                      parameter is not specified, then the "strLoadPath" parameter MUST be specified.
'' strLoadPath          Optional. Specifies the path to the XML document to load and convert to a Recordset.
''                      If this parameter is not specified, the "strXmlSource" parameter MUST be specified.
'' blnSearchAttrib      Optional. If set to TRUE, the attributes of the nodes specified in the "strNodeName"
''                      parameter will be searched instead of the sub-nodes.
'' blnShowErrorMsgs     Optional. If set to TRUE and an error occurs, an error message will be shown to the user.
'' Return_ErrNum        Optional. Returns the error number of any error that occured in the function.
'' Return_ErrDesc       Optional. Returns the error description of any error that occured in the function.
''
'' Return:
'' -------
'' Returns TRUE if the function executed successfully
'' Returns FALSE if the function failed
''
''=============================================================================================================
'Public Function xml_XML_to_RS(ByRef Return_Recordset As ADODB.Recordset, _
'                              ByRef Return_RecordCount As Integer, _
'                              ByVal strNodeName As String, _
'                              ByVal strFieldNames As String, _
'                              Optional ByVal strXmlSource As String, _
'                              Optional ByVal strLoadPath As String, _
'                              Optional ByVal blnSearchAttrib As Boolean = False, _
'                              Optional ByVal blnShowErrorMsgs As Boolean = False, _
'                              Optional ByRef Return_ErrNum As Long, _
'                              Optional ByRef Return_ErrDesc As String) As Boolean
'On Error Resume Next
'
'  Dim xmlResults      As MSXML.DOMDocument
'  Dim lnodList        As MSXML.IXMLDOMNodeList
'  Dim cnodListItem    As MSXML.IXMLDOMNode
'  Dim cnodSubNode     As MSXML.IXMLDOMNode
'  Dim attNodeAttrib   As MSXML.IXMLDOMAttribute
'  Dim strFields()     As String
'  Dim intFieldsCount  As Integer
'  Dim intCounter      As Integer
'
'  ' Erase the return variables
'  Return_ErrNum = 0
'  Return_ErrDesc = ""
'  Return_RecordCount = 0
'  Set Return_Recordset = Nothing
'
'  ' Validate parameters
'  strXmlSource = Trim(strXmlSource)
'  strLoadPath = Trim(strLoadPath)
'  strNodeName = Trim(strNodeName)
'  strFieldNames = Trim(strFieldNames)
'
'  ' Make sure at least one XML source is specified
'  If strXmlSource = "" And strLoadPath = "" Then Exit Function
'
'  ' Make sure there is a node name and field name(s) specified
'  If strNodeName = "" Or strFieldNames = "" Then Exit Function
'
'  ' Get the filters passed
'  If ParseFilters(strFields, intFieldsCount, strFieldNames) = False Then GoTo ExitOut
'
'  ' Load the XML into an object I can retrieve the data from easily
'  Err.Clear
'  If xml_LoadXML_Ex(xmlResults, strXmlSource, strLoadPath, blnShowErrorMsgs, Return_ErrNum, Return_ErrDesc) = False Then GoTo ExitOut
'
'  ' Get a list of all the nodes that match the specified strNodeName parameter
' 'Set lnodList = xmlResults.documentElement.selectNodes("*/" & strNodeName)
'  Set lnodList = xmlResults.documentElement.getElementsByTagName(strNodeName)
'  GoSub CheckError
'
'  ' If there are no nodes matching the one specified in the strNodeNames, exit the function
'  If lnodList.length = 0 Then GoTo ExitOut
'
'  ' Create the recordset that will store the information
'  Set Return_Recordset = New ADODB.Recordset
'  GoSub CheckError
'
'  ' Add the fields specified to the Recordset object
'  For intCounter = 1 To intFieldsCount
'    Return_Recordset.Fields.Append strFields(intCounter), adBSTR, 150
'  Next
'
'  ' Open the Recordset for use
'  Return_Recordset.Open
'
'  ' Loop through the nodes returned in the nodelist and store
'  For Each cnodListItem In lnodList
'
'    ' Add a record to the recordset for each node found
'    Return_Recordset.AddNew
'
'    ' Search the attributes of the nodes, not their sub-nodes
'    If blnSearchAttrib = True Then
'      For Each attNodeAttrib In cnodListItem.Attributes
'        For intCounter = 1 To intFieldsCount
'          If UCase(Trim(attNodeAttrib.Name)) = UCase(Trim(strFields(intCounter))) Then
'            Return_Recordset(Trim(strFields(intCounter))).Value = attNodeAttrib.Text
'          End If
'        Next
'      Next
'
'    ' Search the sub-nodes of the nodes, not their attributes
'    Else
'      For Each cnodSubNode In cnodListItem.childNodes
'        For intCounter = 1 To intFieldsCount
'          If UCase(Trim(cnodSubNode.nodeName)) = UCase(Trim(strFields(intCounter))) Then
'            Return_Recordset(Trim(strFields(intCounter))).Value = cnodSubNode.Text
'            Exit For
'          End If
'        Next
'      Next
'    End If
'  Next
'
'  ' Function executed successfully
'  Return_Recordset.MoveFirst
'  Return_RecordCount = lnodList.length
'  xml_XML_to_RS = True
'
'ExitOut:
'
'  Erase strFields
'  Set xmlResults = Nothing
'  Set lnodList = Nothing
'  Set cnodListItem = Nothing
'  Set cnodSubNode = Nothing
'  Set attNodeAttrib = Nothing
'
'  Exit Function
'
'CheckError:
'
'  ' Check if an error occured
'  Return_ErrNum = Err.Number
'  Return_ErrDesc = Err.Description
'  If Return_ErrNum = 0 Then Return
'  Set Return_Recordset = Nothing
'  Return_RecordCount = 0
'
'  ' An error did occur
'  If blnShowErrorMsgs = True Then _
'    MsgBox "xml_XML_to_RS() caused the following error:" & Chr(13) & Chr(13) & _
'           "Error Number = " & CStr(Return_ErrNum) & Chr(13) & _
'           "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Error " & CStr(Return_ErrNum)
'  Err.Clear
'  Resume ExitOut
'
'End Function
'
'
'
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'
'
'
'' Checks to see if the specified file exists
'Private Function FileExists(ByVal FilePath As String) As Boolean
'On Error GoTo ErrorTrap
'
'  Dim intFileNum As Integer
'
'  ' Validate parameter
'  FilePath = Trim(FilePath)
'  If FilePath = "" Then Exit Function
'
'  ' Get available file number and try to open the file.  If opens as READ-ONLY ok, then the file exists.
'  intFileNum = FreeFile
'  Open FilePath For Input As #intFileNum
'  Close #intFileNum
'
'  ' File exists
'  FileExists = True
'  Exit Function
'
'ErrorTrap:
'
'  ' File doesn't exist, or error opening it
'  Close #intFileNum
'  Err.Clear
'
'End Function
'
'' Loops through the specified node and gets it's nodes.  This function is used by the "xml_RS_to_XML" function
'Private Function NodeLoop(ByRef cnodCurrentNode As MSXML.IXMLDOMNode, _
'                          ByVal strTabChar As String) As String
'On Error Resume Next
'
'  Dim cnodChild       As MSXML.IXMLDOMNode
'  Dim attChildAttribs As MSXML.IXMLDOMAttribute
'  Dim intCounter    As Integer
'
'  ' Validate parameter
'  If cnodCurrentNode Is Nothing Then Exit Function
'  If cnodCurrentNode.hasChildNodes = False Then Exit Function
'
'  ' Loop through the children nodes of the passed node
'  For Each cnodChild In cnodCurrentNode.childNodes
'
'    ' Parse only nodes specified in filter(s)
'    If intFiltersSubCount > 0 Then
'      For intCounter = 1 To intFiltersSubCount
'        If UCase(strFiltersSub(intCounter)) = UCase(cnodChild.nodeName) Then
'          NodeLoop = NodeLoop & strTabChar & "<" & cnodChild.nodeName
'          For Each attChildAttribs In cnodChild.Attributes
'            NodeLoop = NodeLoop & " " & attChildAttribs.Name & "='" & attChildAttribs.Text & "'"
'          Next
'          NodeLoop = NodeLoop & ">" & vbCrLf
'          If cnodChild.hasChildNodes = True Then NodeLoop = NodeLoop & NodeLoop(cnodChild, strTabChar)
'          NodeLoop = NodeLoop & strTabChar & "</" & cnodChild.nodeName & ">" & vbCrLf
'        End If
'      Next
'
'    ' No filters, parse all nodes
'    Else
'      NodeLoop = NodeLoop & strTabChar & "<" & cnodChild.nodeName
'      For Each attChildAttribs In cnodChild.Attributes
'        NodeLoop = NodeLoop & " " & attChildAttribs.Name & "='" & attChildAttribs.Text & "'"
'      Next
'      NodeLoop = NodeLoop & ">" & vbCrLf
'      If cnodChild.hasChildNodes = True Then NodeLoop = NodeLoop & NodeLoop(cnodChild, strTabChar)
'      NodeLoop = NodeLoop & strTabChar & "</" & cnodChild.nodeName & ">" & vbCrLf
'    End If
'  Next
'
'  Set cnodChild = Nothing
'  Set attChildAttribs = Nothing
'
'End Function
'
'' This function goes through a comma delimeted string and returns an array of strings that
'' represent each of the items passed.
''
'' NOTE : The returned array is 1 based, not the default 0 based.  This makes looping easier.
'Private Function ParseFilters(ByRef Return_Filters() As String, _
'                              ByRef Return_Count As Integer, _
'                              ByVal strFilters As String) As Boolean
'On Error Resume Next
'
'  Dim intCounter   As Integer
'  Dim strCharLeft    As String
'  Dim strCharRight   As String
'  Dim strStringSoFar As String
'
'  ' Reset the return variables passed
'  Return_Count = 0
'  Erase Return_Filters
'
'  ' Verify parameters are valid
'  strFilters = Trim(strFilters)
'  If strFilters = "" Then ' No filters
'    ParseFilters = True
'    Exit Function
'  End If
'  If InStr(strFilters, ",") = 0 Then ' Only 1 filter
'    Return_Count = 1
'    ReDim Preserve Return_Filters(1 To 1) As String
'    Return_Filters(1) = strFilters
'    ParseFilters = True
'    Exit Function
'  End If
'
'  ' Loop through the filters passed and seperate them out
'  For intCounter = 1 To Len(strFilters)
'    strCharLeft = Left(strFilters, intCounter)
'    strCharRight = Right(strCharLeft, 1)
'    Select Case strCharRight
'      Case Chr(13), Chr(10), " " ' Take these out
'      Case ","                   ' Field seperator
'        Return_Count = Return_Count + 1
'        ReDim Preserve Return_Filters(1 To Return_Count) As String
'        Return_Filters(Return_Count) = Trim(strStringSoFar)
'        strStringSoFar = ""
'      Case Else                  ' Everything else
'        strStringSoFar = strStringSoFar & strCharRight
'    End Select
'  Next
'
'  ' Get the last one, because normally there isn't a comma at the end of the last filter
'  If Right(strFilters, 1) <> "," Then
'    Return_Count = Return_Count + 1
'    ReDim Preserve Return_Filters(1 To Return_Count) As String
'    Return_Filters(Return_Count) = Trim(strStringSoFar)
'    strStringSoFar = ""
'  End If
'
'  ' Function completed successfully
'  ParseFilters = True
'
'End Function
'
'
