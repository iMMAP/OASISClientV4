Attribute VB_Name = "DBfuncions"
Option Explicit

Public Sub CompactDB()
'    'Microsoft Jet and Replication objects
'    Dim objJE As New JRO.JetEngine, strSource As String, strTarget As String
'
'    DoEvents
'
'    strSource = " "
'    strTarget = " "
'    objJE.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSource & ";", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strTarget & ";Jet OLEDB:Engine Type=4;"
'
'    'Engine type:
'    'Access 97 = 4
'    'Access 2000 = 5
End Sub

Public Sub SendLotusMail()
'    Dim Maildb As Object
'    Dim MailDoc As Object
'    Dim Body As Object
'    Dim Session As Object
'    'Start a session to notes
'    Set Session = CreateObject("Lotus.NotesSession")
'    'This line prompts for password of current ID noted in Notes.INI
'    Call Session.Initialize
'    'or use below to supply password of the current ID
'    'Call Session.Initialize("")
'    'Open the mail database in notes
'    Set Maildb = Session.GETDATABASE("", "c:\notes\data\mail\mymail.nsf")
'
'    If Not Maildb.IsOpen = True Then
'        Call Maildb.Open
'    End If
'
'    'Create the mail document
'    Set MailDoc = Maildb.CREATEDOCUMENT
'    Call MailDoc.ReplaceItemValue("Form", "Memo")
'    'Set the recipient
'    Call MailDoc.ReplaceItemValue("SendTo", "John Doe")
'    'Set subject
'    Call MailDoc.ReplaceItemValue("Subject", "Subject Text")
'    'Create and set the Body content
'    Set Body = MailDoc.CREATERICHTEXTITEM("Body")
'    Call Body.APPENDTEXT("Body text here")
'    'Example to create an attachment (optional)
'    Call Body.ADDNEWLINE(2)
'    Call Body.EMBEDOBJECT(1454, "", "C:\filename", "Attachment")
'    'Example to save the message (optional)
'    MailDoc.SAVEMESSAGEONSEND = True
'    'Send the document
'    'Gets the mail to appear in the Sent items folder
'    Call MailDoc.ReplaceItemValue("PostedDate", Now())
'    Call MailDoc.Send(False)
'    'Clean Up
'    Set Maildb = Nothing
'    Set MailDoc = Nothing
'    Set Body = Nothing
'    Set Session = Nothing
'
'    Note: The Visual Basic programmer needs to set the Reference to use Lotus Domino objects prior to implementing this function. To enable the Lotus Notes classes to appear in the Visual Basic browser, you must execute the following within VB: Select Tools, References and select the checkbox for 'Lotus Notes Automation Classes'.
'
'    The above code is from the IBM support. GETDATABASE given here is pointing to the sample MailDB; you need to change that to your DB.
'
'    You can do that by
'
'    UserName = Session.UserName
'    MailDbName = Left$(UserName, 1) & Right$(UserName, (Len(UserName) - InStr(1, UserName, " "))) & ".nsf"
'    'Open the mail database in notes
'    Set Maildb = Session.GETDATABASE("", MailDbName)
End Sub


Public Function SaveFileToDB(ByVal Filename As String, _
                             rs As Object, _
                             FieldName As String) As Boolean
    '**************************************************************
    'PURPOSE: SAVES DATA FROM BINARY FILE (e.g., .EXE, WORD DOCUMENT
    'CONTROL TO RECORDSET RS IN FIELD NAME FIELDNAME

    'FIELD TYPE MUST BE BINARY (OLE OBJECT IN ACCESS)

    'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

    'SAMPLE USAGE
    'Dim sConn As String
    'Dim oConn As New ADODB.Connection
    'Dim oRs As New ADODB.Recordset
    '
    '
    'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
    '
    'oConn.Open sConn
    'oRs.Open "SELECT * FROM MYTABLE", oConn, adOpenKeyset, _
     adLockOptimistic
    'oRs.AddNew

    'SaveFileToDB "C:\MyDocuments\MyDoc.Doc", oRs, "MyFieldName"
    'oRs.Update
    'oRs.Close
    '**************************************************************

    Dim iFileNum As Integer
    Dim lFileLength As Long

    Dim abBytes() As Byte
    Dim iCtr As Integer

    On Error GoTo ERRORHANDLER

    If Dir(Filename) = "" Then Exit Function
    If Not TypeOf rs Is ADODB.Recordset Then Exit Function

    'read file contents to byte array
    iFileNum = FreeFile
    Open Filename For Binary Access Read As #iFileNum
    lFileLength = LOF(iFileNum)
    ReDim abBytes(lFileLength)
    Get #iFileNum, , abBytes()

    'put byte array contents into db field
    rs.Fields(FieldName).AppendChunk abBytes()
    Close #iFileNum

    SaveFileToDB = True
ERRORHANDLER:
End Function

Public Function LoadFileFromDB(Filename As String, _
                               rs As Object, _
                               FieldName As String) As Boolean
    '************************************************
    'PURPOSE: LOADS BINARY DATA IN RECORDSET RS,
    'FIELD FieldName TO a File Named by the FileName parameter

    'REQUIRES: REFERENCE TO MICROSOFT ACTIVE DATA OBJECTS 2.0 or ABOVE

    'SAMPLE USAGE
    'Dim sConn As String
    'Dim oConn As New ADODB.Connection
    'Dim oRs As New ADODB.Recordset
    '
    '
    'sConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyDb.MDB;Persist Security Info=False"
    '
    'oConn.Open sConn
    'oRs.Open "SELECT * FROM MyTable", oConn, adOpenKeyset,
    ' adLockOptimistic
    'LoadFileFromDB "C:\MyDocuments\MyDoc.Doc",  oRs, "MyFieldName"
    'oRs.Close
    '************************************************
    Dim iFileNum As Integer
    Dim lFileLength As Long
    Dim abBytes() As Byte
    Dim iCtr As Integer

    On Error GoTo ERRORHANDLER

    If Not TypeOf rs Is ADODB.Recordset Then Exit Function

    iFileNum = FreeFile
    Open Filename For Binary As #iFileNum
    lFileLength = LenB(rs(FieldName))

    abBytes = rs(FieldName).GetChunk(lFileLength)
    Put #iFileNum, , abBytes()
    Close #iFileNum
    LoadFileFromDB = True

ERRORHANDLER:
End Function

Private Function FieldIsString(FieldObject As ADODB.Field) As Boolean

    'Input = ADODB.Field Object
    
    'EXAMPLE USAGE
    'After connecting to data source via ADO.
    'Dim myRS As ADODB.Recordset
    'Dim bIsString as boolean
    'Set myRS.ActiveConnection = myADOConnection
    'myRS.Open "SELECT * FROM MyTABLE"
    'bIsString = FieldIsString(myRS.Fields(0))
    
    'could raise an error here

    If Not TypeOf FieldObject Is ADODB.Field Then Exit Function

    Select Case FieldObject.Type

        Case adBSTR, adChar, adVarChar, adWChar, adVarWChar, adLongVarChar, adLongVarWChar
            FieldIsString = True

        Case Else
            FieldIsString = False
    End Select
        
End Function

Public Function TableExists(DatabaseName As String, _
                            TableName As String) As Boolean

'    'DataBaseName is the file/path name of the database
'    'with the field you want to test
'    'tablename is the table which you want to test
'    'if database doesn't exist, an error is raised
'
'    Dim oDB As Database, td As TableDef
'
'    On Error GoTo ErrorHandler
'    Set oDB = Workspaces(0).OpenDatabase(DatabaseName)
'    On Error Resume Next
'
'    Set td = oDB.TableDefs(TableName)
'    TableExists = Err.Number = 0
'    oDB.Close
'
'    Exit Function
'
'ErrorHandler:
'
'    Err.Raise Err.Number
'    Exit Function

End Function

Public Function FieldExists(DatabaseName As String, _
                            TableName As String, _
                            FieldName As String) As Boolean

'    'DataBaseName is the file/path name of the database
'    'with the field you want to test
'    'tablename is the table, fieldname is the field
'    'if database or table does not exist, an error is raised
'
'    Dim oDB As Database
'    Dim td As TableDef
'    Dim f As Field
'
'    On Error GoTo ErrorHandler
'
'    Set oDB = Workspaces(0).OpenDatabase(DatabaseName)
'    Set td = oDB.TableDefs(TableName)
'
'    On Error Resume Next
'    Set f = td.Fields(FieldName)
'    FieldExists = Err.Number = 0
'    oDB.Close
'
'    Exit Function
'
'ErrorHandler:
'
'    If Not oDB Is Nothing Then oDB.Close
'    Err.Raise Err.Number
'    Exit Function

End Function


Public Function prepStringForSQL(ByVal sValue As String) _
As String

Dim sAns As String
sAns = Replace(sValue, Chr(39), "''")
sAns = "'" & sAns & "'"
prepStringForSQL = sAns


End Function

Public Function SynchProfileSettingWithServer(sSettingValue As String, _
                                              sUserGroupName As String, oCnn As ADODB.Connection, Optional sAltLookUpField As String = "")
        '<EhHeader>
        On Error GoTo SynchProfileSettingWithServer_Err
        '</EhHeader>

        Dim RSOasisServer As New ADODB.Recordset
        Dim RSOasisClient As New ADODB.Recordset
        Dim sString As String
         
100     'sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT " & sSettingValue & " FROM " & sUserGroupName & "AppSettings WHERE SettingName = '" & IIf(Len(sAltLookUpField) = 0, "ProfileSettings", sAltLookUpField) & "'")
102     'Set RSOasisServer = OpenSilentHttpCommsRS(sString, True)
        Set RSOasisServer = OpenServerRSCompressed(g_sAppServerPath & "/oasis4.asp", "id", "SELECT " & sSettingValue & " FROM " & sUserGroupName & "AppSettings WHERE SettingName = '" & IIf(Len(sAltLookUpField) = 0, "ProfileSettings", sAltLookUpField) & "'")
        
104     SynchProfileSettingWithServer = False
                    
106     If SafeMoveFirst(RSOasisServer) Then
                        
108         With RSOasisClient

110             .Open "SELECT * FROM AppSettings", oCnn, adOpenDynamic, adLockBatchOptimistic

112             If Not .State = adStateClosed Then
            
                    If Len(sAltLookUpField) = 0 Then
                         .Find "SettingName = 'ProfileSettings'"
                    Else
                        .Find "SettingName = '" & sAltLookUpField & "'"
                    End If
            
116                 If Not .EOF Then

118                     If Not IsNull(RSOasisServer.Fields.Item(sSettingValue).Value) Then
120                         .Fields(sSettingValue).Value = RSOasisServer.Fields.Item(sSettingValue).Value
                        Else
122                         .Fields(sSettingValue).Value = 0
                        End If
        
124                     .UpdateBatch adAffectCurrent
126                     SynchProfileSettingWithServer = True
                    
                    End If

128                 .Close
                
                End If

            End With

130         RSOasisServer.Close

        End If
    
132     Set RSOasisServer = Nothing
134     Set RSOasisClient = Nothing

        '<EhFooter>
        Exit Function

SynchProfileSettingWithServer_Err:
        'MsgBox Err.Description & vbCrLf & _
               "in OASISClient.DBfuncions.SynchProfileSettingWithServer " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function UpdateAppSettingsValue(sSetName As String, _
                                       sValueField As String, _
                                       sValue As String) As Boolean
        '<EhHeader>
        On Error GoTo UpdateAppSettingsValue_Err
        '</EhHeader>
    
        Dim RSUpdater As ADODB.Recordset
100     UpdateAppSettingsValue = True
        
102     Set RSUpdater = New ADODB.Recordset
104     With RSUpdater
            
106         .Open "SELECT * FROM AppSettings", m_Cnn, adOpenDynamic, adLockBatchOptimistic
        
108         If Not sValue = "" Then
110             .Find "SettingName = '" & sSetName & "'"
            Else
112             .Find "SettingName = '" & vbNullString & "'"
            End If

114         If Not .EOF Then
116             .Fields(sValueField).Value = sValue
118             .UpdateBatch adAffectCurrent
120             .Close
            End If

        End With
122     Set RSUpdater = Nothing
        
        '<EhFooter>
        Exit Function

UpdateAppSettingsValue_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.DBfuncions.UpdateAppSettingsValue " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
