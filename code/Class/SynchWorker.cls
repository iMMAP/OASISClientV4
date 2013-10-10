VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SynchWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements Runnable

Private Declare Sub SleepAPI _
                Lib "kernel32" _
                Alias "Sleep" (ByVal dwMilliseconds As Long)
    
Public Event Complete()
Public Event CompleteEX(enmRunnerType As RunnerType)
Public Event Cancelled()
Public Event Status(ByVal i As Long, ByRef Cancel As Boolean)
Public Event StatusEX(ByVal i As Long, ByVal lTotal As Long, ByRef Cancel As Boolean, sMess As String)
Public Event WorkerError(ErrNum As Long, errDesc As String, errSource As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''
Const HTTPREQUEST_PROXYSETTING_PROXY = 2
     
Private RemoteTables As DDTableCollection
Private LocalTables As DDTableCollection
Private bAbortSynchronisation As Boolean
Private sDebugInfo As String
Private lTimeDiffFromServerInSeconds As Long
Private bBatchDownloadOK As Boolean
Private bIsDownloadingBatch As Boolean
Private bLargePolygonDownloadFailed As Boolean
Private sTablesCreated As String
Private m_sWhen As String
Private m_bDebug As Boolean
Private m_bIncidentsV2 As Boolean
Private m_lCountDownloaded As Long
Private bRecordsDownloaded As Boolean

Private OASStringCompression As New OASISStringCompression.OASISCompression

Private Enum SynchNextOperation
    NewOnly = 1
    OldOnly = 2

    DoNotSynch = 3
    EditsOnly = 4
    DeleteOnly = 5
End Enum

Private Enum DownloadOperation
    NewRecords = 1
    NewUpdates = 2
    NewDeletes = 3
End Enum

Private Type SynchOperations
        
     sTableNameClient As String
     sTableNameServer As String
     sSynchHistoryClient As String
     sSynchHistoryServer As String
     sSearchRecordParams As String
     sSearchSynchHistoryParams As String
     ToAddLocal As String
     ToAddServer As String
     ToUpdateLocal As String
     ToUpdateServer As String
     ToDeleteLocal As String
     ToDeleteServer As String
        
End Type

Private Type DownloadUploadStats

    AllDone As Boolean
    NextSynchOperation As SynchNextOperation
    
    sWhenNow As String
    lCountLocal As Long
    sWhenMaxActualLocal As String
    sWhenMinActualLocal As String
    lCountServer As Long
    sWhenMinActualServer As String
    sWhenMaxActualServer As String
    sWhenMinSearched As String
    sWhenMaxSearched As String
    
    lBatchSizeLast As Long
    lBatchSizeNext As Long
    sTablenameLocal As String
    sTableNameServer As String
    sSynchHistoryTableLocal As String
    sSynchHistoryTableServer As String
    lDeletesOnServer As Long
    lDeletesOnClient As Long
    lDeletesPendingOnClient As Long
    ' lCountLocal As Long
    ' NextSynchNextOpeation As SynchNextOpeation
    
End Type

Private Type DDColumn

    Name As String
    Ordinal As Long
    TypeEnum As Long
    Description As String
    MaxLength As Long
        
End Type

Private Type DDTable

    Name As String
    Columns() As DDColumn
    IsTable As Boolean
    IsLinkedTable As Boolean
    IsQuery As Boolean
    'sAutoNum As String
            
End Type

Private Type DDQuery

    Name As String
    sSQL As String
    bNew As Boolean
    bEdited As Boolean
    bDelete As Boolean
    
End Type

Private Type DDTableCollection

    Prefix As String
    Tables() As DDTable
            
End Type

Private Type DDQueryCollection

    Prefix As String
    Queries() As DDQuery
            
End Type

'''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Enum RunnerType
    enmPaddy = 0
    IncidentSynch = 1
    SQLLyrSynch = 2
    GeoMarks = 3
    InternetConnectionCheck = 4
    IncidentNotifier = 5
    IncidentDeleted = 6
    IncidentEdited = 7
    GeoMarksDeleted = 8
    GeoMarksEdited = 9
End Enum

Private Enum SynchOperation
    Add = 0
    Delete = 1
    Update = 2
End Enum

Private m_lInterval As Long
Private m_bRunning As Boolean
Private m_bPaddy As Boolean
Private m_RType As RunnerType
Private m_StrInternetConnectionType As String
Private m_bConnected As Boolean
Private m_StrWebsiteURL As String
Private m_StrLocalConnectionString As String
Private m_StrRemoteTablePrefix As String

Private m_ColNewIncidents As Collection
Private m_ColDeletedIncidents As Collection
Private m_ColEditedIncidents As Collection

Private m_ColNewSynchTableRecords As Collection
Private m_ColDeletedSynchTableRecords As Collection
Private m_ColEditedSynchTableRecords As Collection

Private m_StrSynchTable As String
Private m_StrGUIDField As String
Private WithEvents oServerInterComms As OASISInterComm.IServer
Attribute oServerInterComms.VB_VarHelpID = -1
Private m_frmHiddenSynchWorker As frmHiddenSynchWorker
Private m_sIDs As String
Private m_bUseEvents As Boolean
Private m_bSynchGeoMarks As Boolean
Private m_bSynchSynchLayers As Boolean

Private m_bProxyEnabled As Boolean
Private m_sProxyIP As String
Private m_sProxyPort As String

Private bSQLServerInUse As Boolean
Private m_RSttkGISLayerSQLLocal As ADODB.Recordset
Private m_RSttkGISLayerSQLRemote As ADODB.Recordset

Public Property Get EnableGeoMarkSynch() As Variant
    EnableGeoMarkSynch = m_bSynchGeoMarks
End Property

Public Property Let EnableGeoMarkSynch(ByVal vNewValue As Variant)
    m_bSynchGeoMarks = vNewValue
End Property

Public Property Get EnableSynchTablesSynch() As Variant
    EnableSynchTablesSynch = m_bSynchSynchLayers
End Property

Public Property Let EnableSynchTablesSynch(ByVal vNewValue As Variant)
    m_bSynchSynchLayers = vNewValue
End Property

Public Property Get UseEvents() As Boolean
    UseEvents = m_bUseEvents
End Property

Public Property Let UseEvents(ByVal bValue As Boolean)
    m_bUseEvents = bValue
End Property

Public Property Get GUIDField() As String
    GUIDField = m_StrGUIDField
End Property

Public Property Let GUIDField(ByVal strValue As String)
    m_StrGUIDField = strValue
End Property

Public Property Get SynchTable() As String
    SynchTable = m_StrSynchTable
End Property

Public Property Let SynchTable(ByVal strValue As String)
    m_StrSynchTable = strValue
End Property

Public Property Get EditedIncidents() As Collection
    Set EditedIncidents = m_ColEditedIncidents
End Property

Public Property Get DeletedIncidents() As Collection
    Set DeletedIncidents = m_ColDeletedIncidents
End Property

Public Property Get NewIncidents() As Collection
    Set NewIncidents = m_ColNewIncidents
End Property

Public Property Get EncryptKey() As String
    EncryptKey = g_sKey
End Property

Public Property Let EncryptKey(ByVal strValue As String)
    g_sKey = strValue
End Property

Public Property Let HasEncrypt(b As Boolean)
    g_bHasEncrypt = b
End Property

Public Property Get HasEncrypt() As Boolean
    HasEncrypt = g_bHasEncrypt
End Property

Public Property Get RemoteTablePrefix() As String
    RemoteTablePrefix = m_StrRemoteTablePrefix
End Property

Public Property Let RemoteTablePrefix(ByVal strValue As String)
    m_StrRemoteTablePrefix = strValue
End Property

Public Property Get LocalConnectionString() As String
    LocalConnectionString = m_StrLocalConnectionString
End Property

Public Property Let LocalConnectionString(ByVal strValue As String)
    m_StrLocalConnectionString = strValue
End Property

Public Property Get WebsiteURL() As String
    WebsiteURL = m_StrWebsiteURL
End Property

Public Property Let WebsiteURL(ByVal strValue As String)
    m_StrWebsiteURL = strValue
    m_StrWebsiteURL = LCase(m_StrWebsiteURL)
    m_StrWebsiteURL = Replace(m_StrWebsiteURL, "oasis.asp", "oasis4.asp")
End Property

Public Property Get IsRunning() As Boolean
    IsRunning = m_bRunning
End Property

Public Property Get Connected() As Boolean
    Connected = m_bConnected
End Property

Public Property Get InternetConnectionType() As String
    InternetConnectionType = m_StrInternetConnectionType
End Property

Public Property Get RType() As RunnerType
    RType = m_RType
End Property

Public Property Let RType(rt As RunnerType)
    m_RType = rt
End Property

Public Property Get Paddy() As Boolean
    Paddy = m_bPaddy
End Property

Public Property Let Paddy(ByVal bValue As Boolean)
    m_bPaddy = bValue
End Property

Public Property Get Interval() As Long
    Interval = m_lInterval
End Property

Public Property Let Interval(ByVal lInterval As Long)
    m_lInterval = lInterval
End Property

Private Function CompareSynchHashes(sTableName As String, _
                                    oCN As ADODB.Connection) As Long
        '<EhHeader>
        On Error GoTo CompareSynchHashes_Err
        '</EhHeader>

100     CompareSynchHashes = -1 'this is good - all records match
    
        Dim oRSLocal As New ADODB.Recordset
        Dim oRSServer As ADODB.Recordset
        Dim i As Long
102     i = 0
    
104     UpdateHash sTableName, oCN
106     'SleepAPI 2000
108     oRSLocal.Open "SELECT * FROM [" & sTableName & "Hashed]", oCN, adOpenStatic, adLockReadOnly
        'SleepAPI 2000
110     Set oRSServer = OpenServerRSCompressed(m_StrWebsiteURL, "ID", "[dbo].[GetHashedSynchHistory] '" & sTableName & "'")
        SleepAPI 2000
        i = 0
        
        Do
    
112         i = i + 1

            If oRSServer.EOF And oRSLocal.EOF Then Exit Do

114         If oRSServer.EOF Then
116             CompareSynchHashes = i
118             TxMsg "### Server is missing set: " & i ', True
                Exit Do
120         ElseIf oRSLocal.EOF Then
122             CompareSynchHashes = i
124             TxMsg "### Local is missing set: " & i ', True
                'TxMsg "### oRSServer.Fields(0).Value: " & oRSServer.Fields(0).Value, True
                Exit Do
            Else

126             If Not (Trim(oRSLocal.Fields(0).Value) = Trim(oRSServer.Fields(0).Value)) Then
128                 CompareSynchHashes = i
130                 TxMsg "### oRSLocal Hash Record  (" & i & "): " & oRSLocal.Fields(0).Value, True
132                 TxMsg "### oRSServer Hash Record (" & i & "): " & oRSServer.Fields(0).Value, True
134                 TxMsg "### Mismatch between local and server records: " & ((i * 200) - 199) & " and " & (i * 200) ', True
                    Exit Do
                End If
                
            End If
        
136         oRSLocal.MoveNext
138         oRSServer.MoveNext

140     Loop Until (oRSLocal.EOF And oRSServer.EOF)
        
142     If CompareSynchHashes = -1 Then
144         TxMsg "### Matching server records for: " & sTableName, True
        End If
        
        oRSServer.Close
        oRSLocal.Close

        'TxMsg "SynchWorker.CompareSynchHashes_Err on line (" & Erl & ") " & Err.Description, True
        '<EhFooter>
        Exit Function

CompareSynchHashes_Err:
        TxMsg "SynchWorker.CompareSynchHashes_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Sub UpdateHash(sTableName As String, _
                       oCN As ADODB.Connection)
        '<EhHeader>
        On Error GoTo UpdateHash_Err
        '</EhHeader>

        Dim oMD5 As New MD5Hash
        Dim oRS As New ADODB.Recordset
        Dim oRS2 As New ADODB.Recordset
        Dim sSingleLine As String
        Dim i As Long
        Dim l As Long
    
104     If Not DoesTableExist(oCN, sTableName & "Hashed") Then
106         oCN.Execute "CREATE TABLE [" & sTableName & "Hashed] (Col1 text)"
        Else
108         oCN.Execute "DELETE FROM [" & sTableName & "Hashed]"
        End If

        Do While oCN.State = adStateExecuting
            SleepAPI 2000
        Loop

110     oRS.Open "SELECT * from [" & sTableName & "] order by [stablename], [sID]", oCN, adOpenDynamic, adLockBatchOptimistic
112     oRS2.Open "SELECT * from [" & sTableName & "Hashed]", oCN, adOpenDynamic, adLockBatchOptimistic

116     With oRS

118         l = 1
        
120         Do Until .EOF
    
122             i = 1
124             sSingleLine = ""

126             Do Until i > 200 Or .EOF
128                 sSingleLine = sSingleLine & oMD5.HashBytes(StrConv(.Fields("sID").Value & .Fields("sTableName").Value & .Fields("swhen").Value & CStr(.Fields("sequence").Value) & .Fields("sdelete").Value & .Fields("updates").Value, vbFromUnicode))
130                 i = i + 1
132                 .MoveNext
                Loop

134             sSingleLine = oMD5.HashBytes(StrConv(sSingleLine, vbFromUnicode))
138             l = l + 1

140             oRS2.AddNew
                oRS2.Fields("Col1").Value = sSingleLine
142             oRS2.UpdateBatch adAffectCurrent

            Loop
        
        End With
    
144     oRS2.Close
146     oRS.Close
148     'oCN.Close
150     Set oMD5 = Nothing
    
        '<EhFooter>
        Exit Sub

UpdateHash_Err:
        TxMsg "SynchWorker.UpdateHash_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Sub CheckIfDebugEnhancedEnabled(mCN As ADODB.Connection)
        '<EhHeader>
        On Error GoTo CheckIfDebugEnhancedEnabled_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
100     oRS.Open "SELECT [SettingValue1] FROM [Appsettings] WHERE [SettingName] = 'ShowAdvancedDebug'", mCN
102     m_bDebug = False
    
104     If Not oRS.State = 0 Then
    
106         If Not oRS.EOF Then
        
108             If CStr(Trim(oRS.Fields(0).Value)) = "1" Then m_bDebug = True
        
            End If
            
            oRS.Close
    
        End If
        
        Set oRS = Nothing

        '<EhFooter>
        Exit Sub

CheckIfDebugEnhancedEnabled_Err:
        TxMsg "SynchWorker.CheckIfIncidentsEnabled_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Function OpenServerRSCompressed(sWebsite As String, _
                                        sFunction As String, _
                                        sParameter As String) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo OpenServerRSCompressed_Err
        '</EhHeader>

        Dim sRetVal As String
        
100     bBatchDownloadOK = True
        sRetVal = OpenServerResponseCompressed(sWebsite, sFunction, sParameter)
           
128     If Not sRetVal = "-1" And Not sRetVal = "" Then
130         Set OpenServerRSCompressed = RecordsetFromXMLString(sRetVal)
        Else
132         Set OpenServerRSCompressed = Nothing
        End If
        
        If OpenServerRSCompressed Is Nothing Then
            TxMsg "!!!! OpenServerRSCompressed failed to create a recordset from retval: " & sRetVal, True, True
        Else
            TxMsg ">>>> OpenServerRSCompressed created a recordset from retval: " & sRetVal, True, True
        End If

        '<EhFooter>
        Exit Function

OpenServerRSCompressed_Err:
        TxMsg "!!     -- OpenServerRSCompressed_Err: (" & Erl & ") " & Err.Description, True
        'Resume Next
        '</EhFooter>
End Function

Private Function OpenServerResponseCompressed(sWebsite As String, _
                                              sFunction As String, _
                                              sParameter As String) As String
        '<EhHeader>
        On Error GoTo OpenServerResponseCompressed_Err
        '</EhHeader>

        Dim oHttp As New WinHttpRequest
        Dim sURLEncoded As String
        Dim bByte() As Byte
        Dim sSent As String
        Dim sStringFromUTF As String
        
100     PrepareHttpComms oHttp, sWebsite, False
102     sURLEncoded = sFunction & "=" & sParameter
104     bByte = OASStringCompression.CompressStringToByteArray(sURLEncoded)
106     sSent = OASStringCompression.ConvertByteArrayToString(bByte)

114     TxMsg ">>>> OpenServerResponseCompressed before encryption: " & sURLEncoded, True, True
116     TxMsg ">>>>> OpenServerResponseCompressed after encryption as text: " & sSent, True, True
108     oHttp.send bByte
118     TxMsg ">>>> Data received encrypted as utf-8: " & oHttp.responseBody, True, True

110     sStringFromUTF = OASStringCompression.ConvertByteArrayToString(oHttp.responseBody)
120     TxMsg ">>>> Data received encrypted as text: " & sStringFromUTF, True, True

112     OpenServerResponseCompressed = OASStringCompression.DecompressStringToString(sStringFromUTF)
122     TxMsg ">>>> Data received decrypted: " & OpenServerResponseCompressed, True, True
             
124     oHttp.abort
126     Set oHttp = Nothing
          
        '<EhFooter>
        Exit Function

OpenServerResponseCompressed_Err:
        TxMsg "OpenServerResponseCompressed_Err: (" & Erl & ") " & Err.Description, True
        On Error Resume Next
        oHttp.abort
        Set oHttp = Nothing
        '</EhFooter>
End Function

Private Function SaveServerRSCompressed(sWebsite As String, _
                                        oRS As ADODB.Recordset) As String
        '<EhHeader>
        On Error GoTo SaveServerRSCompressed_Err
        '</EhHeader>
 
        Dim oHttp As New WinHttpRequest
        Dim MsXmlDoc As New MSXML2.DOMDocument
        Dim sDoc As String
        Dim bByte() As Byte
        Dim sSent As String
        
        sWebsite = Replace(sWebsite, "oasis4.asp", "oasisuploader4.asp")
        '  TxMsg ">>>> sWebsite " & sWebsite, True, True
        
100     PrepareHttpComms oHttp, sWebsite, True
102     oRS.Save MsXmlDoc, 1
104     sDoc = CStr(MsXmlDoc.xml)
106     bByte = OASStringCompression.CompressStringToByteArray(sDoc)
108     sSent = OASStringCompression.ConvertByteArrayToString(bByte)
110     oHttp.send bByte
112     SaveServerRSCompressed = OASStringCompression.DecompressStringToString(oHttp.responseText)
        
114     TxMsg ">>>> SaveServerRSCompressed before encryption: " & sDoc, True, True
116     TxMsg ">>>>> SaveServerRSCompressed after encryption: " & sSent, True, True
118     TxMsg ">>>> Data received encrypted: " & oHttp.responseText, True, True
120     TxMsg ">>>> Data received decrypted: " & SaveServerRSCompressed, True, True
             
122     oHttp.abort
124     Set oHttp = Nothing
126     Set MsXmlDoc = Nothing
          
        '<EhFooter>
        Exit Function

SaveServerRSCompressed_Err:
        TxMsg "SaveServerRSCompressed_Err: (" & Erl & ") " & Err.Description, True
        On Error Resume Next
        oHttp.abort
        Set oHttp = Nothing
        Set MsXmlDoc = Nothing
        '</EhFooter>
End Function

Private Function CheckIfIncidentsEnabled(mCN As ADODB.Connection) As Boolean
        '<EhHeader>
        On Error GoTo CheckIfIncidentsEnabled_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
100     oRS.Open "SELECT [SettingValue1] FROM [Appsettings] WHERE [SettingName] = 'ShowSecurityTab'", mCN
102     CheckIfIncidentsEnabled = False
    
104     If Not oRS.State = 0 Then
    
106         If Not oRS.EOF Then
        
108             If CStr(Trim(oRS.Fields(0).Value)) = "1" Then CheckIfIncidentsEnabled = True
        
            End If
    
        End If

        '<EhFooter>
        Exit Function

CheckIfIncidentsEnabled_Err:
        TxMsg "SynchWorker.CheckIfIncidentsEnabled_Err on line (" & Erl & ") " & Err.Description, True
        
        '</EhFooter>
End Function

Public Sub InitComms()
    Set oServerInterComms = New IServer
    Set m_frmHiddenSynchWorker = New frmHiddenSynchWorker
    
    Load m_frmHiddenSynchWorker

    If Not oServerInterComms.IsChannelRegistered(1) Then
        oServerInterComms.ConnectToDataChannel 1, m_frmHiddenSynchWorker.hWnd, True
    End If

End Sub

Private Function FetchServerTime(sWebsite As String) As Boolean
        '<EhHeader>
        On Error GoTo FetchServerTime_Err
        '</EhHeader>

        Dim sResponse As String
        sResponse = OpenServerResponseCompressed(sWebsite, "getservertime", "")
    
        Dim sServerTime As String
        Dim dDate As Date

100     lTimeDiffFromServerInSeconds = -1
        '2010-06-07T01:17:14Z
    
102     'TxMsg "********************************************************************************************"
104     'TxMsg "*** Fetching server time.........."
        
114     If Len(sResponse) = 20 Then
116         FetchServerTime = True
118         sServerTime = sResponse
120         dDate = DateSerial(CInt(Left$(sServerTime, 4)), CInt(Mid$(sServerTime, 6, 2)), CInt(Mid$(sServerTime, 9, 2))) + TimeSerial(CInt(Mid$(sServerTime, 12, 2)), CInt(Mid$(sServerTime, 15, 2)), CInt(Mid$(sServerTime, 18, 2)))
122         lTimeDiffFromServerInSeconds = DateDiff("s", Now(), dDate)
124         TxMsg "*** Server time at " & sServerTime & " is " & lTimeDiffFromServerInSeconds & " seconds out of synch with this local pc"
        Else
126         FetchServerTime = False
128         TxMsg "*** Server time fetch failed - aborting synchronisation."
        End If

134     'TxMsg "********************************************************************************************"

        '<EhFooter>
        Exit Function

FetchServerTime_Err:
        TxMsg "SynchWorker.FetchServerTime_Err on line (" & Erl & ") " & Err.Description, True
        lTimeDiffFromServerInSeconds = -1
        FetchServerTime = False
        
        '</EhFooter>
End Function

Private Sub CreateAccessDatabase(sDatabaseToCreate)

    Dim catNewDB As adox.Catalog
    Set catNewDB = New adox.Catalog
    catNewDB.Create "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sDatabaseToCreate & ";Jet OLEDB:Engine Type=5;"
    ' Engine Type=5 = Access 2000 Database
    ' Engine Type=4 = Access 97 Database
    Set catNewDB = Nothing

End Sub

Private Function CheckIfFileExists(FileName As String) As Boolean
    Dim i As Integer
    
    On Local Error Resume Next
    i = Len(Dir$(FileName$))

    If Err Or i = 0 Then
        CheckIfFileExists = False
    Else
        CheckIfFileExists = True
    End If

    On Local Error GoTo 0
End Function

Public Sub Start()

    If Not m_bRunning Then
        m_bRunning = True
        ' timer to fire the Runnable_Start
        modStart.Start Me
    Else
        ' Make sure this only 1 Version running....
        Err.Raise 32540, App.EXEName & " SynchWorker", "Already running."
    End If

End Sub

Public Sub StartNonAsync()
    'Non Threaded
    Runnable_Start
End Sub

Private Sub Class_Initialize()
    Set m_oAES = New clsAES
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Unload m_frmHiddenSynchWorker
    Set m_frmHiddenSynchWorker = Nothing
    oServerInterComms.DisconnectFromChannel
    Set oServerInterComms = Nothing
End Sub

Private Function IncidentsV2Active(oCNClient As ADODB.Connection) As Boolean

    Dim sFileName As String
    sFileName = oCNClient.Properties.Item("Data Source Name").Value
    sFileName = Left$(sFileName, Len(sFileName) - 15)
    'sFileName = sFileName & "\data\db\DynamicData\IncidentsV2.mdb"
    sFileName = sFileName & "DynamicData\IncidentsV2.mdb"
    IncidentsV2Active = CheckIfFileExists(sFileName)
    TxMsg ">>> Check for filename '" & sFileName & "' returned: " & IncidentsV2Active, True, True
            
End Function

Private Sub Runnable_Start()
        '<EhHeader>
        On Error GoTo Runnable_Start_Err
        '</EhHeader>
        Dim sSQL As String
        Dim bFound As Boolean
        Dim bRetval As Boolean
        Dim iCount As Integer
        Dim i As Long
        Dim bCancel As Boolean
        Dim sDDConnStr As String
        Dim bClientDDInOrder As Boolean
        Dim sDatabaseFullPath As String
        Dim oCN As ADODB.Connection
        Dim oRS As ADODB.Recordset
        Dim DDCnn As ADODB.Connection
        Dim oRSSQLLyr As ADODB.Recordset
        Dim bSkipAttachments As Boolean
        Dim bSkipIncidents As Boolean
        
100     bAbortSynchronisation = True

102     Do Until Not bAbortSynchronisation
            
104         bAbortSynchronisation = False
106         sTablesCreated = ""
        
108         If Not Len(m_StrRemoteTablePrefix) > 0 Then
110             TxMsg "!!!!!! ERROR: m_StrRemoteTablePrefix not set. SynchWorker aborting.", True
112             End
            Else
        
114             If InStr(m_StrLocalConnectionString, "SQLNCLI10") > 0 Then bSQLServerInUse = True
116             TxMsg "****************************************************************************************************************************************", True
118             TxMsg "--- Starting synch engine version: [" & App.Major & "." & App.Minor & "." & App.Revision & "] powered by iMMAP", True
120             TxMsg "****************************************************************************************************************************************", True
                
122             Select Case m_RType

                    Case RunnerType.InternetConnectionCheck

124                     m_StrInternetConnectionType = ""
126                     m_bConnected = CheckInternConnection(m_StrInternetConnectionType)
                    
128                     If Not m_bUseEvents Then
130                         TxMsg "**" & CStr(RT_InternetConnectionCheck)
                        Else
132                         RaiseEvent CompleteEX(InternetConnectionCheck)
                        End If
                
134                 Case RunnerType.enmPaddy
            
136                 Case RunnerType.IncidentNotifier
                    
138                     RaiseEvent CompleteEX(IncidentNotifier)

140                 Case Else

142                     If Len(m_StrLocalConnectionString) > 0 Then
           
144                         Set oCN = New ADODB.Connection
146                         oCN.Open m_StrLocalConnectionString
                            
148                         CheckIfDebugEnhancedEnabled oCN
150                         CheckProxySettings oCN
152                         m_bIncidentsV2 = IncidentsV2Active(oCN)
                            
154                         Set m_ColNewIncidents = New Collection
156                         Set m_ColDeletedIncidents = New Collection

158                         If Not Right$(m_StrWebsiteURL, 1) = "/" Then m_StrWebsiteURL = m_StrWebsiteURL & "/"
        
160                         If FetchServerTime(m_StrWebsiteURL) Then
            
164                             Set m_RSttkGISLayerSQLRemote = OpenServerRSCompressed(m_StrWebsiteURL, "ID", "SELECT * FROM ttkGISLayerSQL")
                                
                                '                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '                                '' SYNCH GEOBOOKMARKS
                                '
                                '162                             If m_bSynchGeoMarks Then
                                '
                                '164                                 ValidateClientTable "GeoBookMarks", "GUID1", True, False, oCN, ""
                                '166                                 ValidateClientTable "GeoBookMarksCategories", "GUID1", True, False, oCN, ""
                                '                                    UpdateRecordsFromServer "GeoBookMarks", m_StrRemoteTablePrefix & "GeoBookMarks", "SynchHistory", m_StrRemoteTablePrefix & "SynchHistory", oCN, m_StrWebsiteURL, "GUID1"
                                '                                    UpdateRecordsFromServer "GeoBookMarksCategories", m_StrRemoteTablePrefix & "GeoBookMarksCategories", "SynchHistory", m_StrRemoteTablePrefix & "SynchHistory", oCN, m_StrWebsiteURL, "GUID1"
                                '180                                 UploadTableChangesToServer "GeoBookMarks", "GUID1", False, m_StrWebsiteURL, oCN, SynchOperation.Add
                                '
                                '                                End If
                                '
                                '                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '                                '' SYNCH INCIDENT CHARTS
                                '
                                '                                If 1 = 3 Then
                                '
                                '                                    If m_bIncidentsV2 Then
                                '182                                     bSkipIncidents = True
                                '                                    Else
                                '                                        bSkipIncidents = Not CheckIfIncidentsEnabled(oCN)
                                '                                    End If
                                '
                                '184                                 If Not bSkipIncidents Then
                                '
                                '186                                     ValidateClientTable "Incidents_ChartSettings", "GUID1", True, False, oCN, ""
                                '
                                '188                                     GetSynchRecordsTOADD "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN
                                '
                                '190                                     If Not SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch And Not SynchHistoryOverview.NextSynchNextOpeation = EditsOnly Then
                                '192                                         GetSynchRecordsTODELETE "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN
                                '                                        End If
                                '
                                '194                                     If SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch Then
                                '196                                         GetSynchRecordsTOUPDATE "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN
                                '                                        End If
                                '
                                '198                                     UploadTableChangesToServer "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN, SynchOperation.Add
                                '200                                     UploadTableChangesToServer "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN, SynchOperation.Delete
                                '202                                     UploadTableChangesToServer "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN, SynchOperation.Update
                                '
                                '                                    End If
                                '
                                '                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '                                    '' SYNCH INCIDENTS
                                '                                    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '
                                '204                                 If Not bSkipIncidents Then
                                '
                                '                                        Set m_RSttkGISLayerSQLLocal = New ADODB.Recordset
                                '                                        m_RSttkGISLayerSQLLocal.Open "SELECT * FROM ttkGISLayerSQL", oCN, adOpenDynamic, adLockBatchOptimistic
                                '                                        UpdatettkGISLaySQL "oincidents", oCN
                                '
                                '206                                     ValidateClientTable "oincidents", "ID", True, True, oCN, ""
                                '
                                '208                                     GetSynchRecordsTOADD "oincidents", "ID", True, m_StrWebsiteURL, oCN
                                '
                                '210                                     Do While SynchHistoryOverview.NextSynchNextOpeation = NewOnly
                                '212                                         GetSynchRecordsTOADD "oincidents", "ID", True, m_StrWebsiteURL, oCN
                                '                                        Loop
                                '
                                '214                                     TxMsg CStr("**" & CStr(RT_IncidentSynch) & m_sIDs), True
                                '
                                '216                                     If Not SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch And Not SynchHistoryOverview.NextSynchNextOpeation = EditsOnly Then
                                '218                                         GetSynchRecordsTODELETE "oincidents", "ID", True, m_StrWebsiteURL, oCN
                                '                                        End If
                                '
                                '220                                     If SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch Then
                                '222                                         GetSynchRecordsTOUPDATE "oincidents", "ID", True, m_StrWebsiteURL, oCN
                                '                                        End If
                                '
                                '224                                     UploadTableChangesToServer "oincidents", "ID", False, m_StrWebsiteURL, oCN, SynchOperation.Add
                                '
                                '                                    End If
                                '
                                '                                End If
                                
                                oCN.Execute "DELETE FROM SynchHistory WHERE stablename <> 'oincidents'"
                                If DoesTableExist(oCN, "SynchHistoryOverview") Then oCN.Execute "DROP TABLE SynchHistoryOverview"
                                
166                             ValidateClientTable "Incidents_ChartSettings", "GUID1", True, False, oCN, m_StrRemoteTablePrefix
168                             ValidateClientTable "oincidents", "ID", True, True, oCN, ""
170                             ValidateClientTable "GeoBookMarks", "GUID1", True, False, oCN, m_StrRemoteTablePrefix
172                             ValidateClientTable "GeoBookMarksCategories", "GUID1", True, False, oCN, m_StrRemoteTablePrefix

174                             If Not DoesTableExist(oCN, m_StrRemoteTablePrefix & "SynchHistory") Then
176                                 oCN.Execute "SELECT * INTO [" & m_StrRemoteTablePrefix & "SynchHistory] FROM [SynchHistory]"
178                                 oCN.Execute "DELETE FROM [" & m_StrRemoteTablePrefix & "SynchHistory]"
                                End If

180                             iCount = 0
                                    
182                             TxMsg "**??############################################################################################", True
184                             TxMsg "**?? Beginning sync for usergroup datasets", True

                                Do
186                                 bRetval = SynchDDWithServer2(m_StrRemoteTablePrefix, m_StrWebsiteURL, oCN)
188                                 iCount = iCount + 1
190                             Loop Until bRetval Or iCount = 50

192                             TxMsg "**?? Ending sync for common datasets", True
194                             TxMsg "**??############################################################################################", True

196                             TxMsg "", True
198                             Set oRS = oCN.OpenSchema(adSchemaTables)
200                             bFound = False
                                
202                             Do Until oRS.EOF

204                                 If oRS.Fields("Table_Name").Value = "oincidents" Then
206                                     bFound = True
                                        Exit Do
                                    End If

208                                 oRS.MoveNext
                                Loop
                                
210                             sSQL = "SELECT oincidents_GEO.UID, oincidents_FEA.UID, oincidents_GEO.XMIN, oincidents_GEO.XMAX, oincidents_GEO.YMIN, oincidents_GEO.YMAX, oincidents_GEO.SHAPETYPE as SHAPETYPEFORINCIDENTS, oincidents_FEA.ID, oincidents_FEA.NAME, oincidents_FEA.TYPE, oincidents_FEA.TARGET, oincidents_FEA.Dead, oincidents_FEA.Affected, oincidents_FEA.Violent, oincidents_FEA.Injured, oincidents_FEA.Incident_DATE, oincidents_FEA.TIME00, oincidents_FEA.LocDesc, oincidents_FEA.Source, oincidents_FEA.Town, oincidents_FEA.District, oincidents_FEA.Province, oincidents_FEA.DESCRIPTION, oincidents_FEA.Scoring, oincidents_FEA.Incident_DATESERIAL" & " FROM oincidents_FEA INNER JOIN oincidents_GEO ON oincidents_FEA.UID = oincidents_GEO.UID"

212                             If Not bFound Then CreateQuery "oincidents", sSQL, oCN, ""
                                
214                             iCount = 0

216                             TxMsg "**??#########################################################################################", True
218                             TxMsg "**?? Beginning sync for common datasets", True

                                Do
220                                 bRetval = SynchDDWithServer2("", m_StrWebsiteURL, oCN)
222                                 iCount = iCount + 1
224                             Loop Until bRetval Or iCount >= 50

226                             TxMsg "**?? Ending sync for common datasets", True
228                             TxMsg "**??########################################################################################", True

                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '' SYNCH DYNAMIC DATA DEFS
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
230                             If 1 = 1 Then 'Currently DynamicDataDefs are synched automatically

232                                 Set oRSSQLLyr = New ADODB.Recordset
                        
234                                 With oRSSQLLyr

236                                     If Not bSQLServerInUse Then
238                                         .Open "SELECT DISTINCT DDDefName, ConnectionString FROM DynamicDataDefs WHERE [Synch] = True", oCN, adOpenDynamic, adLockReadOnly
                                        Else
240                                         .Open "SELECT DISTINCT [DDDefName], CAST([ConnectionString] AS varchar(max)) FROM DynamicDataDefs WHERE [Synch] = 'True'", oCN, adOpenDynamic, adLockReadOnly
                                        End If
                                    
242                                     If Not .BOF Or Not .EOF Then

244                                         .MoveFirst
                                
246                                         Do While Not .EOF
            
248                                             Set DDCnn = New ADODB.Connection

250                                             If bSQLServerInUse Then
                                                    'DDCnn.CursorLocation = adUseClient
252                                                 TxMsg "Connecting to OasisV3 MSSQL Database......"
254                                                 DDCnn.Open m_StrLocalConnectionString
                                                Else

256                                                 sDDConnStr = oCN.Properties.Item("Data Source Name").Value
258                                                 sDDConnStr = Left$(sDDConnStr, Len(sDDConnStr) - 15)
260                                                 sDDConnStr = Replace$(.Fields(1).Value, "\data\db\", sDDConnStr, , , vbTextCompare)
262                                                 DDCnn.CursorLocation = adUseClient
264                                                 TxMsg "Connecting to connection string: " & sDDConnStr, True
266                                                 sDatabaseFullPath = Mid$(sDDConnStr, InStr(sDDConnStr, "Data Source=") + 12, InStr(sDDConnStr, "Mode=") - (InStr(sDDConnStr, "Data Source=") + 12) - 1)
268                                                 DDCnn.ConnectionString = sDDConnStr

270                                                 If Not CheckIfFileExists(sDatabaseFullPath) Then
272                                                     TxMsg "   --- Creating dynamic database: " & sDatabaseFullPath, True
274                                                     CreateAccessDatabase sDatabaseFullPath
                                                    End If
                                                
276                                                 DDCnn.Open sDDConnStr

278                                                 If Not DoesTableExist(DDCnn, "ttkGISLayerSQL") Then
            
280                                                     DDCnn.Execute "CREATE TABLE [ttkGISLayerSQL]  ([NAME] TEXT(255), [XMIN] DOUBLE, [XMAX] DOUBLE, [YMIN] DOUBLE, [YMAX] DOUBLE, [SHAPETYPE] INTEGER)"
282                                                     TxMsg "   --- Creating table [ttkGISLayerSQL]"

                                                    End If
                                                
                                                End If
                                                
284                                             Set m_RSttkGISLayerSQLLocal = New ADODB.Recordset
286                                             m_RSttkGISLayerSQLLocal.Open "SELECT * FROM ttkGISLayerSQL", DDCnn, adOpenDynamic, adLockBatchOptimistic
                                
288                                             bClientDDInOrder = RunDDSchemaCompare(m_StrWebsiteURL, "dd_" & .Fields(0) & "_", DDCnn)
                                                
290                                             If bClientDDInOrder Then
                                                    
292                                                 iCount = 0

294                                                 TxMsg "**??#########################################################################################", True
296                                                 TxMsg "**?? Beginning data sync for dynamic database: " & .Fields(0) & " ......................", True

                                                    Do
298                                                     bRetval = SynchDDWithServer2("dd_" & .Fields(0) & "_", m_StrWebsiteURL, DDCnn)
300                                                     iCount = iCount + 1
302                                                 Loop Until bRetval Or iCount = 50

304                                                 TxMsg "**?? Ending sync for dynamic database: " & "dd_" & .Fields(0) & " ......................", True
306                                                 TxMsg "**??#########################################################################################", True

                                                End If
                                            
308                                             DDCnn.Close
310                                             Set DDCnn = Nothing
312                                             .MoveNext

                                            Loop
                        
314                                         If Len(sTablesCreated) > 0 Then TxMsg "**?? TABLES CREATED: " & sTablesCreated, True
316                                         sTablesCreated = ""
                                            
                                        End If
                            
                                    End With
                        
318                                 Set oRSSQLLyr = Nothing
                    
                                End If
                                
                                '                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '                                '' SYNCH ATTACHMENTS
                                '                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '
                                '302                             bSkipAttachments = True 'IIf(InStr(LCase$(m_StrWebsiteURL), "nomad") > 0 Or InStr(LCase$(m_StrWebsiteURL), "pakistan.pk") > 0 Or InStr(LCase$(m_StrWebsiteURL), "worldvision") > 0, True, False)
                                '                                'bSkipIncidents = IIf(InStr(LCase$(m_StrWebsiteURL), "nomad") > 0 Or InStr(LCase$(m_StrWebsiteURL), "pakistan.pk") > 0 Or InStr(LCase$(m_StrWebsiteURL), "worldvision") > 0, True, False)
                                '
                                '304                             If Not bSkipAttachments Then 'Attachments are synched automatically
                                '
                                '306                                 GetSynchRecordsTOADD "Attachments", "sGUID", False, m_StrWebsiteURL, oCN
                                '308                                 GetSynchRecordsTODELETE "Attachments", "sGUID", False, m_StrWebsiteURL, oCN
                                '310                                 GetSynchRecordsTOUPDATE "Attachments", "sGUID", False, m_StrWebsiteURL, oCN
                                '
                                '312                                 UploadTableChangesToServer "Attachments", "sGUID", False, m_StrWebsiteURL, oCN, SynchOperation.Add
                                '314                                 UploadTableChangesToServer "Attachments", "sGUID", False, m_StrWebsiteURL, oCN, SynchOperation.Delete
                                '316                                 UploadTableChangesToServer "Attachments", "sGUID", False, m_StrWebsiteURL, oCN, SynchOperation.Update
                                '
                                '                                End If
                    
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '' END SYNCH
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
                            End If
                    
                        End If
            
                End Select
        
            End If
    
320         Set oRSSQLLyr = Nothing
            
        Loop

322     TxMsg "**", True
324     TxMsg "**??Synchronisation complete", True
326     m_bRunning = False

        'TxMsg "SynchWorker.Runnable_Start_Err on line (" & Erl & ") " & Err.Description, True
        'Set oRSSQLLyr = Nothing
        'm_bRunning = False
        '<EhFooter>
        Exit Sub

Runnable_Start_Err:
        TxMsg "SynchWorker.Runnable_Start_Err on line (" & Erl & ") " & Err.Description, True
        MsgBox "FATAL ERROR IN SYNCH ENGINE!"
        Set oRSSQLLyr = Nothing
        m_bRunning = False
        '</EhFooter>
End Sub

Private Sub UpdatettkGISLaySQL(ByVal sTableName As String, _
                               oCN As ADODB.Connection)
        '<EhHeader>
        On Error GoTo UpdatettkGISLaySQL_Err
        '</EhHeader>
        
        Dim oRS As New ADODB.Recordset

100     If Right$(sTableName, 4) = "_FEA" Then sTableName = Left$(sTableName, Len(sTableName) - 4)
102     If Right$(sTableName, 4) = "_GEO" Then sTableName = Left$(sTableName, Len(sTableName) - 4)

104     With m_RSttkGISLayerSQLLocal
    
106         If .State = adStateOpen Then
        
108             If Not (.EOF And .BOF) Then .MoveFirst
110             If Not (.EOF And .BOF) Then .Find "[name] = '" & sTableName & "'"
        
112             oRS.Open "SELECT MIN(XMIN), MAX(XMAX), MIN(YMIN), MAX(YMAX), FIRST(SHAPETYPE) FROM [" & sTableName & "_GEO]", oCN, adOpenDynamic, adLockReadOnly
        
114             If .EOF Then .AddNew

116             .Fields("Name").Value = sTableName

118             If oRS.EOF Then
                
120                 .Fields("XMIN").Value = -180
122                 .Fields("XMAX").Value = 180
124                 .Fields("YMIN").Value = -90
126                 .Fields("YMAX").Value = 90

                Else

128                 .Fields("XMIN").Value = oRS.Fields(0).Value
130                 .Fields("XMAX").Value = oRS.Fields(1).Value
132                 .Fields("YMIN").Value = oRS.Fields(2).Value
134                 .Fields("YMAX").Value = oRS.Fields(3).Value

136                 If oCN.Execute("SELECT DISTINCT SHAPETYPE FROM [" & sTableName & "_GEO]").RecordCount = 1 Then
138                     .Fields("SHAPETYPE").Value = oRS.Fields(4).Value
                    Else
140                     .Fields("SHAPETYPE").Value = 0
                    End If
                    
                End If
                
                .UpdateBatch adAffectCurrent
                
142             oRS.Close
144             Set oRS = Nothing
                
                '146             If m_RSttkGISLayerSQLRemote.State = adStateOpen Then
                '
                '148                 If Not (m_RSttkGISLayerSQLRemote.EOF And m_RSttkGISLayerSQLRemote.BOF) Then
                '
                '150                     m_RSttkGISLayerSQLRemote.MoveFirst
                '152                     m_RSttkGISLayerSQLRemote.Find "[name] = '" & sTableName & "'"
                '
                '154                     If Not m_RSttkGISLayerSQLRemote.EOF Then
                '
                '156                         .Fields("Name").Value = m_RSttkGISLayerSQLRemote.Fields("Name").Value
                '158                         .Fields("XMIN").Value = m_RSttkGISLayerSQLRemote.Fields("XMIN").Value
                '160                         .Fields("XMAX").Value = m_RSttkGISLayerSQLRemote.Fields("XMAX").Value
                '162                         .Fields("YMIN").Value = m_RSttkGISLayerSQLRemote.Fields("YMIN").Value
                '164                         .Fields("YMAX").Value = m_RSttkGISLayerSQLRemote.Fields("YMAX").Value
                '166                         .Fields("SHAPETYPE").Value = m_RSttkGISLayerSQLRemote.Fields("SHAPETYPE").Value
                '
                '                        End If
                '
                '                    End If
                '
                '                End If
            
168             .UpdateBatch adAffectCurrent
           
            End If
    
        End With
    
        '<EhFooter>
        Exit Sub

UpdatettkGISLaySQL_Err:
        TxMsg "SynchWorker.UpdatettkGISLaySQL_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Sub ValidateClientTable(ByVal sTableName As String, _
                                ByVal sGUIDFieldName As String, _
                                ByVal bGUIDFieldIsString As String, _
                                ByVal bIsGeoTable As Boolean, _
                                cn As ADODB.Connection, _
                                ByVal sSynchHistoryPrefix As String)
    
    Dim sTableDesc As String
    sTableDesc = IIf(bIsGeoTable, "GEO-table", "table")
    
    'TxMsg "   ------------------------------------------------------------------------------------------------"
    TxMsg "**?? Verifying " & sTableDesc & " [" & sTableName & "]" & " for integrity....."
    'TxMsg "   ------------------------------------------------------------------------------------------------"
    'TxMsg "   --- Verifying " & sTableDesc & " [" & sTableName & "]" & " for duplicate field [" & sGUIDFieldName & "]....."
    
    CheckForClientTableDuplicates sTableName, sGUIDFieldName, bGUIDFieldIsString, bIsGeoTable, cn, sSynchHistoryPrefix
    
    If bIsGeoTable Then
        'TxMsg "   --- Verifying " & sTableDesc & " [" & sTableName & "]" & " for duplicate UID....."
        CheckForClientTableDuplicateUID sTableName, cn, sGUIDFieldName, sSynchHistoryPrefix
        
        'TxMsg "   --- Verifying " & sTableDesc & " [" & sTableName & "]" & " for integrity....."
        ValidateClientGEOTable sTableName, cn, sGUIDFieldName, sSynchHistoryPrefix
    End If
    
    'TxMsg "   --- Verifying " & sTableDesc & " [" & sTableName & "]" & " is in accordance with SynchHistory....."
    CheckForClientTableSynchRules sTableName, sGUIDFieldName, bIsGeoTable, cn, sSynchHistoryPrefix
    
    'TxMsg "   --- Verification of " & sTableDesc & " [" & sTableName & "]" & " complete."
    'TxMsg "   ------------------------------------------------------------------------------------------------"
    
End Sub

Private Sub CheckForClientTableSynchRules(ByVal sTableName As String, _
                                          ByVal sGUIDFieldName As String, _
                                          ByVal bIsGeoTable As Boolean, _
                                          cn As ADODB.Connection, _
                                          ByVal sLocalSynchHistPrefix As String)
        '<EhHeader>
        On Error GoTo CheckForClientTableSynchRules_Err
        '</EhHeader>
    
        Dim sSQL As String
        Dim RSDifferences As New ADODB.Recordset
            
        If Not Right(sTableName, 4) = "_GEO" Then
        
100         If bIsGeoTable = True Then
102             If Not Right(sTableName, 4) = "_FEA" Then sTableName = sTableName & "_FEA"
104             sSQL = "SELECT [" & sTableName & "].[" & sGUIDFieldName & "], [UID] FROM [" & sTableName & "] LEFT OUTER JOIN [" & sLocalSynchHistPrefix & "SynchHistory] ON [" & sTableName & "].[" & sGUIDFieldName & "] = [" & sLocalSynchHistPrefix & "SynchHistory].[sID] Where ([" & sLocalSynchHistPrefix & "SynchHistory].[sID] Is Null)"
            Else
106             sSQL = "SELECT [" & sTableName & "].[" & sGUIDFieldName & "] FROM [" & sTableName & "] LEFT OUTER JOIN [" & sLocalSynchHistPrefix & "SynchHistory] ON [" & sTableName & "].[" & sGUIDFieldName & "] = [" & sLocalSynchHistPrefix & "SynchHistory].[sID] Where ([" & sLocalSynchHistPrefix & "SynchHistory].[sID] Is Null)"
            End If
    
108         RSDifferences.Open sSQL, cn, adOpenStatic, adLockBatchOptimistic
        
110         If RSDifferences.State = adStateOpen Then
        
112             Do Until RSDifferences.EOF
            
114                 TxMsg "!!!!!! [" & sTableName & "].[" & sGUIDFieldName & "] value '" & RSDifferences.Fields(0).Value & "' not found in SynchHistory.  Removing this record...", True
116                 cn.Execute "DELETE FROM [" & sTableName & "] WHERE [" & sGUIDFieldName & "] = '" & RSDifferences.Fields(0).Value & "'"

118                 If bIsGeoTable = True Then
120                     cn.Execute "DELETE FROM [" & Left(sTableName, Len(sTableName) - 4) & "_GEO] WHERE [UID] = " & RSDifferences.Fields(1).Value
                    End If

122                 RSDifferences.MoveNext
        
                Loop
        
124             RSDifferences.Close
        
            End If
        
        End If
    
126     Set RSDifferences = Nothing

        '<EhFooter>
        Exit Sub

CheckForClientTableSynchRules_Err:
        TxMsg "SynchWorker.CheckForClientTableSynchRules_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Sub CheckProxySettings(cn As ADODB.Connection)
        '<EhHeader>
        On Error GoTo CheckProxySettings_Err
        '</EhHeader>
        Dim objShell, bKey
        Dim sKey As String
        Dim sProxyArray() As String
        Dim i As Integer
        
        Dim bAdvancedProxySettings As Boolean
        
100     m_bProxyEnabled = False
102     m_sProxyIP = ""
104     m_sProxyPort = ""
    
106     Set objShell = CreateObject("WScript.Shell")
    
108     If CStr(objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")) = "1" Then
    
110         bKey = objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer")
112         sKey = CStr(bKey)

114         If Len(sKey) > 1 And InStr(sKey, ";") > 0 Then

116             sProxyArray = Split(sKey, ";")
                
118             Do Until i = UBound(sProxyArray)
                
120                 If Left(sProxyArray(i), 5) = "http=" Then
122                     sKey = Replace(sProxyArray(i), "http=", "")
                        Exit Do
                    End If

124                 i = i + 1
                Loop
                
            End If
            
126         sKey = Replace(sKey, "http=", "")
            
128         If Len(sKey) > 1 And InStr(sKey, ":") > 0 Then

130             sProxyArray = Split(sKey, ":")
132             m_bProxyEnabled = True
134             m_sProxyIP = sProxyArray(0)
136             m_sProxyPort = sProxyArray(1)
        
138             TxMsg ">>>>>> Proxy detected at [" & sProxyArray(0) & ":" & sProxyArray(1) & "]", True
    
            End If
    
        End If
    
140     Set objShell = Nothing

        '<EhFooter>
        Exit Sub

CheckProxySettings_Err:
        Resume Next
        'Err.Raise vbObjectError + 100, _
         "OASIS_SynchNG.SynchWorker.CheckProxySettings", _
         "SynchWorker component failure"
        '</EhFooter>
End Sub

Private Sub CheckForClientTableDuplicates(sTableName As String, _
                                          sGUIDFieldName As String, _
                                          bGUIDFieldIsString As String, _
                                          bIsGeoTable As Boolean, _
                                          cn As ADODB.Connection, _
                                          sLocalSynchHistPrefix As String)
        '<EhHeader>
        On Error GoTo CheckForClientTableDuplicates_ERR
        '</EhHeader>
    
        Dim sSQL As String
        Dim sBrackets As String

100     If bGUIDFieldIsString Then sBrackets = "'"
        Dim RSLocalDuplicates As New ADODB.Recordset
    
102     If bIsGeoTable Then
    
            Dim sTableName_GEO As String
            Dim sTableName_FEA As String
            Dim RSLocalDuplicates_BYUID As ADODB.Recordset
        
104         sTableName_GEO = sTableName & "_GEO"
106         sTableName_FEA = sTableName & "_FEA"
108         sSQL = "SELECT [" & sGUIDFieldName & "] FROM [" & sTableName_FEA & "] GROUP BY [" & sGUIDFieldName & "] HAVING (COUNT([" & sGUIDFieldName & "]) > 1)"
110         RSLocalDuplicates.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
        
112         Do Until RSLocalDuplicates.EOF
               
114             TxMsg "!!!!!! Duplicate [" & sGUIDFieldName & "] found in GEOtable [" & sTableName & "] - removing all entries from table and SynchHistory for value: " & RSLocalDuplicates.Fields(0).Value, True
        
116             Set RSLocalDuplicates_BYUID = New ADODB.Recordset
118             sSQL = "SELECT [UID] FROM [" & sTableName_FEA & "] WHERE [" & sGUIDFieldName & "] = " & sBrackets & RSLocalDuplicates.Fields(0).Value & sBrackets
120             RSLocalDuplicates_BYUID.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
        
122             Do Until RSLocalDuplicates_BYUID.EOF
                
124                 cn.Execute "DELETE FROM [" & sTableName_GEO & "] WHERE [UID] = " & RSLocalDuplicates_BYUID.Fields(0).Value
126                 cn.Execute "DELETE FROM [" & sTableName_FEA & "] WHERE [UID] = " & RSLocalDuplicates_BYUID.Fields(0).Value
128                 RSLocalDuplicates_BYUID.MoveNext
                
                Loop
            
130             Set RSLocalDuplicates_BYUID = Nothing
132             cn.Execute "DELETE FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [sID] = '" & RSLocalDuplicates.Fields(0).Value & "'"
134             RSLocalDuplicates.MoveNext
        
            Loop
    
        Else

136         sSQL = "SELECT [" & sGUIDFieldName & "] FROM [" & sTableName & "] GROUP BY [" & sGUIDFieldName & "] HAVING (COUNT([" & sGUIDFieldName & "]) > 1)"
138         RSLocalDuplicates.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
    
140         Do Until RSLocalDuplicates.EOF
    
142             TxMsg "!!!!!! Duplicate [" & sGUIDFieldName & "] found in table [" & sTableName & "] - removing all entries from table and SynchHistory for value: " & RSLocalDuplicates.Fields(0).Value, True
144             cn.Execute "DELETE FROM [" & sTableName & "] WHERE [" & sGUIDFieldName & "] = " & sBrackets & RSLocalDuplicates.Fields(0).Value & sBrackets
146             cn.Execute "DELETE FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [sID] = '" & RSLocalDuplicates.Fields(0).Value & "'"
148             RSLocalDuplicates.MoveNext
        
            Loop
    
        End If
    
150     Set RSLocalDuplicates = Nothing
 
        '<EhFooter>
        Exit Sub

CheckForClientTableDuplicates_ERR:

        TxMsg "SynchWorker.CheckForClientTableDuplicates_ERR: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Function WaitForRSToBecomeEOF(iMins As Integer, _
                                      sSQL As String, _
                                      cn As ADODB.Connection) As ADODB.Recordset

    Dim oRS As New ADODB.Recordset
    Dim iIterations As Integer

    iIterations = 0
    oRS.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
    
    If Not oRS.EOF Then
    
        Do Until (iIterations = iMins) Or (oRS.EOF)
        
            TxMsg "!!!!!! Discrepency found - pausing for 1 minute ( " & iIterations + 1 & "/" & CStr(iMins) & ") to ensure this is not a transaction mid-state", True
            oRS.Close
            Set oRS = New ADODB.Recordset
            SleepAPI 60000
            oRS.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
            iIterations = iIterations + 1
        Loop
    
    End If
    
    Set WaitForRSToBecomeEOF = oRS

End Function

Private Sub ValidateClientGEOTable(sTableName As String, _
                                   cn As ADODB.Connection, _
                                   sUniqueGUID As String, _
                                   Optional sLocalSynchHistPrefix As String = "")
        '<EhHeader>
        On Error GoTo ValidateClientGEOTable_Err
        '</EhHeader>
        
        'Exit Sub
        'Temporarily removing this integrity check since we are having problems
    
        Dim sSQL As String
        Dim sTableName_GEO As String
        Dim sTableName_FEA As String
        Dim oRS As ADODB.Recordset
        Dim iMins As Integer
        
100     If sTableName = "oincidents" Then
102         iMins = 5
        Else
104         iMins = 1
        End If
        
106     sTableName_GEO = sTableName & "_GEO"
108     sTableName_FEA = sTableName & "_FEA"
        
110     sSQL = "SELECT [" & sTableName_FEA & "].UID, [" & sTableName_FEA & "].[" & sUniqueGUID & "] FROM [" & sTableName_FEA & "] LEFT JOIN [" & sTableName_GEO & "] ON [" & sTableName_FEA & "].UID = [" & sTableName_GEO & "].UID WHERE ((([" & sTableName_GEO & "].UID) Is Null));"
112     Set oRS = WaitForRSToBecomeEOF(iMins, sSQL, cn)

114     If Not oRS.EOF Then
        
116         Do Until oRS.EOF
               
118             TxMsg "!!!!!! UID [" & oRS.Fields(0).Value & "] in table [" & sTableName_FEA & "] not found in table [" & sTableName_GEO & "] - removing all entries from table [" & sTableName_FEA & "] and " & sLocalSynchHistPrefix & "SynchHistory", True
        
120             cn.Execute "DELETE FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [sID] = '" & oRS.Fields(1).Value & "'"
122             cn.Execute "DELETE FROM [" & sTableName_FEA & "] WHERE [UID] = " & oRS.Fields(0).Value
124             oRS.MoveNext
        
            Loop
        
        End If

126     sSQL = "SELECT [" & sTableName_GEO & "].UID FROM [" & sTableName_GEO & "] LEFT JOIN [" & sTableName_FEA & "] ON [" & sTableName_GEO & "].UID = [" & sTableName_FEA & "].UID WHERE ((([" & sTableName_FEA & "].UID) Is Null));"
128     Set oRS = WaitForRSToBecomeEOF(iMins, sSQL, cn)

130     If Not oRS.EOF Then

132         Do Until oRS.EOF
               
134             TxMsg "!!!!!! UID [" & oRS.Fields(0).Value & "] in table [" & sTableName_GEO & "] not found in table [" & sTableName_FEA & "] - removing all entries from table [" & sTableName_GEO & "] but cannot remove from " & sLocalSynchHistPrefix & "SynchHistory - hope this does not cause problems", True
        
136             cn.Execute "DELETE FROM [" & sTableName_GEO & "] WHERE [UID] = " & oRS.Fields(0).Value
138             oRS.MoveNext
        
            Loop
        
        End If

140     Set oRS = Nothing
 
        '<EhFooter>
        Exit Sub

ValidateClientGEOTable_Err:
        TxMsg "SynchWorker.ValidateClientGEOTable_Err: (" & Erl & ") " & Err.Description, True
       
        '</EhFooter>
End Sub

Private Sub CheckForClientTableDuplicateUID(sTableName As String, _
                                            cn As ADODB.Connection, _
                                            sUniqueGUID As String, _
                                            Optional sLocalSynchHistPrefix As String = "")
        '<EhHeader>
        On Error GoTo CheckForClientTableDuplicateUID_Err
        '</EhHeader>
    
        Dim sSQL As String
        Dim sTableName_GEO As String
        Dim sTableName_FEA As String
100     sTableName_GEO = sTableName & "_GEO"
102     sTableName_FEA = sTableName & "_FEA"
    
        Dim RSLocalDuplicates As ADODB.Recordset
        Dim RSLocalDuplicates_BYUID As ADODB.Recordset
    
        'CHECK FEA table for duplicates
    
104     Set RSLocalDuplicates = New ADODB.Recordset
106     sSQL = "SELECT [UID] FROM [" & sTableName_FEA & "] GROUP BY [UID] HAVING (COUNT([UID]) > 1)"
108     RSLocalDuplicates.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
        
110     Do Until RSLocalDuplicates.EOF
               
112         TxMsg "!!!!!! Duplicate [UID] found in table [" & sTableName_FEA & "] - removing all entries from FEA and GEO tables and " & sLocalSynchHistPrefix & "SynchHistory for value: " & RSLocalDuplicates.Fields(0).Value, True
        
114         Set RSLocalDuplicates_BYUID = New ADODB.Recordset
116         sSQL = "SELECT [UID], [" & sUniqueGUID & "] FROM [" & sTableName_FEA & "] WHERE [UID] = " & RSLocalDuplicates.Fields(0).Value
118         RSLocalDuplicates_BYUID.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
        
120         Do Until RSLocalDuplicates_BYUID.EOF
                
122             cn.Execute "DELETE FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [sID] = '" & RSLocalDuplicates_BYUID.Fields(1).Value & "'"
124             RSLocalDuplicates_BYUID.MoveNext
                
            Loop
        
126         Set RSLocalDuplicates_BYUID = Nothing
        
128         cn.Execute "DELETE FROM [" & sTableName_GEO & "] WHERE [UID] = " & RSLocalDuplicates.Fields(0).Value
130         cn.Execute "DELETE FROM [" & sTableName_FEA & "] WHERE [UID] = " & RSLocalDuplicates.Fields(0).Value
132         RSLocalDuplicates.MoveNext
        
        Loop
    
        'CHECK GEO table for duplicates
    
134     Set RSLocalDuplicates = New ADODB.Recordset
136     sSQL = "SELECT [UID] FROM [" & sTableName_GEO & "] GROUP BY [UID] HAVING (COUNT([UID]) > 1)"
138     RSLocalDuplicates.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
        
140     Do Until RSLocalDuplicates.EOF
               
142         TxMsg "!!!!!! Duplicate [UID] found in table [" & sTableName_GEO & "] - removing all entries from FEA and GEO tables (but not from " & sLocalSynchHistPrefix & "SynchHistory) for value: " & RSLocalDuplicates.Fields(0).Value, True
144         cn.Execute "DELETE FROM [" & sTableName_GEO & "] WHERE [UID] = " & RSLocalDuplicates_BYUID.Fields(0).Value
146         cn.Execute "DELETE FROM [" & sTableName_FEA & "] WHERE [UID] = " & RSLocalDuplicates_BYUID.Fields(0).Value
148         RSLocalDuplicates.MoveNext
        
        Loop
    
150     Set RSLocalDuplicates = Nothing
 
        '<EhFooter>
        Exit Sub

CheckForClientTableDuplicateUID_Err:
                  
        TxMsg "SynchWorker.CheckForClientTableDuplicateUID_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub
    
Private Function GetSynchTablesToWipe(ByVal sWebsite As String, _
                                      cn As ADODB.Connection, _
                                      sPrefix As String)
        '<EhHeader>
        On Error GoTo GetSynchTablesToWipe_Err
        '</EhHeader>
    
        Dim RSServerSynchHistory As New ADODB.Recordset
        Dim RSLocalSynchHistory As New ADODB.Recordset
        Dim sSQL As String
        Dim bDropFlag As Boolean
    
100     TxMsg "<<<<<<<<<<<<", , True
102     TxMsg "<<< Scanning for tables to be wiped......................", , True

104     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then sWebsite = sWebsite & "/"
106     sSQL = "SELECT * from [" & sPrefix & "SynchHistory] WHERE [sequence] < 0"
    
108     RSLocalSynchHistory.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
110     Set RSServerSynchHistory = OpenServerRSCompressed(sWebsite, "ID", sSQL)
    
112     If Not RSLocalSynchHistory.State = 0 And Not RSServerSynchHistory.State = 0 Then
    
114         With RSLocalSynchHistory
   
116             Do Until RSServerSynchHistory.EOF
    
118                 If Not .EOF Or Not .BOF Then .MoveFirst
120                 bDropFlag = False
122                 .Find "[sTableName] = '" & RSServerSynchHistory.Fields("sTableName").Value & "'"
        
124                 If .EOF Then
126                     .AddNew
128                     .Fields("sID").Value = RSServerSynchHistory.Fields("sID").Value
130                     .Fields("sGUID").Value = RSServerSynchHistory.Fields("sGUID").Value
132                     .Fields("sTableName").Value = RSServerSynchHistory.Fields("sTableName").Value
134                     .Fields("swhen").Value = RSServerSynchHistory.Fields("swhen").Value
136                     .Fields("sStatus").Value = RSServerSynchHistory.Fields("sStatus").Value
                        '.Fields("sequence").Value = RSServerSynchHistory.Fields("sequence").Value
138                     .Fields("sBy").Value = RSServerSynchHistory.Fields("sBy").Value
140                     .Fields("sdelete").Value = RSServerSynchHistory.Fields("sdelete").Value
142                     .Fields("updates").Value = RSServerSynchHistory.Fields("updates").Value
144                     .Fields("noconflict").Value = RSServerSynchHistory.Fields("noconflict").Value
146                     bDropFlag = True
                    Else

148                     If .Fields("sequence").Value > RSServerSynchHistory.Fields("sequence").Value Then bDropFlag = True
                    End If
            
150                 If bDropFlag Then
                
152                     .Fields("sequence").Value = RSServerSynchHistory.Fields("sequence").Value
154                     .UpdateBatch adAffectCurrent
156                     TxMsg "!!! Wiping table: " & .Fields("sTableName").Value
158                     cn.Execute "DELETE FROM [" & .Fields("sTableName").Value & "]"
160                     cn.Execute "DELETE FROM [" & sPrefix & "SynchHistory] WHERE [sequence] > 0 AND [sTableName] = '" & .Fields("sTableName").Value & "'"
            
                        If Right$(.Fields("sTableName").Value, 4) = "_FEA" Then
                            cn.Execute "DELETE FROM [" & Left$(.Fields("sTableName").Value, Len(.Fields("sTableName").Value) - 4) & "_GEO]"
                        End If
            
                    End If

162                 RSServerSynchHistory.MoveNext
                Loop
        
164             .Close
166             RSServerSynchHistory.Close
        
            End With
    
        End If
        
168     TxMsg "<<< End scan for tables to be wiped......................", , True
170     TxMsg "<<<<<<<<<<<<", , True
    
172     Set RSServerSynchHistory = Nothing
174     Set RSLocalSynchHistory = Nothing
                
        '<EhFooter>
        Exit Function

GetSynchTablesToWipe_Err:
        TxMsg "SynchWorker.GetSynchTablesToWipe_Err: (line: " & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Function GetMaxUIDFromServerGEOTable(sTableName As String, _
                                             sWebsite As String) As Long
        '<EhHeader>
        On Error GoTo GetMaxUIDFromServerGEOTable_Err
        '</EhHeader>
        
        'Operation: This function returns the maximum UID in the specified
        '           table
        '
        'Assumption: The table has a UID field and is of type long

        Dim RSServerMaxUID As New ADODB.Recordset
        Dim sSQL As String
    
100     sSQL = "SELECT MAX(UID) FROM [" & sTableName & "]"
102     Set RSServerMaxUID = OpenServerRSCompressed(sWebsite, "ID", sSQL)

104     GetMaxUIDFromServerGEOTable = 0

106     With RSServerMaxUID
       
108         If Not .EOF And Not .BOF Then
110             If Not IsNull(.Fields(0).Value) Then
112                 GetMaxUIDFromServerGEOTable = CLng(.Fields(0).Value)
                End If
            End If
        
114         If .State = adStateOpen Then .Close

        End With
    
116     Set RSServerMaxUID = Nothing

        '<EhFooter>
        Exit Function

GetMaxUIDFromServerGEOTable_Err:
        TxMsg "GetMaxUIDFromServerGEOTable_Err line(" & Erl & ") " & Err.Desc, True
        GetMaxUIDFromServerGEOTable = 0
        '</EhFooter>
End Function

Private Function GetMaxUIDFromGEOTable(sTableName As String, _
                                       oCOnn As ADODB.Connection) As Long
        '<EhHeader>
        On Error GoTo GetMaxUIDFromGEOTable_Err
        '</EhHeader>
        
        'Operation: This function returns the maximum UID in the specified
        '           table
        '
        'Assumption: The table has a UID field and is of type long

        Dim RSLocalMaxUID As ADODB.Recordset
100     Set RSLocalMaxUID = New ADODB.Recordset
    
102     GetMaxUIDFromGEOTable = 0

104     With RSLocalMaxUID

'            .Open "SELECT UID FROM [" & sTableName & "] WHERE UID = 1000000000", oCOnn, adOpenDynamic, adLockBatchOptimistic
'
'            If .EOF Or .BOF Then
'
'                GetMaxUIDFromGEOTable = 999999999
'
'            Else
'                .Close
106             .Open "SELECT MAX(UID) FROM [" & sTableName & "] WHERE UID < 1000000000", oCOnn, adOpenDynamic, adLockBatchOptimistic
                
108             If Not .EOF And Not .BOF Then
110                 If Not IsNull(.Fields(0).Value) Then
112                     GetMaxUIDFromGEOTable = CLng(.Fields(0).Value)
                    End If
                End If
            
'            End If
        
114         If .State = adStateOpen Then .Close

        End With

116     Set RSLocalMaxUID = Nothing

        '<EhFooter>
        Exit Function

GetMaxUIDFromGEOTable_Err:
        TxMsg "GetMaxUIDFromGEOTable_Err line(" & Erl & ") " & Err.Desc, True
        GetMaxUIDFromGEOTable = 0
        
        '</EhFooter>
End Function

Private Function DoesFieldExistInRS(oRS As ADODB.Recordset, _
                                    sField As String) As Boolean

    On Error GoTo hell

    DoesFieldExistInRS = True
    
    Dim fField As ADODB.Field
    Set fField = oRS.Fields(sField)

    If fField Is Nothing Then GoTo hell
    Exit Function
hell:

    DoesFieldExistInRS = False

End Function

Function URLEncode(strWhat As String) As String

    Dim i As Long

    strWhat = Replace$(strWhat, "%", "%25")

    For i = 0 To 32
        strWhat = Replace$(strWhat, Chr$(i), "%" & IIf(i < 16, "0" & Hex$(i), Hex$(i)))
    Next i

    strWhat = Replace$(strWhat, "&", "%26")
    strWhat = Replace$(strWhat, "+", "%2B")

    'Return
    URLEncode = strWhat

End Function

Private Sub PrepareHttpComms(oHttp As WinHttpRequest, _
                             sWebsite As String, _
                             bUpload As Boolean)
        '<EhHeader>
        On Error GoTo PrepareHttpComms_Err
        '</EhHeader>
    
100     oHttp.abort
        Set oHttp = Nothing
        Set oHttp = New WinHttpRequest

102     If m_bProxyEnabled Then oHttp.setProxy HTTPREQUEST_PROXYSETTING_PROXY, m_sProxyIP & ":" & m_sProxyPort, "*.microsoft.com"
104     oHttp.setTimeouts -1, 360000, 360000, 360000
106     oHttp.Option(WinHttpRequestOption_EnableHttp1_1) = False
108     oHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 13056  'dec equivalent to hex 0x3300
110     oHttp.Open "POST", sWebsite & IIf(bUpload, "oasisuploader4", "oasis4") & ".asp", False
        ' TxMsg ">>>>> Server ASP: " & sWebsite & IIf(bUpload, "oasisuploader4", "oasis4") & ".asp", True, True
114     oHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded;charset=UTF-8"
116     oHttp.setRequestHeader "Expires", "0"
118     oHttp.setRequestHeader "Cache-Control", "no-cache"
120     oHttp.setRequestHeader "Pragma", "no-cache"
        'oHttp.SetCredentials sUser, sPwd, HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
    
        '<EhFooter>
        Exit Sub

PrepareHttpComms_Err:
        TxMsg "!!     -- PrepareHttpComms_Err: (" & Erl & ") " & Err.Description, True
        Resume Next
        '</EhFooter>
End Sub

Private Function RecordsetFromXMLString(sXML As String) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo RecordsetFromXMLString_Err
        '</EhHeader>

        Dim oStream As ADODB.Stream
        Dim oRecordset As ADODB.Recordset
        
100     Set oStream = New ADODB.Stream
102     oStream.Open
104     oStream.WriteText sXML   'Give the XML string to the ADO Stream
106     oStream.Position = 0    'Set the stream position to the start

108     Set oRecordset = New ADODB.Recordset
110     oRecordset.CursorLocation = adUseClient
112     oRecordset.LockType = adLockBatchOptimistic
114     oRecordset.CursorType = adOpenDynamic

116     oRecordset.Open oStream    'Open a recordset from the stream
118     oStream.Close

120     Set oStream = Nothing
122     Set RecordsetFromXMLString = oRecordset  'Return the recordset
        'Set oRecordset = Nothing

        '<EhFooter>
        Exit Function

RecordsetFromXMLString_Err:
        TxMsg "RecordsetFromXMLString_Err: (" & Erl & ") " & Err.Description, True
        Set RecordsetFromXMLString = Nothing
        '</EhFooter>
End Function

Private Function SynchDDWithServer2(sTableNamePrefix As String, _
                                    sWebsite As String, _
                                    cn As ADODB.Connection) As Boolean
        '<EhHeader>
        On Error GoTo SynchDDWithServer2_Err
        '</EhHeader>
        
        'return true = all records downloaded for the tableset
        
        Dim oRS_GEO As ADODB.Recordset
        Dim lUID As Long
        Dim sResponse As String
        Dim lMismatchIndex As Long
        Dim lNumRecordsLocal As Long
        Dim lLowVal As Long
        Dim lHighVal As Long
        Dim i As Long
        Dim iTablename As Long
        Dim l As Long
        Dim oRSLocal As ADODB.Recordset
        Dim oRSLocal_GEO As ADODB.Recordset
        Dim oRSServer As ADODB.Recordset
        Dim oRSServer_GEO As ADODB.Recordset
        Dim oRSLocalSynchHistory As ADODB.Recordset
        Dim oRSServerSynchHistory As ADODB.Recordset
        Dim sSynchHistName As String
        Dim bGetServer As Boolean
        Dim bGetLocal As Boolean
        Dim sIDs() As String
        Dim sTableNames() As String
        Dim sTableName As String
        Dim sSQL As String
        Dim sSQLCommand As String
        Dim bServerWinner As Boolean
        Dim bSame As Boolean
        Dim oSynchOperations() As SynchOperations
        Dim sLastOperation As String
        Dim sCurrentOperation As String
        Dim sGUIDName As String
        Dim lNumSynchRecords As Long
        Dim lNumCurrentSynchRecords As Long
        Dim bIsGeoTable As Boolean
    
        Dim oRSMerged As New ADODB.Recordset
100     oRSMerged.Fields.Append "sID", adVarChar, 255
102     oRSMerged.Fields.Append "sGUID", adVarChar, 255
104     oRSMerged.Fields.Append "sTableName", adVarChar, 255
106     oRSMerged.Fields.Append "sWhen", adVarChar, 255
108     oRSMerged.Fields.Append "Sequence", adBigInt
110     oRSMerged.Fields.Append "sDelete", adVarChar, 255
112     oRSMerged.Fields.Append "Updates", adVarChar, 255
114     oRSMerged.Fields.Append "OnServer", adBoolean
116     oRSMerged.Fields.Append "ToAddLocal", adBoolean
118     oRSMerged.Fields.Append "ToUpdateLocal", adBoolean
120     oRSMerged.Fields.Append "ToDeleteLocal", adBoolean
122     oRSMerged.Fields.Append "ToAddServer", adBoolean
124     oRSMerged.Fields.Append "ToUpdateServer", adBoolean
126     oRSMerged.Fields.Append "ToDeleteServer", adBoolean
128     oRSMerged.Open
    
130     sSynchHistName = sTableNamePrefix & "SynchHistory"
132     m_sIDs = ""

        'GetSynchTablesToWipe sWebsite, cn, sTableNamePrefix
134     lMismatchIndex = CompareSynchHashes(sSynchHistName, cn)
136     lNumSynchRecords = lMismatchIndex * 200
        
138     If lMismatchIndex > -1 Then
    
140         SynchDDWithServer2 = False
            'TxMsg "??---- There are about " & lNumSynchRecords & " records to synch in this batch", True
    
142         lLowVal = (lMismatchIndex * 200) - 199
144         lHighVal = lMismatchIndex * 200
146         lNumRecordsLocal = 0
        
148         Set oRSLocal = New ADODB.Recordset
150         oRSLocal.Open "SELECT * FROM [" & sSynchHistName & "] ORDER BY [stablename], [sID]", cn, adOpenDynamic, adLockBatchOptimistic

152         If lLowVal > 1 Then oRSLocal.Move (lLowVal - 1)
        
154         Set oRSServer = OpenServerRSCompressed(sWebsite, "ID", "SELECT * FROM [" & sSynchHistName & "] ORDER BY stablename, sID OFFSET " & (lLowVal - 1) & " ROWS FETCH NEXT 200 ROWS ONLY")

156         Do Until (oRSLocal.EOF) Or (lNumRecordsLocal >= 200)
        
158             With oRSLocal
    
160                 oRSMerged.AddNew
162                 oRSMerged.Fields("sID").Value = .Fields("sID").Value
164                 oRSMerged.Fields("sGUID").Value = .Fields("sGUID").Value
166                 oRSMerged.Fields("sTableName").Value = .Fields("sTableName").Value
168                 oRSMerged.Fields("sWhen").Value = .Fields("sWhen").Value
170                 oRSMerged.Fields("Sequence").Value = .Fields("Sequence").Value
172                 oRSMerged.Fields("sDelete").Value = .Fields("sDelete").Value
174                 oRSMerged.Fields("Updates").Value = .Fields("Updates").Value
176                 oRSMerged.Fields("OnServer").Value = False
    
                End With
                    
178             oRSLocal.MoveNext
180             lNumRecordsLocal = lNumRecordsLocal + 1
            
            Loop
            
182         Set oRSLocal = Nothing
            
184         Do Until (oRSServer.EOF)
        
186             With oRSServer
    
188                 oRSMerged.AddNew
190                 oRSMerged.Fields("sID").Value = .Fields("sID").Value
192                 oRSMerged.Fields("sGUID").Value = .Fields("sGUID").Value
194                 oRSMerged.Fields("sTableName").Value = .Fields("sTableName").Value
196                 oRSMerged.Fields("sWhen").Value = .Fields("sWhen").Value
198                 oRSMerged.Fields("Sequence").Value = .Fields("Sequence").Value
200                 oRSMerged.Fields("sDelete").Value = .Fields("sDelete").Value
202                 oRSMerged.Fields("Updates").Value = .Fields("Updates").Value
204                 oRSMerged.Fields("OnServer").Value = True
    
                End With

206             oRSServer.MoveNext
            Loop
            
208         Set oRSServer = Nothing
        
210         oRSMerged.Sort = "sID"

212         If oRSMerged.RecordCount > 0 Then oRSMerged.MoveFirst
214         ReDim sIDs(201)
216         sIDs(0) = ""
218         iTablename = 1
220         i = 1
        
222         Do Until oRSMerged.EOF Or (i > 200)
        
224             If oRSMerged.Fields("sID").Value <> sIDs(i - 1) Then
226                 sIDs(i) = oRSMerged.Fields("sID").Value
                    'TxMsg CStr(i) & " -- " & sIDs(i), True
228                 i = i + 1
                End If
            
230             oRSMerged.MoveNext
        
            Loop
        
232         bSame = False
234         l = 1

236         Do Until l = i Or oRSMerged.RecordCount = 0
               
238             oRSMerged.MoveFirst
240             oRSMerged.Filter = "[sID] = '" & sIDs(l) & "'"
                'TxMsg CStr(l) & " -- " & sIDs(l) & " --- count: " & oRSMerged.RecordCount, True
            
242             If oRSMerged.RecordCount = 1 Then

                    'either new record or local or server
244                 If oRSMerged.Fields("OnServer").Value = True Then
                
246                     If oRSMerged.Fields("sDelete").Value = "true" Then
                            'Delete record on local
248                         oRSMerged.Fields("ToAddLocal").Value = False
250                         oRSMerged.Fields("ToUpdateLocal").Value = False
252                         oRSMerged.Fields("ToDeleteLocal").Value = True
254                         oRSMerged.Fields("ToAddServer").Value = False
256                         oRSMerged.Fields("ToUpdateServer").Value = False
258                         oRSMerged.Fields("ToDeleteServer").Value = False
                        Else
                            'Add new record to local
260                         oRSMerged.Fields("ToAddLocal").Value = True
262                         oRSMerged.Fields("ToUpdateLocal").Value = False
264                         oRSMerged.Fields("ToDeleteLocal").Value = False
266                         oRSMerged.Fields("ToAddServer").Value = False
268                         oRSMerged.Fields("ToUpdateServer").Value = False
270                         oRSMerged.Fields("ToDeleteServer").Value = False
                        End If
                    
                    Else
                    
272                     If oRSMerged.Fields("sDelete").Value = "true" Then
                            'Delete record on server
274                         oRSMerged.Fields("ToAddLocal").Value = False
276                         oRSMerged.Fields("ToUpdateLocal").Value = False
278                         oRSMerged.Fields("ToDeleteLocal").Value = False
280                         oRSMerged.Fields("ToAddServer").Value = False
282                         oRSMerged.Fields("ToUpdateServer").Value = False
284                         oRSMerged.Fields("ToDeleteServer").Value = True
                        Else
                            'Add new record to server
286                         oRSMerged.Fields("ToAddLocal").Value = False
288                         oRSMerged.Fields("ToUpdateLocal").Value = False
290                         oRSMerged.Fields("ToDeleteLocal").Value = False
292                         oRSMerged.Fields("ToAddServer").Value = True
294                         oRSMerged.Fields("ToUpdateServer").Value = False
296                         oRSMerged.Fields("ToDeleteServer").Value = False
                        End If
                    
                    End If
                
                Else

                    'need to compare both records
                    Dim oRecordLocal As New ADODB.Recordset
                    Dim oRecordServer As New ADODB.Recordset
                
298                 If oRSMerged.Fields("OnServer").Value = True Then
300                     Set oRecordServer = CloneRecordFromRS(oRSMerged)
302                     oRSMerged.MoveNext
304                     Set oRecordLocal = CloneRecordFromRS(oRSMerged)
                    Else
306                     Set oRecordLocal = CloneRecordFromRS(oRSMerged)
308                     oRSMerged.MoveNext
310                     Set oRecordServer = CloneRecordFromRS(oRSMerged)
                    End If
                
312                 bSame = False
314                 oRSMerged.MoveFirst
                
316                 If oRecordLocal.Fields("sequence").Value = oRecordServer.Fields("sequence").Value Then

318                     If oRecordLocal.Fields("sWhen").Value = oRecordServer.Fields("sWhen").Value Then
320                         oRSMerged.Delete
322                         oRSMerged.MoveNext
324                         oRSMerged.Delete
326                     ElseIf oRecordLocal.Fields("sWhen").Value > oRecordServer.Fields("sWhen").Value Then

328                         If oRSMerged.Fields("OnServer").Value Then
330                             oRSMerged.Delete
332                             oRSMerged.MoveNext
                            Else
334                             oRSMerged.MoveNext
336                             oRSMerged.Delete
338                             oRSMerged.MovePrevious
                            End If
                            
340                         If oRecordLocal.Fields("sdelete").Value = "true" Then

342                             oRSMerged.Fields("ToAddLocal").Value = False
344                             oRSMerged.Fields("ToUpdateLocal").Value = False
346                             oRSMerged.Fields("ToDeleteLocal").Value = False
348                             oRSMerged.Fields("ToAddServer").Value = False
350                             oRSMerged.Fields("ToUpdateServer").Value = False
352                             oRSMerged.Fields("ToDeleteServer").Value = True

                            Else
                            
354                             oRSMerged.Fields("ToAddLocal").Value = False
356                             oRSMerged.Fields("ToUpdateLocal").Value = False
358                             oRSMerged.Fields("ToDeleteLocal").Value = False
360                             oRSMerged.Fields("ToAddServer").Value = False
362                             oRSMerged.Fields("ToUpdateServer").Value = True
364                             oRSMerged.Fields("ToDeleteServer").Value = False
                            
                            End If

                        Else

366                         If Not oRSMerged.Fields("OnServer").Value Then
368                             oRSMerged.Delete
370                             oRSMerged.MoveNext
                            Else
372                             oRSMerged.MoveNext
374                             oRSMerged.Delete
376                             oRSMerged.MovePrevious
                            End If

378                         If oRecordServer.Fields("sdelete").Value = "true" Then
                            
380                             oRSMerged.Fields("ToAddLocal").Value = False
382                             oRSMerged.Fields("ToUpdateLocal").Value = False
384                             oRSMerged.Fields("ToDeleteLocal").Value = True
386                             oRSMerged.Fields("ToAddServer").Value = False
388                             oRSMerged.Fields("ToUpdateServer").Value = False
390                             oRSMerged.Fields("ToDeleteServer").Value = False

                            Else

392                             oRSMerged.Fields("ToAddLocal").Value = False
394                             oRSMerged.Fields("ToUpdateLocal").Value = True
396                             oRSMerged.Fields("ToDeleteLocal").Value = False
398                             oRSMerged.Fields("ToAddServer").Value = False
400                             oRSMerged.Fields("ToUpdateServer").Value = False
402                             oRSMerged.Fields("ToDeleteServer").Value = False

                            End If

                        End If
                    
404                 ElseIf oRecordLocal.Fields("sequence").Value > oRecordServer.Fields("sequence").Value Then

406                     If oRSMerged.Fields("OnServer").Value Then
408                         oRSMerged.Delete
410                         oRSMerged.MoveNext
                        Else
412                         oRSMerged.MoveNext
414                         oRSMerged.Delete
416                         oRSMerged.MovePrevious
                        End If

418                     If oRecordLocal.Fields("sdelete") = "true" Then
420                         oRSMerged.Fields("ToAddLocal").Value = False
422                         oRSMerged.Fields("ToUpdateLocal").Value = False
424                         oRSMerged.Fields("ToDeleteLocal").Value = False
426                         oRSMerged.Fields("ToAddServer").Value = False
428                         oRSMerged.Fields("ToUpdateServer").Value = False
430                         oRSMerged.Fields("ToDeleteServer").Value = True
                        Else
432                         oRSMerged.Fields("ToAddLocal").Value = False
434                         oRSMerged.Fields("ToUpdateLocal").Value = False
436                         oRSMerged.Fields("ToDeleteLocal").Value = False
438                         oRSMerged.Fields("ToAddServer").Value = False
440                         oRSMerged.Fields("ToUpdateServer").Value = True
442                         oRSMerged.Fields("ToDeleteServer").Value = False
                        End If
                                
                    Else

444                     If Not oRSMerged.Fields("OnServer").Value Then
446                         oRSMerged.Delete
448                         oRSMerged.MoveNext
                        Else
450                         oRSMerged.MoveNext
452                         oRSMerged.Delete
454                         oRSMerged.MovePrevious
                        End If

456                     If oRecordServer.Fields("sdelete") = "true" Then
458                         oRSMerged.Fields("ToAddLocal").Value = False
460                         oRSMerged.Fields("ToUpdateLocal").Value = False
462                         oRSMerged.Fields("ToDeleteLocal").Value = True
464                         oRSMerged.Fields("ToAddServer").Value = False
466                         oRSMerged.Fields("ToUpdateServer").Value = False
468                         oRSMerged.Fields("ToDeleteServer").Value = False
                        Else
470                         oRSMerged.Fields("ToAddLocal").Value = False
472                         oRSMerged.Fields("ToUpdateLocal").Value = True
474                         oRSMerged.Fields("ToDeleteLocal").Value = False
476                         oRSMerged.Fields("ToAddServer").Value = False
478                         oRSMerged.Fields("ToUpdateServer").Value = False
480                         oRSMerged.Fields("ToDeleteServer").Value = False
                        End If
                                
                    End If

                End If
            
482             l = l + 1
484             oRSMerged.Filter = adFilterNone
            
            Loop
            
486         If oRSMerged.RecordCount = 0 Then
            
488             TxMsg ">>>> No records left - both server and local sets match", True
            
            Else
        
490             oRSMerged.MoveFirst
492             oRSMerged.Filter = adFilterNone
494             oRSMerged.Sort = "sTableName"
            
496             If oRSMerged.RecordCount > 0 Then
498                 oRSMerged.MoveFirst
                    'TxMsg "oRSMerged.Fields(ToUpdateServer).Value: " & oRSMerged.Fields("ToUpdateServer").Value, True
                    'TxMsg "oRSMerged.Fields(ToUpdateLocal).Value: " & oRSMerged.Fields("ToUpdateLocal").Value, True
                End If

                'TxMsg "### " & oRSMerged.RecordCount & " records need attention", True
    
500             ReDim sTableNames(1)
502             iTablename = 1
504             i = 1
        
506             Do Until oRSMerged.EOF
            
508                 If oRSMerged.Fields("sTableName").Value <> sTableNames(iTablename - 1) Then
510                     ReDim Preserve sTableNames(iTablename + 1)
512                     sTableNames(iTablename) = oRSMerged.Fields("sTableName").Value
514                     iTablename = iTablename + 1
                        'TxMsg "### sTableName: " & oRSMerged.Fields("sTableName").Value, True
                    End If
            
516                 oRSMerged.MoveNext
        
                Loop
            
518             iTablename = iTablename - 1
520             ReDim oSynchOperations(1)

                'TxMsg "### tables: " & CStr(iTablename), True
            
                'sTableNames()                      = Tables to be checked
                'oRSMerged.Fields("ToAddLocal")     = Add record from server to local
                'oRSMerged.Fields("ToAddServer")    = Add record from local to server
                'oRSMerged.Fields("ToUpdateLocal")  = Update local record from server
                'oRSMerged.Fields("ToUpdateServer") = Update server record from local
                'oRSMerged.Fields("ToDeleteLocal")  = Delete local record
                'oRSMerged.Fields("ToDeleteServer") = Delete server record
            
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Now fetch records from local which need to be sent to the server
            
522             oRSMerged.Filter = adFilterNone
524             oRSMerged.Filter = "[ToAddServer] = true OR [ToUpdateServer] = true OR [ToAddLocal] = true OR [ToUpdateLocal] = true OR [ToDeleteServer] = true OR [ToDeleteLocal] = true"
526             oRSMerged.Sort = "[sTableName],[ToAddServer] DESC,[ToUpdateServer] DESC, [ToAddLocal] DESC, [ToUpdateLocal] DESC,[ToDeleteServer] DESC,[ToDeleteLocal] DESC"

528             If Not oRSMerged.RecordCount = 0 Then oRSMerged.MoveFirst
                'oRSMerged.Filter = "[ToAddServer] = true OR [ToUpdateServer] = true OR [ToAddLocal] = true OR [ToUpdateLocal] = true OR [ToDeleteServer] = true OR [ToDeleteLocal] = true"
530             i = 1
532             sSQL = ""

                'TxMsg "### before working on commits (oRSMerged.RecordCount) = " & oRSMerged.RecordCount, True

534             If Not oRSMerged.EOF Then
                
                    Do

536                     With oRSMerged
538                         sGUIDName = IIf(InStr(.Fields("sTableName").Value, "_link") > 1, "[GUID2]", IIf(.Fields("sTableName").Value = "oincidents", "[ID]", "[GUID1]"))
                            '  lNumCurrentSynchRecords = lNumCurrentSynchRecords + 1
                    
540                         If Not i = 1 Then .MoveNext
                
542                         If Not oRSMerged.EOF Then
544                             sCurrentOperation = .Fields("sTableName").Value & CStr(.Fields("ToAddLocal").Value) & CStr(.Fields("ToAddServer").Value) & CStr(.Fields("ToUpdateLocal").Value) & CStr(.Fields("ToUpdateServer").Value) & CStr(.Fields("ToDeleteLocal").Value) & CStr(.Fields("ToDeleteServer").Value)
                            Else
546                             sCurrentOperation = "LastRecord"
                            End If

                            'TxMsg "### loop on oRSMerged", True

                            'New table different from last record
548                         If (sCurrentOperation = "LastRecord") Or (Not i = 1 And Not sCurrentOperation = sLastOperation) Then
                    
                                'TxMsg "### working on commits", True
                    
                                'Perform operations so we can perform this operation
550                             If Not i = 1 Then .MovePrevious
552                             If Right(.Fields("sTableName").Value, 4) = "_FEA" Then bIsGeoTable = True
                            
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                'Get local and server recordsets for tables and synchhistory
554                             sSQLCommand = "SELECT * FROM [" & .Fields("sTableName").Value & "] WHERE" & Replace(sSQL, "[sID]", sGUIDName)
556                             Set oRSLocal = New ADODB.Recordset
558                             oRSLocal.Open sSQLCommand, cn, adOpenDynamic, adLockBatchOptimistic
                                If .Fields("sTableName").Value = "Incidents_ChartSettings" Or .Fields("sTableName").Value = "GeoBookmarks" Or .Fields("sTableName").Value = "GeoBookmarksCategories" Then
                                    sSQLCommand = "SELECT * FROM [" & m_StrRemoteTablePrefix & .Fields("sTableName").Value & "] WHERE" & Replace(sSQL, "[sID]", sGUIDName)
                                End If
560                             Set oRSServer = OpenServerRSCompressed(sWebsite, "ID", sSQLCommand)
                          
562                             sSQLCommand = "SELECT * FROM [" & sTableNamePrefix & "SynchHistory] WHERE" & sSQL
564                             Set oRSLocalSynchHistory = New ADODB.Recordset
566                             oRSLocalSynchHistory.Open sSQLCommand, cn, adOpenDynamic, adLockBatchOptimistic
568                             Set oRSServerSynchHistory = OpenServerRSCompressed(sWebsite, "ID", sSQLCommand)
                            
570                             If bIsGeoTable Then

572                                 sSQL = ""

574                                 Do Until oRSServer.EOF
576                                     sSQL = sSQL & " OR [UID]=" & oRSServer.Fields("UID").Value
578                                     oRSServer.MoveNext
                                    Loop

580                                 If Len(sSQL) > 2 Then
582                                     sSQL = Right(sSQL, Len(sSQL) - 4)
584                                     sSQL = "WHERE " & sSQL
                                    End If

586                                 If oRSServer.RecordCount > 0 Then oRSServer.MoveFirst

588                                 sSQLCommand = "SELECT * FROM [" & Left(.Fields("sTableName").Value, Len(.Fields("sTableName").Value) - 4) & "_GEO] " & sSQL
                                    'TxMsg Right(sSQLCommand, 50), True
590                                 Set oRSLocal_GEO = New ADODB.Recordset
592                                 oRSLocal_GEO.Open sSQLCommand, cn, adOpenDynamic, adLockBatchOptimistic
594                                 Set oRSServer_GEO = OpenServerRSCompressed(sWebsite, "ID", sSQLCommand)
                                End If
                                
                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            
596                             If .Fields("ToAddLocal").Value = True Then

                                    'TxMsg "### ToAddLocal", True
                                
                                    'Add to local from server
598                                 If oRSServerSynchHistory.RecordCount > 0 Then oRSServerSynchHistory.MoveFirst

600                                 Do Until oRSServerSynchHistory.EOF
602                                     lNumCurrentSynchRecords = lNumCurrentSynchRecords + 1
604                                     TxMsg "**??(" & lNumCurrentSynchRecords & ") appending record to local table [" & oRSServerSynchHistory.Fields("sTableName").Value & "] with ID: " & oRSServerSynchHistory.Fields("sID").Value
                                
606                                     oRSLocal.Filter = sGUIDName & " = '" & oRSServerSynchHistory.Fields("sID").Value & "'"
                                        'TxMsg sGUIDName & " = '" & oRSServerSynchHistory.Fields("sID").Value & "'", True
608                                     oRSServer.Filter = sGUIDName & " = '" & oRSServerSynchHistory.Fields("sID").Value & "'"
610                                     oRSLocalSynchHistory.Filter = "[sID] = '" & oRSServerSynchHistory.Fields("sID").Value & "'"
                                    
612                                     If Not oRSServer.EOF Then
                                            
614                                         If oRSLocal.EOF Then
616                                             oRSLocal.AddNew

618                                             If bIsGeoTable Then oRSLocal_GEO.AddNew
                                            Else

620                                             If bIsGeoTable Then oRSLocal_GEO.Filter = "UID = " & oRSLocal.Fields("UID").Value
                                            End If

622                                         UpdateRSRecordSingle oRSServer, oRSLocal
                                            
624                                         If bIsGeoTable Then
626                                             oRSServer_GEO.Filter = "UID = " & oRSServer.Fields("UID").Value

628                                             If oRSLocal_GEO.EOF Then oRSLocal_GEO.AddNew
630                                             UpdateRSRecordSingle oRSServer_GEO, oRSLocal_GEO
632                                             Set oRS_GEO = New ADODB.Recordset
634                                             oRS_GEO.Open "SELECT max(UID) from " & .Fields("sTableName").Value, cn, adOpenDynamic, adLockBatchOptimistic

636                                             If oRS_GEO.EOF Then
638                                                 lUID = 1
640                                             ElseIf IsNull(oRS_GEO.Fields(0).Value) Then
642                                                 lUID = 1
                                                Else
644                                                 lUID = oRS_GEO.Fields(0).Value + 1
                                                End If

646                                             oRS_GEO.Close
648                                             oRSLocal.Fields("UID").Value = lUID
650                                             oRSLocal_GEO.Fields("UID").Value = lUID
652                                             oRSLocal_GEO.UpdateBatch
                                            End If

654                                         oRSLocal.UpdateBatch adAffectCurrent
                                            If Right$(.Fields("sTableName").Value, 7) = "Queries" Then CreateQuery sTableNamePrefix & oRSLocal.Fields("QueryName").Value, oRSLocal.Fields("QuerySQL").Value, cn, sWebsite
                                        
656                                         If oRSLocalSynchHistory.EOF Then
658                                             oRSLocalSynchHistory.AddNew
                                            Else
                                                'ERROR: synchhistory record exists - allow override
                                            End If
                                        
660                                         UpdateRSRecordSingle oRSServerSynchHistory, oRSLocalSynchHistory
662                                         oRSLocalSynchHistory.Fields("sBy").Value = "New synch'd down on " & RFC3339DateTime
664                                         oRSLocalSynchHistory.UpdateBatch adAffectCurrent
                            
                                        Else
                                            'ERROR: server record does not exist only in SynchHistory table
                                            '>> may need to wipe the synchhistory row on the server
                                        End If
                                    
666                                     oRSServerSynchHistory.MoveNext
                                    
                                    Loop
                                
                                End If

668                             If .Fields("ToAddServer").Value = True Then

                                    'TxMsg "### ToAddServer", True
                               
                                    'Add to server from local
670                                 If oRSLocalSynchHistory.RecordCount > 0 Then oRSLocalSynchHistory.MoveFirst

672                                 Do Until oRSLocalSynchHistory.EOF
674                                     lNumCurrentSynchRecords = lNumCurrentSynchRecords + 1
676                                     TxMsg "**??(" & lNumCurrentSynchRecords & ") appending record to server table [" & oRSLocalSynchHistory.Fields("sTableName").Value & "] with ID: " & oRSLocalSynchHistory.Fields("sID").Value

678                                     oRSLocal.Filter = sGUIDName & " = '" & oRSLocalSynchHistory.Fields("sID").Value & "'"
680                                     oRSServer.Filter = sGUIDName & " = '" & oRSLocalSynchHistory.Fields("sID").Value & "'"
682                                     oRSServerSynchHistory.Filter = "[sID] = '" & oRSLocalSynchHistory.Fields("sID").Value & "'"
                                    
684                                     If Not oRSLocal.EOF Then
                                            
686                                         If oRSServer.EOF Then
688                                             oRSServer.AddNew
                                            Else
                                                'allow override
                                            End If

690                                         UpdateRSRecordSingle oRSLocal, oRSServer

                                            If bIsGeoTable Then
                                            If Not oRSLocal_GEO.RecordCount = 0 Then oRSLocal_GEO.MoveFirst
694                                             oRSLocal_GEO.Filter = "UID = " & oRSLocal.Fields("UID").Value

                                                sResponse = ""

                                                If Not oRSLocal_GEO.EOF Then

                                                    oRSServer_GEO.Filter = "UID = " & oRSLocal.Fields("UID").Value

696                                                 If oRSServer_GEO.EOF Then oRSServer_GEO.AddNew
698                                                 UpdateRSRecordSingle oRSLocal_GEO, oRSServer_GEO
700                                                 lUID = GetMaxUIDFromServerGEOTable(.Fields("sTableName").Value, sWebsite) + 1
702                                                 oRSServer.Fields("UID").Value = lUID
704                                                 oRSServer_GEO.Fields("UID").Value = lUID
706                                                 sResponse = SaveServerRSCompressed(sWebsite, oRSServer_GEO)

                                                Else
                                                    TxMsg "!!! Corresponding GEO record for UID " & oRSLocal.Fields("UID").Value & " was not found", True
                                                End If

                                            Else
708                                             sResponse = "Data Updated"
                                            End If

710                                         If Not sResponse = "Data Updated" Then

712                                             TxMsg "    --- Error updating oRSServer_GEO: " & sResponse, True

                                            Else

714                                             sResponse = SaveServerRSCompressed(sWebsite, oRSServer)

716                                             If Not sResponse = "Data Updated" Then

718                                                 TxMsg "    --- Error updating oRSServer: " & sResponse, True
                                        
                                                Else

720                                                 If oRSServerSynchHistory.EOF Then
722                                                     oRSServerSynchHistory.AddNew
                                                    Else
                                                        'ERROR: synchhistory record exists - allow override
                                                    End If
                                            
724                                                 oRSLocalSynchHistory.Fields("updates").Value = "true"
726                                                 oRSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
728                                                 UpdateRSRecordSingle oRSLocalSynchHistory, oRSServerSynchHistory
730                                                 oRSServerSynchHistory.Fields("sBy").Value = "New synch'd up on " & CalcServerDateTime(Now())
732                                                 sResponse = SaveServerRSCompressed(sWebsite, oRSServerSynchHistory)

734                                                 If Not sResponse = "Data Updated" Then
736                                                     TxMsg "    --- Error updating oRSServerSynchHistory: " & sResponse, True
                                                    Else
738                                                     oRSLocalSynchHistory.UpdateBatch adAffectCurrent
                                                    End If
                                        
                                                End If
                                        
                                            End If

                                        Else
                                            'ERROR: local record does not exist only in SynchHistory table
                                            '>> may need to wipe the synchhistory row on local
                                        End If
                                    
740                                     oRSLocalSynchHistory.MoveNext
                                    
                                    Loop
                                
                                End If

742                             If .Fields("ToUpdateLocal").Value = True Then
                                    'TxMsg "### ToUpdateLocal", True

                                    'Update local from server
744                                 If oRSServerSynchHistory.RecordCount > 0 Then oRSServerSynchHistory.MoveFirst

746                                 Do Until oRSServerSynchHistory.EOF

748                                     lNumCurrentSynchRecords = lNumCurrentSynchRecords + 1
750                                     TxMsg "**??(" & lNumCurrentSynchRecords & ") updating record to local table [" & oRSServerSynchHistory.Fields("sTableName").Value & "] with ID: " & oRSServerSynchHistory.Fields("sID").Value
                                
752                                     oRSLocal.Filter = sGUIDName & " = '" & oRSServerSynchHistory.Fields("sID").Value & "'"
754                                     oRSServer.Filter = sGUIDName & " = '" & oRSServerSynchHistory.Fields("sID").Value & "'"
756                                     oRSLocalSynchHistory.Filter = "[sID] = '" & oRSServerSynchHistory.Fields("sID").Value & "'"
                                    
758                                     If Not oRSServer.EOF Then
                                            
760                                         If oRSLocal.EOF Then
762                                             oRSLocal.AddNew

764                                             If bIsGeoTable Then oRSLocal_GEO.AddNew
                                            Else

766                                             If bIsGeoTable Then oRSLocal_GEO.Filter = "UID = " & oRSLocal.Fields("UID").Value
                                            End If

768                                         UpdateRSRecordSingle oRSServer, oRSLocal

770                                         If bIsGeoTable Then
772                                             oRSServer_GEO.Filter = "UID = " & oRSServer.Fields("UID").Value

774                                             If oRSLocal_GEO.EOF Then oRSLocal_GEO.AddNew
776                                             UpdateRSRecordSingle oRSServer_GEO, oRSLocal_GEO
778                                             Set oRS_GEO = New ADODB.Recordset
780                                             oRS_GEO.Open "SELECT max(UID) from " & .Fields("sTableName").Value, cn, adOpenDynamic, adLockBatchOptimistic

782                                             If oRS_GEO.EOF Then
784                                                 lUID = 1
                                                Else

786                                                 lUID = oRS_GEO.Fields(0).Value + 1
                                                End If

788                                             oRS_GEO.Close
790                                             oRSLocal.Fields("UID").Value = lUID
792                                             oRSLocal_GEO.Fields("UID").Value = lUID
794                                             oRSLocal_GEO.UpdateBatch
                                            End If

796                                         oRSLocal.UpdateBatch adAffectCurrent
                                            If Right$(.Fields("sTableName").Value, 7) = "Queries" Then CreateQuery sTableNamePrefix & oRSLocal.Fields("QueryName").Value, oRSLocal.Fields("QuerySQL").Value, cn, sWebsite
                                            
798                                         If oRSLocalSynchHistory.EOF Then
800                                             oRSLocalSynchHistory.AddNew
                                            Else
                                                'ERROR: synchhistory record exists - allow override
                                            End If
                                        
802                                         UpdateRSRecordSingle oRSServerSynchHistory, oRSLocalSynchHistory
804                                         oRSLocalSynchHistory.Fields("sBy").Value = "Update synch'd down on " & RFC3339DateTime
806                                         oRSLocalSynchHistory.UpdateBatch adAffectCurrent
                                    
                                        Else
                                            'ERROR: server record does not exist only in SynchHistory table
                                            '>> may need to wipe the synchhistory row on the server
                                        End If
                                    
808                                     oRSServerSynchHistory.MoveNext
                                    
                                    Loop
                                
                                End If

810                             If .Fields("ToUpdateServer").Value = True Then
                                    'TxMsg "### ToUpdateServer", True

                                    'Update server from local
812                                 If oRSLocalSynchHistory.RecordCount > 0 Then oRSLocalSynchHistory.MoveFirst

814                                 Do Until oRSLocalSynchHistory.EOF
816                                     lNumCurrentSynchRecords = lNumCurrentSynchRecords + 1
818                                     TxMsg "**??(" & lNumCurrentSynchRecords & ") updating record to server table [" & oRSLocalSynchHistory.Fields("sTableName").Value & "] with ID: " & oRSLocalSynchHistory.Fields("sID").Value

820                                     oRSLocal.Filter = sGUIDName & " = '" & oRSLocalSynchHistory.Fields("sID").Value & "'"
822                                     oRSServer.Filter = sGUIDName & " = '" & oRSLocalSynchHistory.Fields("sID").Value & "'"
824                                     oRSServerSynchHistory.Filter = "[sID] = '" & oRSLocalSynchHistory.Fields("sID").Value & "'"
                                        
826                                     If Not oRSLocal.EOF Then
                                            
828                                         If oRSServer.EOF Then oRSServer.AddNew
830                                         UpdateRSRecordSingle oRSLocal, oRSServer
                                        
832                                         If bIsGeoTable Then
834                                             oRSLocal_GEO.Filter = "UID = " & oRSLocal.Fields("UID").Value

836                                             If oRSServer_GEO.EOF Then oRSServer_GEO.AddNew
838                                             UpdateRSRecordSingle oRSLocal_GEO, oRSServer_GEO
840                                             lUID = GetMaxUIDFromServerGEOTable(.Fields("sTableName").Value, sWebsite) + 1
842                                             oRSServer.Fields("UID").Value = lUID
844                                             oRSServer_GEO.Fields("UID").Value = lUID
846                                             sResponse = SaveServerRSCompressed(sWebsite, oRSServer_GEO)
                                            Else
848                                             sResponse = "Data Updated"
                                            End If
                                        
850                                         If Not sResponse = "Data Updated" Then

852                                             TxMsg "    --- Error updating oRSServer: " & sResponse, True
                                        
                                            Else
                                        
854                                             sResponse = SaveServerRSCompressed(sWebsite, oRSServer)

856                                             If Not sResponse = "Data Updated" Then

858                                                 TxMsg "    --- Error updating oRSServer: " & sResponse, True
                                        
                                                Else

860                                                 If oRSServerSynchHistory.EOF Then
862                                                     oRSServerSynchHistory.AddNew
                                                    Else
                                                        'ERROR: synchhistory record exists - allow override
                                                    End If
                                            
864                                                 oRSLocalSynchHistory.Fields("updates").Value = "true"
866                                                 oRSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
868                                                 UpdateRSRecordSingle oRSLocalSynchHistory, oRSServerSynchHistory
870                                                 oRSServerSynchHistory.Fields("sBy").Value = "Update synch'd up on " & CalcServerDateTime(Now())
872                                                 sResponse = SaveServerRSCompressed(sWebsite, oRSServerSynchHistory)

874                                                 If Not sResponse = "Data Updated" Then
876                                                     TxMsg "    --- Error updating oRSServerSynchHistory: " & sResponse, True
                                                    Else
878                                                     oRSLocalSynchHistory.UpdateBatch adAffectCurrent
                                                    End If
                                        
                                                End If
                                        
                                            End If
                                        
                                        Else
                                            'ERROR: server record does not exist only in SynchHistory table
                                            '>> may need to wipe the synchhistory row on the server
                                        End If
                                    
880                                     oRSLocalSynchHistory.MoveNext
                                    
                                    Loop
                                
                                End If

882                             If .Fields("ToDeleteLocal").Value = True Then
                                    'TxMsg "### ToDeleteLocal", True
                                
                                    'Delete local
884                                 If oRSServerSynchHistory.RecordCount > 0 Then oRSServerSynchHistory.MoveFirst

886                                 Do Until oRSServerSynchHistory.EOF

888                                     lNumCurrentSynchRecords = lNumCurrentSynchRecords + 1
890                                     TxMsg "**??(" & lNumCurrentSynchRecords & ") deleting record from local table [" & oRSServerSynchHistory.Fields("sTableName").Value & "] with ID: " & oRSServerSynchHistory.Fields("sID").Value

892                                     oRSLocal.Filter = sGUIDName & " = '" & oRSServerSynchHistory.Fields("sID").Value & "'"
894                                     oRSServer.Filter = sGUIDName & " = '" & oRSServerSynchHistory.Fields("sID").Value & "'"
896                                     oRSLocalSynchHistory.Filter = "[sID] = '" & oRSServerSynchHistory.Fields("sID").Value & "'"
                                    
898                                     'If Not oRSServer.EOF Then
                                            'ERROR: server record exist - it should not!
                                            'TODO: clean up the records table
                                            'oRSServer.Delete adAffectCurrent
                                        'End If

900                                     If Not oRSLocal.EOF Then

902                                         If bIsGeoTable Then
904                                             oRSLocal_GEO.Filter = "UID = " & oRSLocal.Fields("UID").Value
906                                             If Not oRSLocal_GEO.EOF Then oRSLocal.Delete adAffectCurrent
908                                             oRSLocal_GEO.UpdateBatch
                                            End If

910                                         oRSLocal.Delete adAffectCurrent
912                                         oRSLocal.UpdateBatch 'adAffectCurrent
913                                         If Right$(.Fields("sTableName").Value, 7) = "Queries" Then DropLocalQuery sTableNamePrefix & oRSLocal.Fields("QueryName").Value, cn
                                            
                                        End If

914                                     If oRSLocalSynchHistory.EOF Then
916                                         oRSLocalSynchHistory.AddNew
                                        Else
                                            'Synchhistory record exists - allow override
                                        End If
                                        
918                                     UpdateRSRecordSingle oRSServerSynchHistory, oRSLocalSynchHistory
920                                     oRSLocalSynchHistory.Fields("sBy").Value = "Delete synch'd down on " & RFC3339DateTime
922                                     oRSLocalSynchHistory.UpdateBatch adAffectCurrent

924                                     oRSServerSynchHistory.MoveNext
                                    
                                    Loop

                                End If

926                             If .Fields("ToDeleteServer").Value = True Then
                                    'TxMsg "### ToDeleteServer", True

                                    'Delete server
928                                 If oRSLocalSynchHistory.RecordCount > 0 Then oRSLocalSynchHistory.MoveFirst

930                                 Do Until oRSLocalSynchHistory.EOF
932                                     lNumCurrentSynchRecords = lNumCurrentSynchRecords + 1
934                                     TxMsg "**??(" & lNumCurrentSynchRecords & ") deleting record from server table [" & oRSLocalSynchHistory.Fields("sTableName").Value & "] with ID: " & oRSLocalSynchHistory.Fields("sID").Value

936                                     oRSLocal.Filter = sGUIDName & " = '" & oRSLocalSynchHistory.Fields("sID").Value & "'"
938                                     oRSServer.Filter = sGUIDName & " = '" & oRSLocalSynchHistory.Fields("sID").Value & "'"
940                                     oRSServerSynchHistory.Filter = "[sID] = '" & oRSLocalSynchHistory.Fields("sID").Value & "'"
                                        
942                                     If Not oRSLocal.EOF Then
                                            'ERROR: local record exists! It should be deleted. Deleting...
944                                         oRSLocal.Delete adAffectCurrent
946                                         oRSLocal.UpdateBatch adAffectCurrent
                                        End If

948                                     If Not oRSServer.EOF Then
                                    
950                                         If bIsGeoTable Then
952                                             oRSServer_GEO.Filter = "UID = " & oRSServer.Fields("UID").Value

954                                             If Not oRSServer_GEO.EOF Then
956                                                 oRSServer_GEO.Delete adAffectCurrent
958                                                 sResponse = SaveServerRSCompressed(sWebsite, oRSServer_GEO)
                                                Else
960                                                 sResponse = "Data Updated"
                                                End If

                                            Else
962                                             sResponse = "Data Updated"
                                            End If
                                    
964                                         If sResponse = "Data Updated" Then
966                                             oRSServer.Delete adAffectCurrent
968                                             sResponse = SaveServerRSCompressed(sWebsite, oRSServer)
                                            End If
                                        
                                        End If
                                        
970                                     If oRSServer.EOF Or sResponse = "Data Updated" Then
                                    
972                                         If oRSServerSynchHistory.EOF Then
974                                             oRSServerSynchHistory.AddNew
                                            Else
                                                'Synchhistory record exists - allow override
                                            End If
                                            
976                                         oRSLocalSynchHistory.Fields("updates").Value = "true"
978                                         oRSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
980                                         UpdateRSRecordSingle oRSLocalSynchHistory, oRSServerSynchHistory
982                                         oRSServerSynchHistory.Fields("sBy").Value = "Delete synch'd up on " & CalcServerDateTime(Now())
984                                         sResponse = SaveServerRSCompressed(sWebsite, oRSServerSynchHistory)

986                                         If Not sResponse = "Data Updated" Then
988                                             TxMsg "    --- Error updating oRSServerSynchHistory: " & sResponse, True
                                            Else
990                                             oRSLocalSynchHistory.UpdateBatch  'adAffectCurrent
                                            End If
                                    
                                        Else
                                    
992                                         TxMsg "    --- Error deleting from oRSServer: " & sResponse, True
                                        End If
                                    
994                                     oRSLocalSynchHistory.MoveNext
                                    
                                    Loop
                                
                                End If
                            
996                             .MoveNext
998                             i = 1
1000                            sSQL = ""
                
                            End If
                        
1002                        If Not .EOF Then
                    
1004                            If i = 1 Then
1006                                sSQL = sSQL & " [sID] = '" & oRSMerged.Fields("sID").Value & "'"
                                Else
1008                                sSQL = sSQL & " OR [sID] = '" & oRSMerged.Fields("sID").Value & "'"
                                End If
                    
                                'Prepare for the next loop
1010                            i = i + 1
1012                            sLastOperation = sCurrentOperation
                        
                            End If
                        
                        End With

1014                Loop Until oRSMerged.EOF

                End If
            
            End If

        Else
1016        SynchDDWithServer2 = True
        End If

        'TxMsg "SynchDDWithServer2_Err: (" & Erl & ") " & Err.Description, True
        '<EhFooter>
        Exit Function

SynchDDWithServer2_Err:
        TxMsg "SynchDDWithServer2_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Sub UpdateRSRecordSingle(RSSource As ADODB.Recordset, _
                                 RSDest As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo UpdateRSRecordSingle_Err
        '</EhHeader>

        Dim uUID As Long
        uUID = -1
        Dim i As Long
100     i = 0
    
102     Do Until i = RSDest.Fields.Count
    
104         If Not RSDest.Fields(i).Name = "WKB" And Not RSDest.Fields(i).Name = "ORIG_ID" Then
106             If DoesFieldExistInRS(RSSource, RSDest.Fields(i).Name) Then

108                 If Not IsNull(RSSource.Fields(RSDest.Fields(i).Name).Value) Then
110                     If Not (Right$(RSDest.Source, 4) = "_FEA" And RSDest.Fields(i).Name = "UID") Then
                        
112                         RSDest.Fields(i).Value = RSSource.Fields(RSDest.Fields(i).Name).Value
114                         ' TxMsg RSDest.Fields(i).Name, True
                    
                        End If
                   
                    End If

116             ElseIf RSDest.Fields(i).Name = "oincidents_FEA.UID" Or RSDest.Fields(i).Name = "oincidents_GEO.UID" Then
                    If uUID = -1 Then uUID = GetMaxUIDFromGEOTable("oincidents_GEO", RSDest.ActiveConnection) + CLng(1)
118                 RSDest.Fields(i).Value = uUID 'RSSource.Fields("UID").Value
120                 'TxMsg RSDest.Fields(i).Name, True
                ElseIf RSDest.Fields(i).Name = "SHAPETYPEFORINCIDENTS" Then
                    RSDest.Fields(i).Value = 2
                End If
            End If

122         i = i + 1
        
        Loop

        '<EhFooter>
        Exit Sub

UpdateRSRecordSingle_Err:
        TxMsg "UpdateRSRecordSingle_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Function IsGeoTablePoly(sTableName As String, _
                                oConnection As ADODB.Connection) As Boolean
        '<EhHeader>
        On Error GoTo IsGeoTablePoly_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
100     IsGeoTablePoly = False

102     With oRS
    
104         .Open "SELECT top 1 [shapetype] from [" & Left(sTableName, Len(sTableName) - 4) & "_GEO] order by [shapetype] DESC", oConnection, adOpenDynamic, adLockBatchOptimistic

106         If Not .EOF Then
108             If Not .Fields(0).Value = 2 Then IsGeoTablePoly = True
            Else
                IsGeoTablePoly = True
            End If
        
110         .Close
    
        End With

        Set oRS = Nothing
        '<EhFooter>
        Exit Function

IsGeoTablePoly_Err:
        TxMsg "IsGeoTablePoly_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Public Function CalcServerDateTime(dPassedDate As Date) As String
    '  Get the current timedate and format it as RFC 3339

    Dim iYear As Integer
    Dim sMonth As String
    Dim sDay As String
    Dim shour As String
    Dim sMinute As String
    Dim sSec As String
    
    Dim dDate As Date
    dDate = dPassedDate + (((CDbl(lTimeDiffFromServerInSeconds) / 24) / 60) / 60)

    iYear = Year(dDate)
    sMonth = Month(dDate)
    sDay = Day(dDate)
    shour = Hour(dDate)
    sMinute = Minute(dDate)
    sSec = Second(dDate)

    If (iYear < 70) Then
        iYear = iYear + 2000
    ElseIf (iYear < 1900) Then
        iYear = iYear + 1900
    End If

    If (CInt(sMonth) <= 9) Then
        sMonth = "0" + sMonth
    End If

    If (CInt(sDay) <= 9) Then
        sDay = "0" + sDay
    End If

    If (CInt(shour) <= 9) Then
        shour = "0" + shour
    End If

    If (CInt(sMinute) <= 9) Then
        sMinute = "0" + sMinute
    End If

    If (CInt(sSec) <= 9) Then
        sSec = "0" + sSec
    End If

    CalcServerDateTime = iYear & "-" & sMonth & "-" & sDay & "T" & shour & ":" & sMinute & ":" & sSec & "Z"
    'TxMsg "Server time logged at: " & CalcServerDateTime

End Function

Public Function CalcServerDateTimeFromRFC3339(sDate As String) As String

    Dim dDate As Date
    TxMsg "Client time passed at: " & sDate

    If Len(sDate) = 20 Then
        dDate = DateSerial(CInt(Left$(sDate, 4)), CInt(Mid$(sDate, 6, 2)), CInt(Mid$(sDate, 9, 2))) + TimeSerial(CInt(Mid$(sDate, 12, 2)), CInt(Mid$(sDate, 15, 2)), CInt(Mid$(sDate, 18, 2)))
        CalcServerDateTimeFromRFC3339 = CalcServerDateTime(dDate)
    Else
        TxMsg "CalcServerDateTimeFromRFC3339 error.  Len of passed string = " & Len(sDate), True
    End If
    
    TxMsg "Server time logged at: " & CalcServerDateTimeFromRFC3339

End Function

Public Function RunDDSchemaCompare(sWebsite As String, _
                                   sPrefix As String, _
                                   oCN As ADODB.Connection) As Boolean
        '<EhHeader>
        On Error GoTo RunDDSchemaCompare_Err
        '</EhHeader>
        
        Dim sUG As String
100     bAbortSynchronisation = False
        
102     If sPrefix = "Incidents_" Then
104         sUG = m_StrRemoteTablePrefix
        Else
106         sUG = ""
        End If
        
108     TxMsg ">>>>>>>>>>>>"
110     TxMsg "**??Beginning schema scan of DD [" & sPrefix & "] ......................", True

112     TxMsg ">>> Getting remote tables ......................", True
114     If sUG = "" Then RemoteTables = GetRemoteTables(sWebsite, sPrefix)

120     TxMsg ">>> Getting local tables ......................", True
122     If sUG = "" Then LocalTables = GetLocalTables(oCN, sPrefix)

128     TxMsg ">>> Comparing tables ......................", True
130     If sUG = "" Then CompareTables sWebsite, sPrefix, oCN
    
132     RunDDSchemaCompare = Not bAbortSynchronisation

134     If sUG = "" Then
136         If UBound(RemoteTables.Tables) = 0 Then
138             TxMsg "   !!! NO TABLES DETECTED ON SERVER - PLEASE CHECK YOUR CONFIG", True
140             RunDDSchemaCompare = False
            End If
        End If

142     TxMsg "<<< Finished schema scan of DD [" & sPrefix & "] ......................"
144     TxMsg "<<<<<<<<<<<<"

        '<EhFooter>
        Exit Function

RunDDSchemaCompare_Err:
        TxMsg "SynchWorker.RunDDSchemaCompare on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Function GetDefinedQueries(oCOnn As ADODB.Connection, _
                                   sPrefix As String) As DDQueryCollection
        '<EhHeader>
        On Error GoTo GetDefinedQueries_Err
        '</EhHeader>

        Dim oRSQueries As New ADODB.Recordset
        Dim ddDatabaseQueries As DDQueryCollection
        Dim sSQL As String

100     sSQL = "SELECT * FROM [" & sPrefix & "Queries] ORDER BY [QueryName]"
102     ddDatabaseQueries.Prefix = sPrefix

104     With oRSQueries

106         .Open sSQL, oCOnn, adOpenDynamic, adLockBatchOptimistic

108         If Not .EOF Or Not .BOF Then .MoveFirst

110         ReDim ddDatabaseQueries.Queries(0)

112         Do Until .EOF

114             ReDim Preserve ddDatabaseQueries.Queries(UBound(ddDatabaseQueries.Queries) + 1)

116             With ddDatabaseQueries.Queries(UBound(ddDatabaseQueries.Queries))

118                 .Name = sPrefix & oRSQueries.Fields("QueryName").Value

120                 If bSQLServerInUse Then
122                     .sSQL = IIf(IsNull(oRSQueries.Fields("QueryMSSQL").Value), "", oRSQueries.Fields("QueryMSSQL").Value)
                    Else
124                     .sSQL = IIf(IsNull(oRSQueries.Fields("QuerySQL").Value), "", oRSQueries.Fields("QuerySQL").Value)
                    End If

                End With

126             .MoveNext
            Loop

128         .Close

        End With
    
130     Set oRSQueries = Nothing
132     GetDefinedQueries = ddDatabaseQueries
    
        '<EhFooter>
        Exit Function

GetDefinedQueries_Err:
        TxMsg "SynchWorker.GetDefinedQueries_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Function GetRemoteTables(sWebsite As String, _
                                 sPrefix As String) As DDTableCollection
        '<EhHeader>
        On Error GoTo GetRemoteTables_Err
        '</EhHeader>

        Dim oRSTables As New ADODB.Recordset
        Dim oRSColumns As New ADODB.Recordset
        
        Dim oRSColumnDesc As New ADODB.Recordset
        Dim ddDatabaseTables As DDTableCollection
        Dim sSQL As String
        
        Dim sIdentityColumn As String
100     sIdentityColumn = "none"
    
102     ddDatabaseTables.Prefix = sPrefix
        'sSQL = sWebsite & "oasis.asp?getDDtables=" & ddDatabaseTables.Prefix
        
104     Set oRSTables = OpenServerRSCompressed(sWebsite, "id", "SELECT OBJECT_NAME(id) as ObjectName, Name as IdentityColumn FROM syscolumns WHERE COLUMNPROPERTY( id ,name, 'IsIdentity') = 1 and OBJECT_NAME(id) = 'oincidents-keithtesting' ORDER BY 1, 2")
        
106     If Not oRSTables Is Nothing Then
108         If oRSTables.State = adStateOpen Then
        
110             If Not oRSTables.EOF Then
            
112                 sIdentityColumn = oRSTables.Fields(1).Value
            
                End If

114             oRSTables.Close
        
            End If
        End If
        
116     Set oRSTables = OpenServerRSCompressed(sWebsite, "getDDtables", ddDatabaseTables.Prefix)
 
118     With oRSTables
    
            On Error GoTo noinet
            '.Open sSQL, , adOpenDynamic, adLockBatchOptimistic
            
            On Error GoTo GetRemoteTables_Err
            '.Filter = "TABLE_TYPE = 'TABLE'"

120         If Not .EOF Or Not .BOF Then .MoveFirst
        
122         ReDim ddDatabaseTables.Tables(0)
124         Set oRSColumns = OpenServerRSCompressed(sWebsite, "getDDcolumnsNEW", ddDatabaseTables.Prefix)
126         Set oRSColumnDesc = OpenServerRSCompressed(sWebsite, "getddcolumnsdesc", ddDatabaseTables.Prefix)

128         Do Until .EOF
           
130             ReDim Preserve ddDatabaseTables.Tables(UBound(ddDatabaseTables.Tables) + 1)

132             With ddDatabaseTables.Tables(UBound(ddDatabaseTables.Tables))

134                 TxMsg "**??Fetching schema of remote table: " & oRSTables.Fields("TABLE_NAME").Value, , True
            
136                 .Name = oRSTables.Fields("TABLE_NAME").Value
138                 .IsTable = True
                    '.sAutoNum = sIdentityColumn

140                 If Left$(.Name, Len(sPrefix & "link")) = sPrefix & "link" Then
142                     .IsLinkedTable = True
                    Else
144                     .IsLinkedTable = False
                    End If

146                 .IsQuery = False
148                 ReDim .Columns(0)
150                 oRSColumns.Sort = "ORDINAL_POSITION"
                    
152                 oRSColumns.Filter = adFilterNone
154                 oRSColumnDesc.Filter = adFilterNone
                    
156                 If Not oRSColumns.EOF Or Not oRSColumns.BOF Then oRSColumns.MoveFirst
158                 If Not oRSColumnDesc.EOF Or Not oRSColumnDesc.BOF Then oRSColumnDesc.MoveFirst
                    
160                 oRSColumns.Filter = "TABLE_NAME = '" & .Name & "'"
162                 oRSColumnDesc.Filter = "TABLE_NAME = '" & .Name & "'"
            
164                 Do Until oRSColumns.EOF

                        If Not oRSColumns.Fields("COLUMN_NAME").Value = "WKB" And Not Left(oRSColumns.Fields("COLUMN_NAME").Value, 1) = "_" Then

166                         oRSColumnDesc.MoveFirst
168                         oRSColumnDesc.Find "[Column Name] = '" & oRSColumns.Fields("COLUMN_NAME").Value & "'"
                
170                         ReDim Preserve .Columns(UBound(.Columns) + 1)
172                         .Columns(UBound(.Columns)).Name = oRSColumns.Fields("COLUMN_NAME").Value
174                         .Columns(UBound(.Columns)).Ordinal = oRSColumns.Fields("ORDINAL_POSITION").Value
176                         .Columns(UBound(.Columns)).TypeEnum = RetDataType(oRSColumns.Fields("DATA_TYPE").Value)
                        
178                         If Not oRSColumnDesc Is Nothing Then
                        
180                             If Not oRSColumnDesc.EOF Then
                        
182                                 If Not IsNull(oRSColumnDesc.Fields("DESCRIPTION").Value) Then
184                                     .Columns(UBound(.Columns)).Description = oRSColumnDesc.Fields("DESCRIPTION").Value
                                    End If
                        
                                End If

                            End If

186                         If Not .Columns(UBound(.Columns)).TypeEnum = 201 And Not .Columns(UBound(.Columns)).TypeEnum = 128 Then
                            
188                             If Not IsNull(oRSColumns.Fields("CHARACTER_MAXIMUM_LENGTH").Value) Then
190                                 .Columns(UBound(.Columns)).MaxLength = oRSColumns.Fields("CHARACTER_MAXIMUM_LENGTH").Value
                                End If
                        
                            End If
                        End If

192                     oRSColumns.MoveNext
                    Loop
            
                End With
            
194             .MoveNext
            Loop
            
196         Set oRSColumns = Nothing
198         Set oRSColumnDesc = Nothing
200         Set oRSTables = Nothing
        
        End With
    
202     GetRemoteTables = ddDatabaseTables

        Exit Function

noinet:
204     TxMsg "   --- No internet connection!"

        '<EhFooter>
        Exit Function

GetRemoteTables_Err:
        TxMsg "SynchWorker.GetRemoteTables_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Function RetDataType(sDataType As String) As Long

    'ref: http://www.w3schools.com/ADO/ado_datatypes.asp
    'some overrides have been applied
    
    sDataType = LCase$(Trim$(sDataType))

    Select Case sDataType

        Case "bigint"
            RetDataType = 3 '20

        Case "binary"
            RetDataType = 128

        Case "timestamp"
            RetDataType = 128

        Case "bit"
            RetDataType = 11

        Case "char"
            RetDataType = 129

        Case "money"
            RetDataType = 6

        Case "smallmoney"
            RetDataType = 6

        Case "datetime"
            RetDataType = 7

        Case "datetime"
            RetDataType = 7 '135

        Case "smalldatetime"
            RetDataType = 7 '135

        Case "float"
            RetDataType = 5

        Case "uniqueidentifier"
            RetDataType = 72

        Case "identity"
            RetDataType = 3

        Case "int"
            RetDataType = 3

        Case "image"
            RetDataType = 128 '205

        Case "text"
            RetDataType = 203 '201

        Case "ntext"
            RetDataType = 130 '203

        Case "decimal"
            RetDataType = 131

        Case "numeric"
            RetDataType = 131

        Case "real"
            RetDataType = 4

        Case "smallint"
            RetDataType = 2

        Case "tinyint"
            RetDataType = 17

        Case "varbinary"
            RetDataType = 204

        Case "varchar"
            RetDataType = 200

        Case "sql_variant"
            RetDataType = 12

        Case "nvarchar"
            RetDataType = 130 '202

        Case "nchar"
            RetDataType = 130
            
        Case Else
                 
            RetDataType = 203
            TxMsg "   !!! Did not find mapping for data type: " & sDataType, True

    End Select

End Function

Public Function HasProperty(obj As Object, strPropName As String) As Boolean
    'Purpose:   Return true if the object has the property.
    Dim varDummy As Variant
    
    On Error Resume Next
    varDummy = obj.Properties(strPropName)
    HasProperty = (Err.Number = 0)
End Function

Function SetPropertyDAO(obj As Object, strPropertyName As String, intType As Integer, varValue As Variant, Optional strErrMsg As String) As Boolean
    On Error GoTo ErrHandler
    'Purpose:   Set a property for an object, creating if necessary.
    'Arguments: obj = the object whose property should be set.
    '           strPropertyName = the name of the property to set.
    '           intType = the type of property (needed for creating)
    '           varValue = the value to set this property to.
    '           strErrMsg = string to append any error message to.
    
    If HasProperty(obj, strPropertyName) Then
        obj.Properties(strPropertyName) = varValue
    Else
        obj.Properties.Append obj.CreateProperty(strPropertyName, intType, varValue)
    End If

    SetPropertyDAO = True

ExitHandler:
    Exit Function

ErrHandler:
    TxMsg "!!!! SetPropertyDAO_err: " & obj.Name & "." & strPropertyName & " not set to " & varValue & ". Error " & Err.Number & " - " & Err.Description & vbCrLf
    Resume ExitHandler
End Function

Private Function GetLocalTables(oCOnn As ADODB.Connection, _
                                sPrefix As String) As DDTableCollection
        '<EhHeader>
        On Error GoTo GetLocalTables_Err
        '</EhHeader>

        Dim oRS As ADODB.Recordset
        Dim oRSColumns As New ADODB.Recordset
        Dim ddDatabaseTables As DDTableCollection
100     ReDim ddDatabaseTables.Tables(0)
        Dim i As Long
        Dim adoxTable As New adox.Table
        Dim adoxColumn As New adox.Column
        Dim adoxCatalog As New adox.Catalog
        Dim sDatabaseFullPath As String
        Dim dbDAO As DAO.Database
        
102     adoxCatalog.ActiveConnection = oCOnn
104     oRSColumns.Open oCOnn.OpenSchema(adSchemaColumns)
        
106     sDatabaseFullPath = Mid$(oCOnn.ConnectionString, InStr(oCOnn.ConnectionString, "Data Source=") + 12, InStr(oCOnn.ConnectionString, "Mode=") - (InStr(oCOnn.ConnectionString, "Data Source=") + 12) - 1)
        
108     For Each adoxTable In adoxCatalog.Tables
    
110         If adoxTable.Type = "TABLE" And Left(adoxTable.Name, Len(sPrefix)) = sPrefix Then
            
112             ReDim Preserve ddDatabaseTables.Tables(UBound(ddDatabaseTables.Tables) + 1)
      
114             With ddDatabaseTables.Tables(UBound(ddDatabaseTables.Tables))
116                 i = 0

118                 Do Until i = adoxTable.Columns.Count
                    
120                     Set adoxColumn = adoxTable.Columns(i)

122                     If adoxColumn.Type = adDate Then
124                         'adoxColumn.Properties("Format").Value = "dd-MMM-yy"

126                         Set dbDAO = OpenDatabase(sDatabaseFullPath)
128                         SetPropertyDAO dbDAO.TableDefs(adoxTable.Name).Fields(adoxColumn.Name), "Format", dbText, "dd-MMM-yy"
130                         Set dbDAO = Nothing

                        End If

132                     i = i + 1
                    Loop
            
134                 .Name = adoxTable.Name
136                 .IsTable = True
138                 .IsLinkedTable = False
140                 .IsQuery = False
                    '.sAutoNum = "none"
142                 TxMsg "**??Fetching schema of local table: " & .Name, , True
                    
144                 ReDim .Columns(0)

146                 oRSColumns.Filter = adFilterNone
                    'If Not oRSColumns.EOF Or Not oRSColumns.BOF Then oRSColumns.MoveFirst
148                 oRSColumns.Filter = "TABLE_NAME = '" & .Name & "'"
            
150                 If bSQLServerInUse Then Set oRS = New ADODB.Recordset
             
152                 Do Until oRSColumns.EOF

154                     If Not oRSColumns.Fields("COLUMN_NAME").Value = "WKB" And Not Left(oRSColumns.Fields("COLUMN_NAME").Value, 1) = "_" Then
                        
156                         ReDim Preserve .Columns(UBound(.Columns) + 1)
158                         .Columns(UBound(.Columns)).Name = oRSColumns.Fields("COLUMN_NAME").Value
160                         .Columns(UBound(.Columns)).Ordinal = oRSColumns.Fields("ORDINAL_POSITION").Value
162                         .Columns(UBound(.Columns)).TypeEnum = oRSColumns.Fields("DATA_TYPE").Value
                        
164                         If bSQLServerInUse Then
166                             oRS.Open "SELECT [value] FROM ::fn_listExtendedProperty ( 'MS_Description','user', 'dbo', 'table', '" & .Name & "', 'column', '" & oRSColumns.Fields("COLUMN_NAME").Value & "')", oCOnn, adOpenDynamic, adLockBatchOptimistic

168                             If Not oRS.EOF Then .Columns(UBound(.Columns)).Description = oRS.Fields("value").Value
170                             oRS.Close
172                         ElseIf Not IsNull(oRSColumns.Fields("DESCRIPTION").Value) Then
174                             .Columns(UBound(.Columns)).Description = oRSColumns.Fields("DESCRIPTION").Value
                            End If

176                         If Not IsNull(oRSColumns.Fields("CHARACTER_MAXIMUM_LENGTH").Value) Then
178                             .Columns(UBound(.Columns)).MaxLength = oRSColumns.Fields("CHARACTER_MAXIMUM_LENGTH").Value
                            End If
                        
                            'If Not oRSColumns.Fields("ISAUTOINCREMENT") Is Nothing Then
                            '   If oRSColumns.Fields("ISAUTOINCREMENT").Value = True Then
                            '      .sAutoNum = oRSColumns.Fields("COLUMN_NAME").Value
                            ' End If
                            '   End If
                        End If

180                     oRSColumns.MoveNext
                    Loop
                    
182                 If bSQLServerInUse Then Set oRS = Nothing
            
                End With
        
            End If

        Next

184     Set adoxCatalog = Nothing
186     Set adoxTable = Nothing
188     Set oRSColumns = Nothing
190     GetLocalTables = ddDatabaseTables
        'TxMsg "SynchWorker.GetLocalTables_Err on line (" & Erl & ") " & Err.Description, True
        '<EhFooter>
        Exit Function

GetLocalTables_Err:
        TxMsg "SynchWorker.GetLocalTables_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Sub CompareTables(sWebsite As String, _
                          sPrefix As String, _
                          oCOnn As ADODB.Connection)
        '<EhHeader>
        On Error GoTo CompareTables_Err
        '</EhHeader>

        Dim iRemoteTableIndex As Long
        Dim iLocalTableIndex As Long
        Dim iRemoteColumnIndex As Long
        Dim iLocalColumnIndex As Long
        Dim RemoteRS As New ADODB.Recordset
        Dim sSQL As String
        
        Dim lTypeLocal As Long
        Dim lTypeServer As Long
        
        Dim sTableNameAltered As String
        Dim sTableName As String
        Dim sColumnName As String
        Dim bGeoTableAltered As Boolean
        
100     iLocalTableIndex = 1
102     TxMsg ">>>>>> Checking local tables incase they have been removed from the server....", False, True
        
104     Do Until iLocalTableIndex > UBound(LocalTables.Tables)
    
106         sTableName = LocalTables.Tables(iLocalTableIndex).Name
108         TxMsg ">>>>>>>> Currently looking at table: " & sTableName, False, True
110         iRemoteTableIndex = LocateIndexOfTable(RemoteTables, sTableName)

            If Right(sTableName, 4) = "_FEA" Then UpdatettkGISLaySQL sTableName, oCOnn
        
112         If iRemoteTableIndex = 0 And Not Right$(sTableName, 6) = "Hashed" Then
114             TxMsg "   !!! Table [" & sTableName & "] has been removed from the server - dropping local copy and removing from SynchHistory", True, True
                DropLocalTable sTableName, oCOnn, sPrefix
                '116             oCOnn.Execute "DROP TABLE [" & sTableName & "]"
118             oCOnn.Execute "DELETE FROM [" & sPrefix & "SynchHistory] WHERE sTableName = '" & sTableName & "'"
            
            End If
        
120         iLocalTableIndex = iLocalTableIndex + 1
    
        Loop

122     iLocalTableIndex = 0
124     iRemoteTableIndex = 1
    
126     TxMsg ">>>>>> Checking local tables incase they are missing or different from server schemas....", False, True
    
128     Do Until iRemoteTableIndex > UBound(RemoteTables.Tables)
    
130         sTableName = RemoteTables.Tables(iRemoteTableIndex).Name
132         iLocalTableIndex = LocateIndexOfTable(LocalTables, sTableName)
134         TxMsg ">>>>>>>> Currently looking at table: " & sTableName, False, True
        
136         If Not iLocalTableIndex = 0 And Not Right(sTableName, 6) = "Hashed" Then
                '
                'If Not RemoteTables.Tables(iRemoteTableIndex).sAutoNum = LocalTables.Tables(iLocalTableIndex).sAutoNum Then
        
                'Tables are on both remote and local
138             iRemoteColumnIndex = 1
            
140             Do Until iRemoteColumnIndex > UBound(RemoteTables.Tables(iRemoteTableIndex).Columns)
            
                    'TxMsg ">>>>>>>>>>> Currently looking at column: " & RemoteTables.Tables(iRemoteTableIndex).Columns(iRemoteColumnIndex).Name, False, False
142                 sColumnName = RemoteTables.Tables(iRemoteTableIndex).Columns(iRemoteColumnIndex).Name
144                 iLocalColumnIndex = LocateIndexOfColumn(LocalTables.Tables(iLocalTableIndex), sColumnName)
            
146                 With RemoteTables.Tables(iRemoteTableIndex).Columns(iRemoteColumnIndex)
                
148                     If Not iLocalColumnIndex = 0 Then

150                         If Not .Description = LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).Description Then
                        
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ' ISSUE:  Column captions are different
                                ' ACTION: Update caption on the local database
                                'TxMsg "   !!! Description mismatch" ', True
                                'bAbortSynchronisation = True
152                             UpdateLocalTableCaption LocalTables.Tables(iLocalTableIndex).Name, LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).Name, .Description, oCOnn

                            End If
                            
154                         If Not (bSQLServerInUse And .TypeEnum = 128) And Not (bSQLServerInUse And .TypeEnum = 200) Then
                            
156                             If Not .TypeEnum = adLongVarBinary And Not LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).TypeEnum = 203 And Not LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).TypeEnum = 130 And Not LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).TypeEnum = 11 Then
158                                 If Not .MaxLength = LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).MaxLength Then
                            
                                        '''''''''''''''''''''''''''''''''''''''''''''''''''''
                                        ' ISSUE:  Column MaxLengths are different
                                        ' ACTION: Drop local table
                                    
160                                     TxMsg "   !!! MaxLength mismatch for column [" & sColumnName & "] of client type [" & LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).TypeEnum & "] and server type [" & .TypeEnum & "]. Length on Server = " & .MaxLength & " and Client = " & LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).MaxLength, True
162                                     bAbortSynchronisation = True
164                                     DropLocalTable LocalTables.Tables(iLocalTableIndex).Name, oCOnn, sPrefix
                                        Exit Do
                                    End If
                                End If
                            
                            End If

166                         If Not .Ordinal = LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).Ordinal Then
                            
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ' ISSUE:  Column Ordinals are different
                                ' ACTION: Drop local table

                                'TxMsg "   !!! Ordinal mismatch", True
                                'bAbortSynchronisation = True
                                'DropLocalTable LocalTables.Tables(iLocalTableIndex).Name, oCOnn, sPrefix
                                'Exit Do
                            End If
                        
168                         lTypeLocal = LocalTables.Tables(iLocalTableIndex).Columns(iLocalColumnIndex).TypeEnum
170                         lTypeServer = .TypeEnum
                            
172                         lTypeLocal = IIf(lTypeLocal = 135, 7, lTypeLocal)
174                         lTypeLocal = IIf(lTypeLocal = 20, 3, lTypeLocal)
176                         lTypeLocal = IIf(lTypeLocal = 201, 203, lTypeLocal)
178                         lTypeLocal = IIf(lTypeLocal = 202, 130, lTypeLocal)
180                         lTypeLocal = IIf(lTypeLocal = 128, 205, lTypeLocal)
182                         lTypeLocal = IIf(lTypeLocal = 200, 129, lTypeLocal)
                            
184                         lTypeServer = IIf(lTypeServer = 135, 7, lTypeServer)
186                         lTypeServer = IIf(lTypeServer = 20, 3, lTypeServer)
188                         lTypeServer = IIf(lTypeServer = 201, 203, lTypeServer)
190                         lTypeServer = IIf(lTypeServer = 202, 130, lTypeServer)
192                         lTypeServer = IIf(lTypeServer = 128, 205, lTypeServer)
194                         lTypeServer = IIf(lTypeServer = 200, 129, lTypeServer)
                            
196                         If lTypeServer = 200 And lTypeLocal = 129 Then lTypeLocal = 200
                            
198                         If Not lTypeServer = lTypeLocal Then
                       
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''
                                ' ISSUE:  Column TypeEnum are different
                                ' ACTION: Update TypeEnum on the local database
                                
200                             TxMsg "   !!! TypeEnum mismatch for column [" & sColumnName & "]. Server = " & lTypeServer & " and Client = " & lTypeLocal, True
202                             bAbortSynchronisation = True
204                             DropLocalTable LocalTables.Tables(iLocalTableIndex).Name, oCOnn, sPrefix
                                Exit Do
                            End If
                        
                        Else
                        
                            '''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ' ISSUE:  Column does not exist in the local database table
                            ' ACTION: Drop local table
                        
206                         If Not sColumnName = "ORIG_ID" And Not sColumnName = "WKB" And Not Left(sColumnName, 1) = "_" Then
208                             TxMsg "   !!! Column missing [" & sColumnName & "]", True
210                             bAbortSynchronisation = True
212                             DropLocalTable LocalTables.Tables(iLocalTableIndex).Name, oCOnn, sPrefix
                                Exit Do
                            End If
                        
                        End If
                
                    End With
               
214                 iRemoteColumnIndex = iRemoteColumnIndex + 1
                Loop

                'Else
                
                '   TxMsg "   !!! autonum mismatch"
                '  bAbortSynchronisation = True
                ' DropLocalTable LocalTables.Tables(iLocalTableIndex).Name, oCOnn, sPrefix
                
                'End If

            Else
        
                '''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' ISSUE:  Table does not exist in the local database
                ' ACTION: Create table on the local database
                If Not Right$(RemoteTables.Tables(iRemoteTableIndex).Name, 6) = "Hashed" Then
216                 bAbortSynchronisation = True
218                 sSQL = "SELECT TOP 1 * FROM [" & RemoteTables.Tables(iRemoteTableIndex).Name & "]"
                    'RemoteRS.Open sSQL, , adOpenDynamic, adLockBatchOptimistic
220                 Set RemoteRS = OpenServerRSCompressed(sWebsite, "ID", sSQL)
222                 CreateTable RemoteTables.Tables(iRemoteTableIndex).Name, RemoteRS, oCOnn, sWebsite
224                 RemoteRS.Close
                End If
            
            End If
        
226         iRemoteTableIndex = iRemoteTableIndex + 1
    
        Loop
        
228     TxMsg ">>>>> Currently checking tables", False, True
        
230     If Not bAbortSynchronisation Then
        
232         iRemoteTableIndex = 1
    
234         Do Until iRemoteTableIndex > UBound(RemoteTables.Tables)
    
236             sTableNameAltered = RemoteTables.Tables(iRemoteTableIndex).Name
238             TxMsg ">>>>>>>> Currently looking at table: " & sTableNameAltered, False, True
            
240             If Right$(sTableNameAltered, 4) = "_FEA" Then
242                 sTableNameAltered = Left$(sTableNameAltered, Len(sTableNameAltered) - 4)
244                 bGeoTableAltered = True
                Else
246                 bGeoTableAltered = False
                End If
                
248             If Not Right$(sTableNameAltered, 4) = "_GEO" And Not Right$(sTableNameAltered, 12) = "SynchHistory" And Not Right$(sTableNameAltered, 6) = "Hashed" Then

                    'TxMsg "validating : " & RemoteTables.Tables(iRemoteTableIndex).Name & " and linked is: " & RemoteTables.Tables(iRemoteTableIndex).IsLinkedTable
250                 ValidateClientTable sTableNameAltered, IIf(RemoteTables.Tables(iRemoteTableIndex).IsLinkedTable, "GUID2", "GUID1"), True, bGeoTableAltered, oCOnn, sPrefix
                End If

252             iRemoteTableIndex = iRemoteTableIndex + 1
    
            Loop
        
        End If

254     Set RemoteRS = Nothing
        'TxMsg "SynchWorker.CompareTables_Err on line (" & Erl & ") " & Err.Description, True
        '<EhFooter>
        Exit Sub

CompareTables_Err:
        TxMsg "SynchWorker.CompareTables_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Public Sub UpdateLocalTableCaption(sTableName As String, _
                                   sFieldName As String, _
                                   sCaption As String, _
                                   oCOnn As ADODB.Connection)
        '<EhHeader>
        On Error GoTo UpdateLocalTableCaption_Err
        '</EhHeader>
        
        Dim bRetry As Boolean
    
        'ref: http://allenbrowne.com/func-ADOX.html#ModifyFieldPropAdox
100     If bSQLServerInUse Then
            ' On Error Resume Next next
            oCOnn.Execute "EXEC sys.sp_addextendedproperty @level0type = N'SCHEMA' ,@level0name = [dbo] ,@level1type = N'TABLE' ,@level1name = [" & sTableName & "] ,@level2type = N'COLUMN' ,@level2name = [" & sFieldName & "] ,@name = N'MS_Description' ,@value = N'" & sCaption & "' ;"

            If bRetry Then oCOnn.Execute "EXEC sys.sp_updateextendedproperty @level0type = N'SCHEMA' ,@level0name = [dbo] ,@level1type = N'TABLE' ,@level1name = [" & sTableName & "] ,@level2type = N'COLUMN' ,@level2name = [" & sFieldName & "] ,@name = N'MS_Description' ,@value = N'" & sCaption & "' ;"

        Else
    
            Dim cat As New adox.Catalog
            Dim col As adox.Column
            Dim prp As adox.Property
101         cat.ActiveConnection = oCOnn
102         Set col = cat.Tables(sTableName).Columns(sFieldName)

104         Set prp = col.Properties("Description")
106         prp.Value = sCaption
110         Set prp = Nothing
112         Set col = Nothing
114         Set cat = Nothing

        End If
    
116     TxMsg "   --- Updating caption on local table [" & sTableName & "] field [" & sFieldName & "] with '" & sCaption & "'"

        '<EhFooter>
        Exit Sub

UpdateLocalTableCaption_Err:

        If Erl = 100 Then
            bRetry = True
            Resume Next
        Else

            TxMsg "SynchWorker.UpdateLocalTableCaption_Err on line (" & Erl & ") " & Err.Description, True
        
        End If
        
        'Err.Raise vbObjectError + 100, "Project1.clsUpdateLocalDbSchemaWithServerSchema.UpdateLocalTableCaption", "clsUpdateLocalDbSchemaWithServerSchema component failure"
        '</EhFooter>
End Sub

Public Sub DropLocalTable(sTableName As String, _
                          oConnection As ADODB.Connection, _
                          sPrefix As String)
        '<EhHeader>
        On Error GoTo DropLocalTable_Err
        '</EhHeader>
        
        Dim sStrippedGeoTableName As String

100     If Not Right$(sTableName, 12) = "SynchHistory" Then

            Dim mcatDB As New adox.Catalog
102         Set mcatDB.ActiveConnection = oConnection
104         mcatDB.Tables.Delete (sTableName)
106         Set mcatDB = Nothing
    
            Dim oRS As New ADODB.Recordset
108         oConnection.Execute "DELETE FROM [" & sPrefix & "SynchHistory] WHERE sTableName = '" & sTableName & "' and [sequence] >= 0"
        
112         TxMsg "   --- Dropping local table [" & sTableName & "]"
            On Error Resume Next
            
114         If Right$(sTableName, 4) = "_FEA" Then
            
116             sStrippedGeoTableName = Left$(sTableName, Len(sTableName) - 4)
118             Set oRS = New ADODB.Recordset
120             oRS.Open "DELETE FROM [" & sStrippedGeoTableName & "_GEO]", oConnection, adOpenDynamic, adLockBatchOptimistic
122             Set oRS = Nothing
            
124         ElseIf Right$(sTableName, 4) = "_GEO" Then
            
126             sStrippedGeoTableName = Left$(sTableName, Len(sTableName) - 4)
128             Set oRS = New ADODB.Recordset
130             oRS.Open "DELETE FROM [" & sStrippedGeoTableName & "_FEA]", oConnection, adOpenDynamic, adLockBatchOptimistic
132             Set oRS = Nothing
            
            End If
    
        End If

        '<EhFooter>
        Exit Sub

DropLocalTable_Err:
        TxMsg "SynchWorker.DropLocalTable_Err on line (" & Erl & ") " & Err.Description, True
        'Err.Raise vbObjectError + 100, "Project1.clsUpdateLocalDbSchemaWithServerSchema.DropLocalQuery", "clsUpdateLocalDbSchemaWithServerSchema component failure"
        '</EhFooter>
End Sub

Public Function DropLocalQuery(sQueryName As String, _
                               oConnection As ADODB.Connection) As Boolean
        '<EhHeader>
        On Error GoTo DropLocalQuery_Err
        '</EhHeader>
100     DropLocalQuery = True
        'Dim mcatDB As New adox.Catalog
102     TxMsg "   --- Attempting to drop local query [" & sQueryName & "]", True

104     If bSQLServerInUse Then
106         oConnection.Execute "DROP VIEW [" & sQueryName & "]"
        Else
108         oConnection.Execute "drop table [" & sQueryName & "]"
        End If

110     'Set mcatDB = Nothing
112     TxMsg "   --- Dropped local query [" & sQueryName & "]", True

        '<EhFooter>
        Exit Function

DropLocalQuery_Err:
        TxMsg "   --- Dropping of local query [" & sQueryName & "] failed with error: " & Err.Description, True
        DropLocalQuery = False
        '</EhFooter>
End Function

Public Function CreateTable(sTableName As String, _
                            oRS As ADODB.Recordset, _
                            oConnection As ADODB.Connection, _
                            sWebsite As String, _
                            Optional bPopulate As Boolean = False)
        '<EhHeader>
        On Error GoTo CreateTable_Err
        '</EhHeader>

        ' if this stops working ref: http://www.dbforums.com/microsoft-access/991186-how-do-i-output-ado-recordset-mdb-table.html
    
        Dim mcatDB As adox.Catalog
        Dim mtblNew As adox.Table
        Dim colField As adox.Column
        Dim x As Long
        Dim bAddAutoUID As Boolean

100     Set mcatDB = New adox.Catalog
102     Set mcatDB.ActiveConnection = oConnection
104     Set mtblNew = New adox.Table
106     mtblNew.Name = sTableName
108     Set mtblNew.ParentCatalog = mcatDB

110     TxMsg "   --- Creating local table [" & sTableName & "]"

112     With oRS

114         For x = 0 To oRS.Fields.Count - 1

116             Set colField = New adox.Column
118             colField.Name = .Fields(x).Name

120             If bSQLServerInUse Then colField.Name = "[" & colField.Name & "]"
                'If bSQLServerInUse And LCase(.Fields(x).Name) = "group" Then colField.Name = "[group]"
                'If bSQLServerInUse And LCase(.Fields(x).Name) = "option" Then colField.Name = "[option]"
                'If bSQLServerInUse And LCase(.Fields(x).Name) = "uid" Then colField.Name = "[UID]"
                
122             colField.Type = .Fields(x).Type
                                
                'not sure if this works properly - might need to fetch the properties seperatly off the server
                'If .Fields(x).Properties("ISAUTOINCREMENT") Then colField.Properties("AutoIncrement") = True

                'overrides
                
124             If Not bSQLServerInUse Then
                
126                 colField.Type = IIf(colField.Type = 135, 7, colField.Type)
128                 colField.Type = IIf(colField.Type = 20, 3, colField.Type)
130                 colField.Type = IIf(colField.Type = 201, 203, colField.Type)
132                 colField.Type = IIf(colField.Type = 202, 130, colField.Type)
134                 colField.Type = IIf(colField.Type = 128, 205, colField.Type)

                End If

                'TxMsg "colField.Name = " & colField.Name & " and colField.Type = " & colField.Type
    
136             If Not colField.Type = adBoolean Then
                
138                 If Not colField.Type = adLongVarWChar Then
                        On Error Resume Next
140                     colField.DefinedSize = IIf(colField.Type = adVarChar, 255, .Fields(x).DefinedSize)
                        On Error GoTo CreateTable_Err
                    End If

142                 colField.Precision = .Fields(x).Precision
144                 colField.NumericScale = .Fields(x).NumericScale

                End If
                
146             If .Fields(x).Name = "UID" And Not Right$(sTableName, 4) = "_FEA" And Not Right$(sTableName, 4) = "_GEO" Then
148                 colField.ParentCatalog = mcatDB
150                 bAddAutoUID = True

152                 If Not bSQLServerInUse Then colField.Properties("Autoincrement").Value = True
                End If
                
154             If Not colField.Type = adBoolean And Not colField.Name = "UID" Then colField.Attributes = adColNullable
                
156             If Not (bSQLServerInUse And bAddAutoUID) Then
158                 mtblNew.Columns.Append colField

160                 'If Not colField.Type = adBoolean And Not bAddAutoUID Then mtblNew.Columns(colField.Name).Attributes = adColNullable
                End If

162         Next x

        End With

        'Note that this creates the table locally but it does not populate the captions - this is done later
164     mcatDB.Tables.Append mtblNew

166     If bSQLServerInUse And bAddAutoUID Then
            'oConnection.Execute "ALTER TABLE [" & sTableName & "] DROP COLUMN [UID]"
168         oConnection.Execute "ALTER TABLE [" & sTableName & "] ADD [UID] INT IDENTITY(1,1)"
        End If

170     If bPopulate Then
172         AddToTable oConnection, oRS, sTableName
        End If

174     Set mtblNew = Nothing
176     Set mcatDB = Nothing
178     Set colField = Nothing

180     If Right$(sTableName, 4) = "_FEA" Then
182         CreateTTKGPReference sTableName, oConnection, sWebsite
        End If
        
184     If Len(sTablesCreated) > 0 Then
186         sTablesCreated = sTablesCreated & ";" & sTableName
        Else
188         sTablesCreated = sTableName
        End If

        '<EhFooter>
        Exit Function

CreateTable_Err:
        TxMsg "SynchWorker.CreateTable_Err for table [" & sTableName & "] on line (" & Erl & ") " & Err.Description, True
        'If Erl = 152 Then TxMsg "!!! Field [" & colField.Name & "] failed adding...", True
        '</EhFooter>
End Function

Public Function AddToTable(oCN As ADODB.Connection, _
                           adoRec As ADODB.Recordset, _
                           sTableName As String)
        '<EhHeader>
        On Error GoTo AddToTable_Err
        '</EhHeader>

        Dim rst As ADODB.Recordset
        Dim fld As ADODB.Field

        'first clear out any old data in the temp table
        'CurrentDb.Execute "DELETE * FROM sTablename"

        'now open a recordset of that table so you can add new records
100     Set rst = New ADODB.Recordset
102     rst.Open "SELECT * FROM [" & sTableName & "]", oCN, adOpenDynamic, adLockBatchOptimistic

        'now loop through the recordset you passed in and add
        'each record to the temp table
        If Not adoRec.EOF Then
        
104         adoRec.MoveFirst

106         Do While Not adoRec.EOF
108             rst.AddNew

110             For Each fld In adoRec.Fields
112                 rst.Fields(fld.Name).Value = fld.Value
                Next

114             rst.UpdateBatch adAffectCurrent
116             adoRec.MoveNext
            Loop
        
        End If

        'cleanup
118     rst.Close
120     Set rst = Nothing

        '<EhFooter>
        Exit Function

AddToTable_Err:
        TxMsg "SynchWorker.AddToTable_Err for table [" & sTableName & "] on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Sub CreateTTKGPReference(sName As String, _
                                 oCnn As ADODB.Connection, _
                                 sWebsite As String)
        '<EhHeader>
        On Error GoTo CreateTTKGPReference_Err
        '</EhHeader>
    
        Dim RSttkGISLayerSQL As New ADODB.Recordset
        Dim sShortName As String
        Dim oRS As ADODB.Recordset

100     If Not DoesTableExist(oCnn, "ttkGISLayerSQL") Then
            
102         oCnn.Execute "CREATE TABLE [ttkGISLayerSQL]  ([NAME] TEXT(255), [XMIN] DOUBLE, [XMAX] DOUBLE, [YMIN] DOUBLE, [YMAX] DOUBLE, [SHAPETYPE] INTEGER)"
104         TxMsg "   --- Creating table [ttkGISLayerSQL]"

        End If
            
106     sShortName = Left$(sName, Len(sName) - 4)
            
108     With RSttkGISLayerSQL
            
110         .Open "SELECT * FROM [ttkGISLayerSQL] WHERE [Name] = '" & sShortName & "'", oCnn, adOpenDynamic, adLockBatchOptimistic
                
112         If .EOF Then
           
114             .AddNew
            End If
            
116         Set oRS = OpenServerRSCompressed(sWebsite, "getExtent", sShortName & "_GEO")

118         If Not oRS.EOF Then

120             If Len(.Fields("XMIN").Value) > 1 Then
122                 .Fields("XMIN").Value = oRS.Fields(0).Value
124                 .Fields("XMAX").Value = oRS.Fields(1).Value
126                 .Fields("YMIN").Value = oRS.Fields(2).Value
128                 .Fields("YMAX").Value = oRS.Fields(3).Value
                Else
130                 .Fields("XMIN") = -180
132                 .Fields("XMAX") = 180
134                 .Fields("YMIN") = -90
136                 .Fields("YMAX") = 90
                End If

            Else

138             .Fields("XMIN") = -180
140             .Fields("XMAX") = 180
142             .Fields("YMIN") = -90
144             .Fields("YMAX") = 90

            End If

146         .Fields("Name") = sShortName
148         .Fields("SHAPETYPE") = 0
150         .UpdateBatch adAffectCurrent
152         TxMsg "   --- Populating table [ttkGISLayerSQL] for GEOTABLE: [" & sName & "]"
           
            ' End If
           
154         .Close
            
        End With
            
156     Set RSttkGISLayerSQL = Nothing

        '<EhFooter>
        Exit Sub

CreateTTKGPReference_Err:
        TxMsg "SynchWorker.CreateTTKGPReference_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Function CreateQuery(sName As String, _
                             sSQL As String, _
                             oCOnn As ADODB.Connection, _
                             sWebsite As String) As Boolean
        '<EhHeader>
        On Error GoTo CreateQuery_Err
        '</EhHeader>
        
        Dim oRSQueries As New ADODB.Recordset
        Dim sError As String
    
        With oRSQueries

            .Open oCOnn.OpenSchema(adSchemaViews)
            .Filter = "TABLE_NAME = '" & sName & "'"
        
            If .EOF Or .BOF Then

                'do nothing
            Else
                TxMsg "!!!   Query already exists!  Dropping", True
                DropLocalQuery sName, oCOnn
            End If

        End With

        Set oRSQueries = Nothing
    
        ''''''''''''''''''''''''''''''''''''

        Dim mcatDB As New adox.Catalog
        Dim mCommand As New Command
        Dim oRS As ADODB.Recordset
        
100     CreateQuery = True
102     mCommand.CommandText = sSQL
        
104     If bSQLServerInUse Then
106         oCOnn.Execute sSQL
        Else
108         Set mcatDB.ActiveConnection = oCOnn
110         mcatDB.Views.Append sName, mCommand
        End If
        
112     Set mcatDB = Nothing
114     Set mCommand = Nothing

136     If Right$(sName, 4) = "_FEA" Then
138         CreateTTKGPReference sName, oCOnn, sWebsite
        End If
            
116     Set oRS = New ADODB.Recordset

        On Error Resume Next

118     oRS.Open "SELECT * FROM [" & sName & "]", oCOnn, adOpenDynamic, adLockBatchOptimistic
        
120     If oRS Is Nothing Then
122         TxMsg "   !!! Note if you are doing a batch download of queries the following error reported may be fixed in this synch cycle....", True
124         TxMsg "   !!! Error creating local query [" & sName & "]: " & Err.Description, True
126     ElseIf oRS.State = 0 Then
128         TxMsg "   !!! Note if you are doing a batch download of queries the following error reported may be fixed in this synch cycle....", True
130         TxMsg "   !!! Error creating local query [" & sName & "]: " & Err.Description, True
        Else
132         TxMsg "**??Created local query [" & sName & "]"
134         oRS.Close

            If Len(sTablesCreated) > 0 Then
                sTablesCreated = sTablesCreated & ";" & sName
            Else
                sTablesCreated = sName
            End If

        End If

140     Set oRS = Nothing

        '<EhFooter>
        Exit Function

CreateQuery_Err:
        sError = Trim(Err.Description)
        TxMsg "   --- Creation of local query [" & sName & "] failed with error: " & sError, True
        
        If Right$(sError, 7) = "exists." Then DropLocalQuery sName, oCOnn
        CreateQuery = False
        '</EhFooter>
End Function

Private Function CloneRS(oRS As ADODB.Recordset, _
                         Optional sExcludedField As String = "") As ADODB.Recordset
        '<EhHeader>
        On Error GoTo CloneRS_Err
        '</EhHeader>

        Dim oRSClone As New ADODB.Recordset
        Dim i As Integer
100     i = 0
    
102     Do Until i = oRS.Fields.Count
    
104         With oRS.Fields(i)
        
106             If Not .Name = sExcludedField Then oRSClone.Fields.Append .Name, .Type, .DefinedSize

            End With

108         i = i + 1
        Loop
    
110     Set CloneRS = oRSClone

        '<EhFooter>
        Exit Function

CloneRS_Err:
        TxMsg "SynchWorker.CloneRS_Err on line (" & Erl & ") " & Err.Description, True
        'Err.Raise vbObjectError + 100, "Project1.clsUpdateLocalDbSchemaWithServerSchema.CloneRS", "clsUpdateLocalDbSchemaWithServerSchema component failure"
        '</EhFooter>
End Function

Private Function CloneRecordFromRS(oRS As ADODB.Recordset) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo CloneRecordFromRS_Err
        '</EhHeader>

        Dim oRClone As New ADODB.Recordset
        Dim i As Integer
100     i = 0
    
102     Do Until i = oRS.Fields.Count
    
104         With oRS.Fields(i)
        
106             oRClone.Fields.Append .Name, .Type, .DefinedSize

            End With

108         i = i + 1
        Loop
    
110     i = 0
112     oRClone.Open
        oRClone.AddNew
    
114     Do Until i = oRS.Fields.Count
    
116         With oRS.Fields(i)
        
118             oRClone.Fields(i).Value = .Value

            End With

120         i = i + 1
        Loop
    
122     Set CloneRecordFromRS = oRClone
        'TxMsg "SynchWorker.CloneRecordFromRS_Err on line (" & Erl & ") " & Err.Description, True
        
        '<EhFooter>
        Exit Function

CloneRecordFromRS_Err:
        TxMsg "SynchWorker.CloneRecordFromRS_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Function LocateIndexOfColumn(oTable As DDTable, _
                                     sItem As String) As Long
        '<EhHeader>
        On Error GoTo LocateIndexOfColumn_Err
        '</EhHeader>

        Dim i As Long
100     i = 1
        LocateIndexOfColumn = 0
     
102     Do Until i > UBound(oTable.Columns)
    
104         If UCase(oTable.Columns(i).Name) = UCase(sItem) Then
106             LocateIndexOfColumn = i
                Exit Do
            End If
        
108         i = i + 1
        Loop

        '<EhFooter>
        Exit Function

LocateIndexOfColumn_Err:
        TxMsg "SynchWorker.LocateIndexOfColumn_Err on line (" & Erl & ") " & Err.Description, True
        'Err.Raise vbObjectError + 100, "Project1.clsUpdateLocalDbSchemaWithServerSchema.LocateIndexOfColumn", "clsUpdateLocalDbSchemaWithServerSchema component failure"
        '</EhFooter>
End Function

Private Function LocateIndexOfTable(oTable As DDTableCollection, _
                                    sItem As String) As Long
        '<EhHeader>
        On Error GoTo LocateIndexOfTable_Err
        '</EhHeader>

        Dim i As Long
100     i = 0
        LocateIndexOfTable = 0
     
102     Do Until i > UBound(oTable.Tables) Or UBound(oTable.Tables) = 0 Or UBound(oTable.Tables) = -1
    
104         If LCase(oTable.Tables(i).Name) = LCase(sItem) Then
106             LocateIndexOfTable = i
                Exit Do
            End If
        
108         i = i + 1
        Loop

        '<EhFooter>
        Exit Function

LocateIndexOfTable_Err:
        LocateIndexOfTable = -1
        TxMsg "SynchWorker.LocateIndexOfTable_Err on line (" & Erl & ") " & Err.Description, True
        'Err.Raise vbObjectError + 100, "Project1.clsUpdateLocalDbSchemaWithServerSchema.LocateIndexOfTable", "clsUpdateLocalDbSchemaWithServerSchema component failure"
        '</EhFooter>
End Function

Private Function DoesTableExist(cn As ADODB.Connection, _
                                sTable As String) As Boolean
    Dim RS As New ADODB.Recordset
    
    On Error GoTo hell
    RS.Open "SELECT * FROM " & sTable, cn, adOpenForwardOnly, adLockReadOnly

    DoesTableExist = True
        
    On Error Resume Next
    RS.Close
    Set RS = Nothing
        
    Exit Function
hell:
    On Error Resume Next
    Err.Clear

    DoesTableExist = False
    Set RS = Nothing
End Function

Private Sub TxMsg(sMsg As String, _
                  Optional bError As Boolean = False, _
                  Optional bDebugging As Boolean = False)

    On Error Resume Next
    Dim bOK As Boolean
    bOK = True

    If bDebugging And Not m_bDebug Then bOK = False
    
    If bOK Then
        If bError Then
            SleepAPI 300
            WriteLog sMsg, bError
        End If

        WriteLog sMsg
        oServerInterComms.TransmitToClient sMsg
    End If
    
End Sub

Private Function AsciiToBinary(txt As String) As String

    Dim result As String
    Dim ch As String
    Dim bin As String
    Dim i As Integer

    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, vbLf, "")

    result = ""

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)

        bin = LongToBinary(Asc(ch), False)
        result = result & Right$(bin, 8)
    Next i

    AsciiToBinary = result
    
End Function

Private Function BinaryToAscii(bin As String) As String

    Dim result As String
    Dim i As Integer
    Dim next_char As String
    Dim ascii As Long

    result = ""

    For i = 1 To Len(bin) + 18 Step 8
        next_char = Mid$(bin, i, 8)

        ascii = BinaryToLong(next_char)
        result = result & Chr$(ascii)
    Next i

    BinaryToAscii = result
    
End Function

Private Function LongToBinary(ByVal long_value As Long, _
                              Optional ByVal separate_bytes As Boolean = True) As String
    Dim hex_string As String
    Dim digit_num As Integer
    Dim digit_value As Integer
    Dim nibble_string As String
    Dim result_string As String
    Dim factor As Integer
    Dim bit As Integer

    ' Convert into hex.
    hex_string = Hex$(long_value)

    ' Zero-pad to a full 8 characters.
    hex_string = Right$(String$(8, "0") & hex_string, 8)

    ' Read the hexadecimal digits
    ' one at a time from right to left.
    For digit_num = 8 To 1 Step -1
        ' Convert this hexadecimal digit into a
        ' binary nibble.
        digit_value = CLng("&H" & Mid$(hex_string, digit_num, 1))

        ' Convert the value into bits.
        factor = 1
        nibble_string = ""

        For bit = 3 To 0 Step -1

            If digit_value And factor Then
                nibble_string = "1" & nibble_string
            Else
                nibble_string = "0" & nibble_string
            End If

            factor = factor * 2
        Next bit

        ' Add the nibble's string to the left of the
        ' result string.
        result_string = nibble_string & result_string
    Next digit_num

    ' Add spaces between bytes if desired.
    If separate_bytes Then
        result_string = Mid$(result_string, 1, 8) & " " & Mid$(result_string, 9, 8) & " " & Mid$(result_string, 17, 8) & " " & Mid$(result_string, 25, 8)
    End If

    ' Return the result.
    LongToBinary = result_string
End Function

Private Function BinaryToLong(ByVal binary_value As String) As Long
    Dim hex_result As String
    Dim nibble_num As Integer
    Dim nibble_value As Integer
    Dim factor As Integer
    Dim bit As Integer

    ' Remove any leading &B if present.
    ' (Note: &B is not a standard prefix, it just
    ' makes some sense.)
    binary_value = UCase$(Trim$(binary_value))

    If Left$(binary_value, 2) = "&B" Then
        binary_value = Mid$(binary_value, 3)
    End If

    ' Strip out spaces in case the bytes are separated
    ' by spaces.
    binary_value = Replace(binary_value, " ", "")

    ' Left pad with zeros so we have a full 32 bits.
    binary_value = Right$(String(32, "0") & binary_value, 32)

    ' Read the bits in nibbles from right to left.
    ' (A nibble is half a byte. No kidding!)
    For nibble_num = 7 To 0 Step -1
        ' Convert this nibble into a hexadecimal string.
        factor = 1
        nibble_value = 0

        ' Read the nibble's bits from right to left.
        For bit = 3 To 0 Step -1

            If Mid$(binary_value, 1 + nibble_num * 4 + bit, 1) = "1" Then
                nibble_value = nibble_value + factor
            End If

            factor = factor * 2
        Next bit

        ' Add the nibble's value to the left of the
        ' result hex string.
        hex_result = Hex$(nibble_value) & hex_result
    Next nibble_num

    ' Convert the result string into a long.
    BinaryToLong = CLng("&H" & hex_result)
End Function
 




