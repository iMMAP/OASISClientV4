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
Private DefinedDDQueries As DDQueryCollection
Private LocalTables As DDTableCollection
Private LocalQueries As DDQueryCollection
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

Private OASStringCompression As New OASISStringCompression.OASISCompression

Private Type SynchHistoryOverviewEnum

    sWhenMinSearched As String
    sWhenMaxSearched As String
    sWhenMinActualLocal As String
    sWhenMaxActualLocal As String
    sWhenMinActualServer As String
    sWhenMaxActualServer As String
    lBatchSizeLast As Long
    lBatchSizeNext As Long
    sTablenameLocal As String
    sTablenameServer As String
    sSynchHistoryTableLocal As String
    sSynchHistoryTableServer As String
    lCountServer As Long
    lCountLocal As Long
    NextSynchNextOpeation As SynchNextOpeation
    
End Type

Private Enum SynchNextOpeation
    NewOnly = 1
    OldOnly = 2
    DoNotSynch = 3
    EditsOnly = 4
End Enum

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

Private SynchHistoryOverview As SynchHistoryOverviewEnum

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

Private lBatchCountToDownload As Long

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

'Private Function SaveServerRS(sWebsite As String, _
'                              oRS As ADODB.Recordset) As String
'
'    Dim oHttp As New WinHttpRequest
'    Dim MsXmlDoc As New MSXML2.DOMDocument
'
'    PrepareHttpComms oHttp, sWebsite, True
'    oRS.Save MsXmlDoc, 1
'    oHttp.send MsXmlDoc
'    SaveServerRS = Trim(CStr(oHttp.responseText))
'    oHttp.abort
'    Set MsXmlDoc = Nothing
'    Set oHttp = Nothing
'
'End Function

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

Private Function CheckIfFileExists(Filename As String) As Boolean
    Dim i As Integer
    
    On Local Error Resume Next
    i = Len(Dir$(Filename$))
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
        
        Dim i As Long
        Dim bCancel As Boolean
        Dim sDDConnStr As String
        Dim bClientDDInOrder As Boolean
        Dim sDatabaseFullPath As String
        Dim oCN As ADODB.Connection
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
                            m_bIncidentsV2 = IncidentsV2Active(oCN)
                            
152                         Set m_ColNewIncidents = New Collection
154                         Set m_ColDeletedIncidents = New Collection

156                         If Not Right$(m_StrWebsiteURL, 1) = "/" Then m_StrWebsiteURL = m_StrWebsiteURL & "/"
        
158                         If FetchServerTime(m_StrWebsiteURL) Then
            
160                             GetNumberOfBatchToDownload m_StrWebsiteURL
                                Set m_RSttkGISLayerSQLRemote = OpenServerRSCompressed(m_StrWebsiteURL, "ID", "SELECT * FROM ttkGISLayerSQL")
                                
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '' SYNCH GEOBOOKMARKS
                    
162                             If m_bSynchGeoMarks Then

164                                 ValidateClientTable "GeoBookMarks", "GUID1", True, False, oCN, ""
166                                 ValidateClientTable "GeoBookMarksCategories", "GUID1", True, False, oCN, ""

168                                 GetSynchRecordsTOADD "GeoBookMarksCategories", "GUID1", False, m_StrWebsiteURL, oCN
170                                 GetSynchRecordsTOADD "GeoBookMarks", "GUID1", False, m_StrWebsiteURL, oCN

172                                 If Not SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch And Not SynchHistoryOverview.NextSynchNextOpeation = EditsOnly Then
174                                     GetSynchRecordsTODELETE "GeoBookMarks", "GUID1", False, m_StrWebsiteURL, oCN
                                    End If

176                                 If SynchHistoryOverview.NextSynchNextOpeation = EditsOnly Then
178                                     GetSynchRecordsTOUPDATE "GeoBookMarks", "GUID1", False, m_StrWebsiteURL, oCN
                                    End If
                                    
180                                 UploadTableChangesToServer "GeoBookMarks", "GUID1", False, m_StrWebsiteURL, oCN, SynchOperation.Add

                                End If

                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '' SYNCH INCIDENT CHARTS
                        
                                If m_bIncidentsV2 Then
182                                 bSkipIncidents = True
                                Else
                                    bSkipIncidents = Not CheckIfIncidentsEnabled(oCN)
                                End If

184                             If Not bSkipIncidents Then

186                                 ValidateClientTable "Incidents_ChartSettings", "GUID1", True, False, oCN, ""

188                                 GetSynchRecordsTOADD "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN

190                                 If Not SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch And Not SynchHistoryOverview.NextSynchNextOpeation = EditsOnly Then
192                                     GetSynchRecordsTODELETE "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN
                                    End If

194                                 If SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch Then
196                                     GetSynchRecordsTOUPDATE "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN
                                    End If
                
198                                 UploadTableChangesToServer "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN, SynchOperation.Add
200                                 UploadTableChangesToServer "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN, SynchOperation.Delete
202                                 UploadTableChangesToServer "Incidents_ChartSettings", "GUID1", False, m_StrWebsiteURL, oCN, SynchOperation.Update
                            
                                End If
                                
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '' SYNCH INCIDENTS
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

204                             If Not bSkipIncidents Then

                                    Set m_RSttkGISLayerSQLLocal = New ADODB.Recordset
                                    m_RSttkGISLayerSQLLocal.Open "SELECT * FROM ttkGISLayerSQL", oCN, adOpenDynamic, adLockBatchOptimistic
                                    UpdatettkGISLaySQL "oincidents", oCN

206                                 ValidateClientTable "oincidents", "ID", True, True, oCN, ""

208                                 GetSynchRecordsTOADD "oincidents", "ID", True, m_StrWebsiteURL, oCN
                                    
210                                 Do While SynchHistoryOverview.NextSynchNextOpeation = NewOnly
212                                     GetSynchRecordsTOADD "oincidents", "ID", True, m_StrWebsiteURL, oCN
                                    Loop

214                                 TxMsg CStr("**" & CStr(RT_IncidentSynch) & m_sIDs), True

216                                 If Not SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch And Not SynchHistoryOverview.NextSynchNextOpeation = EditsOnly Then
218                                     GetSynchRecordsTODELETE "oincidents", "ID", True, m_StrWebsiteURL, oCN
                                    End If

220                                 If SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch Then
222                                     GetSynchRecordsTOUPDATE "oincidents", "ID", True, m_StrWebsiteURL, oCN
                                    End If

224                                 UploadTableChangesToServer "oincidents", "ID", False, m_StrWebsiteURL, oCN, SynchOperation.Add
                                    
                                End If
                    
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                '' SYNCH DYNAMIC DATA DEFS
                                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
226                             If 1 = 1 Then 'Currently DynamicDataDefs are synched automatically

228                                 Set oRSSQLLyr = New ADODB.Recordset
                        
230                                 With oRSSQLLyr

232                                     If Not bSQLServerInUse Then
234                                         .Open "SELECT DISTINCT DDDefName, ConnectionString FROM DynamicDataDefs WHERE [Synch] = True", oCN, adOpenDynamic, adLockReadOnly
                                        Else
236                                         .Open "SELECT DISTINCT [DDDefName], CAST([ConnectionString] AS varchar(max)) FROM DynamicDataDefs WHERE [Synch] = 'True'", oCN, adOpenDynamic, adLockReadOnly
                                        End If
                                    
238                                     If Not .BOF Or Not .EOF Then

240                                         .MoveFirst
                                
242                                         Do While Not .EOF
            
244                                             Set DDCnn = New ADODB.Connection

246                                             If bSQLServerInUse Then
                                                    'DDCnn.CursorLocation = adUseClient
248                                                 TxMsg "Connecting to OasisV3 MSSQL Database......"
250                                                 DDCnn.Open m_StrLocalConnectionString
                                                Else

252                                                 sDDConnStr = oCN.Properties.Item("Data Source Name").Value
254                                                 sDDConnStr = Left$(sDDConnStr, Len(sDDConnStr) - 15)
256                                                 sDDConnStr = Replace$(.Fields(1).Value, "\data\db\", sDDConnStr, , , vbTextCompare)
258                                                 DDCnn.CursorLocation = adUseClient
260                                                 TxMsg "Connecting to connection string: " & sDDConnStr, True
262                                                 sDatabaseFullPath = Mid$(sDDConnStr, InStr(sDDConnStr, "Data Source=") + 12, InStr(sDDConnStr, "Mode=") - (InStr(sDDConnStr, "Data Source=") + 12) - 1)
264                                                 DDCnn.ConnectionString = sDDConnStr

266                                                 If Not CheckIfFileExists(sDatabaseFullPath) Then
268                                                     TxMsg "   --- Creating dynamic database: " & sDatabaseFullPath, True
270                                                     CreateAccessDatabase sDatabaseFullPath
                                                    End If
                                                
272                                                 DDCnn.Open sDDConnStr

                                                    If Not DoesTableExist(DDCnn, "ttkGISLayerSQL") Then
            
                                                        DDCnn.Execute "CREATE TABLE [ttkGISLayerSQL]  ([NAME] TEXT(255), [XMIN] DOUBLE, [XMAX] DOUBLE, [YMIN] DOUBLE, [YMAX] DOUBLE, [SHAPETYPE] INTEGER)"
                                                        TxMsg "   --- Creating table [ttkGISLayerSQL]"

                                                    End If
                                                
                                                End If
                                                
                                                Set m_RSttkGISLayerSQLLocal = New ADODB.Recordset
                                                m_RSttkGISLayerSQLLocal.Open "SELECT * FROM ttkGISLayerSQL", DDCnn, adOpenDynamic, adLockBatchOptimistic
                                
274                                             bClientDDInOrder = RunDDSchemaCompare(m_StrWebsiteURL, "dd_" & .Fields(0) & "_", DDCnn)
                                                
276                                             If bClientDDInOrder Then
278                                                 SynchDDWithServer "dd_" & .Fields(0) & "_", m_StrWebsiteURL, DDCnn
                                                End If
                                            
280                                             DDCnn.Close
282                                             Set DDCnn = Nothing
284                                             .MoveNext

                                            Loop
                        
286                                         If Len(sTablesCreated) > 0 Then TxMsg "**?? TABLES CREATED: " & sTablesCreated, True
288                                         sTablesCreated = ""
                                            
                                        End If
                            
                                    End With
                        
290                                 Set oRSSQLLyr = Nothing
                    
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
    
292         Set oRSSQLLyr = Nothing
            
        Loop

294     TxMsg "**", True
296     TxMsg "**??Synchronisation complete", True
298     m_bRunning = False

        '<EhFooter>
        Exit Sub

Runnable_Start_Err:
        TxMsg "SynchWorker.Runnable_Start_Err on line (" & Erl & ") " & Err.Description, True
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
                
142             oRS.Close
144             Set oRS = Nothing
                
146             If m_RSttkGISLayerSQLRemote.State = adStateOpen Then
        
148                 If Not (m_RSttkGISLayerSQLRemote.EOF And m_RSttkGISLayerSQLRemote.BOF) Then
        
150                     m_RSttkGISLayerSQLRemote.MoveFirst
152                     m_RSttkGISLayerSQLRemote.Find "[name] = '" & sTableName & "'"
                        
154                     If Not m_RSttkGISLayerSQLRemote.EOF Then
                        
156                         .Fields("Name").Value = m_RSttkGISLayerSQLRemote.Fields("Name").Value
158                         .Fields("XMIN").Value = m_RSttkGISLayerSQLRemote.Fields("XMIN").Value
160                         .Fields("XMAX").Value = m_RSttkGISLayerSQLRemote.Fields("XMAX").Value
162                         .Fields("YMIN").Value = m_RSttkGISLayerSQLRemote.Fields("YMIN").Value
164                         .Fields("YMAX").Value = m_RSttkGISLayerSQLRemote.Fields("YMAX").Value
166                         .Fields("SHAPETYPE").Value = m_RSttkGISLayerSQLRemote.Fields("SHAPETYPE").Value
                        
                        End If
                
                    End If
                
                End If
            
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
                                ByVal bIsGeotable As Boolean, _
                                cn As ADODB.Connection, _
                                ByVal sSynchHistoryPrefix As String)
    
    Dim sTableDesc As String
    sTableDesc = IIf(bIsGeotable, "GEO-table", "table")
    
    'TxMsg "   ------------------------------------------------------------------------------------------------"
    TxMsg "**?? Verifying " & sTableDesc & " [" & sTableName & "]" & " for integrity....."
    'TxMsg "   ------------------------------------------------------------------------------------------------"
    'TxMsg "   --- Verifying " & sTableDesc & " [" & sTableName & "]" & " for duplicate field [" & sGUIDFieldName & "]....."
    
    CheckForClientTableDuplicates sTableName, sGUIDFieldName, bGUIDFieldIsString, bIsGeotable, cn, sSynchHistoryPrefix
    
    If bIsGeotable Then
        'TxMsg "   --- Verifying " & sTableDesc & " [" & sTableName & "]" & " for duplicate UID....."
        CheckForClientTableDuplicateUID sTableName, cn, sGUIDFieldName, sSynchHistoryPrefix
        
        'TxMsg "   --- Verifying " & sTableDesc & " [" & sTableName & "]" & " for integrity....."
        ValidateClientGEOTable sTableName, cn, sGUIDFieldName, sSynchHistoryPrefix
    End If
    
    'TxMsg "   --- Verifying " & sTableDesc & " [" & sTableName & "]" & " is in accordance with SynchHistory....."
    CheckForClientTableSynchRules sTableName, sGUIDFieldName, bIsGeotable, cn, sSynchHistoryPrefix
    
    'TxMsg "   --- Verification of " & sTableDesc & " [" & sTableName & "]" & " complete."
    'TxMsg "   ------------------------------------------------------------------------------------------------"
    
End Sub

Private Sub CheckForClientTableSynchRules(ByVal sTableName As String, _
                                          ByVal sGUIDFieldName As String, _
                                          ByVal bIsGeotable As Boolean, _
                                          cn As ADODB.Connection, _
                                          ByVal sLocalSynchHistPrefix As String)
        '<EhHeader>
        On Error GoTo CheckForClientTableSynchRules_Err
        '</EhHeader>
    
        Dim sSQL As String
        Dim RSDifferences As New ADODB.Recordset
            
        If Not Right(sTableName, 4) = "_GEO" Then
        
100         If bIsGeotable = True Then
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

118                 If bIsGeotable = True Then
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
                                          bIsGeotable As Boolean, _
                                          cn As ADODB.Connection, sLocalSynchHistPrefix As String)
        '<EhHeader>
        On Error GoTo CheckForClientTableDuplicates_ERR
        '</EhHeader>
    
        Dim sSQL As String
        Dim sBrackets As String

100     If bGUIDFieldIsString Then sBrackets = "'"
        Dim RSLocalDuplicates As New ADODB.Recordset
    
102     If bIsGeotable Then
    
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
    
Private Sub UpdateSynchHistoryOverview(cn As Connection, ByVal sWebsite As String)
        '<EhHeader>
        On Error GoTo UpdateSynchHistoryOverview_Err
        '</EhHeader>

        Dim oRSLocal As New ADODB.Recordset
        Dim oRSServer As New ADODB.Recordset
        Dim sOperation As String
        Dim sWhenNow As String
        
        sWhenNow = CalcServerDateTime(Now())
    
100     With SynchHistoryOverview
    
102         oRSLocal.Open "SELECT count(*) FROM [" & IIf(.sTablenameLocal = "oincidents", "oincidents_FEA", .sTablenameLocal) & "]", cn, adOpenDynamic, adLockReadOnly
104         .lCountLocal = 0
106         If Not oRSLocal.EOF Then .lCountLocal = oRSLocal.Fields(0).Value
108         oRSLocal.Close
        
110         oRSLocal.Open "SELECT top 1 [sWhen] FROM [" & .sSynchHistoryTableLocal & "] WHERE updates = 'true' AND sDelete = 'false' AND sTablename = '" & .sTablenameLocal & "' AND NOT ([noconflict]='local') ORDER BY [sWhen] DESC", cn, adOpenDynamic, adLockReadOnly
112         .sWhenMaxActualLocal = sWhenNow
114         If Not oRSLocal.EOF Then .sWhenMaxActualLocal = oRSLocal.Fields(0).Value
116         oRSLocal.Close
        
118         oRSLocal.Open "SELECT top 1 [sWhen] FROM [" & .sSynchHistoryTableLocal & "] WHERE updates = 'true' AND sDelete = 'false' AND sTablename = '" & .sTablenameLocal & "' AND NOT ([noconflict]='local') ORDER BY [sWhen]", cn, adOpenDynamic, adLockReadOnly
120         .sWhenMinActualLocal = sWhenNow
122         If Not oRSLocal.EOF Then .sWhenMinActualLocal = oRSLocal.Fields(0).Value
124         oRSLocal.Close
    
126         Set oRSServer = OpenServerRSCompressed(sWebsite, "id", "SELECT count(*) FROM [" & .sTablenameServer & "]")
128         .lCountServer = 0
130         If Not oRSServer.EOF Then .lCountServer = oRSServer.Fields(0).Value
132         oRSServer.Close
        
134         Set oRSServer = OpenServerRSCompressed(sWebsite, "id", "SELECT top 1 [sWhen] FROM [" & .sSynchHistoryTableServer & "] WHERE sDelete = 'false' AND sTablename = '" & .sTablenameServer & "' ORDER BY [sWhen] DESC")
136         .sWhenMaxActualServer = "" 'CalcServerDateTime(Now())
138         If Not oRSServer.EOF Then .sWhenMaxActualServer = oRSServer.Fields(0).Value
140         oRSServer.Close
        
142         Set oRSServer = OpenServerRSCompressed(sWebsite, "id", "SELECT top 1 [sWhen] FROM [" & .sSynchHistoryTableServer & "] WHERE sDelete = 'false' AND sTablename = '" & .sTablenameServer & "' ORDER BY [sWhen]")
144         .sWhenMinActualServer = "" 'CalcServerDateTime(Now())
146         If Not oRSServer.EOF Then .sWhenMinActualServer = oRSServer.Fields(0).Value
148         oRSServer.Close
        
150         If (.lCountLocal = .lCountServer _
                And Not .lCountServer = 0 _
                And Not .sWhenMaxActualLocal = .sWhenMaxActualServer) Then
                
152             .NextSynchNextOpeation = EditsOnly
154             sOperation = "EditsOnly"
            ElseIf .lCountLocal = .lCountServer Then
                .NextSynchNextOpeation = DoNotSynch
                sOperation = "DoNotSynch"
156         ElseIf .sWhenMaxActualLocal = .sWhenMaxActualServer Or .sWhenMaxActualLocal = sWhenNow Then
158             .NextSynchNextOpeation = OldOnly
160             sOperation = "OldOnly"
162         ElseIf .sWhenMinActualLocal = .sWhenMaxActualServer Then
164             .NextSynchNextOpeation = NewOnly
166             sOperation = "NewOnly"
            Else
168             .NextSynchNextOpeation = NewOnly
170             sOperation = "NewOnly - there are old records also but first we want to get all the latest records"
            End If
        
172         TxMsg ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", True, True
174         TxMsg ">>> Local record count : " & .lCountLocal, True, True
176         TxMsg ">>> Local min sWhen    : " & .sWhenMinActualLocal, True, True
178         TxMsg ">>> Local max sWhen    : " & .sWhenMaxActualLocal, True, True
182         TxMsg ">>> Server record count: " & .lCountServer, True, True
184         TxMsg ">>> Server min sWhen   : " & .sWhenMinActualServer, True, True
186         TxMsg ">>> Server max sWhen   : " & .sWhenMaxActualServer, True, True
190         TxMsg "**?? Table [" & .sTablenameLocal & "] -- Operation: " & sOperation & " -- Records: " & .lCountLocal & " on local, " & .lCountServer & " on server", True
192         TxMsg ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", True, True
        
        End With
    
194     Set oRSLocal = Nothing
196     Set oRSServer = Nothing
    
        '<EhFooter>
        Exit Sub

UpdateSynchHistoryOverview_Err:
        TxMsg "SynchWorker.UpdateSynchHistoryOverview_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Function GetSynchRecordsTOADD(ByVal sTableName As String, _
                                      ByVal sGUID As String, _
                                      bIsGeotable As Boolean, _
                                      ByVal sWebsite As String, _
                                      cn As ADODB.Connection, _
                                      Optional sLocalSynchHistPrefix As String = "") As Boolean
        '<EhHeader>
        On Error GoTo GetSynchRecordsTOADD_Err
        '</EhHeader>
        
        'Operation: This sub loads all GUIDs off a server table and all GUIDs off a local table
        '           The list of GUIDs are compared and any missing on the local table
        '           are downloaded from the server table and saved in a table on
        '           the local database by the same name as the server table
        
        Dim sWhenCurrent As String
        m_lCountDownloaded = 0
        
        'Dim RSSynchHistoryOverview As New ADODB.Recordset
        Dim RSServerALLGUIDS As ADODB.Recordset
        Dim RSLocalALLGUIDS As ADODB.Recordset
        Dim RSGUIDSNeedingDownloading As ADODB.Recordset
        
        Dim RSServerNewRecord As New ADODB.Recordset
        Dim RSServerNewRecord_FEA As New ADODB.Recordset
        Dim RSServerNewRecord_GEO As New ADODB.Recordset
        
        Dim RSUpdater As New ADODB.Recordset
        Dim RSUpdater_GEO As New ADODB.Recordset
        Dim RSUpdater_FEA As New ADODB.Recordset
        Dim bIncidents As Boolean
        
        Dim bMasterTableEntryExists As Boolean
        Dim bValidation As Boolean
        Dim sServerDate() As String
        Dim sServerSequence() As Long
        Dim sServerGUID() As String
        Dim sServerDateFound As String
        Dim sServerSequenceFound As Long
        'Dim sServerGUID() As String
        Dim lCurrentRecordInBatch As Long
        Dim sAllGUIDSForDownload As String
        Dim sAllUIDSForDownload As String
        Dim lMaxRecordBatch As Long
        Dim i As Long
        Dim iArrayIndex As Long
        Dim j As Long
        Dim lMAX As Long
        Dim sUGName As String
        
        Dim bCatchFailAddLocalValue_FEA As Boolean
        Dim bCatchFailAddLocalValue_GEO As Boolean
        Dim bCatchFailOpenRSUpdater_FEA As Boolean
        
100     GetSynchRecordsTOADD = False
102     lCurrentRecordInBatch = 0
104     lMaxRecordBatch = lBatchCountToDownload
106     m_sIDs = ""
        
108     If sTableName = "oincidents" Then
110         'TxMsg "**??Scanning for new incidents...", True
112         bIncidents = True
114         sUGName = ""
116     ElseIf sTableName = "Attachments" Then
118         'TxMsg "**??Scanning for new attachments...", True
120         sUGName = ""
        Else

122         If sLocalSynchHistPrefix = "" Then sUGName = m_StrRemoteTablePrefix
            'If sLocalSynchHistPrefix = "" Then sUGName = m_StrRemoteTablePrefix
124         'TxMsg "**??Scanning for new records on server table: [" & sUGName & sTableName & "]", False
        End If
        
'126     RSSynchHistoryOverview.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistoryOverview] WHERE [sTablename] = '" & sTableName & "'", cn, adOpenDynamic, adLockBatchOptimistic
'
'128     If RSSynchHistoryOverview.EOF Then
'130         RSSynchHistoryOverview.AddNew
'132         RSSynchHistoryOverview.Fields("sWhenMinSearched").Value = CalcServerDateTime(Now())
'134         RSSynchHistoryOverview.Fields("sWhenMaxSearched").Value = RSSynchHistoryOverview.Fields("sWhenMinSearched").Value
'136         RSSynchHistoryOverview.Fields("sTablename").Value = sTableName
'138         RSSynchHistoryOverview.UpdateBatch adAffectCurrent
'        End If

        SynchHistoryOverview.sTablenameLocal = sTableName
        SynchHistoryOverview.sTablenameServer = sUGName & sTableName
        SynchHistoryOverview.sSynchHistoryTableLocal = sLocalSynchHistPrefix & "SynchHistory"
        SynchHistoryOverview.sSynchHistoryTableServer = sUGName & sLocalSynchHistPrefix & "SynchHistory"
        
        TxMsg "SynchHistoryOverview.sTablenameLocal: " & SynchHistoryOverview.sTablenameLocal, True, True
        TxMsg "SynchHistoryOverview.sTablenameServer: " & SynchHistoryOverview.sTablenameServer, True, True
        TxMsg "SynchHistoryOverview.sSynchHistoryTableLocal: " & SynchHistoryOverview.sSynchHistoryTableLocal, True, True
        TxMsg "SynchHistoryOverview.sSynchHistoryTableServer: " & SynchHistoryOverview.sSynchHistoryTableServer, True, True
        
140     UpdateSynchHistoryOverview cn, sWebsite
        
142     If Not SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch And Not SynchHistoryOverview.NextSynchNextOpeation = EditsOnly Then
        
            TxMsg "**??Scanning for new records on server table: [" & sUGName & sTableName & "]", True
            
            SynchHistoryOverview.sWhenMinSearched = SynchHistoryOverview.sWhenMinActualLocal
            SynchHistoryOverview.sWhenMaxSearched = SynchHistoryOverview.sWhenMaxActualLocal
            
            
'144         SynchHistoryOverview.sWhenMinSearched = RSSynchHistoryOverview.Fields("sWhenMinSearched").Value
'146         SynchHistoryOverview.sWhenMaxSearched = RSSynchHistoryOverview.Fields("sWhenMaxSearched").Value

            TxMsg ">>> sWhenMinSearched: " & SynchHistoryOverview.sWhenMinSearched, True, True
            TxMsg ">>> sWhenMaxSearched: " & SynchHistoryOverview.sWhenMaxSearched, True, True
            TxMsg ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", True, True
        
148         If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then sWebsite = sWebsite & "/"
150         Set RSLocalALLGUIDS = GetAllGUIDsFromLocalDB(sTableName, sLocalSynchHistPrefix & "SynchHistory", cn, Add)
151         'Set RSServerALLGUIDS = GetAllGUIDsFromServerDB(sUGName & sTableName, sLocalSynchHistPrefix & sUGName & "SynchHistory", m_sWhen, sWebsite, Add)

            If sTableName = "Incidents_ChartSettings" Then
152             Set RSServerALLGUIDS = GetAllGUIDsFromServerDB(sTableName, sLocalSynchHistPrefix & sUGName & "SynchHistory", m_sWhen, sWebsite, Add)
            Else
153             Set RSServerALLGUIDS = GetAllGUIDsFromServerDB(sUGName & sTableName, sLocalSynchHistPrefix & sUGName & "SynchHistory", m_sWhen, sWebsite, Add)
            End If

154         Set RSGUIDSNeedingDownloading = GetRSWithGUIDSFORDOWNLOAD(RSServerALLGUIDS, RSLocalALLGUIDS, Add)
'            RSSynchHistoryOverview.Fields("sWhenMaxSearched").Value = SynchHistoryOverview.sWhenMinSearched
'            RSSynchHistoryOverview.Fields("sWhenMaxSearched").Value = SynchHistoryOverview.sWhenMaxSearched
'            RSSynchHistoryOverview.UpdateBatch adAffectCurrent
        
156         If Not RSGUIDSNeedingDownloading Is Nothing Then

158             With RSGUIDSNeedingDownloading

160                 If Not .EOF Then

162                     .MoveFirst
164                     GetSynchRecordsTOADD = True

166                     TxMsg "    --- " & RSGUIDSNeedingDownloading.RecordCount & " new record(s) are to be downloaded from the OASIS Cloud Server............"
168                     sDebugInfo = "Here is the debug info for function: GetSynchRecordsTOADD"
170                     i = 0

172                     Do Until .EOF

174                         ReDim sServerDate(lMaxRecordBatch)
176                         ReDim sServerSequence(lMaxRecordBatch)
178                         ReDim sServerGUID(lMaxRecordBatch)
                        
180                         iArrayIndex = 0
182                         sAllGUIDSForDownload = "WHERE"
184                         lCurrentRecordInBatch = 0
186                         TxMsg "    --- Processing download list (not downloading right now)............"
                
188                         Do Until .EOF Or (lCurrentRecordInBatch = lMaxRecordBatch)

190                             lCurrentRecordInBatch = lCurrentRecordInBatch + 1
                            
192                             If sAllGUIDSForDownload = "WHERE" Then
194                                 sAllGUIDSForDownload = sAllGUIDSForDownload & " ("
                                Else
196                                 sAllGUIDSForDownload = sAllGUIDSForDownload & " OR "
                                End If
                        
198                             sServerGUID(iArrayIndex) = .Fields("sID").Value
200                             sServerDate(iArrayIndex) = .Fields("sWhen").Value
202                             sServerSequence(iArrayIndex) = .Fields("sequence").Value
204                             sAllGUIDSForDownload = sAllGUIDSForDownload & "[" & sGUID & "] = '" & .Fields(0).Value & "'"
206                             .MoveNext
208                             iArrayIndex = iArrayIndex + 1
                        
                            Loop
                            
210                         If bIncidents Or bIsGeotable Then
212                             If Not sLocalSynchHistPrefix = "" And Right(sTableName, 4) = "_FEA" Then sTableName = Left(sTableName, Len(sTableName) - 4)
214                             RSUpdater_GEO.CursorLocation = adUseClient
216                             RSUpdater_FEA.CursorLocation = adUseClient
                            
218                             RSUpdater_GEO.Open "SELECT * FROM [" & sTableName & "_GEO] ORDER BY [UID]", cn, adOpenDynamic, adLockBatchOptimistic
220                             bCatchFailOpenRSUpdater_FEA = True
222                             RSUpdater_FEA.Open "SELECT * FROM [" & sTableName & "_FEA] " & sAllGUIDSForDownload & ") ORDER BY [" & sGUID & "]", cn, adOpenDynamic, adLockBatchOptimistic
224                             bCatchFailOpenRSUpdater_FEA = False
                            Else
226                             RSUpdater.CursorLocation = adUseClient
                                TxMsg ">>>>>> LOADING LOCALLY: SELECT * FROM [" & sTableName & "] " & sAllGUIDSForDownload & ") ORDER BY [" & sGUID & "]", True, True
228                             RSUpdater.Open "SELECT * FROM [" & sTableName & "] " & sAllGUIDSForDownload & ") ORDER BY [" & sGUID & "]", cn, adOpenDynamic, adLockBatchOptimistic
                            End If
                        
230                         sAllGUIDSForDownload = sAllGUIDSForDownload & ") ORDER BY [sWhen] DESC, [sid]"
                            'TxMsg "sAllGUIDSForDownload: " & sAllGUIDSForDownload, True
                            'iArrayIndex = 0

                            'Get new record from server
232                         If bIncidents Or bIsGeotable Then

234                             TxMsg "    --- Downloading batch............"
236                             Set RSServerNewRecord_FEA = DownloadBatchRecordsFromServer(IIf(bIncidents, "oincidents_FEA", sUGName & sTableName & "_FEA"), sAllGUIDSForDownload, sWebsite, sLocalSynchHistPrefix, sGUID)

238                             If Not bBatchDownloadOK Then GoTo cleanupandexit

240                             If Not RSServerNewRecord_FEA.State = adStateClosed Then
242                                 If Not RSServerNewRecord_FEA.EOF Or Not RSServerNewRecord_FEA.BOF Then

244                                     RSServerNewRecord_FEA.MoveFirst
                                    
246                                     sAllUIDSForDownload = "WHERE"

248                                     Do Until RSServerNewRecord_FEA.EOF
                                        
250                                         If sAllUIDSForDownload = "WHERE" Then
252                                             sAllUIDSForDownload = sAllUIDSForDownload & " ("
                                            Else
254                                             sAllUIDSForDownload = sAllUIDSForDownload & " OR "
                                            End If
                                            
256                                         sAllUIDSForDownload = sAllUIDSForDownload & "[UID] = " & RSServerNewRecord_FEA.Fields("UID").Value
258                                         RSServerNewRecord_FEA.MoveNext
                                            
                                        Loop
                                    
260                                     sAllUIDSForDownload = sAllUIDSForDownload & ")"
262                                     RSServerNewRecord_FEA.MoveFirst
264                                     Set RSServerNewRecord_GEO = DownloadBatchRecordsFromServer(IIf(bIncidents, "oincidents_GEO", sUGName & sTableName & "_GEO"), sAllUIDSForDownload, sWebsite, sLocalSynchHistPrefix, sGUID)

266                                     If Not bBatchDownloadOK Then GoTo cleanupandexit
                                    
                                    End If
                                End If

268                             If Not RSServerNewRecord_FEA.State = adStateClosed And Not RSServerNewRecord_GEO.State = adStateClosed Then

270                                 If Not RSServerNewRecord_FEA.BOF Or Not RSServerNewRecord_FEA.EOF Then RSServerNewRecord_FEA.MoveFirst

272                                 Do Until RSServerNewRecord_FEA.EOF
274                                     i = i + 1
                                        'iArrayIndex = iArrayIndex + 1

276                                     iArrayIndex = 0
                                        
278                                     Do Until iArrayIndex >= UBound(sServerGUID)

280                                         If sServerGUID(iArrayIndex) = RSServerNewRecord_FEA.Fields(sGUID).Value Then
282                                             sServerDateFound = sServerDate(iArrayIndex)
284                                             sServerSequenceFound = sServerSequence(iArrayIndex)
                                            End If

286                                         iArrayIndex = iArrayIndex + 1
                                        Loop
                                                                              
288                                     If Not RSUpdater_FEA.EOF Or Not RSUpdater_FEA.BOF Then RSUpdater_FEA.MoveFirst
290                                     RSUpdater_FEA.Find sGUID & " = '" & RSServerNewRecord_FEA.Fields(sGUID).Value & "'"
                                        'RSUpdater_FEA.Filter = "[" & sGUID & "] = '" & RSServerNewRecord_FEA.Fields(sGUID).Value & "'"
                                        
292                                     If Not RSServerNewRecord_GEO.EOF Or Not RSServerNewRecord_GEO.BOF Then RSServerNewRecord_GEO.MoveFirst
294                                     RSServerNewRecord_GEO.Find "UID = " & RSServerNewRecord_FEA.Fields("UID").Value
                                        'RSServerNewRecord_GEO.Filter = "UID = " & RSServerNewRecord_FEA.Fields("UID").Value
                                    
296                                     If (RSUpdater_FEA.EOF Or RSUpdater_FEA.BOF) Then
    
298                                         lMAX = GetMaxUIDFromGEOTable(sTableName & "_GEO", cn) + CLng(1)
300                                         j = 0
302                                         RSUpdater_FEA.AddNew
                                    
304                                         Do Until j = RSUpdater_FEA.Fields.Count

                                                If Not RSUpdater_FEA.Fields(j).Name = "WKB" And Not RSUpdater_FEA.Fields(j).Name = "ORIG_ID" And Not Left(RSUpdater_FEA.Fields(j).Name, 1) = "_" Then

306                                                 If RSUpdater_FEA.Fields(j).Name = "UID" Then
308                                                     RSUpdater_FEA.Fields(j).Value = lMAX
                                            
                                                    Else
310                                                     bCatchFailAddLocalValue_FEA = True
312                                                     RSUpdater_FEA.Fields(j).Value = RSServerNewRecord_FEA.Fields(RSUpdater_FEA.Fields(j).Name).Value
314                                                     bCatchFailAddLocalValue_FEA = False
                                                    End If
                                                End If

316                                             j = j + 1
                                            Loop
                        
318                                         j = 0
320                                         RSUpdater_GEO.AddNew
                                        
322                                         Do Until j = RSUpdater_GEO.Fields.Count
                                    
324                                             If RSUpdater_GEO.Fields(j).Name = "UID" Then
326                                                 RSUpdater_GEO.Fields(j).Value = lMAX
328                                             ElseIf Not (bIncidents And RSUpdater_GEO.Fields(j).Name = "GEOMETRY") Then
330                                                 bCatchFailAddLocalValue_GEO = True
332                                                 RSUpdater_GEO.Fields(j).Value = RSServerNewRecord_GEO.Fields(RSUpdater_GEO.Fields(j).Name).Value
334                                                 bCatchFailAddLocalValue_GEO = False
                                                End If
                                        
336                                             j = j + 1
                                            Loop
                                    
338                                         RSUpdater_FEA.UpdateBatch adAffectCurrent
340                                         RSUpdater_GEO.UpdateBatch adAffectCurrent
342                                         SynchHistoryAddNew cn, GetGuid, RSServerNewRecord_FEA.Fields(sGUID).Value, "Synched Record", "", sUGName & " UG", sServerDateFound, IIf(sLocalSynchHistPrefix = "", sTableName, sTableName & "_FEA"), True, "true", sLocalSynchHistPrefix, sServerSequenceFound
344                                         sWhenCurrent = sServerDateFound
                
                                            ' Report to client
346                                         If Not m_bUseEvents Then
348                                             If bIncidents Then
                                                    m_lCountDownloaded = m_lCountDownloaded + 1
350                                                 TxMsg CStr("**??New incident " & i & " of " & RSGUIDSNeedingDownloading.RecordCount & " and ID: " & RSServerNewRecord_FEA.Fields(sGUID).Value)
352                                                 m_sIDs = RSServerNewRecord_FEA.Fields(sGUID).Value & "//" & m_sIDs
                                                Else
                                                    'TxMsg CStr("    --- record (" & i & " of " & RSGUIDSNeedingDownloading.RecordCount & ") and UID: " & lMAX & " and  ID: " & RSServerNewRecord_FEA.Fields(sGUID).Value & " added")
                                                    m_lCountDownloaded = m_lCountDownloaded + 1
354                                                 TxMsg CStr("**??New record " & i & " of " & RSGUIDSNeedingDownloading.RecordCount & " for table: " & IIf(sLocalSynchHistPrefix = "", sTableName, sTableName & "_FEA") & " with ID: " & RSServerNewRecord_FEA.Fields(sGUID).Value)
                                                End If
    
                                            Else
356                                             RaiseEvent StatusEX(i, RSGUIDSNeedingDownloading.RecordCount, False, RSServerNewRecord_FEA.Fields(sGUID).Value)
                                            End If
    
                                        Else
                                    
358                                         TxMsg CStr("    --- record (" & i & " of " & RSGUIDSNeedingDownloading.RecordCount & ") with ID: " & RSServerNewRecord_FEA.Fields(sGUID).Value & " is already on local table [" & IIf(sLocalSynchHistPrefix = "", sTableName, sTableName & "_FEA") & "] adjusting SynchHistory accordingly....") ', True
360                                         SynchHistoryAddNew cn, GetGuid, RSServerNewRecord_FEA.Fields(sGUID).Value, "Synched Record", "", sUGName & " UG", sServerDateFound, IIf(sLocalSynchHistPrefix = "", sTableName, sTableName & "_FEA"), True, "true", sLocalSynchHistPrefix, sServerSequenceFound
362                                         sWhenCurrent = sServerDateFound
                                            
                                        End If
                                        
364                                     RSServerNewRecord_FEA.MoveNext
366                                     RSServerNewRecord_GEO.MoveNext
                                    Loop
                
                                Else
368                                 bLargePolygonDownloadFailed = True
                                    'TxMsg "    --- closed RS made the download fail"
370                                 TxMsg "    --- closed RS made the download fail", True
372                                 TxMsg "    --- RSServerNewRecord_FEA.State = adStateClosed: " & (RSServerNewRecord_FEA.State = adStateClosed)
374                                 TxMsg "    --- RSServerNewRecord_GEO.State = adStateClosed: " & (RSServerNewRecord_GEO.State = adStateClosed)
376                                 GoTo cleanupandexit
                                End If
                                
378                             RSUpdater_FEA.Close
380                             RSUpdater_GEO.Close
        
                            Else
                                
382                             Set RSServerNewRecord = DownloadBatchRecordsFromServer(sUGName & sTableName, sAllGUIDSForDownload, sWebsite, sLocalSynchHistPrefix, sGUID)

384                             If Not bBatchDownloadOK Then GoTo cleanupandexit
                            
386                             If Not RSServerNewRecord.State = adStateClosed Then

388                                 If Not RSServerNewRecord.EOF Or Not RSServerNewRecord.BOF Then RSServerNewRecord.MoveFirst

390                                 Do Until RSServerNewRecord.EOF

392                                     iArrayIndex = 0
                                        
394                                     Do Until iArrayIndex >= UBound(sServerGUID)

396                                         If sServerGUID(iArrayIndex) = RSServerNewRecord.Fields(sGUID).Value Then
398                                             sServerDateFound = sServerDate(iArrayIndex)
400                                             sServerSequenceFound = sServerSequence(iArrayIndex)
                                            End If

402                                         iArrayIndex = iArrayIndex + 1
                                        Loop

404                                     i = i + 1
                                        '  iArrayIndex = iArrayIndex + 1

406                                     If Left$(Replace$(sTableName, sLocalSynchHistPrefix, "", , , vbTextCompare), 4) = "link" Then
                                        
408                                         bMasterTableEntryExists = DoesFieldExistInTable(cn, sLocalSynchHistPrefix & "mastertable", "GUID1", RSServerNewRecord.Fields("GUID1").Value)
                                    
                                        Else
                                    
                                            'Override if link table not synching
410                                         bMasterTableEntryExists = True
                                    
                                        End If
                                    
                                        'If Not RSUpdater.EOF Or Not RSUpdater.BOF Then RSUpdater.MoveFirst
                                    
412                                     If bMasterTableEntryExists Then
                                            'RSUpdater.Move
                                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                            'This IF statement block will help prevent duplication of records
                                            'If Not RSUpdater.EOF Or Not RSUpdater.BOF Then RSUpdater.MoveFirst
414                                         RSUpdater.Find "[" & sGUID & "] = '" & RSServerNewRecord.Fields(sGUID).Value & "'"
                                            'RSUpdater.Filter = "[" & sGUID & "] = '" & RSServerNewRecord.Fields(sGUID).Value & "'"

416                                         If RSUpdater.EOF Or RSUpdater.BOF Then

418                                             With RSUpdater
                    
420                                                 j = 0
422                                                 .AddNew

424                                                 Do Until j = .Fields.Count

                                                        If Not .Fields(j).Name = "WKB" And Not .Fields(j).Name = "ORIG_ID" And Not Left(.Fields(j).Name, 1) = "_" Then
426                                                         If Not .Fields(j).Name = "UID" Then
                                                                .Fields(j).Value = RSServerNewRecord.Fields(.Fields(j).Name).Value
                                                            End If
                                                        End If

428                                                     j = j + 1
                                                    Loop

430                                                 .UpdateBatch adAffectCurrent

                                                End With
                                           
432                                             SynchHistoryAddNew cn, GetGuid, RSServerNewRecord.Fields(sGUID).Value, "Synched Record", "", sUGName & " UG", sServerDateFound, sTableName, True, "true", sLocalSynchHistPrefix, sServerSequenceFound
434                                             sWhenCurrent = sServerDateFound

                                                ' Report to client
436                                             If Not m_bUseEvents Then
                                                    m_lCountDownloaded = m_lCountDownloaded + 1
438                                                 TxMsg CStr("**??New record " & i & " of " & RSGUIDSNeedingDownloading.RecordCount & " for table: " & sTableName & " with ID: " & RSServerNewRecord.Fields(sGUID).Value)

                                                    'TxMsg CStr("    --- record (" & i & " of " & RSGUIDSNeedingDownloading.RecordCount & ") and ID: " & RSServerNewRecord.Fields(sGUID).Value)
                                                Else
440                                                 RaiseEvent StatusEX(i, RSGUIDSNeedingDownloading.RecordCount, False, RSServerNewRecord.Fields(sGUID).Value)
                                                End If
                            
                                            Else
                                    
442                                             TxMsg CStr("    --- record (" & i & " of " & RSGUIDSNeedingDownloading.RecordCount & ") with ID: " & RSServerNewRecord.Fields(sGUID).Value & " is already on local table [" & sTableName & "] adjusting SynchHistory accordingly....") ', True
444                                             SynchHistoryAddNew cn, GetGuid, RSServerNewRecord.Fields(sGUID).Value, "Synched Record", "", sUGName & " UG", sServerDateFound, sTableName, True, "true", sLocalSynchHistPrefix, sServerSequenceFound
446                                             sWhenCurrent = sServerDateFound
                                            End If
                                    
                                        Else

448                                         TxMsg CStr("    --- record (" & i & " of " & RSGUIDSNeedingDownloading.RecordCount & ") with ID: " & RSServerNewRecord.Fields(sGUID).Value & " is does not have a mastertable record!  Not adding record to see if this gets resolved after a few synch cycles (if not, contact an oasis admin)..."), True
                                    
                                        End If
                    
450                                     RSServerNewRecord.MoveNext
                                    Loop
                
                                Else

452                                 sDebugInfo = "   --- closed RS made the download fail" & Chr(13) & "RSServerNewRecord.State = adStateClosed: " & (RSServerNewRecord.State = adStateClosed)
454                                 TxMsg sDebugInfo, True
456                                 GoTo cleanupandexit
                                End If
                                
458                             RSUpdater.Close
                
                            End If
                        
460                         If i >= 9000 Then
462                             TxMsg ">>>> Too many rows - next synch cycle will fetch more records for this table....", True
                                Exit Do
                            End If
                
                        Loop
                        
'464                     'If i > 0 Then
'466                     If SynchHistoryOverview.NextSynchNextOpeation = NewOnly Then
'468                         RSSynchHistoryOverview.Fields("sWhenMaxSearched").Value = sWhenCurrent
'                        Else
'470                         RSSynchHistoryOverview.Fields("sWhenMinSearched").Value = sWhenCurrent
                        'End If

'472                     RSSynchHistoryOverview.UpdateBatch adAffectCurrent
                        'End If
                        
'474                     RSSynchHistoryOverview.Close
                    End If
    
                End With
        
            Else
478             GetSynchRecordsTOADD = True 'disable look for update and delete records
            End If
            
            TxMsg "**??Scan complete for new records on the server table: " & sTableName & "......................", True
            
        End If
      
cleanupandexit:
        
480     '
482     Set RSServerALLGUIDS = Nothing
484     Set RSLocalALLGUIDS = Nothing
486     Set RSGUIDSNeedingDownloading = Nothing
488     Set RSServerNewRecord = Nothing
490     Set RSServerNewRecord_FEA = Nothing
492     Set RSServerNewRecord_GEO = Nothing
494     Set RSUpdater = Nothing
496     Set RSUpdater_GEO = Nothing
498     Set RSUpdater_FEA = Nothing
'500     Set RSSynchHistoryOverview = Nothing
        
        '<EhFooter>
        Exit Function

GetSynchRecordsTOADD_Err:
        TxMsg "SynchWorker.GetSynchRecordsTOADD_Err: Fetching new records in table [" & sTableName & "] failed" & Chr(13) & "(" & Erl & ") " & Err.Description, True
        On Error Resume Next
        
        If bCatchFailAddLocalValue_FEA Then
            TxMsg "---   Server FieldName : " & RSUpdater_FEA.Fields(j).Name, True
            TxMsg "---   Server Value     : " & RSServerNewRecord_FEA.Fields(RSUpdater_FEA.Fields(j).Name).Value, True

            If Left(sTableName, Len("oincidents")) = "oincidents" Then TxMsg "---   [ID]: " & RSServerNewRecord_FEA.Fields("ID").Value, True
        End If
        
        If bCatchFailAddLocalValue_GEO Then
            TxMsg "---   Server FieldName : " & RSUpdater_GEO.Fields(j).Name, True
            TxMsg "---   Server Value     : " & RSServerNewRecord_GEO.Fields(RSUpdater_GEO.Fields(j).Name).Value, True
        End If
        
        If bCatchFailOpenRSUpdater_FEA Then
            TxMsg "---   SELECT * FROM [" & sTableName & "_FEA] " & sAllGUIDSForDownload & ") ORDER BY [" & sGUID & "]", True
        End If
        
        GoTo cleanupandexit
        '</EhFooter>
End Function

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
    
100     TxMsg "<<<<<<<<<<<<"
102     TxMsg "<<< Scanning for tables to be wiped......................"

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
                        cn.Execute "DELETE FROM [" & sPrefix & "SynchHistoryOverview] WHERE [sTableName] = '" & .Fields("sTableName").Value & "'"
            
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
        
168     TxMsg "<<< End scan for tables to be wiped......................"
170     TxMsg "<<<<<<<<<<<<"
    
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

            .Open "SELECT UID FROM [" & sTableName & "] WHERE UID = 1000000000", oCOnn, adOpenDynamic, adLockBatchOptimistic
            
            If .EOF Or .BOF Then
            
                GetMaxUIDFromGEOTable = 999999999
            
            Else
                .Close
106             .Open "SELECT MAX(UID) FROM [" & sTableName & "] WHERE UID < 1000000000", oCOnn, adOpenDynamic, adLockBatchOptimistic
                
108             If Not .EOF And Not .BOF Then
110                 If Not IsNull(.Fields(0).Value) Then
112                     GetMaxUIDFromGEOTable = CLng(.Fields(0).Value)
                    End If
                End If
            
            End If
        
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

Private Function GetAllGUIDsFromLocalDB(sTableName As String, _
                                        sSynchHistory As String, _
                                        oCN As ADODB.Connection, _
                                        Operation As SynchOperation) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo GetAllGUIDsFromLocalDB_Err
        '</EhHeader>
                                               
        Dim sSQL As String
        Dim oRS As ADODB.Recordset

100     If Operation = Update Or Operation = Delete Then

102         Select Case Operation
                Case Update
104                 sSQL = "SELECT max([sWhen]) FROM [" & sSynchHistory & "] WHERE ([sTableName]='" & sTableName & "') AND ([updates]='true') AND ([sdelete]='false') AND ([sequence]>1) AND ([noconflict]='false') AND NOT ([noconflict]='local')"
               
106             Case Delete
108                 sSQL = "SELECT max([sWhen]) FROM [" & sSynchHistory & "] WHERE ([sTableName]='" & sTableName & "') AND ([updates]='true') AND ([sdelete]='true') AND ([noconflict]='false') AND NOT ([noconflict]='local')"
            End Select
    
110         TxMsg ">>> GetAllGUIDsFromLocalDB (sWhen): " & sSQL, True, True
112         Set oRS = New ADODB.Recordset
114         oRS.Open sSQL, oCN, adOpenStatic, adLockReadOnly

116         If Not oRS.EOF Then
                m_sWhen = IIf(IsNull(oRS.Fields(0).Value), "", oRS.Fields(0).Value)
                TxMsg ">>> m_sWhen: " & m_sWhen, True, True
            End If
            
118         oRS.Close
120         If Right(sTableName, 8) = "_Queries" Then m_sWhen = ""
    
        End If
    
122     If Operation = Add Then

124         With SynchHistoryOverview
            
126             Select Case .NextSynchNextOpeation
                    
                    Case NewOnly
128                     sSQL = "SELECT [sID], [sequence], [swhen] FROM [" & sSynchHistory & "] GROUP BY [sID], [sequence], [swhen], [sdelete], [sTableName], [updates] HAVING ([sTableName]='" & sTableName & "') AND ([updates]='true') AND ([sdelete]='false') AND ([sWhen] >= '" & .sWhenMaxSearched & "') ORDER BY [sWhen]"
                    
130                 Case OldOnly
132                     sSQL = "SELECT [sID], [sequence], [swhen] FROM [" & sSynchHistory & "] GROUP BY [sID], [sequence], [swhen], [sdelete], [sTableName], [updates] HAVING ([sTableName]='" & sTableName & "') AND ([updates]='true') AND ([sdelete]='false') AND ([sWhen] <= '" & .sWhenMinSearched & "') ORDER BY [sWhen] DESC"

134                 Case Else
136                     sSQL = ""
                End Select
            
            End With
         
138     ElseIf Operation = Update Then
    
140         sSQL = "SELECT [sID], [sequence], [swhen] FROM [" & sSynchHistory & "] GROUP BY [sID], [sequence], [swhen], [sdelete], [sTableName], [updates] HAVING ([sTableName]='" & sTableName & "') AND ([updates]='true') AND ([sWhen] >= '" & m_sWhen & "') AND ([sdelete]='false') AND ([sequence]>1) ORDER BY [sWhen]"
    
142     ElseIf Operation = Delete Then
    
144         sSQL = "SELECT [sID], [sequence], [swhen] FROM [" & sSynchHistory & "] GROUP BY [sID], [sequence], [swhen], [sdelete], [sTableName], [updates] HAVING ([sTableName]='" & sTableName & "') AND ([updates]='true') AND ([sWhen] >= '" & m_sWhen & "') AND ([sdelete]='true') ORDER BY [sWhen]"
                
        End If
    
146     TxMsg ">>> GetAllGUIDsFromLocalDB: " & sSQL, True, True
148     sDebugInfo = Chr(13) & sSQL
150     Set oRS = New ADODB.Recordset
152     oRS.Open sSQL, oCN, adOpenStatic, adLockReadOnly
154     Set GetAllGUIDsFromLocalDB = oRS
        
        '<EhFooter>
        Exit Function

GetAllGUIDsFromLocalDB_Err:
        Set GetAllGUIDsFromLocalDB = Nothing
        TxMsg "GetAllGUIDsFromLocalDB_Err line " & Erl & " ... " & Err.Description, True
        '</EhFooter>
End Function

'Private Function GetAllGUIDsFromServerDB(sTableName As String, _
'                                         sGUIDFieldName As String, _
'                                         sWhereClause As String, _
'                                         sWebsite As String, _
'                                         Optional bMergeCheck As Boolean = False) As ADODB.Recordset
'
'     MsgBox "you should Not Get here!"
'
'
'    'Operation: The function returns a recordset containing a list of GUIDs in
'    '           in the specified table which is located on a server database
'    '
'    'Assumption: The GUID field should be a unique field of type text
'    'Assumption: The where clause MUST contain the command "WHERE" or be empty
'
''    Dim sSQL As String
''    Dim i As Long
''
''    If Not bMergeCheck Then
''        sSQL = "SELECT [" & sGUIDFieldName & "] FROM [" & sTableName & "] " & sWhereClause
''    Else
''        sSQL = "SELECT [" & sGUIDFieldName & "], [sequence], [swhen] FROM [" & sTableName & "] " & sWhereClause
''    End If
''
''    sDebugInfo = Chr(13) & sSQL
''    'TxMsg sSQL
''    Set GetAllGUIDsFromServerDB = OpenServerRSCompressed(sWebsite, "ID", sSQL)
'
'End Function

Private Function GetAllGUIDsFromServerDB(sTableName As String, _
                                         sSynchHistory As String, _
                                         sWhenMin As String, _
                                         sWebsite As String, _
                                         Operation As SynchOperation) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo GetAllGUIDsFromServerDB_Err
        '</EhHeader>
                                               
        Dim sSQL As String
        Dim lRowIndex As Long
        Dim bBreakout As Boolean
        Dim oRS As New ADODB.Recordset
        Dim lCountOfRecords As Long
        
100     lRowIndex = 0
102     bBreakout = False

        Do
108         sSQL = "SELECT * FROM (SELECT TOP 9000 [sID],[sGUID],[sTableName],[swhen],[sStatus],[sequence],[sBy],[sdelete],[updates],[noconflict],ROW_NUMBER() OVER (ORDER BY sWhen DESC) AS 'RowNumber'"
110         sSQL = sSQL & " FROM [" & sSynchHistory & "] WHERE  sTablename = '" & sTableName & "'"
            
112         If Operation = Add Then

114             With SynchHistoryOverview
            
116                 Select Case .NextSynchNextOpeation
                    
                        Case NewOnly
                            sSQL = sSQL & " AND [sdelete] = 'false' AND [swhen] >= '" & .sWhenMaxSearched & "'"
120                     Case OldOnly
122                         sSQL = sSQL & " AND [sdelete] = 'false' AND [swhen] <= '" & IIf(Len(.sWhenMinSearched) > 0, .sWhenMinSearched, .sWhenMaxSearched) & "'"
124                     Case Else
126                         'sSQL = ""

                    End Select
            
                End With

128         ElseIf Operation = Update Then
130             sSQL = sSQL & " AND [sequence] > 1 AND [sdelete] = 'false' AND [sWhen] >= '" & sWhenMin & "'"
132         ElseIf Operation = Delete Then
134             sSQL = sSQL & " AND [sdelete] = 'true' AND [sWhen] >= '" & sWhenMin & "'"
            End If

136         sSQL = sSQL & ") SynchHistoryAlias WHERE RowNumber between " & (1 + lRowIndex) & " and " & (3000 + lRowIndex)
            
            If Operation = Add And SynchHistoryOverview.NextSynchNextOpeation = OldOnly Then
                sSQL = sSQL & " ORDER BY [sWhen] DESC"
            Else
                sSQL = sSQL & " ORDER BY [sWhen]"
            End If

138         TxMsg ">>> GetAllGUIDsFromServerDB: " & sSQL, True, True

140         If Len(sSQL) > 5 Then
142             sDebugInfo = Chr(13) & sSQL
144             Set GetAllGUIDsFromServerDB = OpenServerRSCompressed(sWebsite, "ID", sSQL)
            End If
            
146         If GetAllGUIDsFromServerDB Is Nothing Then
148             bBreakout = True
150         ElseIf Not GetAllGUIDsFromServerDB.State = adStateOpen Then
152             bBreakout = True
            Else

154             If GetAllGUIDsFromServerDB.EOF Then
156                 bBreakout = True
                Else

158                 If lRowIndex = 0 Then
160                     Set oRS = CloneRS(GetAllGUIDsFromServerDB, "RowNumber")
162                     oRS.Open
                    End If
                    
164                 lCountOfRecords = 0

166                 Do Until GetAllGUIDsFromServerDB.EOF
168                     lCountOfRecords = lCountOfRecords + 1
170                     oRS.AddNew
172                     oRS.Fields("sID").Value = GetAllGUIDsFromServerDB.Fields("sID").Value
174                     oRS.Fields("sGUID").Value = GetAllGUIDsFromServerDB.Fields("sGUID").Value
176                     oRS.Fields("sTableName").Value = GetAllGUIDsFromServerDB.Fields("sTableName").Value
178                     oRS.Fields("swhen").Value = GetAllGUIDsFromServerDB.Fields("swhen").Value
180                     oRS.Fields("sequence").Value = GetAllGUIDsFromServerDB.Fields("sequence").Value
182                     oRS.Fields("sBy").Value = GetAllGUIDsFromServerDB.Fields("sBy").Value
184                     oRS.Fields("sdelete").Value = GetAllGUIDsFromServerDB.Fields("sdelete").Value
186                     oRS.Fields("updates").Value = GetAllGUIDsFromServerDB.Fields("updates").Value
188                     oRS.Fields("noconflict").Value = GetAllGUIDsFromServerDB.Fields("noconflict").Value
190                     GetAllGUIDsFromServerDB.MoveNext
                    Loop
                    
                    'If Not oRS.EOF Or Not oRS.BOF Then oRS.MoveFirst
192                 If lCountOfRecords < 3000 Then bBreakout = True
                
                End If

            End If
        
194         lRowIndex = lRowIndex + 3000
            
196     Loop Until bBreakout

198     Set GetAllGUIDsFromServerDB = oRS
        
        '<EhFooter>
        Exit Function

GetAllGUIDsFromServerDB_Err:
        Set GetAllGUIDsFromServerDB = Nothing
        TxMsg "GetAllGUIDsFromServerDB_Err line " & Erl & " ... " & Err.Description, True
        '</EhFooter>
End Function


Private Function GetAllFieldsForGUIDOnServer(sTableName As String, _
                                             sGUIDFieldName As String, _
                                             sGUIDFieldValue As String, _
                                             sWebsite As String) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo GetAllFieldsForGUIDOnServer_Err
        '</EhHeader>
        
        'Operation: The function returns a recordset with one record containing all fields
        '           from the table on the server where the passed GUID matches the one
        '           in that table
        '
        'Assumption: If UID is the GUID field name then it searches for a GUID of type long
        '            otherwise it looks for a GUID of type text

        Dim sSQL As String
        
100     If sGUIDFieldName = "UID" Then
102         sSQL = "SELECT * FROM [" & sTableName & "] WHERE UID = " & sGUIDFieldValue
        Else
104         sSQL = "SELECT * FROM [" & sTableName & "] WHERE [" & sGUIDFieldName & "] = '" & sGUIDFieldValue & "'"
        End If
        
        'TxMsg "GetAllFieldsForGUIDOnServer: " & sSQL
        sDebugInfo = Chr(13) & sSQL

106     Set GetAllFieldsForGUIDOnServer = OpenServerRSCompressed(sWebsite, "ID", sSQL)
        Set GetAllFieldsForGUIDOnServer.ActiveConnection = Nothing

        '<EhFooter>
        Exit Function

GetAllFieldsForGUIDOnServer_Err:
        Set GetAllFieldsForGUIDOnServer = Nothing
        '</EhFooter>
End Function

Private Function DownloadBatchRecordsFromServer(ByVal sTableName As String, _
                                                ByVal sSQL As String, _
                                                ByVal sWebsite As String, _
                                                ByVal sLocalSynchHistPrefix As String, _
                                                ByVal sGUID As String) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo DownloadBatchRecordsFromServer_Err
        '</EhHeader>
        Dim oRS As ADODB.Recordset
        
100     If sTableName = "oincidents_FEA" Then
102         sSQL = "SELECT [UID], [ID], [NAME], [TYPE], [TARGET], [Dead], [Affected], [Violent], [Injured], [Incident_DATE], [TIME00], [Town], [District], [LocDesc], [Source], [Province], [DESCRIPTION], [Scoring], [Incident_DATESERIAL] FROM [oincidents] LEFT JOIN [SynchHistory] ON [oincidents].[ID] = [SynchHistory].[sID] " & sSQL
104     ElseIf sTableName = "oincidents_GEO" Then
106         sSQL = "SELECT [UID], [XMIN], [XMAX], [YMIN], [YMAX], 2 AS [SHAPETYPE] FROM [oincidents] LEFT JOIN [SynchHistory] ON [oincidents].[ID] = [SynchHistory].[sID] " & sSQL
108     ElseIf Right$(sTableName, 4) = "_GEO" Then
110         sSQL = "SELECT * FROM [" & sTableName & "] " & sSQL
        ElseIf sTableName = m_StrRemoteTablePrefix & "Incidents_ChartSettings" Then
111         sSQL = "SELECT [" & sTableName & "].* FROM [" & sTableName & "] LEFT JOIN [" & m_StrRemoteTablePrefix & "SynchHistory] ON [" & sTableName & "].[" & sGUID & "] = [" & m_StrRemoteTablePrefix & "SynchHistory].[sID] WHERE [" & m_StrRemoteTablePrefix & "SynchHistory].[sTableName] = 'Incidents_ChartSettings' " & IIf(Len(sSQL) > 3, Replace$(sSQL, "WHERE", "AND"), "")
        ElseIf Right$(sTableName, 22) = "GeoBookMarksCategories" Or Right$(sTableName, 12) = "GeoBookMarks" Then
112         sSQL = "SELECT [" & sTableName & "].* FROM [" & sTableName & "] LEFT JOIN [" & m_StrRemoteTablePrefix & "SynchHistory] ON [" & sTableName & "].[" & sGUID & "] = [" & m_StrRemoteTablePrefix & "SynchHistory].[sID] WHERE [" & m_StrRemoteTablePrefix & "SynchHistory].[sTableName] = '" & sTableName & "' " & IIf(Len(sSQL) > 3, Replace$(sSQL, "WHERE", "AND"), "")
        Else
113         sSQL = "SELECT [" & sTableName & "].* FROM [" & sTableName & "] LEFT JOIN [" & sLocalSynchHistPrefix & "SynchHistory] ON [" & sTableName & "].[" & sGUID & "] = [" & sLocalSynchHistPrefix & "SynchHistory].[sID] WHERE [" & sLocalSynchHistPrefix & "SynchHistory].[sTableName] = '" & sTableName & "' " & IIf(Len(sSQL) > 3, Replace$(sSQL, "WHERE", "AND"), "")
        End If
               
114     TxMsg "DownloadBatchRecordsFromServer: " & sSQL, True, True
116     sDebugInfo = Chr(13) & sSQL

118     bIsDownloadingBatch = True
120     Set oRS = OpenServerRSCompressed(sWebsite, "ID", sSQL)
        
134     If oRS Is Nothing Then
136         TxMsg "!!!!!!! Batch download failed for table [" & sTableName & "] :-(", True
138         Set DownloadBatchRecordsFromServer = Nothing
        Else
140         Set DownloadBatchRecordsFromServer = oRS
        End If

142     bIsDownloadingBatch = False

        '<EhFooter>
        Exit Function

DownloadBatchRecordsFromServer_Err:
        Set DownloadBatchRecordsFromServer = Nothing
        TxMsg "DownloadBatchRecordsFromServer_Err line " & Erl & " ... " & Err.Description, True
        '</EhFooter>
End Function
    
Private Function GetAllFieldsForIncidentGUIDOnServer_FEA(sID As String, _
                                                         sWebsite As String) As ADODB.Recordset

    Dim sSQL As String
    Dim oRS As New ADODB.Recordset
    
    sSQL = "SELECT [UID], [ID], [NAME], [TYPE], [TARGET], [Dead], [Affected], [Violent], [Injured], [Incident_DATE], [TIME00], [Town], [District], [LocDesc], [Source], [Province], [DESCRIPTION], [Scoring], [Incident_DATESERIAL] FROM [oincidents] WHERE [ID] = '" & sID & "'"
    Set oRS = OpenServerRSCompressed(sWebsite, "ID", sSQL)
    Set GetAllFieldsForIncidentGUIDOnServer_FEA = Nothing
    
    sDebugInfo = Chr(13) & sSQL
    
    If Not oRS.State = adStateClosed Then
   
        If Not oRS.EOF Then
        
            Set GetAllFieldsForIncidentGUIDOnServer_FEA = oRS
   
        End If
        
    End If

End Function

Private Function GetAllFieldsForIncidentGUIDOnServer_GEO(sID As String, _
                                                         sWebsite As String) As ADODB.Recordset

    Dim sSQL As String
    Dim oRS As New ADODB.Recordset
    
    sSQL = "SELECT [UID], [XMIN], [XMAX], [YMIN], [YMAX], 2 AS [SHAPETYPE] FROM [oincidents] WHERE [ID] = '" & sID & "'"
    Set oRS = OpenServerRSCompressed(sWebsite, "ID", sSQL)
    Set GetAllFieldsForIncidentGUIDOnServer_GEO = Nothing
    
    sDebugInfo = Chr(13) & sSQL
    
    If Not oRS.State = adStateClosed Then
   
        If Not oRS.EOF Then
            Set GetAllFieldsForIncidentGUIDOnServer_GEO = oRS
        End If
    End If

End Function

Private Function GetAllFieldsForGUIDOnLocal(sTableName As String, _
                                            sGUIDFieldName As String, _
                                            sGUIDFieldValue As String, _
                                            cn As ADODB.Connection) As ADODB.Recordset

    Dim RSLocalGUIDs As ADODB.Recordset
    Dim sSQL As String
    Set RSLocalGUIDs = New ADODB.Recordset
    Set GetAllFieldsForGUIDOnLocal = Nothing

    With RSLocalGUIDs

        If sGUIDFieldName = "UID" Then
            sSQL = "SELECT * FROM [" & sTableName & "] WHERE UID = " & sGUIDFieldValue
        Else
            sSQL = "SELECT * FROM [" & sTableName & "] WHERE [" & sGUIDFieldName & "] = '" & sGUIDFieldValue & "'"
        End If
            
        sDebugInfo = Chr(13) & sSQL

        RSLocalGUIDs.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
    
        If .State = adStateOpen Then
            Set GetAllFieldsForGUIDOnLocal = RSLocalGUIDs
        Else
            TxMsg "RS did not load", True
        End If
        
    End With
    

End Function

'SELECT SynchHistory.[sID], SynchHistory.[sequence], SynchHistory.[swhen]
'From SynchHistory
'GROUP BY SynchHistory.[sID], SynchHistory.[sequence], SynchHistory.[swhen], SynchHistory.[sTableName]
'HAVING (((SynchHistory.[sTableName])='oincidents') AND ((Max(SynchHistory.swhen))="1"));

Private Function GetRSWithGUIDSFORDOWNLOAD(RSServerALLGUIDS As ADODB.Recordset, _
                                           RSLocalALLGUIDS As ADODB.Recordset, _
                                           Operation As SynchOperation) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo GetRSWithGUIDSFORDOWNLOAD_Err
        '</EhHeader>
        
        Dim lCount As Long
        Dim sWhenCurrent As String

100     If RSServerALLGUIDS Is Nothing Then
102         Set GetRSWithGUIDSFORDOWNLOAD = Nothing
104         TxMsg "    ... RSServerALLGUIDS Is Nothing!", True
106     ElseIf RSLocalALLGUIDS Is Nothing Then
108         Set GetRSWithGUIDSFORDOWNLOAD = RSServerALLGUIDS
110         TxMsg "    ... RSLocalALLGUIDS Is Nothing!", True
        Else
        
112         With RSLocalALLGUIDS

114             If .State = adStateOpen And RSServerALLGUIDS.State = adStateOpen Then

116                 If (Not .EOF Or Not .BOF) And (Not RSServerALLGUIDS.EOF Or Not RSServerALLGUIDS.BOF) Then

118                     RSLocalALLGUIDS.MoveFirst
                        
122                     Do Until .EOF Or .BOF
                            
124                         If Not RSServerALLGUIDS.EOF Or Not RSServerALLGUIDS.BOF Then RSServerALLGUIDS.MoveFirst
126                         RSServerALLGUIDS.Find "[sID] = '" & .Fields(0).Value & "'"
                            
128                         If Not RSServerALLGUIDS.EOF Then

130                             If Operation = Update Then

132                                 If RSServerALLGUIDS.Fields("sequence").Value < RSLocalALLGUIDS.Fields("sequence").Value Then

134                                     RSServerALLGUIDS.Delete adAffectCurrent

136                                 ElseIf RSServerALLGUIDS.Fields("sequence").Value = RSLocalALLGUIDS.Fields("sequence").Value Then

138                                     If RSServerALLGUIDS.Fields("swhen").Value <= RSLocalALLGUIDS.Fields("swhen").Value Then

140                                         RSServerALLGUIDS.Delete adAffectCurrent
                                        Else
                                            'TxMsg "   ------ NOT RSServerALLGUIDS.Fields('swhen').Value < RSLocalALLGUIDS.Fields('swhen').Value"
                                            'TxMsg "   ------ RSServerALLGUIDS.Fields('swhen').Value: " & RSServerALLGUIDS.Fields("swhen").Value
                                            'TxMsg "   ------ RSLocalALLGUIDS.Fields('swhen').Value:  " & RSLocalALLGUIDS.Fields("swhen").Value
                                        End If

                                    End If

                                Else
142                                 RSServerALLGUIDS.Delete adAffectCurrent
                                End If

                            End If
                           
144                         sWhenCurrent = RSLocalALLGUIDS.Fields("swhen").Value
146                         .MoveNext
                        Loop

150                     If Not RSServerALLGUIDS.EOF Or Not RSServerALLGUIDS.BOF Then RSServerALLGUIDS.MoveFirst
                        
152                     If Operation = Add And RSServerALLGUIDS.RecordCount = 0 Then
                        
154                         TxMsg "!!!!!! Some records were not downloaded.  If this is a reoccurring message then it may be possible that all records on the server table do not exist on its corresponding SynchHistory table on the server.  If indeed this message reoccurs then please contact your OASIS administrator", True

156                         If SynchHistoryOverview.NextSynchNextOpeation = NewOnly Then
158                             SynchHistoryOverview.sWhenMaxSearched = sWhenCurrent
                            Else
160                             SynchHistoryOverview.sWhenMinSearched = sWhenCurrent
                            End If
                            
                        Else
162                         TxMsg ">>>> " & RSServerALLGUIDS.RecordCount & " record(s) are left after looking locally....", True, True
                        End If

                    End If
                    
164                 Set GetRSWithGUIDSFORDOWNLOAD = RSServerALLGUIDS
                    
                Else
166                 Set GetRSWithGUIDSFORDOWNLOAD = Nothing
                End If

            End With

        End If

        '<EhFooter>
        Exit Function

GetRSWithGUIDSFORDOWNLOAD_Err:
        TxMsg "GetRSWithGUIDSFORDOWNLOAD line " & Erl & " ... " & Err.Description, True
        Resume Next
        '</EhFooter>
End Function

Private Sub UploadTableChangesToServer(sLocalTableWithNewRecords As String, _
                                       sLocalTableGUID As String, _
                                       bIsGeo As Boolean, _
                                       sWebsite As String, _
                                       cn As ADODB.Connection, _
                                       SynchOp As SynchOperation, _
                                       Optional sLocalSynchHistPrefix As String = "")
        '<EhHeader>
        On Error GoTo UploadTableChangesToServer_Err
        '</EhHeader>
        
        'Operation:     This sub identifies any pending records for sending to the server
        '               Once identified they are uploaded to the server table and
        '               SynchHistory is updated on the local and server database accordingly
        '
        'Assumption:    Field UID must be in the oincidents_FEA and oincidents_GEO table (if used)
        'Note:          If sLocalTableWithNewRecords is geotable it should not have _FEA or _GEO entensionn in veriable
        
        'Dim MsXmlHttp As New MSXML2.ServerXMLHTTP40
        'Dim MsXmlDoc As New MSXML2.DOMDocument
        
        Dim sResponse As String
        Dim sField As String
        Dim RSServerSynchHistory As New ADODB.Recordset
        Dim RSServerTableToBeUpdated As New ADODB.Recordset
        Dim RSServerTableToBeUpdated_FEA As New ADODB.Recordset
        Dim RSServerTableToBeUpdated_GEO As New ADODB.Recordset
        
        Dim RSLocalSynchHistory As New ADODB.Recordset
        Dim RSLocalTableRecordsToBeUploaded As New ADODB.Recordset
        Dim RSLocalTableRecordsToBeUploaded_FEA As New ADODB.Recordset
        Dim RSLocalTableRecordsToBeUploaded_GEO As New ADODB.Recordset

        Dim j As Integer
        Dim i As Integer
        Dim sSQL As String
        Dim bSuccess As Boolean
        Dim sUserGroupPrefix As String
        Dim bSSEValidation As Boolean
        Dim bNewWasAdded As Boolean
        Dim lUID As Long
        Dim bUploadFailed As Boolean
        Dim sGuidToDelete As String
        Dim sServerTableName As String
100     bUploadFailed = False
        'Dim sID As String

102     If sLocalSynchHistPrefix = "" And Not sLocalTableWithNewRecords = "Attachments" Then
104         sUserGroupPrefix = m_StrRemoteTablePrefix
        End If
    
106     TxMsg ">>>>>>>>>>>>"
108     TxMsg ">>> Beginning scan for pending records to commit to server from table [" & sLocalTableWithNewRecords & "] ......................"

110     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then sWebsite = sWebsite & "/"
112     m_sIDs = ""
    
114     If Not sLocalSynchHistPrefix = "" And bIsGeo Then
    
            'Not incidents table & is geo table
    
116         If SynchOp = Delete Then
118             RSLocalSynchHistory.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [updates] = 'false' AND [sdelete] = 'true' AND [sTableName] = '" & sLocalTableWithNewRecords & "_FEA'", cn, adOpenDynamic, adLockBatchOptimistic
120         ElseIf SynchOp = Add Then
122             RSLocalSynchHistory.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [updates] = 'false' AND [sdelete] = 'false' AND [sTableName] = '" & sLocalTableWithNewRecords & "_FEA'", cn, adOpenDynamic, adLockBatchOptimistic
124         ElseIf SynchOp = Update Then
126             RSLocalSynchHistory.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [updates] = 'false' AND [sequence] > 1 AND [sdelete] = 'false' AND [sTableName] = '" & sLocalTableWithNewRecords & "_FEA'", cn, adOpenDynamic, adLockBatchOptimistic
            End If

        Else

128         If SynchOp = Delete Then
130             RSLocalSynchHistory.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [updates] = 'false' AND [sdelete] = 'true' AND [sTableName] = '" & sLocalTableWithNewRecords & "'", cn, adOpenDynamic, adLockBatchOptimistic
132         ElseIf SynchOp = Add Then
134             RSLocalSynchHistory.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [updates] = 'false' AND [sdelete] = 'false' AND [sTableName] = '" & sLocalTableWithNewRecords & "'", cn, adOpenDynamic, adLockBatchOptimistic
136         ElseIf SynchOp = Update Then
138             RSLocalSynchHistory.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE [updates] = 'false' AND [sequence] > 1 AND [sdelete] = 'false' AND [sTableName] = '" & sLocalTableWithNewRecords & "'", cn, adOpenDynamic, adLockBatchOptimistic
            End If
        
        End If
        
140     If Not RSLocalSynchHistory.State = adStateClosed Then

142         If RSLocalSynchHistory.EOF Then
                'TxMsg "dude8) RSLocalSynchHistory.EOF"
                'TxMsg RSLocalSynchHistory.Source
            Else
     
                ' Records are available for upload to the server
144             TxMsg "    Records are available for commit to the server..........."

146             Do Until RSLocalSynchHistory.EOF
                    
148                 bUploadFailed = False

                    ' Point to the SynchHistory table to be updated on the server
150                 If sLocalTableWithNewRecords = "oincidents" Then
152                     sSQL = "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE sID = '" & RSLocalSynchHistory.Fields("sID").Value & "'"
                    Else
154                     sSQL = "SELECT * FROM [" & sLocalSynchHistPrefix & sUserGroupPrefix & "SynchHistory] WHERE sID = '" & RSLocalSynchHistory.Fields("sID").Value & "'"
                    End If

156                 Set RSServerSynchHistory = OpenServerRSCompressed(sWebsite, "ID", sSQL)
                   
158                 If sLocalTableWithNewRecords = "oincidents" Then
                
                        ' Point to the table to be updated on the server
160                     Set RSServerTableToBeUpdated = New ADODB.Recordset
162                     sSQL = "SELECT * FROM oincidents WHERE [" & sLocalTableGUID & "] = '" & RSLocalSynchHistory.Fields("sID").Value & "'"
164                     Set RSServerTableToBeUpdated = OpenServerRSCompressed(sWebsite, "ID", sSQL)

166                 ElseIf Not bIsGeo Then
    
                        ' Point to the table to be updated on the server
168                     Set RSServerTableToBeUpdated = New ADODB.Recordset
170                     sSQL = "SELECT " & GetAllFieldNamesForLocalTable(sLocalTableWithNewRecords, cn) & " FROM [" & sUserGroupPrefix & sLocalTableWithNewRecords & "] WHERE [" & sLocalTableGUID & "] = '" & RSLocalSynchHistory.Fields("sID").Value & "'"
    'TxMsg sLocalTableWithNewRecords & "==== " & sSQL, True
172                     Set RSServerTableToBeUpdated = OpenServerRSCompressed(sWebsite, "ID", sSQL)
                       
                    Else

                        ' Point to the FEA and GEO table to be updated on the server
174                     Set RSServerTableToBeUpdated_FEA = New ADODB.Recordset
176                     Set RSServerTableToBeUpdated_GEO = New ADODB.Recordset
178                     sSQL = "SELECT " & GetAllFieldNamesForLocalTable(sLocalTableWithNewRecords & "_FEA", cn) & " FROM [" & sUserGroupPrefix & sLocalTableWithNewRecords & "_FEA] WHERE [" & sLocalTableGUID & "] = '" & RSLocalSynchHistory.Fields("sID").Value & "'"
180                     Set RSServerTableToBeUpdated_FEA = OpenServerRSCompressed(sWebsite, "ID", sSQL)

182                     If Not RSServerTableToBeUpdated_FEA.EOF Then
184                         sSQL = "SELECT * FROM [" & sUserGroupPrefix & sLocalTableWithNewRecords & "_GEO] WHERE UID = " & RSServerTableToBeUpdated_FEA.Fields("UID").Value
186                         Set RSServerTableToBeUpdated_GEO = OpenServerRSCompressed(sWebsite, "ID", sSQL)
                        Else
188                         sSQL = "SELECT * FROM [" & sUserGroupPrefix & sLocalTableWithNewRecords & "_GEO] WHERE UID = " & -66666666
190                         Set RSServerTableToBeUpdated_GEO = OpenServerRSCompressed(sWebsite, "ID", sSQL)
                        End If
                    
                    End If
                    
192                 TxMsg "RSServerTableToBeUpdated SQL: " & sSQL

194                 If Not sLocalTableWithNewRecords = "oincidents" And bIsGeo Then

196                     If SynchOp = Delete Then

                            ''' DELETE GEO-RECORD ON SERVER
                        
198                         i = i + 1
200                         TxMsg "    --- (" & i & ") deleting id: " & RSLocalSynchHistory.Fields("sID").Value

202                         If Not RSServerTableToBeUpdated_FEA.EOF Then RSServerTableToBeUpdated_FEA.Delete adAffectCurrent
204                         If Not RSServerTableToBeUpdated_GEO.EOF Then RSServerTableToBeUpdated_GEO.Delete adAffectCurrent

206                         j = 0

208                         If RSServerSynchHistory.EOF Then RSServerSynchHistory.AddNew
210                         RSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
212                         RSLocalSynchHistory.Fields("updates").Value = "true"
                            RSLocalSynchHistory.Fields("noconflict").Value = "false"

214                         Do Until j = RSLocalSynchHistory.Fields.Count
216                             RSServerSynchHistory.Fields(j).Value = RSLocalSynchHistory.Fields(j).Value
218                             j = j + 1
                            Loop
                        
220                     ElseIf SynchOp = Update Then

                            ''' UPDATE GEO-RECORD ON SERVER
                            ''' This code block conforms with the SSE spec protocol (http://msdn.microsoft.com/en-us/xml/bb510102.aspx)
                            
222                         bSSEValidation = True

224                         If Not RSServerSynchHistory.EOF Then
                               
226                             If RSLocalSynchHistory.Fields("sequence").Value < RSServerSynchHistory.Fields("sequence").Value Then
                                        
228                                 bSSEValidation = False

230                             ElseIf RSLocalSynchHistory.Fields("sequence").Value = RSServerSynchHistory.Fields("sequence").Value Then

232                                 If CalcServerDateTimeFromRFC3339(RSLocalSynchHistory.Fields("swhen").Value) <= RSServerSynchHistory.Fields("swhen").Value Then
234                                     bSSEValidation = False
                                    End If
                                    
                                End If
                            
                            Else
236                             RSServerSynchHistory.AddNew
                            End If
                                    
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        
238                         If Not bSSEValidation Then
                        
240                             TxMsg "    !!! Record: " & RSLocalSynchHistory.Fields("sID").Value & " flagged on client database for upload is older than the server version - updating local SynchHistory accordingly", True
                                'RSLocalSynchHistory.Delete adAffectCurrent
242                             RSLocalSynchHistory.Fields("updates") = "true"
244                             RSLocalSynchHistory.UpdateBatch
                        
                            Else
    
                                ' Get FEA table on local database
246                             Set RSLocalTableRecordsToBeUploaded_FEA = GetAllFieldsForGUIDOnLocal(sLocalTableWithNewRecords & "_FEA", sLocalTableGUID, RSLocalSynchHistory.Fields("sID").Value, cn)

248                             If RSLocalTableRecordsToBeUploaded_FEA.State = adStateClosed Then
                                    'TxMsg "dude3) RSLocalTableRecordsToBeUploaded_FEA.State = adStateClosed"
                                Else
                            
250                                 If RSLocalTableRecordsToBeUploaded_FEA.EOF Then
                                
252                                     TxMsg "    --- Record: " & RSLocalSynchHistory.Fields("sID").Value & " does not exist on the local database - removing entry from SynchHistory", True
254                                     RSLocalSynchHistory.Delete adAffectCurrent
256                                     RSLocalSynchHistory.UpdateBatch
                                    
                                    Else

                                        ' Get GEO table on local database
258                                     Set RSLocalTableRecordsToBeUploaded_GEO = GetAllFieldsForGUIDOnLocal(sLocalTableWithNewRecords & "_GEO", "UID", RSLocalTableRecordsToBeUploaded_FEA.Fields("UID").Value, cn)

260                                     If RSLocalTableRecordsToBeUploaded_GEO.State = adStateClosed Then
                                            'TxMsg "dude2) RSLocalTableRecordsToBeUploaded_GEO.State = adStateClosed"
                                        Else
                                    
262                                         If RSLocalTableRecordsToBeUploaded_GEO.EOF Then
                                        
264                                             TxMsg "    --- Record: " & RSLocalSynchHistory.Fields("sID").Value & " altough FEA record exists the GEO record does not exist on the local database - removing entry from FEA and SynchHistory", True
266                                             RSLocalSynchHistory.Delete adAffectCurrent
268                                             RSLocalSynchHistory.UpdateBatch
                                            
                                                'FEA record without GEO! ......................................
270                                             RSLocalTableRecordsToBeUploaded_FEA.Delete adAffectCurrent
272                                             RSLocalTableRecordsToBeUploaded_FEA.UpdateBatch adAffectCurrent
                                        
                                            Else

                                                ' All checks came up good and now you can upload pending records
274                                             i = i + 1
276                                             TxMsg "    --- (" & i & ") uploading id: " & RSLocalSynchHistory.Fields("sID").Value
                                        
278                                             If RSServerTableToBeUpdated_FEA.EOF Then
280                                                 RSServerTableToBeUpdated_FEA.AddNew
282                                                 RSServerTableToBeUpdated_GEO.AddNew
284                                                 bNewWasAdded = True
286                                                 lUID = GetMaxUIDFromServerGEOTable(sUserGroupPrefix & sLocalTableWithNewRecords & "_GEO", sWebsite) + 1
                                                Else
288                                                 bNewWasAdded = False
                                                End If

                                                ' Update the FEA record to the server
290                                             j = 0

292                                             Do Until j = RSLocalTableRecordsToBeUploaded_FEA.Fields.Count

294                                                 sField = RSLocalTableRecordsToBeUploaded_FEA.Fields(j).Name
                                             
296                                                 If DoesFieldExistInRS(RSServerTableToBeUpdated_FEA, sField) Then

298                                                     If bNewWasAdded And sField = "UID" Then
300                                                         RSServerTableToBeUpdated_FEA.Fields(sField).Value = lUID
302                                                     ElseIf Not sField = "UID" Then

304                                                         If Len(RSLocalTableRecordsToBeUploaded_FEA.Fields(sField).Value) > 0 Then RSServerTableToBeUpdated_FEA.Fields(sField).Value = RSLocalTableRecordsToBeUploaded_FEA.Fields(sField).Value
                                                        Else
306                                                         RSServerTableToBeUpdated_FEA.Fields("UID").Value = RSServerTableToBeUpdated_FEA.Fields("UID").OriginalValue
                                                        End If

                                                    Else
308                                                     TxMsg "!!!! Skipping update of field [" & sField & "] on server since it only exists on the client", True
                                                    End If

310                                                 j = j + 1
                                                Loop
                            
                                                ' Update the GEO record to the server
312                                             j = 0

314                                             Do Until j = RSLocalTableRecordsToBeUploaded_GEO.Fields.Count

316                                                 If bNewWasAdded And RSServerTableToBeUpdated_GEO.Fields(j).Name = "UID" Then
318                                                     RSServerTableToBeUpdated_GEO.Fields(j).Value = lUID

320                                                 ElseIf Not RSServerTableToBeUpdated_GEO.Fields(j).Name = "UID" Then

322                                                     If Not (sLocalTableWithNewRecords = "oincidents" And RSServerTableToBeUpdated_GEO.Fields(j).Name = "GEOMETRY") Then
324                                                         RSServerTableToBeUpdated_GEO.Fields(j).Value = RSLocalTableRecordsToBeUploaded_GEO.Fields(j).Value
                                                        End If
                                                    End If
                                                
326                                                 j = j + 1
                                                Loop
                                    
                                                ' Add the SynchHistory record to the server
328                                             j = 0

330                                             If RSServerSynchHistory.EOF Or RSServerSynchHistory.BOF Then RSServerSynchHistory.AddNew
332                                             RSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
334                                             RSLocalSynchHistory.Fields("updates").Value = "true"
                                                RSLocalSynchHistory.Fields("noconflict").Value = "false"

336                                             Do Until j = RSLocalSynchHistory.Fields.Count
338                                                 RSServerSynchHistory.Fields(j).Value = RSLocalSynchHistory.Fields(j).Value
340                                                 j = j + 1
                                                Loop

                                            End If

                                        End If
                                    
                                    End If
                                End If
                            
                            End If

342                     ElseIf SynchOp = Add And RSServerTableToBeUpdated_FEA.EOF Then

                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''' ADD GEO-RECORD TO SERVER
                            '''

                            ' Get FEA table on local database
344                         Set RSLocalTableRecordsToBeUploaded_FEA = GetAllFieldsForGUIDOnLocal(sLocalTableWithNewRecords & "_FEA", sLocalTableGUID, RSLocalSynchHistory.Fields("sID").Value, cn)

346                         If Not RSLocalTableRecordsToBeUploaded_FEA.State = adStateClosed Then

348                             If RSLocalTableRecordsToBeUploaded_FEA.EOF Then

350                                 TxMsg "    --- Record " & RSLocalSynchHistory.Fields("sID").Value & " does not exist on the local database - updating local SynchHistory accordingly", True
352                                 RSLocalSynchHistory.Delete adAffectCurrent
354                                 RSLocalSynchHistory.UpdateBatch
                                
                                Else
                                
                                    ' Get GEO table on local database
356                                 Set RSLocalTableRecordsToBeUploaded_GEO = GetAllFieldsForGUIDOnLocal(sLocalTableWithNewRecords & "_GEO", "UID", RSLocalTableRecordsToBeUploaded_FEA.Fields("UID").Value, cn)

358                                 If RSLocalTableRecordsToBeUploaded_GEO.State = adStateClosed Then
                                
                                        'T'xMsg "dude1) RSLocalTableRecordsToBeUploaded_GEO.State = adStateClosed"
                                
                                    Else

360                                     If RSLocalTableRecordsToBeUploaded_GEO.EOF Then
                                    
362                                         TxMsg "    --- Record: " & RSLocalSynchHistory.Fields("sID").Value & " altough FEA record exists the GEO record does not exist on the local database - removing entry from FEA and SynchHistory", True
364                                         RSLocalSynchHistory.Delete adAffectCurrent
366                                         RSLocalSynchHistory.UpdateBatch
                                        
                                            'FEA record without GEO! ......................................
368                                         RSLocalTableRecordsToBeUploaded_FEA.Delete adAffectCurrent
370                                         RSLocalTableRecordsToBeUploaded_FEA.UpdateBatch adAffectCurrent
                                    
                                        Else

                                            ' All checks came up good and now you can upload pending records
372                                         i = i + 1
374                                         TxMsg "--- (" & i & ") uploading id: " & RSLocalSynchHistory.Fields("sID").Value
376                                         lUID = GetMaxUIDFromServerGEOTable(sUserGroupPrefix & sLocalTableWithNewRecords & "_GEO", sWebsite) + 1
                                        
                                            ' Add the FEA record to the server
378                                         RSServerTableToBeUpdated_FEA.AddNew
380                                         j = 0

382                                         Do Until j = RSLocalTableRecordsToBeUploaded_FEA.Fields.Count
                                            
384                                             sField = RSLocalTableRecordsToBeUploaded_FEA.Fields(j).Name
                                             
386                                             If DoesFieldExistInRS(RSServerTableToBeUpdated_FEA, sField) Then
                                            
388                                                 If Not sField = "ORIG_ID" Then
390                                                     If sField = "UID" Then
392                                                         RSServerTableToBeUpdated_FEA.Fields(sField).Value = lUID
                                                        Else
394                                                         If Len(RSLocalTableRecordsToBeUploaded_FEA.Fields(sField).Value) > 0 Then RSServerTableToBeUpdated_FEA.Fields(sField).Value = RSLocalTableRecordsToBeUploaded_FEA.Fields(sField).Value
                                                        End If
                                                    End If
                                                
                                                Else
396                                                 TxMsg "!!!! Skipping update of field [" & sField & "] on server since it only exists on the client", True
                                                End If

                                                ''''
                                            
                                                'RSServerTableToBeUpdated_FEA.Fields(j).Value = RSLocalTableRecordsToBeUploaded_FEA.Fields(j).Value
398                                             j = j + 1
                                            Loop
                            
                                            ' Add the GEO record to the server
400                                         RSServerTableToBeUpdated_GEO.AddNew
402                                         j = 0

404                                         Do Until j = RSLocalTableRecordsToBeUploaded_GEO.Fields.Count
                                            
                                                ''''
                                        
406                                             If RSServerTableToBeUpdated_GEO.Fields(j).Name = "UID" Then
408                                                 RSServerTableToBeUpdated_GEO.Fields(j).Value = lUID
                                                Else
410                                                 RSServerTableToBeUpdated_GEO.Fields(j).Value = RSLocalTableRecordsToBeUploaded_GEO.Fields(j).Value
                                                End If
                                                
                                                ''''
                                                
                                                'Else
                                                '   TxMsg "!!!! Skipping update of field [" & sField & "] on server since it only exists on the client", True
                                                'End If
                                            
                                                'RSServerTableToBeUpdated_GEO.Fields(j).Value = RSLocalTableRecordsToBeUploaded_GEO.Fields(j).Value
412                                             j = j + 1
                                            Loop
                                    
                                            ' Add the SynchHistory record to the server
414                                         j = 0
416                                         RSServerSynchHistory.AddNew
418                                         RSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
420                                         RSLocalSynchHistory.Fields("updates").Value = "true"
                                            RSLocalSynchHistory.Fields("noconflict").Value = "false"
                                            
422                                         Do Until j = RSLocalSynchHistory.Fields.Count
424                                             RSServerSynchHistory.Fields(j).Value = RSLocalSynchHistory.Fields(j).Value

426                                             j = j + 1
                                            Loop
                                            
428                                         RSServerSynchHistory.Fields("updates").Value = "true"
                                            
                                        End If

                                    End If
                                    
                                End If
                            End If

                        Else
                                
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''' GEO-RECORD ALREADY EXISTS ON THE SERVER - CANNOT ADD
                            '''
430                         i = i + 1
432                         TxMsg "    --- add aborted: Record " & i & " with id: " & RSLocalSynchHistory.Fields("sID").Value & " is already on the server - adjusting SynchHistory accordingly" ', True
434                         RSLocalSynchHistory.Fields("updates").Value = "true"
436                         RSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
                            RSLocalSynchHistory.Fields("noconflict").Value = "false"
                        
                        End If

                    Else
                    
438                     If SynchOp = Add And RSServerTableToBeUpdated.EOF Then
                            
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''' ADD RECORD TO SERVER
                            '''
                        
                            'TxMsg "add RECORD ON SERVER"
                                
                            ' Get the table on local database
440                         If Not sLocalTableWithNewRecords = "oincidents" Then
442                             Set RSLocalTableRecordsToBeUploaded = GetAllFieldsForGUIDOnLocal(sLocalTableWithNewRecords, sLocalTableGUID, RSLocalSynchHistory.Fields("sID").Value, cn)
                            Else
444                             Set RSLocalTableRecordsToBeUploaded = GetAllFieldsForGUIDOnLocal("oincidents_FEA", sLocalTableGUID, RSLocalSynchHistory.Fields("sID").Value, cn)
                            End If

446                         If RSLocalTableRecordsToBeUploaded.State = adStateClosed Then

448                             TxMsg "!!! RSLocalTableRecordsToBeUploaded.State = adStateClosed ", True
                            Else

450                             If RSLocalTableRecordsToBeUploaded.EOF Then
                            
452                                 TxMsg "    --- Record " & RSLocalSynchHistory.Fields("sID").Value & " does not exist on the local database - updating local SynchHistory accordingly", True
454                                 RSLocalSynchHistory.Delete adAffectCurrent
456                                 RSLocalSynchHistory.UpdateBatch
                            
                                Else

458                                 bSuccess = False
                                
460                                 If Not sLocalTableWithNewRecords = "oincidents" Then

                                        ' If the table is not incidents then all checks came up good
                                        ' and now you can upload pending records
462                                     RSServerTableToBeUpdated.AddNew
464                                     j = 0

466                                     Do Until j = RSLocalTableRecordsToBeUploaded.Fields.Count

468                                         sField = RSLocalTableRecordsToBeUploaded.Fields(j).Name
                                             
470                                         If DoesFieldExistInRS(RSServerTableToBeUpdated, sField) Then

472                                             If Not sField = "UID" Then
474                                                 If Len(RSLocalTableRecordsToBeUploaded.Fields(sField).Value) > 0 Then RSServerTableToBeUpdated.Fields(sField).Value = RSLocalTableRecordsToBeUploaded.Fields(sField).Value
                                                Else
476                                                 TxMsg "Skipping UID update on server since it should be an autonum field....."
                                                End If
                                            
                                            Else
478                                             TxMsg "!!!! Skipping update of field [" & sField & "] on server since it only exists on the client", True
                                            End If

480                                         j = j + 1
                                        Loop

482                                     bSuccess = True
                            
                                    Else
                                
                                        ' The table is incidents - we need to get the local GEO table too
484                                     Set RSLocalTableRecordsToBeUploaded_FEA = RSLocalTableRecordsToBeUploaded
486                                     Set RSLocalTableRecordsToBeUploaded_GEO = GetAllFieldsForGUIDOnLocal(sLocalTableWithNewRecords & "_GEO", "UID", RSLocalTableRecordsToBeUploaded.Fields("UID").Value, cn)

488                                     If Not RSLocalTableRecordsToBeUploaded_GEO.State = adStateClosed Then
490                                         If Not RSLocalTableRecordsToBeUploaded_GEO.EOF Then
                                        
                                                ' All checks came up good and now you can upload pending records
492                                             With RSServerTableToBeUpdated
                                            
494                                                 lUID = GetMaxUIDFromServerGEOTable("oincidents", sWebsite) + 1
                                
496                                                 .AddNew

498                                                 'If InStr(LCase(sWebsite), "afghanistan") = 0 Then

500                                                     '.Fields("UID").Value = lUID 'RSLocalTableRecordsToBeUploaded_FEA.Fields("UID").Value
                                                        
                                                    'Else
                                                    
502                                                     TxMsg "Bypassing update of server UID this should be done automatically on the server using 'identity' spec on the column."
                                                        
                                                    'End If
                                                    
504                                                 .Fields("ID").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("id").Value
506                                                 .Fields("NAME").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Name").Value
508                                                 .Fields("TYPE").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Type").Value
510                                                 .Fields("TARGET").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Target").Value
512                                                 .Fields("Dead").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Dead").Value
514                                                 .Fields("Affected").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Affected").Value
516                                                 .Fields("Violent").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Violent").Value
518                                                 .Fields("Injured").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Injured").Value
520                                                 .Fields("Incident_DATE").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Incident_DATE").Value
522                                                 .Fields("TIME00").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("TIME00").Value
524                                                 .Fields("LocDesc").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("LocDesc").Value
526                                                 .Fields("Source").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Source").Value
528                                                 .Fields("Town").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Town").Value
530                                                 .Fields("District").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("District").Value
532                                                 .Fields("Province").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Province").Value
534                                                 .Fields("DESCRIPTION").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Description").Value
536                                                 .Fields("Scoring").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Scoring").Value
538                                                 .Fields("Incident_DATESERIAL").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Incident_DATESERIAL").Value
    
                                                    '.Fields("UID").Value = lUID 'RSLocalTableRecordsToBeUploaded_GEO.Fields("UID").Value
540                                                 .Fields("XMIN").Value = RSLocalTableRecordsToBeUploaded_GEO.Fields("XMin").Value
542                                                 .Fields("XMAX").Value = RSLocalTableRecordsToBeUploaded_GEO.Fields("XMax").Value
544                                                 .Fields("YMIN").Value = RSLocalTableRecordsToBeUploaded_GEO.Fields("YMin").Value
546                                                 .Fields("YMAX").Value = RSLocalTableRecordsToBeUploaded_GEO.Fields("YMax").Value
                 
                                                End With
                                            
548                                             bSuccess = True
                        
                                            End If
                                        End If
                                
                                    End If
                                
550                                 If bSuccess Then
                                
552                                     i = i + 1
554                                     TxMsg "--- (" & i & ") uploading id: " & RSLocalSynchHistory.Fields("sID").Value
                                            
556                                     j = 0
558                                     RSServerSynchHistory.AddNew
560                                     RSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
562                                     RSLocalSynchHistory.Fields("updates").Value = "true"
                                        RSLocalSynchHistory.Fields("noconflict").Value = "false"
    
564                                     Do Until j = RSLocalSynchHistory.Fields.Count
566                                         RSServerSynchHistory.Fields(j).Value = RSLocalSynchHistory.Fields(j).Value
568                                         j = j + 1
                                        Loop
                                        
570                                     If RSServerSynchHistory.Fields("sTableName").Value = "GeoBookMarks" Then
572                                         RSServerSynchHistory.Fields("sTableName").Value = m_StrRemoteTablePrefix & RSServerSynchHistory.Fields("sTableName").Value
                                        End If
                                
                                    End If
                        
                                End If

                            End If
                        
574                     ElseIf SynchOp = Update Then

                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''' UPDATE RECORD TO SERVER
                            '''
                        
                            'TxMsg "update RECORD ON SERVER"
                                
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            '' This code block conforms with the SSE spec protocol
                            '' http://msdn.microsoft.com/en-us/xml/bb510102.aspx
                            ''
                               
576                         bSSEValidation = True

578                         If Not RSServerSynchHistory.EOF Then
                               
580                             If RSLocalSynchHistory.Fields("sequence").Value < RSServerSynchHistory.Fields("sequence").Value Then
                                        
582                                 bSSEValidation = False
                                
584                             ElseIf RSLocalSynchHistory.Fields("sequence").Value = RSServerSynchHistory.Fields("sequence").Value Then
                            
586                                 If CalcServerDateTimeFromRFC3339(RSLocalSynchHistory.Fields("swhen").Value) <= RSServerSynchHistory.Fields("swhen").Value Then
588                                     bSSEValidation = False
                                    End If
                                    
                                End If
                            
                            Else
590                             RSServerSynchHistory.AddNew
                            End If
                                    
                            ''
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
592                         If Not bSSEValidation Then
                        
594                             TxMsg "    !!! Record: " & RSLocalSynchHistory.Fields("sID").Value & " flagged on client database for upload is older than the server version - updating local SynchHistory accordingly", True
                                'RSLocalSynchHistory.Delete adAffectCurrent
596                             RSLocalSynchHistory.Fields("updates") = "true"
598                             RSLocalSynchHistory.UpdateBatch
                        
                            Else

600                             If RSServerTableToBeUpdated.EOF Then
602                                 RSServerTableToBeUpdated.AddNew
                                End If
                        
                                ' Get the table on local database
604                             If Not sLocalTableWithNewRecords = "oincidents" Then
606                                 Set RSLocalTableRecordsToBeUploaded = GetAllFieldsForGUIDOnLocal(sLocalTableWithNewRecords, sLocalTableGUID, RSLocalSynchHistory.Fields("sID").Value, cn)
                                Else
608                                 Set RSLocalTableRecordsToBeUploaded = GetAllFieldsForGUIDOnLocal("oincidents_FEA", sLocalTableGUID, RSLocalSynchHistory.Fields("sID").Value, cn)
                                End If

610                             If Not RSLocalTableRecordsToBeUploaded.State = adStateClosed Then

612                                 If RSLocalTableRecordsToBeUploaded.EOF Then

                                        'does not exist
614                                     TxMsg "    --- Record: " & RSLocalSynchHistory.Fields("sID").Value & " does not exist on the local database - removing entry from SynchHistory", True
616                                     RSLocalSynchHistory.Delete adAffectCurrent
618                                     RSLocalSynchHistory.UpdateBatch
                                
                                    Else

620                                     bSuccess = False
                                
622                                     If Not sLocalTableWithNewRecords = "oincidents" Then

                                            ' If the table is not incidents then all checks came up good
                                            ' and now you can upload pending records
624                                         j = 0

626                                         Do Until j = RSLocalTableRecordsToBeUploaded.Fields.Count

628                                             sField = RSLocalTableRecordsToBeUploaded.Fields(j).Name
                                                
630                                             If DoesFieldExistInRS(RSServerTableToBeUpdated, sField) Then

632                                                 If Not sField = "UID" Then 'And Not InStr(LCase(sWebsite), "nomad") > 0 Then
634                                                     If Len(RSLocalTableRecordsToBeUploaded.Fields(sField).Value) > 0 Then RSServerTableToBeUpdated.Fields(sField).Value = RSLocalTableRecordsToBeUploaded.Fields(sField).Value
                                                    Else
636                                                     TxMsg "Skipping UID update on server since it should be an autonum field....."
                                                    End If
                                                
                                                Else
638                                                 TxMsg "!!!! Skipping update of field [" & sField & "] on server since it only exists on the client", True
                                                End If

640                                             j = j + 1
                                            Loop

642                                         bSuccess = True
                            
                                        Else
                                
                                            ' The table is incidents - we need to get the local GEO table too
644                                         Set RSLocalTableRecordsToBeUploaded_FEA = RSLocalTableRecordsToBeUploaded
646                                         Set RSLocalTableRecordsToBeUploaded_GEO = GetAllFieldsForGUIDOnLocal(sLocalTableWithNewRecords & "_GEO", "UID", RSLocalTableRecordsToBeUploaded.Fields("UID").Value, cn)

648                                         If Not RSLocalTableRecordsToBeUploaded_GEO.State = adStateClosed Then
650                                             If Not RSLocalTableRecordsToBeUploaded_GEO.EOF Then
                                        
                                                    ' All checks came up good and now you can upload pending records
652                                                 With RSServerTableToBeUpdated
                                
                                                        '.Fields("UID").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("UID").Value
654                                                     .Fields("ID").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("id").Value
656                                                     .Fields("NAME").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Name").Value
658                                                     .Fields("TYPE").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Type").Value
660                                                     .Fields("TARGET").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Target").Value
662                                                     .Fields("Dead").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Dead").Value
664                                                     .Fields("Affected").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Affected").Value
666                                                     .Fields("Violent").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Violent").Value
668                                                     .Fields("Injured").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Injured").Value
670                                                     .Fields("Incident_DATE").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Incident_DATE").Value
672                                                     .Fields("TIME00").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("TIME00").Value
674                                                     .Fields("LocDesc").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("LocDesc").Value
676                                                     .Fields("Source").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Source").Value
678                                                     .Fields("Town").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Town").Value
680                                                     .Fields("District").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("District").Value
682                                                     .Fields("Province").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Province").Value
684                                                     .Fields("DESCRIPTION").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Description").Value
686                                                     .Fields("Scoring").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Scoring").Value
688                                                     .Fields("Incident_DATESERIAL").Value = RSLocalTableRecordsToBeUploaded_FEA.Fields("Incident_DATESERIAL").Value
    
                                                        '.Fields("UID").Value = RSLocalTableRecordsToBeUploaded_GEO.Fields("UID").Value
690                                                     .Fields("XMIN").Value = RSLocalTableRecordsToBeUploaded_GEO.Fields("XMin").Value
692                                                     .Fields("XMAX").Value = RSLocalTableRecordsToBeUploaded_GEO.Fields("XMax").Value
694                                                     .Fields("YMIN").Value = RSLocalTableRecordsToBeUploaded_GEO.Fields("YMin").Value
696                                                     .Fields("YMAX").Value = RSLocalTableRecordsToBeUploaded_GEO.Fields("YMax").Value
                 
                                                    End With
                                            
698                                                 bSuccess = True
                        
                                                End If
                                            End If
                                
                                        End If
                                
700                                     If bSuccess Then
                                
702                                         i = i + 1
704                                         TxMsg "    --- (" & i & ") uploading id: " & RSLocalSynchHistory.Fields("sID").Value
                                            
706                                         j = 0

708                                         If RSServerSynchHistory.EOF Or RSServerSynchHistory.BOF Then RSServerSynchHistory.AddNew
710                                         RSLocalSynchHistory.Fields("updates").Value = "true"
712                                         RSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
                                            RSLocalSynchHistory.Fields("noconflict").Value = "false"
                                            
714                                         Do Until j = RSLocalSynchHistory.Fields.Count
716                                             RSServerSynchHistory.Fields(j).Value = RSLocalSynchHistory.Fields(j).Value
718                                             j = j + 1
                                            Loop
                                
                                        End If
                        
                                    End If

                                End If
                            
                            End If

720                     ElseIf SynchOp = Delete Then

                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''' DELETE RECORD ON SERVER
                            '''
                        
                            'TxMsg "DELETE RECORD ON SERVER"
                                
722                         i = i + 1
724                         TxMsg "    --- (" & i & ") deleting id: " & RSLocalSynchHistory.Fields("sID").Value, True
726                         sGuidToDelete = RSLocalSynchHistory.Fields("sID").Value

728                         If Not RSServerTableToBeUpdated.EOF Then
730                             RSServerTableToBeUpdated.Delete adAffectCurrent
                            End If
                        
                            ' Add the SynchHistory record to the server
732                         j = 0

734                         If RSServerSynchHistory.EOF Or RSServerSynchHistory.BOF Then RSServerSynchHistory.AddNew
736                         RSLocalSynchHistory.Fields("updates").Value = "true"
738                         RSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
                            RSLocalSynchHistory.Fields("noconflict").Value = "false"

740                         Do Until j = RSLocalSynchHistory.Fields.Count
742                             RSServerSynchHistory.Fields(j).Value = RSLocalSynchHistory.Fields(j).Value
744                             j = j + 1
                            Loop
                        
                        Else
                            
                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            ''' RECORD ALREADY EXISTS ON SERVER
                            '''
                        
746                         i = i + 1

748                         If SynchOp = Delete Then
750                             TxMsg "    --- delete ABORTED: Record " & i & " with id: " & RSLocalSynchHistory.Fields("sID").Value & "is not on the server - adjusting SynchHistory accordingly" ', True
752                         ElseIf SynchOp = Add Then
754                             TxMsg "    --- add ABORTED: Record " & i & " with id: " & RSLocalSynchHistory.Fields("sID").Value & " is already on the server - adjusting SynchHistory accordingly" ', True
756                         ElseIf SynchOp = Update Then
758                             TxMsg "    --- update ABORTED: Record " & i & " with id: " & RSLocalSynchHistory.Fields("sID").Value & " is not on the server - adjusting SynchHistory accordingly" ', True
                            Else
                                'TxMsg "undude"
                            End If

760                         RSLocalSynchHistory.Fields("updates").Value = "true"
762                         RSLocalSynchHistory.Fields("sWhen").Value = CalcServerDateTime(Now())
                            RSLocalSynchHistory.Fields("noconflict").Value = "false"
764                         RSLocalSynchHistory.UpdateBatch adAffectCurrent

                        End If
                
                    End If
                
766                 If Not bUploadFailed Then
                        '''''

                        'Commit to server
768                     If sLocalTableWithNewRecords = "oincidents" Or Not bIsGeo Then

770                         RSServerTableToBeUpdated.Filter = adFilterPendingRecords
                            'MsXmlHttp.Open "POST", sWebsite & "oasisuploader.asp", False
                            'RSServerTableToBeUpdated.Save MsXmlDoc, 1

                            'If m_bProxyEnabled Then MsXmlHttp.setProxy 2, m_sProxyIP & ":" & m_sProxyPort
                            'MsXmlHttp.send MsXmlDoc
772                         sResponse = SaveServerRSCompressed(sWebsite, RSServerTableToBeUpdated)

774                         If Not sResponse = "Data Updated" Then
776                             bUploadFailed = True
778                             TxMsg "    --- Error updating RSServerTableToBeUpdated: " & sResponse, True
                                '746                         ElseIf InStr(LCase(sWebsite), "nomad") > 0 Then
                                '
                                '748                             TxMsg "NOMAD commit to i-mogene server....."
                                '
                                '                                If SynchOp = Delete Then
                                '                                    bUploadFailed = Not CommitXML2Nomad("http://nomad.i-mogene.org/NomadWs/" & OASIS2NomadTable(sUserGroupPrefix & sLocalTableWithNewRecords) & "/" & sGuidToDelete, "", "oasis", "oasisoasis", m_bProxyEnabled, m_sProxyIP & ":" & m_sProxyPort, True)
                                '                                Else
                                '                                    bUploadFailed = Not ConvertOASIS2Nomad(RSServerTableToBeUpdated, OASIS2NomadTable(sUserGroupPrefix & sLocalTableWithNewRecords), OASIS2NomadTable(sUserGroupPrefix & sLocalTableWithNewRecords) & "_", sLocalTableWithNewRecords, Replace(LCase(m_StrLocalConnectionString), "oasisclient.mdb", "\dynamicdata\NOMADOASIS.mdb"), "http://nomad.i-mogene.org/NomadWs", OASIS2NomadTable(sUserGroupPrefix & sLocalTableWithNewRecords), "oasis", "oasisoasis", m_bProxyEnabled, m_sProxyIP & ":" & m_sProxyPort)
                                '                                End If
                                '
                                '                                If bUploadFailed Then TxMsg "!!!   Commit to i-mogen server failed!"
                        
                            End If
                        
780                         If RSServerTableToBeUpdated.State = adStateOpen Then RSServerTableToBeUpdated.Close

                        Else
782                         RSServerTableToBeUpdated_FEA.Filter = adFilterPendingRecords
                            'MsXmlHttp.Open "POST", sWebsite & "oasisuploader.asp", 0
                            'RSServerTableToBeUpdated_FEA.Save MsXmlDoc, 1

                            'If m_bProxyEnabled Then MsXmlHttp.setProxy 2, m_sProxyIP & ":" & m_sProxyPort
                            'MsXmlHttp.send MsXmlDoc
784                         sResponse = SaveServerRSCompressed(sWebsite, RSServerTableToBeUpdated_FEA)
786                         RSServerTableToBeUpdated_FEA.Close

788                         If Not sResponse = "Data Updated" Then
790                             bUploadFailed = True
792                             TxMsg "    --- Error updating RSServerTableToBeUpdated_FEA: " & sResponse, True
                            Else
794                             RSServerTableToBeUpdated_GEO.Filter = adFilterPendingRecords
                                'MsXmlHttp.Open "POST", sWebsite & "oasisuploader.asp", 0
                                'RSServerTableToBeUpdated_GEO.Save MsXmlDoc, 1

                                'If m_bProxyEnabled Then MsXmlHttp.setProxy 2, m_sProxyIP & ":" & m_sProxyPort
                                'MsXmlHttp.send MsXmlDoc
796                             sResponse = SaveServerRSCompressed(sWebsite, RSServerTableToBeUpdated_GEO)
798                             RSServerTableToBeUpdated_GEO.Close

800                             If Not sResponse = "Data Updated" Then
802                                 bUploadFailed = True
804                                 TxMsg "    --- Error updating RSServerTableToBeUpdated_GEO: " & sResponse, True
                                End If
                            End If
                    
                        End If
                    End If
                        
806                 If Not bUploadFailed Then

808                     RSLocalSynchHistory.UpdateBatch adAffectCurrent
                        
                        '794                     If Not InStr(UCase(sWebsite), "/NOMAD") > 0 Then
                        
810                     RSServerSynchHistory.Filter = adFilterPendingRecords
812                     sResponse = SaveServerRSCompressed(sWebsite, RSServerSynchHistory)
                        '                        Else
                        '
                        '806                         TxMsg "Skipping SynchHistory update for server on NOMAD - server db trigger takes care of this"
                        '                        End If

814                     RSServerSynchHistory.Close

816                     If Not sResponse = "Data Updated" Then
818                         TxMsg "    --- Error updating RSServerSynchHistory: " & sResponse, True
                        End If
                
                    Else
820                     RSServerSynchHistory.Close
                    End If

822                 RSLocalSynchHistory.MoveNext
                
                Loop

            End If
            
824         TxMsg "--- Finished commiting pending records to server from table " & sLocalTableWithNewRecords & "......................"
    
        End If
        
        'Set MsXmlHttp = Nothing
        'Set MsXmlDoc = Nothing
826     Set RSServerSynchHistory = Nothing
828     Set RSServerTableToBeUpdated = Nothing
830     Set RSServerTableToBeUpdated_FEA = Nothing
832     Set RSServerTableToBeUpdated_GEO = Nothing
834     Set RSLocalSynchHistory = Nothing
836     Set RSLocalTableRecordsToBeUploaded = Nothing
838     Set RSLocalTableRecordsToBeUploaded_FEA = Nothing
840     Set RSLocalTableRecordsToBeUploaded_GEO = Nothing
        'TxMsg "UploadTableChangesToServer line(" & Erl & ") " & Err.Description
       
        
        '<EhFooter>
        Exit Sub

UploadTableChangesToServer_Err:

            TxMsg "UploadTableChangesToServer_Err: (" & Erl & ") " & Err.Description, True, True
      
        '</EhFooter>
End Sub

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

Private Sub GetSynchRecordsTODELETE(ByVal sTableName As String, _
                                    ByVal sGUID As String, _
                                    bIsGeotable As Boolean, _
                                    sWebsite As String, _
                                    cn As ADODB.Connection, _
                                    Optional sLocalSynchHistPrefix As String = "")
        '<EhHeader>
        On Error GoTo GetSynchRecordsTODELETE_Err
        '</EhHeader>
                                      
        'Operation: This sub fetches a list of GUIDs from the server SynchHistory table and
        '           also from the local SynchHistory table.  Both lists of GUIDs are compared
        '           and any missing on the local table denote records needing deletion on the
        '           local table.  These records are deleted locally.

        Dim RSServerALLGUIDSFORDELETION As New ADODB.Recordset
        Dim RSLocalALLGUIDSDELETED As New ADODB.Recordset
        Dim RSGUIDSNeedingDeleting As New ADODB.Recordset
    
        Dim RSUpdater As New ADODB.Recordset
        Dim RSUpdater_GEO As New ADODB.Recordset
        Dim RSUpdater_FEA As New ADODB.Recordset
        Dim bIncidents As Boolean
    
        Dim i As Long
        Dim j As Long
        Dim lIndex As Long
        Dim sUserGroupPrefix As String
        Dim lCountOfNonRecords As Long
        Dim sQueryName As String
        
100     lCountOfNonRecords = 0
        
102     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then sWebsite = sWebsite & "/"
104     If sTableName = "oincidents" Then bIncidents = True
106     If sLocalSynchHistPrefix = "" And Not bIncidents And Not sTableName = "Attachments" Then
108         sUserGroupPrefix = m_StrRemoteTablePrefix
        End If

        ' And Not sTableName = "GeoBookMarks" And Not sTableName = "GeoBookmarksCategories" And Not sTableName = "Incidents_ChartSettings"
        
110     'TxMsg ">>>>>>>>>>>>"
112     TxMsg "**??Scaning for deletions flagged on server table [" & sUserGroupPrefix & sTableName & "] ......................", True
        'Set RSServerALLGUIDSFORDELETION = GetAllGUIDsFromServerDB(sLocalSynchHistPrefix & sUserGroupPrefix & "SynchHistory", "sID", "WHERE sdelete = 'true' AND [sTableName] = '" & sTableName & "'", sWebsite)

116     'Set RSLocalALLGUIDSDELETED = GetAllGUIDsFromLocalDB(sLocalSynchHistPrefix & "SynchHistory", "sID", "WHERE [updates] = 'true' AND [sdelete] = 'true' AND [sTableName] = '" & sTableName & "' ORDER BY [sWhen] DESC", cn)
        Set RSLocalALLGUIDSDELETED = GetAllGUIDsFromLocalDB(sTableName, sLocalSynchHistPrefix & "SynchHistory", cn, Delete)
        Set RSServerALLGUIDSFORDELETION = GetAllGUIDsFromServerDB(sTableName, sLocalSynchHistPrefix & sUserGroupPrefix & "SynchHistory", m_sWhen, sWebsite, Delete)
118     Set RSGUIDSNeedingDeleting = GetRSWithGUIDSFORDOWNLOAD(RSServerALLGUIDSFORDELETION, RSLocalALLGUIDSDELETED, Delete)
    
120     If bIsGeotable Then

122         If Right$(sTableName, 4) = "_FEA" Or Right$(sTableName, 4) = "_GEO" Then
124             sTableName = Left$(sTableName, Len(sTableName) - 4)
            End If
            
126         RSUpdater_GEO.Open "SELECT * FROM [" & sTableName & "_GEO]", cn, adOpenDynamic, adLockBatchOptimistic
128         RSUpdater_FEA.Open "SELECT * FROM [" & sTableName & "_FEA]", cn, adOpenDynamic, adLockBatchOptimistic

        Else

130         RSUpdater.Open "SELECT * FROM [" & sTableName & "]", cn, adOpenDynamic, adLockBatchOptimistic

        End If

        'Cycle through the RSGUIDSNeedingDeleting and get each record
132     With RSGUIDSNeedingDeleting
    
134         If Not .State = adStateClosed Then

136             If (bIsGeotable And Not RSUpdater_FEA.State = adStateClosed And Not RSUpdater_GEO.State = adStateClosed) Or (Not bIsGeotable And Not RSUpdater.State = adStateClosed) Then
        
138                 If Not .BOF Or Not .EOF Then .MoveFirst
                
140                 Do Until .EOF

142                     If i = 0 Then TxMsg .RecordCount & " records identifed for deletion......................"
144                     i = i + 1
                        
                        'Get new record from server
146                     If bIncidents Or bIsGeotable Then


148                         If Not RSUpdater_FEA.EOF Or Not RSUpdater_FEA.BOF Then
                                RSUpdater_FEA.MoveFirst
150                             RSUpdater_FEA.Find "[" & sGUID & "] = '" & .Fields(0).Value & "'"
                            End If

152                         If Not RSUpdater_FEA.EOF Then
                        
154                             If Not RSUpdater_GEO.EOF Or Not RSUpdater_GEO.BOF Then RSUpdater_GEO.MoveFirst
156                             RSUpdater_GEO.Filter = "UID = " & RSUpdater_FEA.Fields("UID").Value

158                             If Not RSUpdater_GEO.EOF Then
160                                 RSUpdater_GEO.Delete adAffectCurrent
162                                 RSUpdater_GEO.UpdateBatch
                                End If

164                             RSUpdater_FEA.Delete adAffectCurrent
166                             RSUpdater_FEA.UpdateBatch

                                ' Report to client
168                             If Not m_bUseEvents Then
170                                 TxMsg CStr("**??deleted record " & i & " of " & .RecordCount & " and ID: " & .Fields(0).Value & " from " & sTableName)
                                Else
172                                 RaiseEvent StatusEX(i, .RecordCount, False, .Fields(0).Value)
                                End If

                            Else
                            
174                             lCountOfNonRecords = lCountOfNonRecords + 1

176                             If lCountOfNonRecords = 11 Then TxMsg "    --- THERE SEEMS TO BE A LOT OF THESE..... SPEEDING UP THE PROCESS...."
178                             If lCountOfNonRecords > 10 Then
180                                 TxMsg CStr("    --- could not locate ID: " & .Fields(0).Value & " for deletion - updating SynchHistory accordingly!"), False
                                Else
182                                 TxMsg CStr("    --- could not locate ID: " & .Fields(0).Value & " for deletion - updating SynchHistory accordingly!"), True
                                End If
                            
                            End If
                            
184                         'RSUpdater_FEA.Filter = ""
186                         'RSUpdater_GEO.Filter = ""

188                         SynchHistoryDelete cn, GetGuid, .Fields(0).Value, "Synched Record", "", sUserGroupPrefix & " UG", RSGUIDSNeedingDeleting.Fields("sWhen").Value, IIf(sLocalSynchHistPrefix = "", sTableName, sTableName & "_FEA"), True, "true", sLocalSynchHistPrefix
        
                        Else

                            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                            'This IF statement block will help prevent duplication of records
190                         If Not RSUpdater.EOF Or Not RSUpdater.BOF Then
                                RSUpdater.MoveFirst
192                             RSUpdater.Find sGUID & " = '" & .Fields(0).Value & "'"
                            End If
                            
194                         If Not RSUpdater.EOF And Not RSUpdater.BOF Then

                                sQueryName = ""
                                If Right(sTableName, 8) = "_Queries" Then
                                    TxMsg "!!! Since this is the queries table the corresponding query will also be deleted....", True
                                    sQueryName = RSUpdater.Fields("QueryName").Value
                                End If

196                             With RSUpdater
198                                 .Delete adAffectCurrent
200                                 .UpdateBatch
                                End With

                                ' Report to client
202                             If Not m_bUseEvents Then
204                                 TxMsg CStr("**??deleted record " & i & " of " & .RecordCount & " and ID: " & .Fields(0).Value & " from " & sTableName)
                                Else
206                                 RaiseEvent StatusEX(i, .RecordCount, False, .Fields(0).Value)
                                End If
                            
                            Else
                            
208                             lCountOfNonRecords = lCountOfNonRecords + 1

210                             If lCountOfNonRecords = 11 Then TxMsg "    --- THERE SEEMS TO BE A LOT OF THESE..... SPEEDING UP THE PROCESS...."
212                             If lCountOfNonRecords > 10 Then
214                                 TxMsg CStr("    --- could not locate ID: " & .Fields(0).Value & " for deletion - updating SynchHistory accordingly!"), False
                                Else
216                                 TxMsg CStr("    --- could not locate ID: " & .Fields(0).Value & " for deletion - updating SynchHistory accordingly!"), True
                                End If
                            
                            End If

218                         'RSUpdater.Filter = ""
220                         SynchHistoryDelete cn, GetGuid, .Fields(0).Value, "Synched Record", "", sUserGroupPrefix & " UG", RSGUIDSNeedingDeleting.Fields("sWhen").Value, sTableName, True, "true", sLocalSynchHistPrefix, sQueryName
                
                        End If
                    
222                     .MoveNext
                      
                    Loop
        
                Else
224                 TxMsg "GetSynchRecordsTODELETE_Err: closed state of destination RS made the delete operation fail", True
                End If
                
            End If
    
        End With
        
226     TxMsg "**??Finished deletion of records on local table " & sTableName & "......................", True
228     'TxMsg "<<<<<<<<<<<<"
        
230     Set RSServerALLGUIDSFORDELETION = Nothing
232     Set RSLocalALLGUIDSDELETED = Nothing
234     Set RSGUIDSNeedingDeleting = Nothing
    
236     Set RSUpdater = Nothing
238     Set RSUpdater_GEO = Nothing
240     Set RSUpdater_FEA = Nothing
    
        'TxMsg "GetSynchRecordsTODELETE_Err: (" & Erl & ") " & Err.Description
        '<EhFooter>
        Exit Sub

GetSynchRecordsTODELETE_Err:
        TxMsg "GetSynchRecordsTODELETE_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Sub GetSynchRecordsTOUPDATE(ByVal sTableName As String, _
                                    ByVal sGUID As String, _
                                    bIsGeotable As Boolean, _
                                    sWebsite As String, _
                                    cn As ADODB.Connection, _
                                    Optional sLocalSynchHistPrefix As String = "")
        '<EhHeader>
        On Error GoTo GetSynchRecordsTOUPDATE_Err
        '</EhHeader>
                                      
        'Operation: This sub fetches a list of GUIDs from the server SynchHistory table and
        '           also from the local SynchHistory table.  Both lists of GUIDs are compared
        '           and any missing on the local table denote records needing deletion on the
        '           local table.  These records are deleted locally.
    
        Dim RSServerALLGUIDSFOREDIT As New ADODB.Recordset
        Dim RSLocalALLGUIDSEDITED As New ADODB.Recordset
        Dim RSGUIDSNeedingEditing As New ADODB.Recordset
    
        Dim RSServerNewRecord As New ADODB.Recordset
        Dim RSServerNewRecord_FEA As New ADODB.Recordset
        Dim RSServerNewRecord_GEO As New ADODB.Recordset

        Dim RSUpdater As New ADODB.Recordset
        Dim RSUpdater_GEO As New ADODB.Recordset
        Dim RSUpdater_FEA As New ADODB.Recordset

        Dim bIncidents As Boolean
        Dim bAddedNew As Boolean
        Dim bRSServerNewRecord_FEAisEOF As Boolean
        Dim bRSServerNewRecord_GEOisEOF As Boolean
        
        Dim lUID As Long
        
        Dim i As Long
        Dim j As Long
        Dim jField As Integer
        Dim sUserGroupPrefix As String
        
100     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then sWebsite = sWebsite & "/"
102     If sTableName = "oincidents" Then bIncidents = True
104     If sLocalSynchHistPrefix = "" And Not bIncidents And Not sTableName = "Attachments" Then
106         sUserGroupPrefix = m_StrRemoteTablePrefix
        End If

        ' And Not sTableName = "GeoBookMarks" And Not sTableName = "GeoBookmarksCategories" And Not sTableName = "Incidents_ChartSettings"
        
108     'TxMsg ">>>>>>>>>>>>"
110     TxMsg "**??Beginning scan for updates flagged on server table " & sUserGroupPrefix & sTableName & "......................", True

        'Set RSServerALLGUIDSFOREDIT = GetAllGUIDsFromServerDB(sLocalSynchHistPrefix & sUserGroupPrefix & "SynchHistory", "sID", "WHERE sdelete = 'false' AND sequence > 1 AND [sTableName] = '" & sTableName & "'", sWebsite, True)
    '114     Set RSLocalALLGUIDSEDITED = GetAllGUIDsFromLocalDB(sLocalSynchHistPrefix & "SynchHistory", "sID", "WHERE [updates] = 'true' AND [sequence] > 1 AND [sTableName] = '" & sTableName & "' ORDER BY [sWhen] DESC", cn, True)
112     Set RSLocalALLGUIDSEDITED = GetAllGUIDsFromLocalDB(sTableName, sLocalSynchHistPrefix & "SynchHistory", cn, Update)
118     Set RSServerALLGUIDSFOREDIT = GetAllGUIDsFromServerDB(sTableName, sLocalSynchHistPrefix & sUserGroupPrefix & "SynchHistory", m_sWhen, sWebsite, Update)
122     Set RSGUIDSNeedingEditing = GetRSWithGUIDSFORDOWNLOAD(RSServerALLGUIDSFOREDIT, RSLocalALLGUIDSEDITED, Update)

124     If bIsGeotable Then

126         If Right$(sTableName, 4) = "_FEA" Or Right$(sTableName, 4) = "_GEO" Then
128             sTableName = Left$(sTableName, Len(sTableName) - 4)
            End If
            
130         RSUpdater_GEO.Open "SELECT * FROM [" & sTableName & "_GEO]", cn, adOpenDynamic, adLockBatchOptimistic
132         RSUpdater_FEA.Open "SELECT * FROM [" & sTableName & "_FEA]", cn, adOpenDynamic, adLockBatchOptimistic

        Else

134         RSUpdater.Open "SELECT * FROM [" & sTableName & "]", cn, adOpenDynamic, adLockBatchOptimistic

        End If

        'Cycle through the RSGUIDSNeedingEditing and get each record
136     With RSGUIDSNeedingEditing
    
138         If Not .State = adStateClosed Then

140             If (bIsGeotable And Not RSUpdater_FEA.State = adStateClosed And Not RSUpdater_GEO.State = adStateClosed) Or (Not bIsGeotable And Not RSUpdater.State = adStateClosed) Then
        
142                 If Not .EOF Or Not .BOF Then .MoveFirst
                
144                 Do Until .EOF

146                     If i = 0 Then TxMsg "    " & .RecordCount & " record(s) identifed for updating ......................"
148                     i = i + 1
                        
                        'Get new record from server
150                     If bIncidents Or bIsGeotable Then
                         
152                         If bIncidents Then
154                             Set RSServerNewRecord_FEA = GetAllFieldsForIncidentGUIDOnServer_FEA(.Fields(0).Value, sWebsite)
                            Else
156                             Set RSServerNewRecord_FEA = GetAllFieldsForGUIDOnServer(sUserGroupPrefix & sTableName & "_FEA", sGUID, .Fields(0).Value, sWebsite)
                            End If
                            
158                         bRSServerNewRecord_GEOisEOF = True
160                         bRSServerNewRecord_FEAisEOF = True
                      
162                         If Not RSServerNewRecord_FEA.State = adStateClosed Then

164                             If Not RSServerNewRecord_FEA.EOF Or Not RSServerNewRecord_FEA.BOF Then
166                                 RSServerNewRecord_FEA.MoveFirst

168                                 If bIncidents Then

170                                     Set RSServerNewRecord_GEO = GetAllFieldsForIncidentGUIDOnServer_GEO(.Fields(0).Value, sWebsite)
                                    Else
                                   
172                                     Set RSServerNewRecord_GEO = GetAllFieldsForGUIDOnServer(sUserGroupPrefix & sTableName & "_GEO", "UID", RSServerNewRecord_FEA.Fields("UID").Value, sWebsite)
                                    End If
                                    
174                                 bRSServerNewRecord_GEOisEOF = RSServerNewRecord_GEO.EOF
176                                 bRSServerNewRecord_FEAisEOF = False
                                
                                Else
178                                 bRSServerNewRecord_GEOisEOF = True
180                                 bRSServerNewRecord_FEAisEOF = True
                                End If

                            Else
                            
182                             TxMsg CStr("    ---(at line 152) RSServerNewRecord_FEA.State = adStateClosed!"), True
                            
                            End If

184                         If Not RSUpdater_FEA.EOF Or Not RSUpdater_FEA.BOF Then RSUpdater_FEA.MoveFirst
186                         RSUpdater_FEA.Filter = "[" & sGUID & "] = '" & .Fields(0).Value & "'"
                           
188                         If Not bRSServerNewRecord_FEAisEOF And Not bRSServerNewRecord_GEOisEOF Then
                      
190                             If RSUpdater_FEA.EOF Then
                        
192                                 RSUpdater_FEA.AddNew
194                                 RSUpdater_GEO.AddNew
196                                 bAddedNew = True
198                                 lUID = GetMaxUIDFromGEOTable(sTableName & "_GEO", cn) + CLng(1)
                            
                                Else
                         
200                                 bAddedNew = False

202                                 If Not RSUpdater_GEO.EOF Or Not RSUpdater_GEO.BOF Then RSUpdater_GEO.MoveFirst
204                                 RSUpdater_GEO.Filter = "UID = " & RSUpdater_FEA.Fields("UID").Value
206                                 lUID = RSUpdater_FEA.Fields("UID").Value
                                
                                End If
                            
208                             If bAddedNew Or Not RSUpdater_GEO.EOF Then
              
210                                 jField = 0 ' why was the 7????
    
212                                 Do Until jField >= RSUpdater_GEO.Fields.Count
                                
214                                     If RSUpdater_GEO.Fields(jField).Name = "UID" Then
    
216                                         RSUpdater_GEO.Fields(jField).Value = lUID
    
218                                     ElseIf Not RSUpdater_GEO.Fields(jField).Name = "UID" And Not RSUpdater_GEO.Fields(jField).Name = "ORIG_ID" Then

220                                         If Not (bIncidents And RSUpdater_GEO.Fields(jField).Name = "GEOMETRY") Then
222                                             RSUpdater_GEO.Fields(jField).Value = RSServerNewRecord_GEO.Fields(RSUpdater_GEO.Fields(jField).Name).Value
                                            End If
                                            
                                        End If
                                     
224                                     jField = jField + 1
                                    Loop
    
226                                 RSUpdater_GEO.UpdateBatch adAffectCurrent
                            
228                                 jField = 0

230                                 Do Until jField >= RSUpdater_FEA.Fields.Count
                               
232                                     If Not RSUpdater_FEA.Fields(jField).Name = "UID" Then

234                                         RSUpdater_FEA.Fields(jField).Value = RSServerNewRecord_FEA.Fields(RSUpdater_FEA.Fields(jField).Name).Value
                                        Else
236                                         RSUpdater_FEA.Fields("UID").Value = lUID
                                        End If
                                    
238                                     jField = jField + 1
                                    Loop

240                                 RSUpdater_FEA.UpdateBatch adAffectCurrent
                                Else

242                                 TxMsg "Check out on local: " & RSUpdater_GEO.Filter

                                End If

244                             RSUpdater_FEA.Filter = ""
246                             RSUpdater_GEO.Filter = ""

248                             If Not sLocalSynchHistPrefix = "" Then
250                                 sTableName = sTableName & "_FEA"
                                End If

252                             SynchHistoryEdit cn, GetGuid, .Fields(0).Value, "Synched Record", "", sUserGroupPrefix & " UG", RSGUIDSNeedingEditing.Fields("sWhen").Value, sTableName, True, "true", sLocalSynchHistPrefix, RSGUIDSNeedingEditing.Fields("sequence").Value

                                ' Report to client
254                             If Not m_bUseEvents Then
256                                 TxMsg CStr("**??updated record " & i & " of " & .RecordCount & " and ID: " & .Fields(0).Value & " from " & sTableName)
                                Else
258                                 RaiseEvent StatusEX(i, .RecordCount, False, .Fields(0).Value)
                                End If
                            
260                             If Not sLocalSynchHistPrefix = "" Then
262                                 sTableName = Left$(sTableName, Len(sTableName) - 4)
                                End If

                            Else
264                             TxMsg CStr("    --- could not locate ID: " & .Fields(0).Value & " on the server for update!"), True

266                             If Not sLocalSynchHistPrefix = "" Then
268                                 sTableName = sTableName & "_FEA"
                                End If

270                             SynchHistoryDeleteFromServer sWebsite, .Fields(0).Value, sTableName, sLocalSynchHistPrefix

272                             If Not sLocalSynchHistPrefix = "" Then
274                                 sTableName = Left$(sTableName, Len(sTableName) - 4)
                                End If
                            End If
        
                        Else
                    
276                         Set RSServerNewRecord = GetAllFieldsForGUIDOnServer(sUserGroupPrefix & sTableName, sGUID, .Fields(0).Value, sWebsite)

278                         If Not RSUpdater.EOF Or Not RSUpdater.BOF Then RSUpdater.MoveFirst
280                         RSUpdater.Filter = sGUID & " = '" & .Fields(0).Value & "'"
                        
282                         If Not RSServerNewRecord.EOF Then
                                
284                             If RSUpdater.EOF Then
286                                 RSUpdater.AddNew
                                End If

288                             jField = 0

290                             Do Until jField = RSUpdater.Fields.Count
                                    'TxMsg "was: " & RSUpdater.Fields(jField).Value
                                    'TxMsg "will be: " & RSServerNewRecord.Fields(RSUpdater.Fields(jField).Name).Value

292                                 If Not RSUpdater.Fields(jField).Name = "UID" And Not (RSUpdater.Fields(jField).Name = "ID" And sTableName = "GeoBookMarks") Then
    
294                                     If RSUpdater.Fields(jField).Name = "OCTSetting" Then
296                                         RSUpdater.Fields(jField).Value = AsciiToBinary(RSServerNewRecord.Fields(RSUpdater.Fields(jField).Name).Value)
                                        ElseIf Not RSUpdater.Fields(jField).Name = "ORIG_ID" Then
298                                         RSUpdater.Fields(jField).Value = RSServerNewRecord.Fields(RSUpdater.Fields(jField).Name).Value
                                        End If


                                    End If

300                                 jField = jField + 1
                                Loop

302                             RSUpdater.UpdateBatch adAffectCurrent
304                             RSUpdater.Filter = ""
306                             SynchHistoryEdit cn, GetGuid, .Fields(0).Value, "Synched Record", "", sUserGroupPrefix & " UG", RSGUIDSNeedingEditing.Fields("sWhen").Value, sTableName, True, "true", sLocalSynchHistPrefix, RSGUIDSNeedingEditing.Fields("sequence").Value

                                ' Report to client
308                             If Not m_bUseEvents Then
310                                 TxMsg CStr("**??updated record " & i & " of " & .RecordCount & " and ID: " & .Fields(0).Value & " from " & sTableName)
                                Else
312                                 RaiseEvent StatusEX(i, .RecordCount, False, .Fields(0).Value)
                                End If

                            Else
                        
314                             TxMsg CStr("    --- could not locate ID: " & .Fields(0).Value & " on the server for update!"), True
316                             SynchHistoryDeleteFromServer sWebsite, .Fields(0).Value, sTableName, sLocalSynchHistPrefix
                            
                            End If
                
                        End If
                    
318                     .MoveNext
                      
                    Loop
        
                Else
320                 TxMsg "GetSynchRecordsTODELETE_Err: closed state of destination RS made the delete operation fail", True
                End If
                
            End If
    
        End With
        
322     TxMsg "**??Finished update of records on local table " & sTableName & "......................", True
324     'TxMsg "<<<<<<<<<<<<"
    
326     Set RSServerALLGUIDSFOREDIT = Nothing
328     Set RSLocalALLGUIDSEDITED = Nothing
330     Set RSGUIDSNeedingEditing = Nothing
    
332     Set RSServerNewRecord = Nothing
334     Set RSServerNewRecord_FEA = Nothing
336     Set RSServerNewRecord_GEO = Nothing
    
338     Set RSUpdater = Nothing
340     Set RSUpdater_GEO = Nothing
342     Set RSUpdater_FEA = Nothing
    
        'TxMsg "GetSynchRecordsTOUPDATE_Err: (" & Erl & ") " & Err.Description
        '<EhFooter>
        Exit Sub

GetSynchRecordsTOUPDATE_Err:
        TxMsg "GetSynchRecordsTOUPDATE_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

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

Private Sub PrepareHttpComms(oHttp As WinHttpRequest, sWebsite As String, bUpload As Boolean)
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

Private Sub GetNumberOfBatchToDownload(sWebsite As String)

    Dim sResponse As String
    sResponse = OpenServerResponseCompressed(sWebsite, "batchcounttodownload", "")
    
    If Len(sResponse) > 0 Then
        lBatchCountToDownload = CLng(sResponse)
        TxMsg "*** " & lBatchCountToDownload & " records will be downloaded in batch"
    Else
        lBatchCountToDownload = 200
        TxMsg "*** compression.txt is not working on server.  The default " & lBatchCountToDownload & " records will be downloaded in batch", True
    End If
    
    'TxMsg "********************************************************************************************"

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

Private Sub SynchDDWithServer(sTableNamePrefix As String, _
                              sWebsite As String, _
                              cn As ADODB.Connection)
        '<EhHeader>
        On Error GoTo SynchDDWithServer_Err
        '</EhHeader>

        Dim sLocalTableName As String
        Dim RSLocalSynchHistory As New ADODB.Recordset
        Dim RSServerSynchHistory As New ADODB.Recordset
        Dim bAddNew As Boolean
        Dim bUpdate As Boolean
        Dim bDelete As Boolean
        Dim sLocalTableGUID As String
        Dim bIsGeo As Boolean
        Dim sSQL As String
        Dim sTableNames() As String
        Dim i As Long
        Dim sRetVal As String
        Dim SynchOp As SynchOperation
        Dim sTableNameSmall As String
        Dim lOldNumBatchDownload As Long
        Dim bNewRecordsFetched As Boolean
    
100     TxMsg ">>>>>>>>>>>>"
102     TxMsg ">>> Beginning synchronisation for dynamic database: " & sTableNamePrefix & "......................"

104     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then sWebsite = sWebsite & "/"
106     m_sIDs = ""

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'UPLOAD DATA...............................
        '
        'RSLocalSynchHistory.Open "SELECT * FROM [" & sTableNamePrefix & "SynchHistory] WHERE [updates] = 'false' ORDER BY [sdelete] DESC", cn, adOpenDynamic, adLockBatchOptimistic
108     RSLocalSynchHistory.Open "SELECT * FROM [" & sTableNamePrefix & "SynchHistory] WHERE [updates] = 'false' ORDER BY [sTableName] DESC", cn, adOpenDynamic, adLockBatchOptimistic
        
110     If Not RSLocalSynchHistory.State = adStateClosed Then

112         If RSLocalSynchHistory.EOF = False Then

114             TxMsg "Records are pending update to server..........."

116             Do Until RSLocalSynchHistory.EOF

118                 If RSLocalSynchHistory.Fields("sdelete").Value = "true" Then
120                     SynchOp = Delete
                    Else

122                     If RSLocalSynchHistory.Fields("sequence").Value > 1 Then
124                         SynchOp = Update
126                     ElseIf RSLocalSynchHistory.Fields("sequence").Value = 1 Then
128                         SynchOp = Add
                        End If
                    End If

130                 sLocalTableName = RSLocalSynchHistory.Fields("sTableName").Value
132                 sTableNameSmall = Right$(sLocalTableName, Len(sLocalTableName) - Len(sTableNamePrefix))

134                 If Left$(sTableNameSmall, 4) = "link" Then
136                     sLocalTableGUID = "GUID2"
                    Else
138                     sLocalTableGUID = "GUID1"
                    End If

140                 If Left$(Replace(sLocalTableName, sTableNamePrefix, "", , , vbTextCompare), 4) = "link" Then sLocalTableGUID = "GUID2"
142                 bIsGeo = False

144                 If Right$(sLocalTableName, 4) = "_FEA" Then
146                     bIsGeo = True
148                     sLocalTableName = Left$(sLocalTableName, Len(sLocalTableName) - 4)
                    End If
                    
150                 UploadTableChangesToServer sLocalTableName, sLocalTableGUID, bIsGeo, sWebsite, cn, SynchOp, sTableNamePrefix
152                 RSLocalSynchHistory.MoveFirst
154                 RSLocalSynchHistory.Requery
                
                Loop

            End If

        End If
        
156     Set RSLocalSynchHistory = Nothing

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' DOWNLOAD DATA
        
158     GetSynchTablesToWipe sWebsite, cn, sTableNamePrefix
        
160     sRetVal = GetTableNames(cn, sTableNamePrefix)
162     sTableNames = Split(sRetVal, ",", , vbTextCompare)
        
164     i = UBound(sTableNames)

166     Do Until i = -1 'UBound(sTableNames) + 1
          
168         sTableNameSmall = Right$(sTableNames(i), Len(sTableNames(i)) - Len(sTableNamePrefix))

170         If Left$(sTableNameSmall, 4) = "link" Then
172             sLocalTableGUID = "GUID2"
            Else
174             sLocalTableGUID = "GUID1"
            End If
            
            'If Not Right(sTableNames(i), 13) = "ChartSettings" And Not Right(sTableNames(i), 7) = "Queries" Then
                'sTableNames(i) = sTableNames(i) & "_GEO"
            'End If
            
176         If Not Right$(sTableNames(i), 12) = "SynchHistory" And Not Right$(sTableNames(i), 20) = "SynchHistoryOverview" And Not Right$(sTableNames(i), 4) = "_GEO" Then

178             If Right$(sTableNames(i), 4) = "_FEA" Then
180                 bIsGeo = True
                    ' sTableNames(i) = Left$(sTableNames(i), Len(sTableNames(i)) - 4)
                Else
182                 bIsGeo = False
                End If
                

                
186             lOldNumBatchDownload = lBatchCountToDownload
188             bLargePolygonDownloadFailed = False
                
190             bNewRecordsFetched = GetSynchRecordsTOADD(sTableNames(i), sLocalTableGUID, bIsGeo, sWebsite, cn, sTableNamePrefix)
                If SynchHistoryOverview.NextSynchNextOpeation = NewOnly And Not bNewRecordsFetched Then GetSynchRecordsTODELETE sTableNames(i), sLocalTableGUID, bIsGeo, sWebsite, cn, sTableNamePrefix
                
'                If SynchHistoryOverview.NextSynchNextOpeation = NewOnly And m_lCountDownloaded = 0 Then
'                    TxMsg "!!!! Next synch operation is still set to NewOnly but no records were downloaded!  Adjusting SynchHistoryOverview", True, True
'                    cn.Execute "UPDATE [" & sTableNamePrefix & "SynchHistoryOverview] SET sWhenMaxSearched = sWhenMinSearched WHERE [sTablename] = '" & sTableNames(i) & "'"
'
'                End If
'
'192             Do While SynchHistoryOverview.NextSynchNextOpeation = NewOnly 'And bNewRecordsFetched
'194                 bNewRecordsFetched = GetSynchRecordsTOADD(sTableNames(i), sLocalTableGUID, bIsGeo, sWebsite, cn, sTableNamePrefix)
'                Loop
                                                  
            'need to check this logic
                If Not SynchHistoryOverview.NextSynchNextOpeation = DoNotSynch And Not SynchHistoryOverview.NextSynchNextOpeation = EditsOnly Then
                    GetSynchRecordsTODELETE sTableNames(i), sLocalTableGUID, bIsGeo, sWebsite, cn, sTableNamePrefix
                End If
            '''''''''
                
196             If bLargePolygonDownloadFailed Then
                
198                 If bIsGeo And lBatchCountToDownload > 5 Then
200                     If IsGeoTablePoly(sTableNames(i), cn) Then
202                         lBatchCountToDownload = 5
204                         TxMsg "!!!! Detected polygon / polyline download issue - setting batch download to 5 a time"
                        End If
                    End If
                
206                 bNewRecordsFetched = GetSynchRecordsTOADD(sTableNames(i), sLocalTableGUID, bIsGeo, sWebsite, cn, sTableNamePrefix)
208                 bLargePolygonDownloadFailed = False
                    
                End If
                
210             If SynchHistoryOverview.NextSynchNextOpeation = EditsOnly Then GetSynchRecordsTOUPDATE sTableNames(i), sLocalTableGUID, bIsGeo, sWebsite, cn, sTableNamePrefix
                
212             If Right$(sTableNames(i), 8) = "_Queries" Then
214                 If CompareQueries(sWebsite, sTableNamePrefix, cn) = False Then
216                     TxMsg "!!!!!! Update of queries failed", True
                    Else
218                     'TxMsg ">>>>>> Update of queries succeeded", True
                    End If
                End If
                
220             lBatchCountToDownload = lOldNumBatchDownload
                
            End If

222         i = i - 1
        Loop
        
        TxMsg "**??Updating extents for queries......................", True
        FixLocalQueriesExtents cn, sTableNamePrefix
        
224     TxMsg "**??Finished synchronisation of dynamic database: " & sTableNamePrefix & "......................", True
226     TxMsg ">>>>>>>>>>>>"
        '<EhFooter>
        Exit Sub

SynchDDWithServer_Err:
        TxMsg "SynchDDWithServer_Err: (" & Erl & ") " & Err.Description, True
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

Private Function GetTableNames(oCOnn As ADODB.Connection, _
                               sTablePreFix As String)
        '<EhHeader>
        On Error GoTo GetTableNames_Err
        '</EhHeader>
 
        Dim oDB As adox.Catalog
        Dim itbl As adox.Table
 
100     Set oDB = New adox.Catalog
102     Set itbl = New adox.Table
104     Set oDB.ActiveConnection = oCOnn
106     GetTableNames = ""

108     For Each itbl In oDB.Tables

110         If InStr(1, itbl.Name, sTablePreFix) > 0 Then
    
112             If bSQLServerInUse Then
114                 If itbl.Type = "TABLE" Then GetTableNames = GetTableNames & itbl.Name & ","
                Else
                
116                 If Not itbl.Properties("Jet OLEDB:Create Link") = True Then
                        'get linked table
118                     If itbl.Type = "TABLE" Then GetTableNames = GetTableNames & itbl.Name & ","
                    End If
                
                End If
                
            End If

        Next
        
        'GetTableNames = GetTableNames & sTablePrefix & "ChartSettings,"
        
120     Set oDB = Nothing
122     Set itbl = Nothing

124     If Len(GetTableNames) > 0 Then GetTableNames = Left$(GetTableNames, Len(GetTableNames) - 1)

        '<EhFooter>
        Exit Function

GetTableNames_Err:
        TxMsg "GetTableNames_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Public Function SynchHistoryEdit(oCN As ADODB.Connection, _
                                 sNewGUID As String, _
                                 sID As String, _
                                 sTitle As String, _
                                 sDescription As String, _
                                 sBy As String, _
                                 sRFC3339DateTime As String, _
                                 sTableName As String, _
                                 bIsGeoLayer As Boolean, _
                                 supdates As String, _
                                 Optional sLocalSynchHistPrefix As String = "", _
                                 Optional lSequence As Long = 0)
        '<EhHeader>
        On Error GoTo SynchHistoryEdit_Err
        '</EhHeader>
        
        Dim i As Long
        Dim j As Long

100     SynchHistoryEdit = False
    
        Dim oRS As New ADODB.Recordset
        Dim oRSQuery As ADODB.Recordset
102     oRS.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "' ORDER by [sequence] DESC", oCN, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If Not .EOF Then

                    '.Fields("sID").Value = sID
110                 .Fields("sGUID").Value = sNewGUID
112                 .Fields("sTableName").Value = sTableName
114                 .Fields("swhen").Value = sRFC3339DateTime
116                 .Fields("sStatus").Value = "pending"

118                 If lSequence = 0 Then
120                     .Fields("sequence").Value = CLng(.Fields("sequence").Value) + 1
                    Else
122                     .Fields("sequence").Value = lSequence
                    End If
                    
124                 .Fields("sBy").Value = "[" & sTitle & "] " & sDescription & " (" & sBy & ")"
126                 .Fields("sdelete").Value = "false"
128                 .Fields("updates").Value = supdates
130                 .Fields("noconflict").Value = "false"
132                 .UpdateBatch adAffectCurrent

134                 If Right(sTableName, 8) = "_Queries" Then
                    
136                     Set oRSQuery = New ADODB.Recordset

138                     With oRSQuery
                        
140                         .Open "SELECT * FROM [" & sTableName & "] WHERE GUID1 = '" & sID & "'", oCN, adOpenDynamic, adLockBatchOptimistic

142                         If Not .EOF Then

144                             i = LocateIndexOfQuery(LocalQueries, sLocalSynchHistPrefix & .Fields("QueryName").Value)
146                             j = LocateIndexOfQuery(DefinedDDQueries, sLocalSynchHistPrefix & .Fields("QueryName").Value)
                    
148                             If i > 0 And j > 0 Then
                                
152                                 DefinedDDQueries.Queries(j).bEdited = True
154                                 TxMsg ">>>>>> Query [" & sLocalSynchHistPrefix & .Fields("QueryName").Value & "] flagged for edit", True
                                
156                             ElseIf j > 0 Then
                                
160                                 DefinedDDQueries.Queries(j).bNew = True
162                                 TxMsg ">>>>>> Query [" & sLocalSynchHistPrefix & .Fields("QueryName").Value & "] flagged for addition", True
                                    
                                Else
                                    TxMsg "!!!!!! Query [" & sLocalSynchHistPrefix & .Fields("QueryName").Value & "] is not in the local database or in [" & sLocalSynchHistPrefix & "Queries] table - edit failed", True
                                End If

                                'CompareQueries sWebsite, sPrefix, oCN
                                'DropLocalQuery sLocalSynchHistPrefix & .Fields("QueryName").Value, oCN, sLocalSynchHistPrefix
                                'CreateQuery sLocalSynchHistPrefix & .Fields("QueryName").Value, .Fields("QuerySQL").Value, oCN
                                
                            End If

164                         .Close
                        End With
                        
166                     Set oRSQuery = Nothing
                        
                    End If
                    
168                 If .RecordCount > 1 Then
                    
170                     TxMsg "!!!!!!!!!!!! (SynchHistoryEdit) Something dodgy is going on the table: [" & sLocalSynchHistPrefix & "SynchHistory].  There should be only 1 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  The additional records will be deleted.  If this error message persists please contact an OASIS Developer", True
                        
172                     .MoveNext

174                     Do Until .EOF
176                         .Delete adAffectCurrent
178                         .UpdateBatch adAffectCurrent
180                         .MoveNext
                        Loop
                            
                    End If
                    
182                 SynchHistoryEdit = True
            
                Else
                      
184                 TxMsg "!!!!!!!!!!!! (SynchHistoryEdit) Something dodgy is going on the table: [" & sLocalSynchHistPrefix & "SynchHistory].  There should be a record with sID: " & sID & " for tablename [" & sTableName & "] but it is missing!!!  This record will be created.  If this error message persists please contact an OASIS Developer" ', True
186                 SynchHistoryAddNew oCN, sNewGUID, sID, sTitle, sDescription, sBy, sRFC3339DateTime, sTableName, bIsGeoLayer, supdates, sLocalSynchHistPrefix
188                 SynchHistoryEdit = SynchHistoryEdit(oCN, sNewGUID, sID, sTitle, sDescription, sBy, sRFC3339DateTime, sTableName, bIsGeoLayer, supdates, sLocalSynchHistPrefix)
        
                End If

            End If
        
        End With
        
190     Set oRS = Nothing
        
        '<EhFooter>
        Exit Function

SynchHistoryEdit_Err:
        SynchHistoryEdit = False
        Set oRS = Nothing
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

Public Function SynchHistoryDelete(oCN As ADODB.Connection, _
                                   sNewGUID As String, _
                                   sID As String, _
                                   sTitle As String, _
                                   sDescription As String, _
                                   sBy As String, _
                                   sRFC3339DateTime As String, _
                                   sTableName As String, _
                                   bIsGeoLayer As Boolean, _
                                   supdates As String, _
                                   Optional sLocalSynchHistPrefix As String = "", _
                                   Optional sQueryToDelete As String = "")
        '<EhHeader>
        On Error GoTo SynchHistoryDelete_Err
        '</EhHeader>
        
100     SynchHistoryDelete = False
        
        Dim oRS As New ADODB.Recordset
        Dim oRSQuery As New ADODB.Recordset
        Dim j As Long
102     oRS.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "' ORDER by [sequence] DESC", oCN, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If Not .EOF Then

                    '.Fields("sID").Value = sID
110                 .Fields("sGUID").Value = sNewGUID
112                 .Fields("sTableName").Value = sTableName
114                 .Fields("swhen").Value = sRFC3339DateTime
116                 .Fields("sStatus").Value = "pending"
118                 .Fields("sequence").Value = CLng(.Fields("sequence").Value) + 1
120                 .Fields("sBy").Value = "[" & sTitle & "] " & sDescription & " (" & sBy & ")"
122                 .Fields("sdelete").Value = "true"
124                 .Fields("updates").Value = supdates
126                 .Fields("noconflict").Value = "false"
128                 .UpdateBatch adAffectCurrent

130                 If Len(sQueryToDelete) > 0 Then

                        j = LocateIndexOfQuery(DefinedDDQueries, sLocalSynchHistPrefix & .Fields("QueryName").Value)
                             
                        If j > 0 Then
                                
                            DefinedDDQueries.Queries(j).bDelete = True
                            TxMsg ">>>>>> Query [" & sLocalSynchHistPrefix & .Fields("QueryName").Value & "] flagged for deletion", True
                                    
                        End If
                        
                    End If
                    
146                 If .RecordCount > 1 Then
                    
148                     TxMsg "!!!!!!!!!!!! (SynchHistoryDelete) Something dodgy is going on the table: [" & sLocalSynchHistPrefix & "SynchHistory].  There should be only 1 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  The additional records will be deleted.  If this error message persists please contact an OASIS Developer", True
                        
150                     .MoveNext

152                     Do Until .EOF
154                         .Delete adAffectCurrent
156                         .UpdateBatch adAffectCurrent
158                         .MoveNext
                        Loop
                            
                    End If
                    
160                 SynchHistoryDelete = True
            
                Else
            
                    'TxMsg "!!!!!!!!!!!! (SynchHistoryDelete) Something dodgy is going on the table: [" & sLocalSynchHistPrefix & "SynchHistory].  There should be a record with sID: " & sID & " for tablename [" & sTableName & "] but it is missing!!!  This record will be created.  If this error message persists please contact an OASIS Developer"
162                 SynchHistoryAddNew oCN, sNewGUID, sID, sTitle, sDescription, sBy, sRFC3339DateTime, sTableName, bIsGeoLayer, supdates, sLocalSynchHistPrefix
164                 SynchHistoryDelete = SynchHistoryDelete(oCN, sNewGUID, sID, sTitle, sDescription, sBy, sRFC3339DateTime, sTableName, bIsGeoLayer, supdates, sLocalSynchHistPrefix)
            
                End If
    
            End If
        
        End With
        
166     Set oRS = Nothing
        
        '<EhFooter>
        Exit Function

SynchHistoryDelete_Err:
        SynchHistoryDelete = False
        Set oRS = Nothing
        '</EhFooter>
End Function

Public Function SynchHistoryAddNew(oCN As ADODB.Connection, _
                                   sNewGUID As String, _
                                   sID As String, _
                                   sTitle As String, _
                                   sDescription As String, _
                                   sBy As String, _
                                   sRFC3339DateTime As String, _
                                   ByVal sTableName As String, _
                                   bIsGeoLayer As Boolean, _
                                   supdates As String, _
                                   Optional sLocalSynchHistPrefix As String = "", _
                                   Optional lSequenceOverride = 1)
        '<EhHeader>
        On Error GoTo SynchHistoryAddNew_Err
        '</EhHeader>

100     SynchHistoryAddNew = False

        Dim i As Long
        Dim j As Long
        Dim oRS As New ADODB.Recordset
        Dim oRSQuery As ADODB.Recordset
        
102     oRS.Open "SELECT * FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE sID = '" & sID & "' AND sTableName = '" & sTableName & "'", oCN, adOpenDynamic, adLockBatchOptimistic

104     With oRS
 
106         If .State = adStateOpen Then
    
108             If .RecordCount > 0 Then
                    
110                 TxMsg "    !!! (SynchHistoryAddNew) Something dodgy is going on the table: [" & sLocalSynchHistPrefix & "SynchHistory].  There should be 0 record with sID: " & sID & " for tablename [" & sTableName & "] but there are actually " & .RecordCount & "!!!  These records will be deleted.  If this error message persists please contact an OASIS Developer", True

112                 Do Until .EOF
114                     .Delete adAffectCurrent
116                     .UpdateBatch adAffectCurrent
118                     .MoveNext
                    Loop
                            
                End If
    
120             If .EOF Then
                    
122                 .AddNew
124                 .Fields("sID").Value = sID
126                 .Fields("sGUID").Value = sNewGUID
128                 .Fields("sTableName").Value = sTableName
130                 .Fields("swhen").Value = sRFC3339DateTime
132                 .Fields("sStatus").Value = "pending"
134                 .Fields("sequence").Value = lSequenceOverride
136                 .Fields("sBy").Value = "[" & sTitle & "] " & sDescription & " (" & sBy & ")"
138                 .Fields("sdelete").Value = "false"
140                 .Fields("updates").Value = supdates
142                 .Fields("noconflict").Value = "false"
144                 .UpdateBatch adAffectCurrent
146                 SynchHistoryAddNew = True

148                 If Right(sTableName, 8) = "_Queries" Then
                    
150                     Set oRSQuery = New ADODB.Recordset

152                     With oRSQuery
                        
154                         .Open "SELECT * FROM [" & sTableName & "] WHERE GUID1 = '" & sID & "'", oCN, adOpenDynamic, adLockBatchOptimistic

156                         If Not .EOF Then

158                             i = LocateIndexOfQuery(LocalQueries, sLocalSynchHistPrefix & .Fields("QueryName").Value)
160                             j = LocateIndexOfQuery(DefinedDDQueries, sLocalSynchHistPrefix & .Fields("QueryName").Value)
                    
162                             If i > 0 And j > 0 Then
                                
164                                 DefinedDDQueries.Queries(j).bEdited = True
166                                 TxMsg ">>>>>> Query [" & sLocalSynchHistPrefix & .Fields("QueryName").Value & "] flagged for edit", True
                                
168                             ElseIf j > 0 Then
                                
170                                 DefinedDDQueries.Queries(j).bNew = True
172                                 TxMsg ">>>>>> Query [" & sLocalSynchHistPrefix & .Fields("QueryName").Value & "] flagged for addition", True
                                    
                                Else
                                
174                                 ReDim Preserve DefinedDDQueries.Queries(UBound(DefinedDDQueries.Queries) + 1)
176                                 DefinedDDQueries.Queries(UBound(DefinedDDQueries.Queries)).Name = sLocalSynchHistPrefix & .Fields("QueryName").Value

178                                 If bSQLServerInUse Then
180                                     DefinedDDQueries.Queries(UBound(DefinedDDQueries.Queries)).sSQL = .Fields("QueryMSSQL").Value
                                    Else
182                                     DefinedDDQueries.Queries(UBound(DefinedDDQueries.Queries)).sSQL = .Fields("QuerySQL").Value
                                    End If

184                                 DefinedDDQueries.Queries(UBound(DefinedDDQueries.Queries)).bNew = True
                                
186                                 TxMsg ">>>>>> Query [" & sLocalSynchHistPrefix & .Fields("QueryName").Value & "] flagged for addition", True

                                End If

                                'CreateQuery sLocalSynchHistPrefix & .Fields("QueryName").Value, .Fields("QuerySQL").Value, oCN

                            End If

188                         .Close
                        End With
                        
190                     Set oRSQuery = Nothing
                        
                    End If
                    
                End If
    
            Else
            
192             TxMsg "    !!! (SynchHistoryAddNew) Table [" & sLocalSynchHistPrefix & "SynchHistory] failed to open", True
            
            End If
        
        End With
        
194     Set oRS = Nothing
        Set oRSQuery = Nothing
        
        '<EhFooter>
        Exit Function

SynchHistoryAddNew_Err:
        SynchHistoryAddNew = False
        Set oRS = Nothing
        '</EhFooter>
End Function

Public Sub SynchHistoryDeleteFromServer(sWebsite As String, _
                                        sGUID As String, _
                                        sTableName As String, _
                                        sLocalSynchHistPrefix As String)
        '<EhHeader>
        On Error GoTo SynchHistoryDeleteFromServer_Err
        '</EhHeader>

        Dim oRS As New ADODB.Recordset
        Dim sSQL As String
        Dim sResponse As String

100     sSQL = "SELECT [sID] FROM [" & sLocalSynchHistPrefix & "SynchHistory] WHERE sID = '" & sGUID & "' AND sTableName = '" & sTableName & "'"
102     Set oRS = OpenServerRSCompressed(sWebsite, "ID", sSQL)

104     If Not oRS.EOF Or Not oRS.BOF Then
106         oRS.Delete adAffectCurrent
            sResponse = SaveServerRSCompressed(sWebsite, oRS)
        
114         If sResponse = "Data Updated" Then
116             TxMsg "    --- The invalid update flagged on server has been removed", True
            Else
                TxMsg "    --- The invalid update flagged on server could not be removed: " & sResponse, True
            End If
        End If
        
118     If oRS.State = adStateOpen Then oRS.Close
120     Set oRS = Nothing
    
        '<EhFooter>
        Exit Sub

SynchHistoryDeleteFromServer_Err:
        TxMsg "SynchHistoryDeleteFromServer_Err: (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Sub

Private Function DoesFieldExistInTable(oCN As ADODB.Connection, _
                                       sTableName As String, _
                                       sFieldName As String, _
                                       sFieldValue As String) As Boolean

    Dim oRS As New ADODB.Recordset
    Dim sSQL As String
    
    sSQL = "SELECT [" & sFieldName & "] FROM [" & sTableName & "] WHERE [" & sFieldName & "] = '" & sFieldValue & "'"
    
    'TxMsg "DoesFieldExistInTable: " & sSQL
    oRS.Open sSQL, oCN, adOpenDynamic, adLockBatchOptimistic

    DoesFieldExistInTable = False
    
    If oRS.State = adStateOpen Then
    
        If Not oRS.EOF Then DoesFieldExistInTable = True
        oRS.Close
        
    End If
    
    Set oRS = Nothing

End Function

'Private Function RemoveWhereValueExists(mConn As ADODB.Connection, _
'                                        sValue As String, _
'                                        sFieldName As String, _
'                                        sTableName As String) As Boolean
'
'    Dim sTableNameStripped As String
'    Dim bIsMasterTable As Boolean
'    Dim bIsDDDefTable As Boolean
'    Dim bIsLinkedTable As Boolean
'    Dim sTitle As String
'    Dim oRS As New ADODB.Recordset
'
'    sTableNameStripped = Right$(sTableName, Len(sTableName) - Len(DDDefCurrent.prefix))
'
'    If Left$(sTableNameStripped, 2) = "dd" Then
'        bIsDDDefTable = True
'        sTitle = "DD ddDef "
'    End If
'
'    If Left$(sTableNameStripped, 4) = "link" Then
'        bIsLinkedTable = True
'        sTitle = "DD linktable "
'    End If
'
'    If sTableNameStripped = "mastertable" Then
'        bIsMasterTable = True
'        sTitle = "DD mastertable "
'    End If
'
'    With oRS
'
'        .Open "SELECT * FROM [" & sTableName & "] WHERE [" & sFieldName & "] = '" & sValue & "'", mConn, adOpenDynamic, adLockBatchOptimistic
'
'        Do Until .EOF
'            'SynchHistoryDelete mConn, GetGuid, .Fields(1).Value, sTitle & "Delete", g_sRemoteTablePrefix, g_sUserName, CalcServerDateTime(now()), sTableName, False, "false", DDDefCurrent.prefix
'
'            .Delete adAffectCurrent
'            .UpdateBatch adAffectCurrent
'
'            .MoveNext
'        Loop
'
'    End With
'
'    Set oRS = Nothing
'
'End Function

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

116     TxMsg ">>> Getting defined queries from [" & sPrefix & "Queries] table ......................", True
118     DefinedDDQueries = GetDefinedQueries(oCN, sPrefix)

120     TxMsg ">>> Getting local tables ......................", True
122     If sUG = "" Then LocalTables = GetLocalTables(oCN, sPrefix)

124     TxMsg ">>> Getting local queries ......................", True
126     LocalQueries = GetLocalQueries(oCN, sPrefix)
        
128     TxMsg ">>> Comparing tables ......................", True
130     If sUG = "" Then CompareTables sWebsite, sPrefix, oCN
        'CompareQueries sWebsite, sPrefix, oCN
    
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

134                 TxMsg "**??Fetching schema of remote table: " & oRSTables.Fields("TABLE_NAME").Value
            
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

Private Function GetLocalTables(oCOnn As ADODB.Connection, _
                                sPrefix As String) As DDTableCollection
        '<EhHeader>
        On Error GoTo GetLocalTables_Err
        '</EhHeader>

        Dim oRS As ADODB.Recordset
        Dim oRSColumns As New ADODB.Recordset
        Dim ddDatabaseTables As DDTableCollection
100     ReDim ddDatabaseTables.Tables(0)
        
        Dim adoxTable As New adox.Table
        Dim adoxCatalog As New adox.Catalog
102     adoxCatalog.ActiveConnection = oCOnn
104     oRSColumns.Open oCOnn.OpenSchema(adSchemaColumns)
        
106     For Each adoxTable In adoxCatalog.Tables
    
108         If adoxTable.Type = "TABLE" And Left(adoxTable.Name, Len(sPrefix)) = sPrefix Then
            
110             ReDim Preserve ddDatabaseTables.Tables(UBound(ddDatabaseTables.Tables) + 1)
      
112             With ddDatabaseTables.Tables(UBound(ddDatabaseTables.Tables))
            
114                 .Name = adoxTable.Name
116                 .IsTable = True
118                 .IsLinkedTable = False
120                 .IsQuery = False
                    '.sAutoNum = "none"
122                 TxMsg "**??Fetching schema of local table: " & .Name
                    
124                 ReDim .Columns(0)

126                 oRSColumns.Filter = adFilterNone
128                 'If Not oRSColumns.EOF Or Not oRSColumns.BOF Then oRSColumns.MoveFirst
130                 oRSColumns.Filter = "TABLE_NAME = '" & .Name & "'"
            
                    If bSQLServerInUse Then Set oRS = New ADODB.Recordset
             
132                 Do Until oRSColumns.EOF

                        If Not oRSColumns.Fields("COLUMN_NAME").Value = "WKB" And Not Left(oRSColumns.Fields("COLUMN_NAME").Value, 1) = "_" Then
                        
134                         ReDim Preserve .Columns(UBound(.Columns) + 1)
136                         .Columns(UBound(.Columns)).Name = oRSColumns.Fields("COLUMN_NAME").Value
138                         .Columns(UBound(.Columns)).Ordinal = oRSColumns.Fields("ORDINAL_POSITION").Value
140                         .Columns(UBound(.Columns)).TypeEnum = oRSColumns.Fields("DATA_TYPE").Value
                        
142                         If bSQLServerInUse Then
146                             oRS.Open "SELECT [value] FROM ::fn_listExtendedProperty ( 'MS_Description','user', 'dbo', 'table', '" & .Name & "', 'column', '" & oRSColumns.Fields("COLUMN_NAME").Value & "')", oCOnn, adOpenDynamic, adLockBatchOptimistic

148                             If Not oRS.EOF Then .Columns(UBound(.Columns)).Description = oRS.Fields("value").Value
150                             oRS.Close
                            ElseIf Not IsNull(oRSColumns.Fields("DESCRIPTION").Value) Then
154                             .Columns(UBound(.Columns)).Description = oRSColumns.Fields("DESCRIPTION").Value
                            End If

156                         If Not IsNull(oRSColumns.Fields("CHARACTER_MAXIMUM_LENGTH").Value) Then
158                             .Columns(UBound(.Columns)).MaxLength = oRSColumns.Fields("CHARACTER_MAXIMUM_LENGTH").Value
                            End If
                        
                            'If Not oRSColumns.Fields("ISAUTOINCREMENT") Is Nothing Then
                            '   If oRSColumns.Fields("ISAUTOINCREMENT").Value = True Then
                            '      .sAutoNum = oRSColumns.Fields("COLUMN_NAME").Value
                            ' End If
                            '   End If
                        End If

160                     oRSColumns.MoveNext
                    Loop
                    
                    If bSQLServerInUse Then Set oRS = Nothing
            
                End With
        
            End If

        Next

162     Set adoxCatalog = Nothing
164     Set adoxTable = Nothing
166     Set oRSColumns = Nothing
168     GetLocalTables = ddDatabaseTables
        'TxMsg "SynchWorker.GetLocalTables_Err on line (" & Erl & ") " & Err.Description, True
        '<EhFooter>
        Exit Function

GetLocalTables_Err:
        TxMsg "SynchWorker.GetLocalTables_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Function FixLocalQueriesExtents(oCOnn As ADODB.Connection, _
                                        sPrefix As String) As DDQueryCollection

    Dim oRSQueries As New ADODB.Recordset
    
    If Not DoesTableExist(oCOnn, "ttkGISLayerSQL") Then
            
        oCOnn.Execute "CREATE TABLE [ttkGISLayerSQL]  ([NAME] TEXT(255), [XMIN] DOUBLE, [XMAX] DOUBLE, [YMIN] DOUBLE, [YMAX] DOUBLE, [SHAPETYPE] INTEGER)"
        TxMsg "   --- Creating table [ttkGISLayerSQL]"

    End If

    With oRSQueries

        .Open oCOnn.OpenSchema(adSchemaViews)
        .Filter = "TABLE_NAME LIKE '" & sPrefix & "%'"

        If Not .BOF Or Not .EOF Then .MoveFirst

        Do Until .EOF

            If Right(oRSQueries.Fields("TABLE_NAME").Value, 4) = "_FEA" Then UpdatettkGISLaySQL oRSQueries.Fields("TABLE_NAME").Value, oCOnn
            .MoveNext
        Loop

    End With

    Set oRSQueries = Nothing

End Function


Private Function GetLocalQueries(oCOnn As ADODB.Connection, _
                                 sPrefix As String) As DDQueryCollection
        '<EhHeader>
        On Error GoTo GetLocalQueries_Err
        '</EhHeader>

        Dim oRSQueries As New ADODB.Recordset
        Dim ddQueries As DDQueryCollection

100     ddQueries.Prefix = sPrefix
102     ReDim ddQueries.Queries(0)

104     With oRSQueries

106         .Open oCOnn.OpenSchema(adSchemaViews)
108         .Filter = "TABLE_NAME LIKE '" & sPrefix & "%'"

110         If Not .BOF Or Not .EOF Then .MoveFirst

112         Do Until .EOF

114             ReDim Preserve ddQueries.Queries(UBound(ddQueries.Queries) + 1)

116             'TxMsg ">>> Query [" & oRSQueries.Fields("TABLE_NAME").Value & "] located in array at index: " & UBound(ddQueries.Queries), True

118             ddQueries.Queries(UBound(ddQueries.Queries)).Name = oRSQueries.Fields("TABLE_NAME").Value

'120             If Not IsNull(oRSQueries.Fields("VIEW_DEFINITION").Value) Then
'122                 ddQueries.Queries(UBound(ddQueries.Queries)).sSQL = Trim(oRSQueries.Fields("VIEW_DEFINITION").Value)
'                Else
'124                 TxMsg "!!! WARNING..... Query [" & oRSQueries.Fields("TABLE_NAME").Value & "] definition is NULL!!!", True
'                End If

126             .MoveNext

            Loop

        End With

128     Set oRSQueries = Nothing
130     GetLocalQueries = ddQueries

        '<EhFooter>
        Exit Function

GetLocalQueries_Err:
        TxMsg "SynchWorker.GetLocalQueries_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

Private Function CompareQueries(sWebsite As String, _
                                sPrefix As String, _
                                oCOnn As ADODB.Connection) As Boolean
        '<EhHeader>
        On Error GoTo CompareQueries_Err
        '</EhHeader>

        On Error Resume Next
        Dim oRS As ADODB.Recordset
        Dim i As Long
        Dim iLoops As Long
        Dim bError As Boolean
        Dim bRetVal As Boolean
        
        'TxMsg "**??Creating queries/views in the client database...", True
        
100     iLoops = 1
102     bError = True
        
104     Do Until iLoops > 5 Or bError = False

106         If iLoops = 5 Then On Error GoTo CompareQueries_Err
108         bError = False
            
            i = 1
110         If Not iLoops = 1 Then TxMsg ">>> Attempting loop #" & iLoops & " of 5 for query creation...............................................", True

114         Do Until i > UBound(DefinedDDQueries.Queries) Or UBound(DefinedDDQueries.Queries) = 0 Or UBound(DefinedDDQueries.Queries) = -1

                
                
116             With DefinedDDQueries.Queries(i)

118                 bRetVal = True
                    
                    If (.bNew Or .bEdited Or .bDelete) And iLoops = 1 Then TxMsg ">>> Attempting loop #" & iLoops & " of 5 for query creation...............................................", True
                    
120                 If .bNew Then

122                     bRetVal = CreateQuery(.Name, .sSQL, oCOnn, sWebsite)

124                 ElseIf .bEdited Then
126                     bRetVal = DropLocalQuery(.Name, oCOnn)
128                     bRetVal = CreateQuery(.Name, .sSQL, oCOnn, sWebsite)
130                 ElseIf .bDelete Then
132                     bRetVal = DropLocalQuery(.Name, oCOnn)
                    End If
                    
134                 If bError = False Then bError = Not bRetVal
                
136                 If bRetVal Then
138                     DefinedDDQueries.Queries(i).bDelete = False
140                     DefinedDDQueries.Queries(i).bNew = False
142                     DefinedDDQueries.Queries(i).bEdited = False
                    End If

                End With

144             i = i + 1
        
            Loop
  
148         If bError Then

                TxMsg ">>> Loop #" & iLoops & " reached the midway point...............................................", True
150             i = UBound(DefinedDDQueries.Queries)

152             Do Until i = 0 Or UBound(DefinedDDQueries.Queries) = 0 Or UBound(DefinedDDQueries.Queries) = -1

154                 With DefinedDDQueries.Queries(i)

156                     bRetVal = True

158                     If .bNew Then
160                         bRetVal = CreateQuery(.Name, .sSQL, oCOnn, sWebsite)
162                     ElseIf .bEdited Then
164                         bRetVal = DropLocalQuery(.Name, oCOnn)
166                         bRetVal = CreateQuery(.Name, .sSQL, oCOnn, sWebsite)
168                     ElseIf .bDelete Then
170                         bRetVal = DropLocalQuery(.Name, oCOnn)
                        End If
                    
172                     If bError = False Then bError = Not bRetVal
                    
174                     If iLoops = 5 And Not Len(.sSQL) > 0 Then bRetVal = False
                    
176                     If bRetVal Then

178                         DefinedDDQueries.Queries(i).bDelete = False
180                         DefinedDDQueries.Queries(i).bNew = False
182                         DefinedDDQueries.Queries(i).bEdited = False

184                     ElseIf iLoops = 5 Then
                    
186                         Set oRS = New ADODB.Recordset
188                         oRS.Open "SELECT [GUID1] from [" & sPrefix & "Queries] WHERE [QueryName] = '" & Replace(.Name, sPrefix, "") & "'", oCOnn, adOpenDynamic, adLockBatchOptimistic

190                         If Not oRS.EOF Then
192                             oCOnn.Execute "DELETE FROM [" & sPrefix & "SynchHistory] WHERE [sTableName] = '" & sPrefix & "Queries' AND [sID] = '" & oRS.Fields(0).Value & "'"
                                oCOnn.Execute "DELETE FROM [" & sPrefix & "SynchHistoryOverview] WHERE [sTableName] = '" & sPrefix & "Queries'"
                                'oCOnn.Execute "DELETE FROM [" & sPrefix & "Queries] WHERE QueryName = '" & .Name & "'"
194                             oRS.Delete adAffectCurrent
196                             oRS.UpdateBatch
198                             oRS.Close
                            End If

200                         Set oRS = Nothing
                        
                        End If

                    End With

202                 i = i - 1
        
                Loop
        
            End If
            
204         iLoops = iLoops + 1
        Loop
        
206     CompareQueries = Not bError
        'TxMsg "**??Completed creating queries/views in the client database...", True

        '<EhFooter>
        Exit Function

CompareQueries_Err:
        TxMsg "SynchWorker.CompareQueries_Err on line (" & Erl & ") " & Err.Description, True
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
        
112         If iRemoteTableIndex = 0 Then
114             TxMsg "   !!! Table [" & sTableName & "] has been removed from the server - dropping local copy and removing from SynchHistory", True
                DropLocalTable sTableName, oCOnn, sPrefix
'116             oCOnn.Execute "DROP TABLE [" & sTableName & "]"
118             oCOnn.Execute "DELETE FROM [" & sPrefix & "SynchHistory] WHERE sTableName = '" & sTableName & "'"
                oCOnn.Execute "DELETE FROM [" & sPrefix & "SynchHistoryOverview] WHERE sTableName = '" & sTableName & "'"
            
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
        
136         If Not iLocalTableIndex = 0 Then
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
            
216             bAbortSynchronisation = True
218             sSQL = "SELECT TOP 1 * FROM [" & RemoteTables.Tables(iRemoteTableIndex).Name & "]"
                'RemoteRS.Open sSQL, , adOpenDynamic, adLockBatchOptimistic
220             Set RemoteRS = OpenServerRSCompressed(sWebsite, "ID", sSQL)
222             CreateTable RemoteTables.Tables(iRemoteTableIndex).Name, RemoteRS, oCOnn, sWebsite
224             RemoteRS.Close
            
            End If
        
226         iRemoteTableIndex = iRemoteTableIndex + 1
    
        Loop
        
228    TxMsg ">>>>> Currently checking tables", False, True
        
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
                
248             If Not Right$(sTableNameAltered, 4) = "_GEO" And Not Right$(sTableNameAltered, 12) = "SynchHistory" And Not Right$(sTableNameAltered, 20) = "SynchHistoryOverview" Then

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
            oConnection.Execute "DELETE FROM [" & sPrefix & "SynchHistoryOverview] WHERE sTableName = '" & sTableName & "'"
        
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
                             oCOnn As ADODB.Connection, sWebsite As String) As Boolean
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

136         If Right$(sName, 4) = "_FEA" Then
138             CreateTTKGPReference sName, oCOnn, sWebsite
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

Private Function CloneRS(oRS As ADODB.Recordset, Optional sExcludedField As String = "") As ADODB.Recordset
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

Private Function LocateIndexOfQuery(oQuery As DDQueryCollection, _
                                    sItem As String) As Long
        '<EhHeader>
        On Error GoTo LocateIndexOfQuery_Err
        '</EhHeader>

        Dim i As Long
        LocateIndexOfQuery = 0
100     i = 0
     
102     Do Until i > UBound(oQuery.Queries) Or UBound(oQuery.Queries) = 0
    
104         If oQuery.Queries(i).Name = sItem Then
106             LocateIndexOfQuery = i
                Exit Do
            End If
        
108         i = i + 1
        Loop
        
        'TxMsg ">>>      Located query [" & sItem & "] in array at index: " & LocateIndexOfQuery, True
        '<EhFooter>
        Exit Function

LocateIndexOfQuery_Err:
        TxMsg "!!!!!! Error locating query [" & sItem & "] in array, Error message: " & Err.Description, True
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
    Dim Rs As New ADODB.Recordset
    
    On Error GoTo hell
    Rs.Open "SELECT * FROM " & sTable, cn, adOpenForwardOnly, adLockReadOnly

    DoesTableExist = True
        
    On Error Resume Next
    Rs.Close
    Set Rs = Nothing
        
    Exit Function
hell:
    On Error Resume Next
    Err.Clear

    DoesTableExist = False
    Set Rs = Nothing
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

Private Function GetAllFieldNamesForLocalTable(sTableName As String, _
                                               cn As ADODB.Connection) As String
        '<EhHeader>
        On Error GoTo GetAllFieldNamesForLocalTable_Err
        '</EhHeader>

        Dim RSLocal As ADODB.Recordset
100     Set RSLocal = New ADODB.Recordset
        Dim iCol As Long
        Dim sSQL As String
102     GetAllFieldNamesForLocalTable = ""
    
104     With RSLocal

106         sSQL = "SELECT top 1 * FROM [" & sTableName & "]"
108         .Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
    
110         If .State = adStateOpen Then
        
112             iCol = 0

114             Do Until iCol = .Fields.Count
             
116                 If Not .Fields(iCol).Name = "WKB" And Not .Fields(iCol).Name = "ORIG_ID" And Not Left(.Fields(iCol).Name, 1) = "_" Then
118                     If Not GetAllFieldNamesForLocalTable = "" Then GetAllFieldNamesForLocalTable = GetAllFieldNamesForLocalTable & ","
120                     GetAllFieldNamesForLocalTable = GetAllFieldNamesForLocalTable & "[" & .Fields(iCol).Name & "]"
                    End If
                
122                 iCol = iCol + 1
                
                Loop
            
124             .Close
        
            End If
        
        End With

126     Set RSLocal = Nothing

        '<EhFooter>
        Exit Function

GetAllFieldNamesForLocalTable_Err:
        TxMsg "SynchWorker.GetAllFieldNamesForLocalTable_Err on line (" & Erl & ") " & Err.Description, True
        '</EhFooter>
End Function

























