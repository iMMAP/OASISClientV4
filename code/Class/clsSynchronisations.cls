VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSynchronisations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'
'   - Busy                  returns free/busy flag
'   - Cancel                cancels asyncronious operation
'   - ChannelCount          returns channel count
'   - Channels              returns collection of channels
'   - Clear                 clears all data
'   - ClearLastError        clears last error
'   - GetItemByPath         returns either channel or item object (path is n/n)
'   - LastErrDescription    last error description
'   - LastErrNumber         last error number
'   - LastErrSource         last error source
'   - LoadAllChannels       load all channels that have URLs in locations collection
'   - LoadChannel           load channel from location or XML node
'   - LoadFromCache         loads all channels from local cache
'   - Locations             returns collection of channel locations
'   - NewLocations          returns collection of all channel locations that
'                           have not been downloaded
'   - RemoveChannel         removes channel
'   - Save                  saves channel as several XML formats
'                           Note that RSS 1.0 format is not valid RSS 1.0 because
'                           it actually saves multiple channels in one document
'                           <rdf:RDF>
'                               <channel>...</channel>
'                               <channel>...</channel>
'                               ...
'                           </rdf:RDF>
'   - SaveAsOCS             saves channel as OCS - Open Content Syndication format
'   - Serialize             returns XML DOM of all channels
'   - SetLastError          sets last error

Option Explicit

Public Enum SaveFormat
    sfRSS = 1 ' RSS 1.0, http://groups.yahoo.com/group/rss-dev/files/namespace.html
    sfOCS = 2 ' Open Content Syndication, http://internetalchemy.org/ocs/index.html
End Enum

Public Event Progress(ByVal lProgress As Long, ByVal lMax As Long, ByVal sURL As String)
Private moLocations As Collection
Private moChannels As Collection
Private mbBusy As Boolean
Private mlLastErrNum As Long
Private msLastErrDescr As String
Private msLastErrSource As String

Public Function LoadFromCache(ByVal sPath As String) As Boolean
    
    Dim oXML As MSXML2.DOMDocument
    Dim oNodes As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMAttribute
    Dim oChannel As clsSynchChannel
    Dim sLocation As String
    
    Set oXML = New MSXML2.DOMDocument
    LoadFromCache = oXML.Load(sPath)
    If LoadFromCache Then
        Set oNodes = oXML.selectNodes("rdf:RDF/channel")
        If Not oNodes Is Nothing Then
            On Error Resume Next
            For Each oNode In oNodes
                Set oAttr = oNode.selectSingleNode("@rdf:about")
                If Not oAttr Is Nothing Then
                    sLocation = oAttr.Text
                    Set oChannel = New clsSynchChannel
                    If oChannel.ParseChannel(oNode, sLocation, moChannels.Count + 1) Then
                        moChannels.Add oChannel, sLocation
                        On Error Resume Next
                        moLocations.Add sLocation, sLocation
                    End If
                End If
            Next oNode
        End If
    Else
        Clear
        SetLastError -1, "Cannot load news from cache" & vbCrLf & sPath & vbCrLf & _
                    oXML.parseError.reason, "News.LoadFromCache"
    End If
    
End Function

Public Sub ParseSynchXML(sXML As String)
    Dim oSynchXML As New MSXML2.DOMDocument
    Dim oAttr As MSXML2.IXMLDOMAttribute
    Dim oNodes As MSXML2.IXMLDOMNodeList
    Dim oElement As MSXML2.IXMLDOMElement
    
    sXML = "<?xml version=""1.0"" encoding=""UTF-16""?>" & vbCrLf & sXML
    sXML = Replace$(sXML, """", """" & """")
    
    m_frmDebug.DebugPrint oSynchXML.loadXML(sXML)
    
    Set oNodes = oSynchXML.selectNodes("//sync")
    
    For Each oAttr In oNodes
        m_frmDebug.DebugPrint oAttr.baseName
    Next
    
    For Each oElement In oNodes
        
    Next
    
End Sub


Public Function LoadFromCacheEx(ByVal sPath As String) As Boolean
    
    Dim oXML As MSXML2.DOMDocument
    Dim oNodes As MSXML2.IXMLDOMNodeList
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oAttr As MSXML2.IXMLDOMAttribute
    Dim oElement As MSXML2.IXMLDOMElement
    
    Dim oSynchNode As IXMLDOMElement
    
    Dim oChannel As clsSynchChannel
    Dim sLocation As String
    
    Set oXML = New MSXML2.DOMDocument
    LoadFromCacheEx = oXML.Load(sPath)
    If LoadFromCacheEx Then
        'Get The RSS Feed Items
        Set oNodes = oXML.selectNodes("//channel/item")
        If Not oNodes Is Nothing Then
            On Error Resume Next
            
            For Each oNode In oNodes
                'Loop Through the Items sub Elements
                For Each oElement In oNode.childNodes
                    m_frmDebug.DebugPrint oElement.nodeValue
                    m_frmDebug.DebugPrint oElement.xml
                    m_frmDebug.DebugPrint oElement.baseName
                    m_frmDebug.DebugPrint oElement.nodeName
                    
                    If LCase$(oElement.nodeName) = "sx:sync" Then
                        'get the Synch Details
                        m_frmDebug.DebugPrint oElement.attributes.getNamedItem("id").Text
                        m_frmDebug.DebugPrint oElement.attributes.getNamedItem("updates").Text
                    
                        For Each oSynchNode In oElement.childNodes
                            'Get the History
                            m_frmDebug.DebugPrint oSynchNode.baseName
                            m_frmDebug.DebugPrint oSynchNode.attributes.getNamedItem("when").Text
                            m_frmDebug.DebugPrint oSynchNode.attributes.getNamedItem("sequence").Text
                            m_frmDebug.DebugPrint oSynchNode.attributes.getNamedItem("by").Text
                        Next oSynchNode
                        
                        'ParseSynchXML oElement.Xml
                        
                        
                    End If
                    
                    m_frmDebug.DebugPrint oElement.nodeType
                    m_frmDebug.DebugPrint oElement.tagName
                    m_frmDebug.DebugPrint oElement.Text
                    
                Next
                
                'Set oAttr = oNode.nextSibling.cl
            
                'Do While Not oAttr Is Nothing
                
                 '   m_frmDebug.DebugPrint oNode
                'Loop
            
                Set oAttr = oNode.selectSingleNode("title")
                
                
                m_frmDebug.DebugPrint ""
                
                
                If Not oAttr Is Nothing Then
                    sLocation = oAttr.Text
                    Set oChannel = New clsSynchChannel
                    If oChannel.ParseChannel(oNode, sLocation, moChannels.Count + 1) Then
                        moChannels.Add oChannel, sLocation
                        On Error Resume Next
                        moLocations.Add sLocation, sLocation
                    End If
                End If
            Next oNode
        End If
    Else
        Clear
        SetLastError -1, "Cannot load news from cache" & vbCrLf & sPath & vbCrLf & _
                    oXML.parseError.reason, "News.LoadFromCache"
    End If
    
End Function


Public Function LoadChannel(ByVal sLocation As String, ByVal lChannelIndex As Long, _
                Optional ByVal oNode As MSXML2.IXMLDOMNode = Nothing) As Boolean
    
    Dim oChannel As clsSynchChannel
    
On Error GoTo errHandler
    
    mbBusy = True
    gbCancelled = False
    ClearLastError
    
    Set oChannel = New clsSynchChannel
    LoadChannel = oChannel.Load(sLocation, lChannelIndex, oNode)
    If LoadChannel Then
        moChannels.Add oChannel, sLocation
        On Error Resume Next
        moLocations.Add sLocation, sLocation
        On Error GoTo errHandler
    End If
    
    mbBusy = False
    
    Exit Function
errHandler:
    mbBusy = False
    SetLastError Err.Number, Err.Description, Err.Source
    Err.Clear
End Function

Public Function LoadAllChannels() As Boolean
    
    Dim idx As Long
    Dim sLocation As String
    
On Error Resume Next
    Clear
    ClearLastError
    gbCancelled = False
    
    For idx = 1 To moLocations.Count
        sLocation = moLocations(idx)
        RaiseEvent Progress(idx, moLocations.Count, sLocation)
        If Not LoadChannel(sLocation, moChannels.Count + 1) Then LoadAllChannels = False
        If gbCancelled Then
            gbCancelled = False
            Exit For
        End If
    Next idx
    LoadAllChannels = True
    
End Function

Public Function LoadNewChannels() As Boolean
    
    Dim idx As Long
    Dim sLocation As String
    Dim oChannel As clsSynchChannel
    Dim lChannelCount As Long
    
On Error Resume Next
    ClearLastError
    gbCancelled = False
    '
    ' Go thru loaded channels to see if they exist in
    ' locations collection, if they don't that means they
    ' need to be removed
    lChannelCount = moChannels.Count
    For Each oChannel In moChannels
        Err.Clear
        sLocation = moLocations(oChannel.Location)
        If Err.Number <> 0 Then moChannels.Remove oChannel.Location
    Next oChannel
    '
    ' Go thru locations collection, check if they already loaded
    ' if they're not then load them
    LoadNewChannels = True
    For idx = 1 To moLocations.Count
        sLocation = moLocations(idx)
        Set oChannel = moChannels(sLocation)
        If oChannel Is Nothing Then
            RaiseEvent Progress(idx, moLocations.Count, sLocation)
            If Not LoadChannel(sLocation, lChannelCount + 1) Then _
                LoadNewChannels = False
            If gbCancelled Then
                gbCancelled = False
                Exit For
            End If
            lChannelCount = lChannelCount + 1
        End If
        Set oChannel = Nothing
    Next idx
    
End Function

Public Sub RemoveChannel(ByVal sLocation As String)
    
    On Error Resume Next
    moChannels.Remove sLocation
    moLocations.Remove sLocation
    
End Sub

Public Sub Cancel()
    gbCancelled = True
End Sub

Public Property Get Busy() As Boolean
    Busy = mbBusy
End Property

Public Property Get Channels() As Collection
    
    Set Channels = moChannels
    
End Property

Public Property Get NewLocations() As Collection
    
    Dim v As Variant
    Dim sLocation As String
    Dim col As Collection
    Dim oChannel As clsSynchChannel
    Dim bFound As Boolean
    
    Set col = New Collection
    For Each v In moLocations
        sLocation = CStr(v)
        bFound = False
        For Each oChannel In moChannels
            If oChannel.Location = sLocation Then
                bFound = True
                Exit For
            End If
        Next oChannel
        If Not bFound Then col.Add sLocation, sLocation
    Next v
    Set NewLocations = col
    
End Property

Public Property Get Locations() As Collection
    Set Locations = moLocations
End Property

Public Property Set Locations(ByVal NewVal As Collection)
    Set moLocations = NewVal
End Property

Public Property Get ChannelCount() As Long
    ChannelCount = moChannels.Count
End Property

Public Function GetItemByPath(ByVal sPath As String) As Object
    
    Dim oChannel As clsSynchChannel
    Dim oItem As clsSynchItem
    Dim ar As Variant
    Dim lChannel As Long
    Dim lItem As Long
    
    ar = Split(sPath, "/")
    If IsArray(ar) Then
        lChannel = ar(0)
        For Each oChannel In moChannels
            If oChannel.Index = lChannel Then
                Exit For
            End If
        Next oChannel
        If Not oChannel Is Nothing Then
            If UBound(ar) = 1 Then
                If Len(ar(1)) > 0 Then
                    lItem = ar(1)
                    For Each oItem In oChannel.Items
                        If oItem.Index = lItem Then
                            Set GetItemByPath = oItem
                            Exit For
                        End If
                    Next oItem
                Else
                    Set GetItemByPath = oChannel
                End If
            Else
                Set GetItemByPath = oChannel
            End If
        End If
    End If
    
End Function

Public Sub Save(ByVal sPath As String, Optional ByVal lFormat As SaveFormat = sfRSS)
    Dim oXML As MSXML2.DOMDocument
    
    Select Case lFormat
        Case sfRSS
            Set oXML = Serialize
            oXML.Save sPath
        Case sfOCS
            SaveAsOCS sPath
    End Select
End Sub

Private Function SaveAsOCS(ByVal sFileName As String) As Boolean
    
    Dim oXML As MSXML2.DOMDocument
    Dim oXSL As MSXML2.DOMDocument
    Dim oOut As MSXML2.DOMDocument
    
    Set oXSL = New MSXML2.DOMDocument
    If oXSL.Load(App.Path & XSL_FOLDER & "rsstoocs.xsl") Then
        Set oXML = Serialize
        Set oOut = New MSXML2.DOMDocument
        oXML.transformNodeToObject oXSL, oOut
        If Not oOut Is Nothing Then
            oOut.Save sFileName
            SaveAsOCS = True
        End If
    End If
    
End Function

Public Sub Clear()
    If Not moChannels Is Nothing Then
        While moChannels.Count > 0
            moChannels.Remove 1
        Wend
    Else
        Set moChannels = New Collection
    End If
End Sub

Public Function Serialize() As MSXML2.DOMDocument
    
    Dim oXML As MSXML2.DOMDocument
    Dim oNode As MSXML2.IXMLDOMNode
    Dim oChannel As clsSynchChannel
    
    Set oXML = New MSXML2.DOMDocument
    Set oNode = oXML.createNode(NODE_ELEMENT, "rdf:RDF", RDF_SCHEMA)
    oXML.appendChild oNode
    
    For Each oChannel In moChannels
        oChannel.Serialize oXML
    Next oChannel
    
    Set Serialize = oXML
    
End Function

Public Property Get LastErrNumber() As Long
    LastErrNumber = mlLastErrNum
End Property

Public Property Get LastErrDescription() As String
    LastErrDescription = msLastErrDescr
End Property

Public Property Get LastErrSource() As String
    LastErrSource = msLastErrSource
End Property

Private Sub SetLastError(ByVal lErr As String, ByVal sErr As String, ByVal sSrc As String)
    mlLastErrNum = lErr
    msLastErrDescr = sErr
    msLastErrSource = sSrc
End Sub

Private Sub ClearLastError()
    mlLastErrNum = 0
    msLastErrDescr = ""
    msLastErrSource = ""
End Sub

Private Sub Class_Initialize()
    
    Clear
    Set moLocations = New Collection
    Set moChannels = New Collection
    
End Sub

Private Sub Class_Terminate()
    Set moLocations = Nothing
    Set moChannels = Nothing
End Sub


