VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSynchChannel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'   - AddNewsItem   - adds news item to the channel
'   - ChangeDate    - returns/sets last change date of the channel
'   - Count         - returns number of news items in the channel
'   - Description   - returns/sets description of the channel
'   - ForEach       - enumerates news items
'   - Index         - channel index
'   - Items         - returns collection of news items
'   - Link          - returns/sets link property of the channel
'   - Load          - loads channel from location or XML node
'                     It can process RSS, scriptingNews and moreoverNews
'                     formats, it has not been tested throughly and
'                     might not work with all news sources.
'   - Location      - returns/sets physical location of the channel file
'   - ParseChannel  - parses XML channel node
'   - Path          - returns path to this channel like 2/ , 3/
'                     the path is always unique and can be used to get
'                     news channel thru clsNews.GetItemByPath() function.
'   - Serialize     - generates XML DOM of the channel
'   - Title         - returns/sets channel title

Option Explicit

Private moItems As Collection
Private msTitle As String
Private msLink As String
Private msDescr As String
'New Non Optional
Private msLanguage As String
Private msCopyRight As String
Private msmanagingEditor As String
Private msWebMaster As String
Private msPubDate As String
Private mslastBuildDate As String
Private mscategory As String
Private msgenerator As String
Private msdocs As String
Private mscloud As String
Private msttl As String
Private msimage As String
Private msrating As String
Private mstextInput As String
Private msskipHours As String
Private msskipDays As String


Private mlIndex As Long
Private msChangeDate As String
Private msLocation As String
Private mbCancelled As Boolean

Public Function Load(ByVal sLocation As String, ByVal lChannelIndex As Long, _
                Optional ByVal oNode As MSXML2.IXMLDOMNode = Nothing) As Boolean
    
    Dim oDoc As MSXML2.DOMDocument
    Dim oDocOut As MSXML2.DOMDocument
    Dim oXSL As MSXML2.DOMDocument
    Dim oNodes As MSXML2.IXMLDOMNodeList
    
    gbCancelled = False
    
    If oNode Is Nothing Then
        Set oDoc = New MSXML2.DOMDocument
        oDoc.validateOnParse = False
        oDoc.async = True
        If oDoc.Load(sLocation) Then
            Do While Not oDoc.readyState = 4
                DoEvents
                If gbCancelled Then Exit Do
            Loop
            With oDoc.parseError
                If .errorCode <> 0 Then
                    Err.Raise .errorCode, "SynchChannel.Load", _
                            "Failed to load document" & vbCrLf & _
                            .URL & vbCrLf & vbCrLf & .reason & vbCrLf & .srcText
                End If
            End With
            Set oNode = oDoc.documentElement
        End If
    End If
    
    If Not oNode Is Nothing Then
        Set oXSL = New MSXML2.DOMDocument
        Set oDocOut = New MSXML2.DOMDocument
        oXSL.async = False
        Select Case oNode.nodeName
            Case "rss"
                If oXSL.Load(App.Path & XSL_FOLDER & "rss091.xsl") Then _
                    oDoc.transformNodeToObject oXSL, oDocOut
            Case "scriptingNews"
                If oXSL.Load(App.Path & XSL_FOLDER & "scriptingNews2.xsl") Then _
                    oDoc.transformNodeToObject oXSL, oDocOut
            Case "rdf:RDF"
                If oXSL.Load(App.Path & XSL_FOLDER & "rdf.xsl") Then _
                    oDoc.transformNodeToObject oXSL, oDocOut
            Case "moreovernews"
                If oXSL.Load(App.Path & XSL_FOLDER & "moreover.xsl") Then _
                    oDoc.transformNodeToObject oXSL, oDocOut
        End Select
        With oDocOut.parseError
            If .errorCode <> 0 Then
                Err.Raise .errorCode, "SynchChannel.Load", _
                        "Failed to parse document" & vbCrLf & _
                        .URL & vbCrLf & vbCrLf & .reason & vbCrLf & .srcText
            End If
        End With
        Set oNode = oDocOut.documentElement
        Set oNodes = oNode.selectNodes("channel")
        If Not oNodes Is Nothing Then
            For Each oNode In oNodes
                Load = ParseChannel(oNode, sLocation, lChannelIndex)
            Next oNode
        Else
            Err.Raise -1, "SynchChannel.Load", _
                    "The document does not contain any channels." & _
                    vbCrLf & sLocation
        End If
    Else
        Err.Raise -1, "SynchChannel.Load", "The document is empty." & _
                vbCrLf & sLocation
    End If
    
End Function

Public Function ParseChannel(ByVal oNode As IXMLDOMNode, _
                            ByVal sLocation As String, _
                            ByVal lChannelIndex As Long) As Boolean
    
    Dim oItems As MSXML2.IXMLDOMNodeList
    Dim oItem As IXMLDOMNode
    Dim oEl As IXMLDOMElement
    
On Error GoTo errHandler
    
    ' Add rdf:about attribute to <channel>
    Set oItem = oNode.selectSingleNode("@rdf:about")
    If oItem Is Nothing Then
        Set oItem = oNode.ownerDocument.createNode(NODE_ATTRIBUTE, "rdf:about", RDF_SCHEMA)
        oItem.Text = sLocation
        Set oEl = oNode
        oEl.setAttributeNode oItem
    Else
        oItem.Text = sLocation
    End If
    
    ' Check if we have lastBuildDate, add it if it's not found
    Set oItem = oNode.selectSingleNode("lastBuildDate")
    If oItem Is Nothing Then
        Set oItem = oNode.ownerDocument.createNode(NODE_ELEMENT, "lastBuildDate", RSS_SCHEMA)
        oNode.insertBefore oItem, oNode.selectSingleNode("title")
    ElseIf Len(oItem.Text) = 0 Then
        oItem.Text = Now
    End If
    
    Set oItem = oNode.selectSingleNode("title")
    If Not oItem Is Nothing Then msTitle = oItem.Text
    Set oItem = oNode.selectSingleNode("link")
    If Not oItem Is Nothing Then msLink = oItem.Text
    Set oItem = oNode.selectSingleNode("lastBuildDate")
    If Not oItem Is Nothing Then msChangeDate = oItem.Text
    Set oItem = oNode.selectSingleNode("description")
    If Not oItem Is Nothing Then msDescr = oItem.Text
    msLocation = oNode.selectSingleNode("@rdf:about").Text
    mlIndex = lChannelIndex
    
    Set oItems = oNode.selectNodes("item")
    If Not oItems Is Nothing Then
        For Each oItem In oItems
           ' AddNewsItem oItem
        Next oItem
    End If
    
    ParseChannel = True
    
    Exit Function
    
errHandler:
    ParseChannel = False
End Function

Public Property Get Location() As String
    Location = msLocation
End Property
Public Property Let Location(ByVal NewVal As String)
    msLocation = NewVal
End Property

Public Property Get Title() As String
    Title = msTitle
End Property
Public Property Let Title(ByVal NewVal As String)
    msTitle = NewVal
End Property

Public Property Get Description() As String
    Description = msDescr
End Property
Public Property Let Description(ByVal NewVal As String)
    msTitle = NewVal
End Property

Public Property Get Link() As String
    Link = msLink
End Property
Public Property Let Link(ByVal NewVal As String)
    msLink = NewVal
End Property

Public Property Get Index() As Long
    Index = mlIndex
End Property

Public Property Get Path() As String
    Path = mlIndex & "/"
End Property

Public Property Get ChangeDate() As String
    ChangeDate = msChangeDate
End Property
Public Property Let ChangeDate(ByVal NewVal As String)
    msChangeDate = NewVal
End Property

Public Property Get Count() As Long
    Count = moItems.Count
End Property

Public Property Get ForEach() As IUnknown
    Set ForEach = moItems.[_NewEnum]
End Property

Public Function AddSynchItem(ByVal oNode As MSXML2.IXMLDOMNode) As Boolean
    
    Dim oSynchItm As clsSynchItem
    
    Set oSynchItm = New clsSynchItem
    oSynchItm.ParseItem oNode, mlIndex, moItems.Count + 1
    moItems.Add oSynchItm
    
End Function

Public Sub Serialize(ByRef oXML As MSXML2.DOMDocument)
    
    Dim oItem As clsSynchItem
    Dim oChannel As MSXML2.IXMLDOMElement
    Dim oNode As MSXML2.IXMLDOMNode
    
    Set oChannel = oXML.createNode(NODE_ELEMENT, "channel", RSS_SCHEMA)
    Set oNode = oXML.createNode(NODE_ATTRIBUTE, "rdf:about", RDF_SCHEMA)
    oNode.Text = msLocation
    oChannel.setAttributeNode oNode
    Set oNode = oXML.createNode(NODE_ELEMENT, "lastBuildDate", RSS_SCHEMA)
    oNode.Text = msChangeDate
    oChannel.appendChild oNode
    Set oNode = oXML.createNode(NODE_ELEMENT, "title", RSS_SCHEMA)
    oNode.Text = msTitle
    oChannel.appendChild oNode
    Set oNode = oXML.createNode(NODE_ELEMENT, "link", RSS_SCHEMA)
    oNode.Text = msLink
    oChannel.appendChild oNode
    Set oNode = oXML.createNode(NODE_ELEMENT, "description", RSS_SCHEMA)
    oNode.Text = msDescr
    oChannel.appendChild oNode
    Set oChannel = oXML.documentElement.appendChild(oChannel)
    
    For Each oItem In moItems
        oItem.Serialize oChannel
    Next oItem
    
End Sub

Public Property Get Items() As Collection
    Set Items = moItems
End Property

Private Sub Class_Initialize()
    Set moItems = New Collection
End Sub

Private Sub Class_Terminate()
    Set moItems = Nothing
End Sub


