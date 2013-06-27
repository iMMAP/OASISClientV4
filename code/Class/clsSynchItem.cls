VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSynchItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'   - ChangeDate    - returns/sets last change date of the news item
'   - Description   - returns/sets news item description
'   - Index         - returns index of this item in the channel
'   - ParseItem     - parses XML node to extract news item properties
'   - Path          - returns path to this item like 1/2 or 4/1
'                     the path is always unique and can be used to get
'                     news item thru clsNews.GetItemByPath() function.
'   - Serialize     - generates XML DOM news item
'   - Title         - returns/sets title of the news items
'   - URL           - returns/sets URL to resource linked by the news item

Option Explicit

Private msTitle As String
Private msURL As String
Private msDescr As String

'New Optional
Private msauthor As String
Private mscategory As String
Private mscomments As String
Private msenclosure As String
Private msguid As String
Private msPubDate As String
Private mssource As String


Private mlChannelIndex As Long
Private mlIndex As Long
Private mdtChangeDate As Date
Private moSequences As Collection

Public Property Get Count() As Long
    Count = moSequences.Count
End Property

Public Property Get ForEach() As IUnknown
    Set ForEach = moSequences.[_NewEnum]
End Property

Public Property Get Items() As Collection
    Set Items = moSequences
End Property

Public Function ParseItem(ByVal oNode As MSXML2.IXMLDOMNode, _
                    ByVal lChannelIndex As Long, _
                    ByVal lItemIndex As Long) As Boolean
    
    Dim oItem As MSXML2.IXMLDOMNode
    
    Set oItem = oNode.selectSingleNode("title")
    If Not oItem Is Nothing Then msTitle = oItem.Text
    Set oItem = oNode.selectSingleNode("link")
    If Not oItem Is Nothing Then msURL = oItem.Text
    Set oItem = oNode.selectSingleNode("description")
    If Not oItem Is Nothing Then msDescr = oItem.Text
    mlChannelIndex = lChannelIndex
    mlIndex = lItemIndex
    
End Function

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

Public Property Get Index() As Long
    Index = mlIndex
End Property

Public Property Get Path() As String
    Path = mlChannelIndex & "/" & mlIndex
End Property

Public Property Get URL() As String
    URL = msURL
End Property
Public Property Let URL(ByVal NewVal As String)
    msURL = NewVal
End Property

Public Property Get ChangeDate() As String
    ChangeDate = mdtChangeDate
End Property
Public Property Let ChangeDate(ByVal NewVal As String)
    mdtChangeDate = NewVal
End Property

Public Sub Serialize(ByRef oChannel As MSXML2.IXMLDOMNode)
    'Dim oItem As clsSynchSequence
    Dim oItem As MSXML2.IXMLDOMElement
    Dim oNode As MSXML2.IXMLDOMNode
    
    Set oItem = oChannel.ownerDocument.createNode(NODE_ELEMENT, "item", RSS_SCHEMA)
    
    Set oNode = oChannel.ownerDocument.createNode(NODE_ELEMENT, "title", RSS_SCHEMA)
    oNode.Text = msTitle
    oItem.appendChild oNode
    Set oNode = oChannel.ownerDocument.createNode(NODE_ELEMENT, "link", RSS_SCHEMA)
    oNode.Text = msURL
    oItem.appendChild oNode
    Set oNode = oChannel.ownerDocument.createNode(NODE_ELEMENT, "link", RSS_SCHEMA)
    oNode.Text = msURL
    oItem.appendChild oNode
    Set oNode = oChannel.ownerDocument.createNode(NODE_ELEMENT, "description", RSS_SCHEMA)
    oNode.Text = msDescr
    oItem.appendChild oNode
    
    oChannel.appendChild oItem
    
    Dim osItem As Object
    
    For Each osItem In moSequences
        osItem.Serialize oChannel
    Next osItem
    
End Sub

Private Sub Class_Initialize()
    Set moSequences = New Collection
End Sub

Private Sub Class_Terminate()
    Set moSequences = Nothing
End Sub

