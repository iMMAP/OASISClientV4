VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmDynamicContent 
   Caption         =   "Dynamic Content"
   ClientHeight    =   7500
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   2895
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrFrequency 
      Interval        =   6000
      Left            =   840
      Top             =   3540
   End
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   7500
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2895
      _cx             =   5106
      _cy             =   13229
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   1
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmDynamicContent.frx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7320
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   2715
         _cx             =   4789
         _cy             =   12912
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   2
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Content|Settings"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   0   'False
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin C1SizerLibCtl.C1Elastic C1ElSettings 
            Height          =   6945
            Left            =   3360
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   330
            Width           =   2625
            _cx             =   4630
            _cy             =   12250
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            Version         =   801
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   3
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmDynamicContent.frx":0037
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame FraGeoRSS 
               Caption         =   "Geo RSS"
               Height          =   5160
               Left            =   90
               TabIndex        =   16
               Top             =   1245
               Width           =   2445
               Begin C1SizerLibCtl.C1Elastic elGEORSS 
                  Height          =   3300
                  Left            =   60
                  TabIndex        =   17
                  TabStop         =   0   'False
                  Top             =   300
                  Width           =   2325
                  _cx             =   4101
                  _cy             =   5821
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Enabled         =   -1  'True
                  Appearance      =   0
                  MousePointer    =   0
                  Version         =   801
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  FloodColor      =   6553600
                  ForeColorDisabled=   -2147483631
                  Caption         =   ""
                  Align           =   0
                  AutoSizeChildren=   0
                  BorderWidth     =   6
                  ChildSpacing    =   4
                  Splitter        =   0   'False
                  FloodDirection  =   0
                  FloodPercent    =   0
                  CaptionPos      =   0
                  WordWrap        =   -1  'True
                  MaxChildSize    =   0
                  MinChildSize    =   0
                  TagWidth        =   0
                  TagPosition     =   0
                  Style           =   0
                  TagSplit        =   2
                  PicturePos      =   4
                  CaptionStyle    =   0
                  ResizeFonts     =   0   'False
                  GridRows        =   0
                  GridCols        =   0
                  Frame           =   1
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   ""
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin VB.CheckBox chkExcludeNon 
                     Caption         =   "Exclude Non Geocoded Feeds"
                     Height          =   375
                     Left            =   120
                     TabIndex        =   23
                     Top             =   570
                     Width           =   2115
                  End
                  Begin VB.CheckBox chkGMLGeoRSS 
                     Caption         =   "GML GeoRSS Format Support"
                     Height          =   495
                     Left            =   120
                     TabIndex        =   22
                     Top             =   1380
                     Width           =   1935
                  End
                  Begin VB.CheckBox chkSimpleGeoRSS 
                     Caption         =   "Simple GeoRSS Format Support"
                     Height          =   375
                     Left            =   120
                     TabIndex        =   21
                     Top             =   975
                     Width           =   1515
                  End
                  Begin VB.CheckBox chkAutoGeocode 
                     Caption         =   "Auto Geocode"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   20
                     Top             =   285
                     Width           =   2115
                  End
                  Begin VB.CheckBox chkDetectGeo 
                     Caption         =   "Detect Geo RSS"
                     Height          =   255
                     Left            =   120
                     TabIndex        =   19
                     Top             =   0
                     Width           =   2115
                  End
                  Begin VB.ComboBox ComCountry 
                     Height          =   315
                     ItemData        =   "frmDynamicContent.frx":0085
                     Left            =   120
                     List            =   "frmDynamicContent.frx":008C
                     Sorted          =   -1  'True
                     Style           =   2  'Dropdown List
                     TabIndex        =   18
                     Top             =   2100
                     Width           =   2235
                  End
                  Begin VB.Label lblFilterCountry 
                     Caption         =   "Filter Country:"
                     Height          =   315
                     Left            =   120
                     TabIndex        =   24
                     Top             =   1860
                     Width           =   975
                  End
               End
            End
            Begin VB.CommandButton cmdApply 
               Caption         =   "Apply"
               Height          =   390
               Left            =   90
               TabIndex        =   14
               Top             =   6465
               Width           =   2445
            End
            Begin VB.Frame FraGeneral 
               Caption         =   "General"
               Height          =   1095
               Left            =   90
               TabIndex        =   10
               Top             =   90
               Width           =   2445
               Begin VB.CheckBox chkAutoUpdate 
                  Caption         =   "Auto Update"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   15
                  Top             =   480
                  Value           =   1  'Checked
                  Width           =   1995
               End
               Begin VB.CheckBox chkAutoSave 
                  Caption         =   "Auto Save Viewed Feeds"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  TabIndex        =   13
                  Top             =   720
                  Value           =   2  'Grayed
                  Visible         =   0   'False
                  Width           =   2175
               End
               Begin VB.TextBox txtSec 
                  Height          =   285
                  Left            =   1620
                  TabIndex        =   12
                  Text            =   "90"
                  Top             =   180
                  Width           =   615
               End
               Begin VB.Label lblUpdateFrequency 
                  Caption         =   "Update Frequency:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   11
                  Top             =   240
                  Width           =   1395
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   6945
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   2625
            _cx             =   4630
            _cy             =   12250
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   -1  'True
            Appearance      =   0
            MousePointer    =   0
            Version         =   801
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   6
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmDynamicContent.frx":009E
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.ListBox lstHeadlines 
               Height          =   2985
               Left            =   90
               TabIndex        =   5
               Top             =   3750
               Width           =   2445
            End
            Begin VB.ComboBox cboCategory 
               Height          =   315
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   345
               Width           =   2445
            End
            Begin VB.ListBox lstFeeds 
               Height          =   2400
               Left            =   90
               TabIndex        =   3
               Top             =   975
               Width           =   2445
            End
            Begin VB.Label lblHeadlines 
               AutoSize        =   -1  'True
               Caption         =   "Feed headlines"
               Height          =   255
               Left            =   90
               TabIndex        =   8
               Top             =   3435
               Width           =   2445
            End
            Begin VB.Label lblCategory 
               AutoSize        =   -1  'True
               Caption         =   "Category(s)"
               Height          =   195
               Left            =   90
               TabIndex        =   7
               Top             =   90
               Width           =   2445
            End
            Begin VB.Label lblSubscribed 
               AutoSize        =   -1  'True
               Caption         =   "Available OASIS Feeds"
               Height          =   195
               Left            =   90
               TabIndex        =   6
               Top             =   720
               Width           =   2445
            End
         End
      End
   End
End
Attribute VB_Name = "frmDynamicContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event CategoryClick()
Public Event FeedsClick(sURL As String, bUseGeo As Boolean, oItems As MSXML2.IXMLDOMNodeList, sCountryISO As String)
Public Event HeadlinesClick(oNode As MSXML2.IXMLDOMNode)
Public Event StatusMessage(sMess As String)

Private m_StrCategory As String
Private m_StrFeed As String
Private m_StrHeadline As String

Private m_LonCategoryID As Long

Private sWebsite As String

' Global FileSystemObject settings
Private FSys As New FileSystemObject
Private FSysFile As Object
Private FSysFolder As Object

Private strURL As String
Private strFeed As String
Private strPubDate As String
Private strHeadlines As String
Private FeedURL As String
Private strFeedImage As String
Private m_RSFeed As ADODB.Recordset
Private oRSS As MSXML2.DOMDocument
Private oItemList() As MSXML2.IXMLDOMNode
Private IsoCountry As Dictionary
Private lngINtervall As Long

Public Property Get CategoryId() As Long
    CategoryId = m_LonCategoryID
End Property

Public Property Get ImageURL() As String
    ImageURL = strFeedImage
End Property

Public Property Let CategoryId(ByVal LonValue As Long)
    m_LonCategoryID = LonValue
End Property

Public Property Get Headline() As String
    Headline = m_StrHeadline
End Property

Public Property Let Headline(ByVal StrValue As String)
    m_StrHeadline = StrValue
End Property

Public Property Get Feed() As String
    Feed = m_StrFeed
End Property

Public Property Let Feed(ByVal StrValue As String)
    m_StrFeed = StrValue
End Property

Public Property Get Category() As String
    Category = m_StrCategory
End Property

Public Property Let Category(ByVal StrValue As String)
    m_StrCategory = StrValue
End Property

Private Sub chkAutoUpdate_Click()
    tmrFrequency.Enabled = IIf(chkAutoUpdate.value = vbChecked, True, False)
End Sub

Private Sub CheckOfflineBrowsing()
Dim fs As New FileSystemObject
    
   If Not fs.FolderExists(g_sAppPath & "\data\user\Exports\Feeds") Then
        fs.CreateFolder g_sAppPath & "\data\user\Exports\Feeds"
    'g_sAppPath & \data\user\Exports\Feeds
    Else
    
    End If
    
End Sub

Private Sub cmdApply_Click()
    Dim m_LonCategoryID As Long

    m_LonCategoryID = 0
    On Error Resume Next
    lngINtervall = CLng(txtSec.Text * 1000)
    
    If lngINtervall = 0 Then lngINtervall = 90000
    
    tmrFrequency.Interval = lngINtervall

End Sub

Private Sub ComCountry_Click()
    'DebugPrint IsoCountry.Item(ComCountry.List(ComCountry.ListIndex))
End Sub

Private Sub lstFeeds_Click()
        '<EhHeader>
        On Error GoTo lstFeeds_Click_Err
        Dim sCountryISO As String
        
        If lstFeeds.ListIndex = -1 Then Exit Sub
        
        sCountryISO = IsoCountry.Item(ComCountry.List(ComCountry.ListIndex))
        
100     m_RSFeed.MoveFirst
        
102     m_RSFeed.Find "FeedID = " & lstFeeds.ItemData(lstFeeds.ListIndex)
        
108     RaiseEvent FeedsClick("", IIf(chkAutoGeocode.value = vbChecked, True, False), GetRSS, sCountryISO)
        
        '<EhFooter>
        Exit Sub

lstFeeds_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicContent.lstFeeds_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Function GetRSS() As MSXML2.IXMLDOMNodeList
        '<EhHeader>
        On Error GoTo GetRSS_Err
        '</EhHeader>
        Dim RSFeedIMG As New ADODB.Recordset
    
100     lstHeadlines.Clear
102     strHeadlines = ""
104     strURL = ""
106     strFeed = ""
108     strPubDate = ""

110     DoEvents
    
112     RaiseEvent StatusMessage("Please wait getting feeds...")

114     DoEvents
    
        Dim oItems As MSXML2.IXMLDOMNodeList
        Dim i As Integer
        Dim oNode As IXMLDOMNode
    
116     Set oRSS = New MSXML2.DOMDocument
118     oRSS.async = False

        If chkAutoGeocode.value = vbChecked Then
            If ComCountry.List(ComCountry.ListIndex) = "-------NONE-------" Then
                Call oRSS.Load("http://ws.geonames.org/rssToGeoRSS?feedUrl=" & m_RSFeed.Fields("FeedURL").value & IIf(chkExcludeNon.value = vbChecked, "", "&addUngeocodedItems=true"))
            Else
                Call oRSS.Load("http://ws.geonames.org/rssToGeoRSS?feedUrl=" & m_RSFeed.Fields("FeedURL").value & IIf(chkExcludeNon.value = vbChecked, "", "&addUngeocodedItems=true") & "&country=" & IsoCountry.Item(ComCountry.List(ComCountry.ListIndex)))
            End If
        Else
120         Call oRSS.Load(m_RSFeed.Fields("FeedURL").value)
        End If
        
122     Set oItems = oRSS.selectNodes("rss/channel/item")

124     i = -1
    
126     ReDim oItemList(oItems.Length)
    
128     For Each oNode In oItems
130         i = i + 1
132         lstHeadlines.AddItem oNode.selectSingleNode("title").Text
134         Set oItemList(i) = oNode
136     Next oNode
    
138     RaiseEvent StatusMessage("Retrieved " & lstHeadlines.ListCount & " feeds.")

140     RSFeedIMG.Open "SELECT FeedImageURL FROM Feeds WHERE FeedID = " & lstFeeds.ItemData(lstFeeds.ListIndex), m_Cnn, adOpenDynamic, adLockOptimistic
142     DebugPrint "SELECT FeedImageURL FROM Feeds WHERE FeedID = " & lstFeeds.ItemData(lstFeeds.ListIndex)

144     If RSFeedIMG.EOF Then
146         strFeedImage = ""
        Else

148         If Not IsNull(RSFeedIMG.Fields(0).value) Then
150             strFeedImage = RSFeedIMG.Fields(0).value
            End If
        End If
'strFeedImage = "http://best-apps.t3.com/wp-content/uploads/2011/05/bbc-news.png"
152     RSFeedIMG.Close
154     Set RSFeedIMG = Nothing
        
156     DoEvents

        Set GetRSS = oItems 'm_RSFeed.Fields("FeedURL").Value

        '<EhFooter>
        Exit Function

GetRSS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicContent.GetRSS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub lstHeadlines_Click()

        '</EhHeader>

        Dim oNode As MSXML2.IXMLDOMNode
100     Set oNode = oItemList(lstHeadlines.ListIndex)

        RaiseEvent HeadlinesClick(oNode)
End Sub

Public Sub Init(Optional bUpdate As Boolean)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>

100     sWebsite = g_sAppServerPath
    
102     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
104         sWebsite = sWebsite & "/"
        End If

106     sWebsite = sWebsite & "oasis4.asp"

108     If bUpdate Then
110         If UpdateFeedGroups Then UpdateFeeds
        End If

124     OpenFeed
        LoadISOnum
        CheckOfflineBrowsing
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicContent.Init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function WriteFeeds()
        '<EhHeader>
        On Error GoTo WriteFeeds_Err
        '</EhHeader>
    
        ' This is the HTML that will display the feed.
100     Open App.Path & "\HTML\Feeds.html" For Output As #1
102     Print #1, "<html>"
104     Print #1, "<head>"
106     Print #1, "<title>" & strHeadlines & "</title>"
108     Print #1, "<style type=""text/css"">"
110     Print #1, "<!--"
112     Print #1, "body,td,th {color: #383C45;font-family: Verdana, Arial, Helvetica, sans-serif;}"
114     Print #1, "body {background-color: #F4F5F9;margin-left: 5px;margin-top: 5px;margin-right: 5px;margin-bottom: 5px;}"
116     Print #1, "a:link {color: #2F4D8B;text-decoration: none;}"
118     Print #1, "a:visited {text-decoration: none;color: #2F4D8B;}"
120     Print #1, "a:hover {text-decoration: underline;color: #2F4D8B;}"
122     Print #1, "a:active {text-decoration: none;color: #2F4D8B;}"
124     Print #1, ".style2 {font-size: xx-small;color: #797C83;}"
126     Print #1, "-->"
128     Print #1, "</style></head>"
130     Print #1, "<body><table width=""100%"">"
132     Print #1, "<tr>"
134     Print #1, "<td width=""2%""><img src=""" & strFeedImage & """ border=""0""></td>"
        '    Print #1, "<td width=""98%""><a href=" & strURL & " target=""_blank""><strong>" & strHeadlines & "</strong></a></td>"
136     Print #1, "<td width=""98%""><a href=" & strURL & "><strong>" & strHeadlines & "</strong></a></td>"
138     Print #1, "</tr>"
140     Print #1, "<tr>"
142     Print #1, "<td>&nbsp;</td>"
144     Print #1, "<td><span class=""style2""><strong>Published Date:</strong> " & strPubDate & "</span></td>"
146     Print #1, "</tr>"
148     Print #1, "<tr>"
150     Print #1, "<td>&nbsp;</td>"
152     Print #1, "<td>" & strFeed & "</td>"
154     Print #1, "</tr>"
156     Print #1, "</table>"
158     Print #1, "</body>"
160     Print #1, "</html>"
162     Close #1

        '<EhFooter>
        Exit Function

WriteFeeds_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicContent.WriteFeeds " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub WriteHTML2()
        '<EhHeader>
        On Error GoTo WriteHTML2_Err
        '</EhHeader>
        On Error Resume Next

100     Kill App.Path & "\HTML\RSSIntro.html"
102     Open App.Path & "\HTML\RSSIntro.html" For Output As #1
104     Print #1, "<html>"
106     Print #1, "<head>"
108     Print #1, "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
110     Print #1, "<title>OASIS Dynamic Content Feeder</title>"
112     Print #1, "<style type=""text/css"">"
114     Print #1, "<!--"
116     Print #1, "body,td,th {color: #383C45;font-family: Verdana, Arial, Helvetica, sans-serif;}"
118     Print #1, "body {background-color: #F4F5F9;margin-left: 5px;margin-top: 5px;margin-right: 5px;margin-bottom: 5px;}"
120     Print #1, "a:link {color: #2F4D8B;text-decoration: none;}"
122     Print #1, "a:visited {text-decoration: none;color: #2F4D8B;}"
124     Print #1, "a:hover {text-decoration: underline;color: #2F4D8B;}"
126     Print #1, "a:active {text-decoration: none;color: #2F4D8B;}"
128     Print #1, ".style2 {font-size: xx-small;color: #797C83;}"
130     Print #1, "-->"
132     Print #1, "</style>"
134     Print #1, "</head>"

136     Print #1, "<body>"
138     Print #1, "<table width=""100%"">"
140     Print #1, "<tr>"
142     Print #1, "    <td width=""100%""><strong><center> <h1> OASIS </h1></center></strong></td>"
144     Print #1, "</tr>"
146     Print #1, "<tr>"
148     Print #1, "    <td width=""100%""><center><h4> - Information Matters - </h4></center></td>"
150     Print #1, "</tr>"
152     Print #1, "<tr>"
154     Print #1, "    <td width=""100%""><center><h4> - OASIS Dynamic Content Module - </h4></center></td>"
156     Print #1, "</tr>"
158     Print #1, "<tr>"
160     Print #1, "    <td width=""100%""><hr></td>"
162     Print #1, "</tr>"
164     Print #1, "<tr>"
166     Print #1, "    <td width=""100%""><center><strong>QUICK START:</strong></center></td>"
168     Print #1, "</tr>"
170     Print #1, "<tr>"
172     Print #1, "    <td width=""100%""><center>1) Chooose your Category from the dropdown box. <br>"
174     Print #1, "    2) Click on the available feed topics. Now the available content headlines is displayed.<br>"
176     Print #1, "    3) Click on the content topic. Description of the topic is shown in the main window.<br>"
178     Print #1, "    4) Click on the feed description link to view full story.</center></td>"
180     Print #1, "</tr>"
182     Print #1, "<tr>"
184     Print #1, "    <td width=""100%""><hr></td>"
186     Print #1, "</tr>"
188     Print #1, "<tr>"
190     Print #1, "    <td width=""100%""> <center> <h6> - OASIS is developed by iMMAP - </h6></center></td>"
192     Print #1, "</tr>"
194     Print #1, "<tr>"
196     Print #1, "    <td width=""100%""> <center> <a href=""http://www.immap.org"" target=""_blank""> <span class=""style2"">www.immap.org</span></a></center></td>"
198     Print #1, "</tr>"

200     Print #1, "<tr>"
202     Print #1, "<td>&nbsp;</td>"
204     Print #1, "<td>&nbsp;</td>"
206     Print #1, "</tr>"
208     Print #1, "</table>"
210     Print #1, "</body>"
212     Print #1, "</html>"
214     Close #1

        '<EhFooter>
        Exit Sub

WriteHTML2_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicContent.WriteHTML2 " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CheckHTML()
        '<EhHeader>
        On Error GoTo CheckHTML_Err
        '</EhHeader>

        ' Make sure the HTML directory is there.
100     If FSys.FolderExists(App.Path & "\HTML") Then
            Exit Function
        Else
102         FSys.CreateFolder (App.Path & "\HTML")
        End If

        '<EhFooter>
        Exit Function

CheckHTML_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicContent.CheckHTML " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub UpdateFeeds()
        '<EhHeader>
        On Error GoTo UpdateFeeds_Err
        '</EhHeader>
        Dim rsRemote As ADODB.Recordset
        Dim RS As ADODB.Recordset
        Dim j As Integer
        Dim i As Integer
        'Dim RSUpdater As ADODB.Recordset
        
        'Now Check the Dynamic Content version
100     If Not g_sRemoteTablePrefix = "" Then
                
102         Set rsRemote = New ADODB.Recordset
    
104         'rsRemote.Open sWebsite & "?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & "Feeds"), , adOpenDynamic, adLockOptimistic
            Set rsRemote = OpenServerRSCompressed(sWebsite, "ID", "SELECT * FROM " & g_sRemoteTablePrefix & "Feeds")

106         If Not rsRemote.State = 0 Then

108             m_Cnn.Execute "delete from Feeds"

110             If rsRemote.EOF And rsRemote.Bof Then
                    Exit Sub
                End If
            
112             Set RS = New ADODB.Recordset
                
114             rsRemote.MoveFirst
    
116             RS.Open "SELECT * FROM Feeds", m_Cnn, adOpenDynamic, adLockBatchOptimistic
    
118             Do While Not rsRemote.EOF
120                 RS.AddNew
    
122                 For j = 1 To rsRemote.Fields.Count - 1

                        If DoFieldExists(RS, RS.Fields.Item(j).Name) Then
124                         If Len(rsRemote.Fields.Item(RS.Fields.Item(j).Name).value) > 0 Then
126                             'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value

                                If RS.Fields.Item(j).Type = adBoolean And UCase$(rsRemote.Fields(RS.Fields.Item(j).Name).value) = "NO" Then
                                    RS.Fields.Item(j).value = False
                                ElseIf RS.Fields.Item(j).Type = adBoolean And UCase$(rsRemote.Fields(RS.Fields.Item(j).Name).value) = "YES" Then
                                    RS.Fields.Item(j).value = True
                                Else
                                    RS.Fields.Item(j).value = rsRemote.Fields(RS.Fields.Item(j).Name).value
                                End If
                                
                            End If
                        End If

                    Next
                    
128                 If Not bSQLServerInUse Then RS.Fields.Item("FeedID").value = i
                    
130                 RS.UpdateBatch
132                 i = i + 1
134                 rsRemote.MoveNext
                Loop
    
136             rsRemote.Close
138             RS.Close
    
140             SynchProfileSettingWithServer "SettingValue5", g_sRemoteTablePrefix, m_Cnn
            End If
                
        End If
    
        '<EhFooter>
        Exit Sub

UpdateFeeds_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicContent.UpdateFeeds " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Function UpdateFeedGroups() As Boolean
        '<EhHeader>
        On Error GoTo UpdateFeedGroups_Err
        '</EhHeader>
        Dim rsRemote As ADODB.Recordset
        Dim RS As ADODB.Recordset
        Dim j As Integer
        
        'Now Check the Dynamic Content version
100     If Not g_sRemoteTablePrefix = "" Then
                
102         Set rsRemote = New ADODB.Recordset
            On Error Resume Next
            
104         'rsRemote.Open sWebsite & "?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & "FeedGroups"), , adOpenDynamic, adLockOptimistic
            Set rsRemote = OpenServerRSCompressed(sWebsite, "ID", "SELECT * FROM " & g_sRemoteTablePrefix & "FeedGroups")
            On Error GoTo UpdateFeedGroups_Err
            
106         If Not rsRemote.State = 0 Then

108             m_Cnn.Execute "delete from Groups"
110             DebugPrint Err.Description

112             If rsRemote.EOF And rsRemote.Bof Then
                    Exit Function
                End If
            
114             Set RS = New ADODB.Recordset
                
116             rsRemote.MoveFirst
    
118             RS.Open "SELECT * FROM Groups", m_Cnn, adOpenDynamic, adLockBatchOptimistic
    
120             Do While Not rsRemote.EOF
122                 RS.AddNew
    
124                 For j = 0 To RS.Fields.Count - 1

126                     If Len(rsRemote.Fields.Item(RS.Fields.Item(j).Name).value) > 0 Then
128                         'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
                            RS.Fields.Item(j).value = rsRemote.Fields(RS.Fields.Item(j).Name).value
                        End If

                    Next

130                 RS.UpdateBatch adAffectCurrent
132                 rsRemote.MoveNext
                Loop
    
134             rsRemote.Close
136             RS.Close
    
138             Set rsRemote = Nothing
140             Set RS = Nothing
            End If
                
        End If
        
142     UpdateFeedGroups = True
        
        '<EhFooter>
        Exit Function

UpdateFeedGroups_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicContent.UpdateFeedGroups " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cboCategory_Click()
        '<EhHeader>
        On Error GoTo cboCategory_Click_Err

        '</EhHeader>
100     If cboCategory.List(cboCategory.ListIndex) = "---Choose your topic---" Then Exit Sub

102     m_StrCategory = cboCategory.List(cboCategory.ListIndex)
104     m_LonCategoryID = cboCategory.ItemData(cboCategory.ListIndex)

        ' Tidy up a bit.
106     lstHeadlines.Clear
108     strHeadlines = ""
110     strURL = ""
112     strFeed = ""
114     strPubDate = ""

116     DoEvents

118     Set m_RSFeed = New ADODB.Recordset
    
120     m_RSFeed.Open "SELECT * FROM Feeds WHERE GroupID =" & cboCategory.ItemData(cboCategory.ListIndex), m_Cnn, adOpenDynamic, adLockOptimistic
    
122     lstFeeds.Clear
    
124     Do While Not m_RSFeed.EOF
126         lstFeeds.AddItem m_RSFeed.Fields("FeedName").value
128         lstFeeds.ItemData(lstFeeds.ListCount - 1) = m_RSFeed.Fields("FeedID").value
130         m_RSFeed.MoveNext
        Loop
    
132     RaiseEvent CategoryClick
    
        '<EhFooter>
        Exit Sub

cboCategory_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicContent.cboCategory_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function OpenFeed()
        '<EhHeader>
        On Error GoTo OpenFeed_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset
    
100     RS.Open "SELECT * FROM Groups ORDER BY GroupText", m_Cnn
    
102     cboCategory.Clear
    
104     cboCategory.AddItem "---Choose your topic---"
        
106     If Not RS.EOF And Not RS.Bof Then

108         RS.MoveFirst
    
110         Do While Not RS.EOF
112             cboCategory.AddItem RS.Fields("GroupText").value
114             cboCategory.ItemData(cboCategory.ListCount - 1) = RS.Fields("GroupID").value
116             RS.MoveNext
            Loop
    
118         RS.Close
        
        End If

120     cboCategory.ListIndex = 0

        Exit Function

OpenFeed_Err:
122     Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.OpenFeed", "RSSBrowser component failure"
        '</EhFooter>
        '<EhFooter>
        Exit Function

End Function


Private Sub Form_Load()
    
    
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
    
    m_LonCategoryID = 0
    On Error Resume Next
    lngINtervall = CLng(txtSec.Text * 1000)
    
    If lngINtervall = 0 Then lngINtervall = 90000
    
    tmrFrequency.Interval = lngINtervall
    ComCountry.ListIndex = 0
End Sub


Private Sub LoadISOnum()

    Set IsoCountry = New Dictionary
    ComCountry.Clear
    ComCountry.AddItem "-------NONE-------"
    
    With IsoCountry
        .Add "URUGUAY", "UY"
        ComCountry.AddItem "URUGUAY"
        .Add "UZBEKISTAN", "UZ"
        ComCountry.AddItem "UZBEKISTAN"
        .Add "VANUATU", "VU"
        ComCountry.AddItem "VANUATU"
        .Add "VENEZUELA, BOLIVARIAN REPUBLIC OF", "VE"
        ComCountry.AddItem "VENEZUELA, BOLIVARIAN REPUBLIC OF"
        .Add "VIET NAM", "VN"
        ComCountry.AddItem "VIET NAM"
        .Add "VIRGIN ISLANDS, BRITISH", "VG"
        ComCountry.AddItem "VIRGIN ISLANDS, BRITISH"
        .Add "VIRGIN ISLANDS, U.S.", "VI"
        ComCountry.AddItem "VIRGIN ISLANDS, U.S."
        .Add "WALLIS AND FUTUNA", "WF"
        ComCountry.AddItem "WALLIS AND FUTUNA"
        .Add "WESTERN SAHARA", "EH"
        ComCountry.AddItem "WESTERN SAHARA"
        .Add "YEMEN", "YE"
        ComCountry.AddItem "YEMEN"
        .Add "ZAMBIA", "ZM"
        ComCountry.AddItem "ZAMBIA"
        .Add "ZIMBABWE", "ZW"
        ComCountry.AddItem "ZIMBABWE"
        .Add "TUNISIA", "TN"
        ComCountry.AddItem "TUNISIA"
        .Add "TURKEY", "TR"
        ComCountry.AddItem "TURKEY"
        .Add "TURKMENISTAN", "TM"
        ComCountry.AddItem "TURKMENISTAN"
        .Add "TURKS AND CAICOS ISLANDS", "TC"
        ComCountry.AddItem "TURKS AND CAICOS ISLANDS"
        .Add "TUVALU", "TV"
        ComCountry.AddItem "TUVALU"
        .Add "UGANDA", "UG"
        ComCountry.AddItem "UGANDA"
        .Add "UKRAINE", "UA"
        ComCountry.AddItem "UKRAINE"
        .Add "UNITED ARAB EMIRATES", "AE"
        ComCountry.AddItem "UNITED ARAB EMIRATES"
        .Add "UNITED KINGDOM", "GB"
        ComCountry.AddItem "UNITED KINGDOM"
        .Add "UNITED STATES", "US"
        ComCountry.AddItem "UNITED STATES"
        .Add "UNITED STATES MINOR OUTLYING ISLANDS", "UM"
        ComCountry.AddItem "UNITED STATES MINOR OUTLYING ISLANDS"
        .Add "SURINAME", "SR"
        ComCountry.AddItem "SURINAME"
        .Add "SVALBARD AND JAN MAYEN", "SJ"
        ComCountry.AddItem "SVALBARD AND JAN MAYEN"
        .Add "SWAZILAND", "SZ"
        ComCountry.AddItem "SWAZILAND"
        .Add "SWEDEN", "SE"
        ComCountry.AddItem "SWEDEN"
        .Add "SWITZERLAND", "CH"
        ComCountry.AddItem "SWITZERLAND"
        .Add "SYRIAN ARAB REPUBLIC", "SY"
        ComCountry.AddItem "SYRIAN ARAB REPUBLIC"
        .Add "TAIWAN, PROVINCE OF CHINA", "TW"
        ComCountry.AddItem "TAIWAN, PROVINCE OF CHINA"
        .Add "TAJIKISTAN", "TJ"
        ComCountry.AddItem "TAJIKISTAN"
        .Add "TANZANIA, UNITED REPUBLIC OF", "TZ"
        ComCountry.AddItem "TANZANIA, UNITED REPUBLIC OF"
        .Add "THAILAND", "TH"
        ComCountry.AddItem "THAILAND"
        .Add "TIMOR-LESTE", "TL"
        ComCountry.AddItem "TIMOR-LESTE"
        .Add "TOGO", "TG"
        ComCountry.AddItem "TOGO"
        .Add "TOKELAU", "TK"
        ComCountry.AddItem "TOKELAU"
        .Add "TONGA", "TO"
        ComCountry.AddItem "TONGA"
        .Add "TRINIDAD AND TOBAGO", "TT"
        ComCountry.AddItem "TRINIDAD AND TOBAGO"
        .Add "SAINT BARTHÉLEMY", "BL"
        ComCountry.AddItem "SAINT BARTHÉLEMY"
        .Add "SAINT HELENA", "SH"
        ComCountry.AddItem "SAINT HELENA"
        .Add "SAINT KITTS AND NEVIS", "KN"
        ComCountry.AddItem "SAINT KITTS AND NEVIS"
        .Add "SAINT LUCIA", "LC"
        ComCountry.AddItem "SAINT LUCIA"
        .Add "SAINT MARTIN", "MF"
        ComCountry.AddItem "SAINT MARTIN"
        .Add "SAINT PIERRE AND MIQUELON", "PM"
        ComCountry.AddItem "SAINT PIERRE AND MIQUELON"
        .Add "SAINT VINCENT AND THE GRENADINES", "VC"
        ComCountry.AddItem "SAINT VINCENT AND THE GRENADINES"
        .Add "SAMOA", "WS"
        ComCountry.AddItem "SAMOA"
        .Add "SAN MARINO", "SM"
        ComCountry.AddItem "SAN MARINO"
        .Add "SAO TOME AND PRINCIPE", "ST"
        ComCountry.AddItem "SAO TOME AND PRINCIPE"
        .Add "SAUDI ARABIA", "SA"
        ComCountry.AddItem "SAUDI ARABIA"
        .Add "SENEGAL", "SN"
        ComCountry.AddItem "SENEGAL"
        .Add "SERBIA", "RS"
        ComCountry.AddItem "SERBIA"
        .Add "SEYCHELLES", "SC"
        ComCountry.AddItem "SEYCHELLES"
        .Add "SIERRA LEONE", "SL"
        ComCountry.AddItem "SIERRA LEONE"
        .Add "SINGAPORE", "SG"
        ComCountry.AddItem "SINGAPORE"
        .Add "SLOVAKIA", "SK"
        ComCountry.AddItem "SLOVAKIA"
        .Add "SLOVENIA", "SI"
        ComCountry.AddItem "SLOVENIA"
        .Add "SOLOMON ISLANDS", "SB"
        ComCountry.AddItem "SOLOMON ISLANDS"
        .Add "SOMALIA", "SO"
        ComCountry.AddItem "SOMALIA"
        .Add "SOUTH AFRICA", "ZA"
        ComCountry.AddItem "SOUTH AFRICA"
        .Add "SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS", "GS"
        ComCountry.AddItem "SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS"
        .Add "SPAIN", "ES"
        ComCountry.AddItem "SPAIN"
        .Add "SRI LANKA", "LK"
        ComCountry.AddItem "SRI LANKA"
        .Add "SUDAN", "SD"
        ComCountry.AddItem "SUDAN"
        .Add "NIUE", "NU"
        ComCountry.AddItem "NIUE"
        .Add "NORFOLK ISLAND", "NF"
        ComCountry.AddItem "NORFOLK ISLAND"
        .Add "NORTHERN MARIANA ISLANDS", "MP"
        ComCountry.AddItem "NORTHERN MARIANA ISLANDS"
        .Add "NORWAY", "NO"
        ComCountry.AddItem "NORWAY"
        .Add "OMAN", "OM"
        ComCountry.AddItem "OMAN"
        .Add "PAKISTAN", "PK"
        ComCountry.AddItem "PAKISTAN"
        .Add "PALAU", "PW"
        ComCountry.AddItem "PALAU"
        .Add "PALESTINIAN TERRITORY, OCCUPIED", "PS"
        ComCountry.AddItem "PALESTINIAN TERRITORY, OCCUPIED"
        .Add "PANAMA", "PA"
        ComCountry.AddItem "PANAMA"
        .Add "PAPUA NEW GUINEA", "PG"
        ComCountry.AddItem "PAPUA NEW GUINEA"
        .Add "PARAGUAY", "PY"
        ComCountry.AddItem "PARAGUAY"
        .Add "PERU", "PE"
        ComCountry.AddItem "PERU"
        .Add "PHILIPPINES", "PH"
        ComCountry.AddItem "PHILIPPINES"
        .Add "PITCAIRN", "PN"
        ComCountry.AddItem "PITCAIRN"
        .Add "POLAND", "PL"
        ComCountry.AddItem "POLAND"
        .Add "PORTUGAL", "PT"
        ComCountry.AddItem "PORTUGAL"
        .Add "PUERTO RICO", "PR"
        ComCountry.AddItem "PUERTO RICO"
        .Add "QATAR", "QA"
        ComCountry.AddItem "QATAR"
        .Add "RÉUNION", "RE"
        ComCountry.AddItem "RÉUNION"
        .Add "ROMANIA", "RO"
        ComCountry.AddItem "ROMANIA"
        .Add "RUSSIAN FEDERATION", "RU"
        ComCountry.AddItem "RUSSIAN FEDERATION"
        .Add "RWANDA", "RW"
        ComCountry.AddItem "RWANDA"
        .Add "MOLDOVA, REPUBLIC OF", "MD"
        ComCountry.AddItem "MOLDOVA, REPUBLIC OF"
        .Add "MONACO", "MC"
        ComCountry.AddItem "MONACO"
        .Add "MONGOLIA", "MN"
        ComCountry.AddItem "MONGOLIA"
        .Add "MONTENEGRO", "ME"
        ComCountry.AddItem "MONTENEGRO"
        .Add "MONTSERRAT", "MS"
        ComCountry.AddItem "MONTSERRAT"
        .Add "MOROCCO", "MA"
        ComCountry.AddItem "MOROCCO"
        .Add "MOZAMBIQUE", "MZ"
        ComCountry.AddItem "MOZAMBIQUE"
        .Add "MYANMAR", "MM"
        ComCountry.AddItem "MYANMAR"
        .Add "NAMIBIA", "NA"
        ComCountry.AddItem "NAMIBIA"
        .Add "NAURU", "NR"
        ComCountry.AddItem "NAURU"
        .Add "NEPAL", "NP"
        ComCountry.AddItem "NEPAL"
        .Add "NETHERLANDS", "NL"
        ComCountry.AddItem "NETHERLANDS"
        .Add "NETHERLANDS ANTILLES", "AN"
        ComCountry.AddItem "NETHERLANDS ANTILLES"
        .Add "NEW CALEDONIA", "NC"
        ComCountry.AddItem "NEW CALEDONIA"
        .Add "NEW ZEALAND", "NZ"
        ComCountry.AddItem "NEW ZEALAND"
        .Add "NICARAGUA", "NI"
        ComCountry.AddItem "NICARAGUA"
        .Add "NIGER", "NE"
        ComCountry.AddItem "NIGER"
        .Add "NIGERIA", "NG"
        ComCountry.AddItem "NIGERIA"
        .Add "HONDURAS", "HN"
        ComCountry.AddItem "HONDURAS"
        .Add "HONG KONG", "HK"
        ComCountry.AddItem "HONG KONG"
        .Add "HUNGARY", "HU"
        ComCountry.AddItem "HUNGARY"
        .Add "ICELAND", "IS"
        ComCountry.AddItem "ICELAND"
        .Add "INDIA", "IN"
        ComCountry.AddItem "INDIA"
        .Add "INDONESIA", "ID"
        ComCountry.AddItem "INDONESIA"
        .Add "IRAN, ISLAMIC REPUBLIC OF", "IR"
        ComCountry.AddItem "IRAN, ISLAMIC REPUBLIC OF"
        .Add "IRAQ", "IQ"
        ComCountry.AddItem "IRAQ"
        .Add "IRELAND", "IE"
        ComCountry.AddItem "IRELAND"
        .Add "ISLE OF MAN", "IM"
        ComCountry.AddItem "ISLE OF MAN"
        .Add "ISRAEL", "IL"
        ComCountry.AddItem "ISRAEL"
        .Add "ITALY", "IT"
        ComCountry.AddItem "ITALY"
        .Add "JAMAICA", "JM"
        ComCountry.AddItem "JAMAICA"
        .Add "JAPAN", "JP"
        ComCountry.AddItem "JAPAN"
        .Add "JERSEY", "JE"
        ComCountry.AddItem "JERSEY"
        .Add "JORDAN", "JO"
        ComCountry.AddItem "JORDAN"
        .Add "KAZAKHSTAN", "KZ"
        ComCountry.AddItem "KAZAKHSTAN"
        .Add "KENYA", "KE"
        ComCountry.AddItem "KENYA"
        .Add "KIRIBATI", "KI"
        ComCountry.AddItem "KIRIBATI"
        .Add "KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF", "KP"
        ComCountry.AddItem "KOREA, DEMOCRATIC PEOPLE'S REPUBLIC OF"
        .Add "KOREA, REPUBLIC OF", "KR"
        ComCountry.AddItem "KOREA, REPUBLIC OF"
        .Add "KUWAIT", "KW"
        ComCountry.AddItem "KUWAIT"
        .Add "KYRGYZSTAN", "KG"
        ComCountry.AddItem "KYRGYZSTAN"
        .Add "LAO PEOPLE'S DEMOCRATIC REPUBLIC", "LA"
        ComCountry.AddItem "LAO PEOPLE'S DEMOCRATIC REPUBLIC"
        .Add "LATVIA", "LV"
        ComCountry.AddItem "LATVIA"
        .Add "LEBANON", "LB"
        ComCountry.AddItem "LEBANON"
        .Add "LESOTHO", "LS"
        ComCountry.AddItem "LESOTHO"
        .Add "LIBERIA", "LR"
        ComCountry.AddItem "LIBERIA"
        .Add "LIBYAN ARAB JAMAHIRIYA", "LY"
        ComCountry.AddItem "LIBYAN ARAB JAMAHIRIYA"
        .Add "LIECHTENSTEIN", "LI"
        ComCountry.AddItem "LIECHTENSTEIN"
        .Add "LITHUANIA", "LT"
        ComCountry.AddItem "LITHUANIA"
        .Add "LUXEMBOURG", "LU"
        ComCountry.AddItem "LUXEMBOURG"
        .Add "MACAO", "MO"
        ComCountry.AddItem "MACAO"
        .Add "MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF", "MK"
        ComCountry.AddItem "MACEDONIA, THE FORMER YUGOSLAV REPUBLIC OF"
        .Add "MADAGASCAR", "MG"
        ComCountry.AddItem "MADAGASCAR"
        .Add "MALAWI", "MW"
        ComCountry.AddItem "MALAWI"
        .Add "MALAYSIA", "MY"
        ComCountry.AddItem "MALAYSIA"
        .Add "MALDIVES", "MV"
        ComCountry.AddItem "MALDIVES"
        .Add "MALI", "ML"
        ComCountry.AddItem "MALI"
        .Add "MALTA", "MT"
        ComCountry.AddItem "MALTA"
        .Add "MARSHALL ISLANDS", "MH"
        ComCountry.AddItem "MARSHALL ISLANDS"
        .Add "MARTINIQUE", "MQ"
        ComCountry.AddItem "MARTINIQUE"
        .Add "MAURITANIA", "MR"
        ComCountry.AddItem "MAURITANIA"
        .Add "MAURITIUS", "MU"
        ComCountry.AddItem "MAURITIUS"
        .Add "MAYOTTE", "YT"
        ComCountry.AddItem "MAYOTTE"
        .Add "MEXICO", "MX"
        ComCountry.AddItem "MEXICO"
        .Add "MICRONESIA, FEDERATED STATES OF", "FM"
        ComCountry.AddItem "MICRONESIA, FEDERATED STATES OF"
        .Add "CONGO, THE DEMOCRATIC REPUBLIC OF THE", "CD"
        ComCountry.AddItem "CONGO, THE DEMOCRATIC REPUBLIC OF THE"
        .Add "COOK ISLANDS", "CK"
        ComCountry.AddItem "COOK ISLANDS"
        .Add "COSTA RICA", "CR"
        ComCountry.AddItem "COSTA RICA"
        .Add "CÔTE D'IVOIRE", "CI"
        ComCountry.AddItem "CÔTE D'IVOIRE"
        .Add "CROATIA", "HR"
        ComCountry.AddItem "CROATIA"
        .Add "CUBA", "CU"
        ComCountry.AddItem "CUBA"
        .Add "CYPRUS", "CY"
        ComCountry.AddItem "CYPRUS"
        .Add "CZECH REPUBLIC", "CZ"
        ComCountry.AddItem "CZECH REPUBLIC"
        .Add "DENMARK", "DK"
        ComCountry.AddItem "DENMARK"
        .Add "DJIBOUTI", "DJ"
        ComCountry.AddItem "DJIBOUTI"
        .Add "DOMINICA", "DM"
        ComCountry.AddItem "DOMINICA"
        .Add "DOMINICAN REPUBLIC", "DO"
        ComCountry.AddItem "DOMINICAN REPUBLIC"
        .Add "ECUADOR", "EC"
        ComCountry.AddItem "ECUADOR"
        .Add "EGYPT", "EG"
        ComCountry.AddItem "EGYPT"
        .Add "EL SALVADOR", "SV"
        ComCountry.AddItem "EL SALVADOR"
        .Add "EQUATORIAL GUINEA", "GQ"
        ComCountry.AddItem "EQUATORIAL GUINEA"
        .Add "ERITREA", "ER"
        ComCountry.AddItem "ERITREA"
        .Add "ESTONIA", "EE"
        ComCountry.AddItem "ESTONIA"
        .Add "ETHIOPIA", "ET"
        ComCountry.AddItem "ETHIOPIA"
        .Add "FALKLAND ISLANDS (MALVINAS)", "FK"
        ComCountry.AddItem "FALKLAND ISLANDS (MALVINAS)"
        .Add "FAROE ISLANDS", "FO"
        ComCountry.AddItem "FAROE ISLANDS"
        .Add "FIJI", "FJ"
        ComCountry.AddItem "FIJI"
        .Add "FINLAND", "FI"
        ComCountry.AddItem "FINLAND"
        .Add "FRANCE", "FR"
        ComCountry.AddItem "FRANCE"
        .Add "FRENCH GUIANA", "GF"
        ComCountry.AddItem "FRENCH GUIANA"
        .Add "FRENCH POLYNESIA", "PF"
        ComCountry.AddItem "FRENCH POLYNESIA"
        .Add "FRENCH SOUTHERN TERRITORIES", "TF"
        ComCountry.AddItem "FRENCH SOUTHERN TERRITORIES"
        .Add "GABON", "GA"
        ComCountry.AddItem "GABON"
        .Add "GAMBIA", "GM"
        ComCountry.AddItem "GAMBIA"
        .Add "GEORGIA", "GE"
        ComCountry.AddItem "GEORGIA"
        .Add "GERMANY", "DE"
        ComCountry.AddItem "GERMANY"
        .Add "GHANA", "GH"
        ComCountry.AddItem "GHANA"
        .Add "GIBRALTAR", "GI"
        ComCountry.AddItem "GIBRALTAR"
        .Add "GREECE", "GR"
        ComCountry.AddItem "GREECE"
        .Add "GREENLAND", "GL"
        ComCountry.AddItem "GREENLAND"
        .Add "GRENADA", "GD"
        ComCountry.AddItem "GRENADA"
        .Add "GUADELOUPE", "GP"
        ComCountry.AddItem "GUADELOUPE"
        .Add "GUAM", "GU"
        ComCountry.AddItem "GUAM"
        .Add "GUATEMALA", "GT"
        ComCountry.AddItem "GUATEMALA"
        .Add "GUERNSEY", "GG"
        ComCountry.AddItem "GUERNSEY"
        .Add "GUINEA", "GN"
        ComCountry.AddItem "GUINEA"
        .Add "GUINEA-BISSAU", "GW"
        ComCountry.AddItem "GUINEA-BISSAU"
        .Add "GUYANA", "GY"
        ComCountry.AddItem "GUYANA"
        .Add "HAITI", "HT"
        ComCountry.AddItem "HAITI"
        .Add "HEARD ISLAND AND MCDONALD ISLANDS", "HM"
        ComCountry.AddItem "HEARD ISLAND AND MCDONALD ISLANDS"
        .Add "HOLY SEE (VATICAN CITY STATE)", "VA"
        ComCountry.AddItem "HOLY SEE (VATICAN CITY STATE)"
        .Add "AFGHANISTAN", "AF"
        ComCountry.AddItem "AFGHANISTAN"
        .Add "ÅLAND ISLANDS", "AX"
        ComCountry.AddItem "ÅLAND ISLANDS"
        .Add "ALBANIA", "AL"
        ComCountry.AddItem "ALBANIA"
        .Add "ALGERIA", "DZ"
        ComCountry.AddItem "ALGERIA"
        .Add "AMERICAN SAMOA", "AS"
        ComCountry.AddItem "AMERICAN SAMOA"
        .Add "ANDORRA", "AD"
        ComCountry.AddItem "ANDORRA"
        .Add "ANGOLA", "AO"
        ComCountry.AddItem "ANGOLA"
        .Add "ANGUILLA", "AI"
        ComCountry.AddItem "ANGUILLA"
        .Add "ANTARCTICA", "AQ"
        ComCountry.AddItem "ANTARCTICA"
        .Add "ANTIGUA AND BARBUDA", "AG"
        ComCountry.AddItem "ANTIGUA AND BARBUDA"
        .Add "ARGENTINA", "AR"
        ComCountry.AddItem "ARGENTINA"
        .Add "ARMENIA", "AM"
        ComCountry.AddItem "ARMENIA"
        .Add "ARUBA", "AW"
        ComCountry.AddItem "ARUBA"
        .Add "AUSTRALIA", "AU"
        ComCountry.AddItem "AUSTRALIA"
        .Add "AUSTRIA", "AT"
        ComCountry.AddItem "AUSTRIA"
        .Add "AZERBAIJAN", "AZ"
        ComCountry.AddItem "AZERBAIJAN"
        .Add "BAHAMAS", "BS"
        ComCountry.AddItem "BAHAMAS"
        .Add "BAHRAIN", "BH"
        ComCountry.AddItem "BAHRAIN"
        .Add "BANGLADESH", "BD"
        ComCountry.AddItem "BANGLADESH"
        .Add "BARBADOS", "BB"
        ComCountry.AddItem "BARBADOS"
        .Add "BELARUS", "BY"
        ComCountry.AddItem "BELARUS"
        .Add "BELGIUM", "BE"
        ComCountry.AddItem "BELGIUM"
        .Add "BELIZE", "BZ"
        ComCountry.AddItem "BELIZE"
        .Add "BENIN", "BJ"
        ComCountry.AddItem "BENIN"
        .Add "BERMUDA", "BM"
        ComCountry.AddItem "BERMUDA"
        .Add "BHUTAN", "BT"
        ComCountry.AddItem "BHUTAN"
        .Add "BOLIVIA, PLURINATIONAL STATE OF", "BO"
        ComCountry.AddItem "BOLIVIA, PLURINATIONAL STATE OF"
        .Add "BOSNIA AND HERZEGOVINA", "BA"
        ComCountry.AddItem "BOSNIA AND HERZEGOVINA"
        .Add "BOTSWANA", "BW"
        ComCountry.AddItem "BOTSWANA"
        .Add "BOUVET ISLAND", "BV"
        ComCountry.AddItem "BOUVET ISLAND"
        .Add "BRAZIL", "BR"
        ComCountry.AddItem "BRAZIL"
        .Add "BRITISH INDIAN OCEAN TERRITORY", "IO"
        ComCountry.AddItem "BRITISH INDIAN OCEAN TERRITORY"
        .Add "BRUNEI DARUSSALAM", "BN"
        ComCountry.AddItem "BRUNEI DARUSSALAM"
        .Add "BULGARIA", "BG"
        ComCountry.AddItem "BULGARIA"
        .Add "BURKINA FASO", "BF"
        ComCountry.AddItem "BURKINA FASO"
        .Add "BURUNDI", "BI"
        ComCountry.AddItem "BURUNDI"
        .Add "CAMBODIA", "KH"
        ComCountry.AddItem "CAMBODIA"
        .Add "CAMEROON", "CM"
        ComCountry.AddItem "CAMEROON"
        .Add "CANADA", "CA"
        ComCountry.AddItem "CANADA"
        .Add "CAPE VERDE", "CV"
        ComCountry.AddItem "CAPE VERDE"
        .Add "CAYMAN ISLANDS", "KY"
        ComCountry.AddItem "CAYMAN ISLANDS"
        .Add "CENTRAL AFRICAN REPUBLIC", "CF"
        ComCountry.AddItem "CENTRAL AFRICAN REPUBLIC"
        .Add "CHAD", "TD"
        ComCountry.AddItem "CHAD"
        .Add "CHILE", "CL"
        ComCountry.AddItem "CHILE"
        .Add "CHINA", "CN"
        ComCountry.AddItem "CHINA"
        .Add "CHRISTMAS ISLAND", "CX"
        ComCountry.AddItem "CHRISTMAS ISLAND"
        .Add "COCOS (KEELING) ISLANDS", "CC"
        ComCountry.AddItem "COCOS (KEELING) ISLANDS"
        .Add "COLOMBIA", "CO"
        ComCountry.AddItem "COLOMBIA"
        .Add "COMOROS", "KM"
        ComCountry.AddItem "COMOROS"
        .Add "CONGO", "CG"
        ComCountry.AddItem "CONGO"
    End With

    FindIndexStrEx ComCountry, "-------NONE-------"

    'ComCountry.ListIndex = 0

End Sub

Private Sub tmrFrequency_Timer()
Exit Sub
    If lstFeeds.ListCount > 0 Then
        lstFeeds_Click
    End If
End Sub
