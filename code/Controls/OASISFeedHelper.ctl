VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.UserControl OASISFeedHelper 
   ClientHeight    =   8835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3255
   ScaleHeight     =   8835
   ScaleWidth      =   3255
   Begin C1SizerLibCtl.C1Elastic C1Elastic2 
      Height          =   8835
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3255
      _cx             =   5741
      _cy             =   15584
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
      Appearance      =   4
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
      GridRows        =   6
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"OASISFeedHelper.ctx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.ListBox lstFeeds 
         Height          =   2400
         Left            =   90
         TabIndex        =   3
         Top             =   975
         Width           =   3075
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Text            =   "C:\RSS"
         Top             =   345
         Width           =   3075
      End
      Begin VB.ListBox lstHeadlines 
         Height          =   4935
         Left            =   90
         TabIndex        =   1
         Top             =   3750
         Width           =   3075
      End
      Begin VB.Label lblSubscribed 
         AutoSize        =   -1  'True
         Caption         =   "Available OASIS Feeds"
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   720
         Width           =   3075
      End
      Begin VB.Label lblCategory 
         AutoSize        =   -1  'True
         Caption         =   "Category(s)"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Width           =   3075
      End
      Begin VB.Label lblHeadlines 
         AutoSize        =   -1  'True
         Caption         =   "Feed headlines"
         Height          =   255
         Left            =   90
         TabIndex        =   4
         Top             =   3435
         Width           =   3075
      End
   End
End
Attribute VB_Name = "OASISFeedHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event CategoryClick()
Public Event FeedsClick()
Public Event HeadlinesClick()

Private m_StrCategory As String
Private m_StrFeed As String
Private m_StrHeadline As String

Private m_LonCategoryID As Long

Public Property Get CategoryId() As Long
    CategoryId = m_LonCategoryID
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

Private Sub cboCategory_Click()
    m_StrCategory = cboCategory.List(cboCategory.ListIndex)
    m_LonCategoryID = cboCategory.ItemData(cboCategory.ListIndex)
    RaiseEvent CategoryClick
End Sub

Private Sub lstFeeds_Click()
    RaiseEvent FeedsClick
End Sub

Private Sub lstHeadlines_Click()
    RaiseEvent HeadlinesClick
End Sub

Private Sub UserControl_Initialize()
    m_LonCategoryID = 0
End Sub

Public Sub Init()
    OpenFeed
End Sub

Private Function OpenFeed()
        '<EhHeader>
        On Error GoTo OpenFeed_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset
    
100     RS.Open "SELECT * FROM Groups ORDER BY GroupText", m_Cnn
    
102     cboCategory.Clear
    
104     cboCategory.AddItem "---Choose your topic---"
        
        If Not RS.EOF And Not RS.Bof Then

106         RS.MoveFirst
    
108         Do While Not RS.EOF
110             cboCategory.AddItem RS.Fields("GroupText").value
112             cboCategory.ItemData(cboCategory.ListCount - 1) = RS.Fields("GroupID").value
114             RS.MoveNext
            Loop
    
116         RS.Close
    
118         cboCategory.ListIndex = 0
        
        End If
    
        Exit Function

OpenFeed_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.OpenFeed", "RSSBrowser component failure"
        '</EhFooter>
End Function
