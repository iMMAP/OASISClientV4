VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.UserControl OASISRSSBrowser 
   ClientHeight    =   10065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10620
   ScaleHeight     =   10065
   ScaleWidth      =   10620
   ToolboxBitmap   =   "OasisRSSBrowser.ctx":0000
   Begin MSComctlLib.ImageList ImgList 
      Left            =   4560
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OasisRSSBrowser.ctx":0312
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OasisRSSBrowser.ctx":08AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OasisRSSBrowser.ctx":0C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OasisRSSBrowser.ctx":0FE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OasisRSSBrowser.ctx":137A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OasisRSSBrowser.ctx":1714
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OasisRSSBrowser.ctx":1AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OasisRSSBrowser.ctx":1E48
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "OasisRSSBrowser.ctx":23E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   9765
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1746
            MinWidth        =   1411
            TextSave        =   "02-Feb-14"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12462
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   741
      BandCount       =   1
      FixedOrder      =   -1  'True
      _CBWidth        =   10620
      _CBHeight       =   420
      _Version        =   "6.7.9782"
      Caption1        =   "URL"
      Child1          =   "cboAddress"
      MinWidth1       =   600
      MinHeight1      =   360
      Width1          =   9120
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar ToolBar2 
         Height          =   330
         Left            =   10560
         TabIndex        =   3
         Top             =   15
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   582
         ButtonWidth     =   1138
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgList"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Go"
               Key             =   "Go"
               Object.ToolTipText     =   "Go to RSS feed"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboAddress 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   435
         TabIndex        =   2
         Top             =   30
         Width           =   10095
      End
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9345
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   10620
      _cx             =   18733
      _cy             =   16484
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   0
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
      _GridInfo       =   $"OasisRSSBrowser.ctx":277C
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab c1ContentTab 
         Height          =   9165
         Left            =   90
         TabIndex        =   6
         Top             =   90
         Width           =   10440
         _cx             =   18415
         _cy             =   16166
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
         Caption         =   "Content|Map|Tab&3"
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
         Flags(2)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   8790
            Left            =   11385
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   330
            Width           =   10350
            _cx             =   18256
            _cy             =   15505
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
            Appearance      =   4
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
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic2 
               Height          =   3405
               Left            =   0
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   4080
               _cx             =   7197
               _cy             =   6006
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
               Align           =   0
               AutoSizeChildren=   4
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
               GridRows        =   0
               GridCols        =   0
               Frame           =   3
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
               Begin VB.ListBox lstHeadlines 
                  Height          =   1620
                  Left            =   90
                  TabIndex        =   12
                  Top             =   90
                  Width           =   3900
               End
               Begin VB.ComboBox cboCategory 
                  Height          =   315
                  Left            =   90
                  TabIndex        =   11
                  Text            =   "C:\RSS"
                  Top             =   1920
                  Width           =   3900
               End
               Begin VB.ListBox lstFeeds 
                  Height          =   255
                  Left            =   90
                  TabIndex        =   10
                  Top             =   2550
                  Width           =   3900
               End
               Begin VB.Label lblHeadlines 
                  AutoSize        =   -1  'True
                  Caption         =   "Feed headlines"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   15
                  Top             =   2865
                  Width           =   3900
               End
               Begin VB.Label lblCategory 
                  AutoSize        =   -1  'True
                  Caption         =   "Category(s)"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   14
                  Top             =   2295
                  Width           =   3900
               End
               Begin VB.Label lblSubscribed 
                  AutoSize        =   -1  'True
                  Caption         =   "Available OASIS Feeds"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   13
                  Top             =   3120
                  Width           =   3900
               End
            End
         End
         Begin SHDocVwCtl.WebBrowser webFeeds 
            Height          =   8790
            Left            =   45
            TabIndex        =   7
            Top             =   330
            Width           =   10350
            ExtentX         =   18256
            ExtentY         =   15505
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
            AutoArrange     =   0   'False
            NoClientEdge    =   0   'False
            AlignLeft       =   0   'False
            NoWebView       =   0   'False
            HideFileNames   =   0   'False
            SingleClick     =   0   'False
            SingleSelection =   0   'False
            NoFolders       =   0   'False
            Transparent     =   0   'False
            ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
            Location        =   ""
         End
         Begin C1SizerLibCtl.C1Elastic elMap 
            Height          =   8790
            Left            =   11085
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   330
            Width           =   10350
            _cx             =   18256
            _cy             =   15505
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
            BorderWidth     =   1
            ChildSpacing    =   0
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
            GridRows        =   2
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"OasisRSSBrowser.ctx":27B4
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin TatukGIS_XDK10.XGIS_ViewerWnd GIS_RSS 
               Height          =   8385
               Left            =   15
               TabIndex        =   21
               Top             =   390
               Width           =   10320
               BigExtentMargin =   -10
               RestrictedDrag  =   -1  'True
               CachedPaint     =   -1  'True
               IncrementalPaint=   -1  'True
               FullPaint       =   -1  'True
               CodePage        =   0
               OutCodePage     =   0
               CharSet         =   0
               UseRTree        =   0   'False
               PrinterTileSize =   2700
               PrintTitle      =   ""
               PrintSubtitle   =   ""
               PrintFooter     =   ""
               BeginProperty PrintTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty PrintSubtitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty PrintFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PrintTitleFontColor=   -16777208
               PrintSubtitleFontColor=   -16777208
               PrintFooterFontColor=   -16777208
               SelectionColor  =   16777215
               SelectionPattern=   "OasisRSSBrowser.ctx":27F7
               SelectionTransparency=   100
               SelectionWidth  =   100
               SelectionOutlineOnly=   0   'False
               OldCachedPaint  =   0   'False
               PrinterModeDraft=   0   'False
               PrinterModeForceBitmap=   0   'False
               GDIType         =   0
               ScaleAsFloat    =   1
               Mode            =   0
               BorderStyle     =   1
               CursorForDrag   =   0
               CursorForSelect =   0
               CursorForZoom   =   0
               CursorForEdit   =   0
               MinZoomSize     =   -5
               ScrollBars      =   0
               AutoCenter      =   0   'False
               Align           =   0
               Ctl3D           =   -1  'True
               Object.Visible         =   -1  'True
               Cursor          =   16
               DoubleBuffered  =   0   'False
               ModeMouseButton =   0
               CursorForUserDefined=   0
               View3D          =   0   'False
            End
            Begin C1SizerLibCtl.C1Elastic elMapTools 
               Height          =   375
               Left            =   15
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   15
               Width           =   10320
               _cx             =   18203
               _cy             =   661
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
               GridRows        =   0
               GridCols        =   0
               Frame           =   3
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
               Begin VB.CommandButton cmdGEORSS 
                  Height          =   375
                  Left            =   60
                  MaskColor       =   &H000000FF&
                  Picture         =   "OasisRSSBrowser.ctx":2859
                  Style           =   1  'Graphical
                  TabIndex        =   20
                  Top             =   0
                  UseMaskColor    =   -1  'True
                  Width           =   375
               End
               Begin VB.CommandButton cmdNavMode 
                  Height          =   375
                  Left            =   420
                  MaskColor       =   &H00FF00FF&
                  Picture         =   "OasisRSSBrowser.ctx":2C51
                  Style           =   1  'Graphical
                  TabIndex        =   19
                  Top             =   0
                  UseMaskColor    =   -1  'True
                  Width           =   375
               End
               Begin VB.CommandButton cmdPan 
                  Height          =   375
                  Left            =   780
                  MaskColor       =   &H00FF00FF&
                  Picture         =   "OasisRSSBrowser.ctx":2F95
                  Style           =   1  'Graphical
                  TabIndex        =   18
                  Top             =   0
                  UseMaskColor    =   -1  'True
                  Width           =   375
               End
            End
         End
      End
   End
   Begin VB.Label lblDefault 
      AutoSize        =   -1  'True
      Caption         =   "To open this feed in your default browser, click here."
      Height          =   195
      Left            =   7020
      TabIndex        =   4
      Top             =   8460
      Width           =   4545
   End
End
Attribute VB_Name = "OASISRSSBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private oRSS As MSXML2.DOMDocument
Private oItemList() As MSXML2.IXMLDOMNode

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Private strURL As String
Private strFeed As String
Private strPubDate As String
Private strHeadlines As String
Private FeedURL As String
Private strFeedImage As String

' Global FileSystemObject settings
Private FSys As New FileSystemObject
Private FSysFile As Object
Private FSysFolder As Object

'Private WithEvents m_frmSysTray As frmSysTray
Private CN As New ADODB.Connection
Private m_RSFeed As ADODB.Recordset
Private m_StrUserGroupPrefix As String
Private m_StrServerURL As String
Private oLyr As New TatukGIS_XDK10.XGIS_LayerVector
Private m_LonCategoryID As Long

Public Property Get CategoryId() As Long
    CategoryId = m_LonCategoryID
End Property

Public Property Let CategoryId(ByVal LonValue As Long)
    m_LonCategoryID = LonValue
End Property

Public Property Let ImageURL(ByVal sURL As String)
    strFeedImage = sURL
End Property

Public Property Get ServerURL() As String
    ServerURL = m_StrServerURL
End Property

Public Property Let ServerURL(ByVal StrValue As String)
    m_StrServerURL = StrValue
End Property

Public Property Get UserGroupPrefix() As String
    UserGroupPrefix = m_StrUserGroupPrefix
End Property

Public Property Let UserGroupPrefix(ByVal StrValue As String)
    m_StrUserGroupPrefix = StrValue
End Property

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
100     If Not m_StrUserGroupPrefix = "" Then
                
102         Set rsRemote = New ADODB.Recordset
    
104         'rsRemote.Open m_StrServerURL & "?ID=" & CheckEncrypt("SELECT * FROM " & m_StrUserGroupPrefix & "Feeds"), , adOpenDynamic, adLockOptimistic
            'Set rsRemote = OpenSilentHttpCommsRS(m_StrServerURL & "?ID=SELECT * FROM " & m_StrUserGroupPrefix & "Feeds", False)
            Set rsRemote = OpenServerRSCompressed(m_StrServerURL, "ID", "SELECT * FROM " & m_StrUserGroupPrefix & "Feeds")

106         If Not rsRemote.State = 0 Then

108             CN.Execute "delete from Feeds"

110             If rsRemote.EOF And rsRemote.Bof Then
                    Exit Sub
                End If
            
112             Set RS = New ADODB.Recordset
                
114             rsRemote.MoveFirst
    
116             RS.Open "SELECT * FROM Feeds", CN, adOpenDynamic, adLockBatchOptimistic
    
118             Do While Not rsRemote.EOF
120                 RS.AddNew
    
122                 For j = 1 To rsRemote.Fields.Count - 1

124                     If Not IsNull(rsRemote.Fields.Item(j).value) Then
126                         'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
                            RS.Fields.Item(j).value = rsRemote.Fields(RS.Fields.Item(j).Name).value
                        End If

                    Next
                    
                    RS.Fields.Item("FeedID").value = i
                    
128                 RS.UpdateBatch
                    i = i + 1
130                 rsRemote.MoveNext
                Loop
    
132             rsRemote.Close
134             RS.Close
    
                '136             Set rsRemote = Nothing
                '138             Set RS = Nothing
                '
                '140             Set rsRemote = New ADODB.Recordset
                '
                '142             rsRemote.Open m_StrServerURL & "?ID=" & CheckEncrypt("SELECT SettingValue9 FROM " & m_StrUserGroupPrefix & "AppSettings WHERE SettingName = 'ProfileSettings'"), , adOpenDynamic, adLockOptimistic
                '
                '144             If Not rsRemote.State = 0 Then
                '
                '146                 Set RSUpdater = New ADODB.Recordset
                '148                 With RSUpdater
                '
                '150                     .Open "SELECT * FROM AppSettings", cn, adOpenDynamic, adLockBatchOptimistic
                '152                     .Find "SettingName = 'ProfileSettings'"
                '
                '154                     If Not .EOF Then
                '156                         .Fields("SettingValue9").Value = IIf(IsNull(rsRemote.Fields.Item("SettingValue9").Value), "null", "'" & rsRemote.Fields.Item("SettingValue9").Value & "'")
                '158                         .UpdateBatch adAffectCurrent
                '160                         .Close
                '                        End If
                '
                '                    End With
                '162                 Set RSUpdater = Nothing
                '
                '164                 rsRemote.Close
                '166                 Set rsRemote = Nothing
                ' End If
                
                SynchProfileSettingWithServer "SettingValue5", m_StrUserGroupPrefix, m_Cnn
            End If
                
        End If
    
        '<EhFooter>
        Exit Sub

UpdateFeeds_Err:
        Err.Raise vbObjectError + 100, "OASISClient.OASISRSSBrowser.UpdateFeeds", "OASISRSSBrowser component failure"
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
100     If Not m_StrUserGroupPrefix = "" Then
                
102         Set rsRemote = New ADODB.Recordset
            
            'Set rsRemote = OpenSilentHttpCommsRS(m_StrServerURL & "?ID=SELECT * FROM " & m_StrUserGroupPrefix & "FeedGroups", False)
104         Set rsRemote = OpenServerRSCompressed(m_StrServerURL, "ID", "SELECT * FROM " & m_StrUserGroupPrefix & "FeedGroups")
            'rsRemote.Open m_StrServerURL & "?ID=" & CheckEncrypt("SELECT * FROM " & m_StrUserGroupPrefix & "FeedGroups"), , adOpenDynamic, adLockOptimistic
    
106         If Not rsRemote.State = 0 Then

108             CN.Execute "delete from Groups"
110             DebugPrint Err.Description

112             If rsRemote.EOF And rsRemote.Bof Then
                    Exit Function
                End If
            
114             Set RS = New ADODB.Recordset
                
116             rsRemote.MoveFirst
    
118             RS.Open "SELECT * FROM Groups", CN, adOpenDynamic, adLockBatchOptimistic
    
120             Do While Not rsRemote.EOF
122                 RS.AddNew
    
124                 For j = 0 To rsRemote.Fields.Count - 1
126                     If Not IsNull(rsRemote.Fields.Item(j).value) Then
128                         'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
                            RS.Fields.Item(j).value = rsRemote.Fields(RS.Fields.Item(j).Name).value
                        End If
                    Next
130                  RS.UpdateBatch
132                 rsRemote.MoveNext
                Loop
    
134             rsRemote.Close
136             RS.Close
    
138             Set rsRemote = Nothing
140             Set RS = Nothing
            End If
                
        End If
        
        UpdateFeedGroups = True
        
        '<EhFooter>
        Exit Function

UpdateFeedGroups_Err:
        MsgBox vbObjectError + 100, _
                  "OASISRssTool.RSSBrowser.UpdateFeedGroups", _
                  "RSSBrowser component failure"
            Resume Next
        '</EhFooter>
End Function

Public Sub Init2(sConnectionString As String, _
                 Optional bUpdate As Boolean)
        '<EhHeader>
        On Error GoTo Init2_Err
        '</EhHeader>
100     CN.CursorLocation = g_sGlobalCursorLocation 'This was adUseServer
102     CN.Open sConnectionString
        
104     oLyr.AddField "Link", TatukGIS_XDK10.XgisFieldTypeString, 255, 0
106     oLyr.AddField "Title", TatukGIS_XDK10.XgisFieldTypeString, 255, 0
108     oLyr.AddField "Description", TatukGIS_XDK10.XgisFieldTypeString, 255, 0
        GIS_RSS.Open g_sAppPath & "\data\user\Maps\DefaultMap.TTKGP", False
110     GIS_RSS.Add oLyr

        '104     If bUpdate Then
        '106         If UpdateFeedGroups Then UpdateFeeds
        '        End If
114     StatusBar.Panels(1).Width = 100

116     Call CheckHTML

        WriteHTML2

118     webFeeds.Navigate g_sAppPath & "\Data\HTML\RSSIntro.html"

        '<EhFooter>
        Exit Sub

Init2_Err:
       'Err.Raise vbObjectError + 100, _
       '           "OASISClient.OASISRSSBrowser.Init2", _
       '           "OASISRSSBrowser component failure"
                  Err.Clear
        '</EhFooter>
End Sub

Public Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
    
100     CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Oasisclient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"  '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb"
                     
        ' Set a few bits up.
102     CoolBar2.Bands(2).MinWidth = 700
104     StatusBar.Panels(1).Width = 100
        'fileFeeds.Pattern = "*"
    
        'TODO
        'cboCategory.AddItem "C:\RSS"
    
        ' Check the HTML directory is there.
106     Call CheckHTML
        ' Call SystemTray to enable system tray goodness.
108     Call SystemTray
        ' Write the HTML file incase it has been deleted.
110     Call WriteHTML2 'Call WriteHTML
        ' Write the HTML file incase it has been delete.
112     Call WriteFeeds
        ' Get the directory where all the saved feeds will be stored.
114     Call GetDirectory
        ' Now populate the Category Combo with the sub directories in C:\RSS
116     Call FillCategory
    
        ' Now navigate to the default HTML page.
        '
    
118     Call OpenFeed
120     Call GetRSS
122     webFeeds.Navigate g_sAppPath & "\Data\HTML\RSSIntro.html" 'App.Path & "\HTML\RSS.html"
        '<EhFooter>
        Exit Sub

Init_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.Init", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Function GetRSS()
        '<EhHeader>
        On Error GoTo GetRSS_Err
        '</EhHeader>
    
        ' This just makes sure everything is nice and clean.
100     lstHeadlines.Clear
102     webFeeds.Navigate "about:blank"
104     strHeadlines = ""
106     strURL = ""
108     strFeed = ""
110     strPubDate = ""

112     DoEvents
    
        ' Disbale fileFeeds and let the user know we are getting the feeds.
        'fileFeeds.Enabled = False
114     StatusBar.Panels(2).Text = "Please wait getting feeds..."

116     DoEvents
    
        Dim oItems As MSXML2.IXMLDOMNodeList
        Dim i As Integer
        Dim oNode As IXMLDOMNode
    
118     Set oRSS = New MSXML2.DOMDocument
120     oRSS.async = False
122     oRSS.Load (cboAddress.Text)
    
124     Set oItems = oRSS.selectNodes("rss/channel/item")

126     i = -1
    
128     ReDim oItemList(oItems.Length)
    
130     For Each oNode In oItems
132         i = i + 1
134         lstHeadlines.AddItem oNode.selectSingleNode("title").Text
136         Set oItemList(i) = oNode
138     Next oNode
    
        ' Let the user know we are done.
        'fileFeeds.Enabled = True
140     StatusBar.Panels(2).Text = "Retrieved " & lstHeadlines.ListCount & " feeds."

        Dim m_RSFeed As New ADODB.Recordset
142     m_RSFeed.Open "SELECT FeedImageURL FROM Feeds WHERE FeedID = " & lstFeeds.ItemData(lstFeeds.ListIndex), CN, adOpenDynamic, adLockOptimistic
        'm_RSFeed.Find "GroupID =" & lstFeeds.ItemData(lstFeeds.ListIndex)
        DebugPrint "SELECT FeedImageURL FROM Feeds WHERE FeedID = " & lstFeeds.ItemData(lstFeeds.ListIndex)

144     If m_RSFeed.EOF Then
146         strFeedImage = ""
        Else

            If Not IsNull(m_RSFeed.Fields(0).value) Then
148             strFeedImage = m_RSFeed.Fields(0).value
            End If
        End If

150     m_RSFeed.Close
152     Set m_RSFeed = Nothing

154     DoEvents
        '<EhFooter>
        Exit Function

GetRSS_Err:
        Err.Raise vbObjectError + 100, "OASISClient.OASISRSSBrowser.GetRSS", "OASISRSSBrowser component failure"
        '</EhFooter>
End Function

Private Function GetHeadlines()
        '<EhHeader>
        On Error GoTo GetHeadlines_Err
        '</EhHeader>

        Dim oNode As MSXML2.IXMLDOMNode
100     Set oNode = oItemList(lstHeadlines.ListIndex)

102     strHeadlines = oNode.selectSingleNode("title").Text
104     strURL = oNode.selectSingleNode("link").Text
106     strFeed = oNode.selectSingleNode("description").Text
        
        Dim oTestNode As MSXML2.IXMLDOMNode
108     Set oTestNode = oNode.selectSingleNode("pubDate|dc:date")

110     If oTestNode Is Nothing Then
    
            ' If the Node does not exist we simply display No Information.
112         strPubDate = "No Information"
        
        Else
    
            ' Next statement finds the first node matching any of the tags in the list
114         strPubDate = oNode.selectSingleNode("pubDate|dc:date").Text
        
        End If
       
        ' Now we write the info to the HTML page.
116     Call WriteFeeds

        ' Now we display the web page to the user.
118     webFeeds.Navigate g_sAppPath & "\data\user\Exports\Feeds\Feeds.html"  'App.Path & "\HTML\Feeds.html"
        
        '<EhFooter>
        Exit Function

GetHeadlines_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.GetHeadlines", "RSSBrowser component failure"
        '</EhFooter>
End Function

Private Function OpenFeed()
        '<EhHeader>
        On Error GoTo OpenFeed_Err
        '</EhHeader>

        ' This just put the address of you favorite feeds in the address bar
        ' when you open them.
        Dim InStream As TextStream
   
        Dim RS As New ADODB.Recordset
    
100     RS.Open "SELECT * FROM Groups ORDER BY GroupText", CN
    
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

        '    Call CheckAddress
    
        Exit Function
    
        ' Set InStream = FSys.OpenTextFile(frmMain.cboCategory & "\" & fileFeeds.FileName)
    
        Dim noth As String
120     FeedURL = InStream.ReadLine
122     InStream.Close '< -EOF
        
124     cboAddress.Text = FeedURL
126     cboAddress.Text = Replace(cboAddress.Text, Chr(10), "")
128     cboAddress.Text = Replace(cboAddress.Text, Chr(13), "")
130     cboAddress.Text = Trim(cboAddress.Text)
        '  fileFeeds.Refresh
    
        ' Now we check that the address has gone in ok.
132     Call CheckAddress
    
        '<EhFooter>
        Exit Function

OpenFeed_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.OpenFeed", "RSSBrowser component failure"
        '</EhFooter>
End Function

Private Function CheckHTML()
        '<EhHeader>
        On Error GoTo CheckHTML_Err
        '</EhHeader>

        ' Make sure the HTML directory is there.
100     If FSys.FolderExists(g_sAppPath & "\Data\HTML") Then
            Exit Function
        Else
102         FSys.CreateFolder (g_sAppPath & "\Data\HTML")
        End If

        '<EhFooter>
        Exit Function

CheckHTML_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.CheckHTML", "RSSBrowser component failure"
        '</EhFooter>
End Function

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
        On Error Resume Next
100     m_RSFeed.Close
102     CN.Close
104     Set m_RSFeed = Nothing
106     Set CN = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.Form_Unload", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Sub cmdGEORSS_Click()
    GIS_RSS.Mode = TatukGIS_XDK10.XgisSelect
End Sub

Private Sub cmdNavMode_Click()
    GIS_RSS.Mode = TatukGIS_XDK10.XgisZoomEx
End Sub

Private Sub cmdPan_Click()
    GIS_RSS.Mode = TatukGIS_XDK10.XgisDrag
End Sub

Private Sub GIS_RSS_OnContextPopup(translated As Boolean, ByVal mousePos As TatukGIS_XDK10.IXPoint, handled As Boolean)
    DebugPrint ""
End Sub

Private Sub GIS_RSS_OnMouseDown(translated As Boolean, _
                                ByVal Button As TatukGIS_XDK10.XMouseButton, _
                                ByVal Shift As TatukGIS_XDK10.XShiftState, _
                                ByVal X As Long, _
                                ByVal Y As Long)
        Dim shp As TatukGIS_XDK10.XGIS_Shape
        Dim ptg As TatukGIS_XDK10.XGIS_Point

        Set ptg = GIS_RSS.ScreenToMap(GisUtils.POINT(X, Y))

        If GIS_RSS.Mode = TatukGIS_XDK10.XgisSelect Then
118         Set shp = oLyr.Locate(ptg, 5 / GIS_RSS.zoom) ' 5 pixels precision
            
120         If Not shp Is Nothing Then
                webFeeds.Navigate shp.GetField("Link")
                c1ContentTab.CurrTab = 0
            End If
        End If

End Sub

Private Sub lstFeeds_Click()
        '<EhHeader>
        On Error GoTo lstFeeds_Click_Err
        '</EhHeader>
100     m_RSFeed.MoveFirst
102     m_RSFeed.Find "FeedID = " & lstFeeds.ItemData(lstFeeds.ListIndex)
        
104     cboAddress.Text = m_RSFeed.Fields("FeedURL").value
106     Call GetRSS
108     webFeeds.Navigate g_sAppPath & "\Data\HTML\RSSIntro.html"
        '<EhFooter>
        Exit Sub

lstFeeds_Click_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.lstFeeds_Click", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Sub lstFeeds_ItemCheck(Item As Integer)
    'Me.Caption = "VVAFs RSS Reader - " & cboCategory.List(cboCategory.ListIndex) & "\" & lstFeeds.List(Item)
End Sub

Private Sub lstHeadlines_Click()
        '<EhHeader>
        On Error GoTo lstHeadlines_Click_Err
        '</EhHeader>
    
        ' This gets the headlines and displayed them in webFeeds.
100     Call GetHeadlines
    
        '<EhFooter>
        Exit Sub

lstHeadlines_Click_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.lstHeadlines_Click", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Function GetDirectory()
    '
    '    ' Make sure that C:\RSS is there, if not create it.
    '    If FSys.FolderExists(cboCategory) Then
    '        fileFeeds.Path = cboCategory
    '    Else
    '        FSys.CreateFolder ("C:\RSS")
    '        fileFeeds.Path = cboCategory
    '    End If
    
End Function

Private Sub FillCategory()

End Sub

Private Function DeleteFeed()


    
End Function

Private Function DeleteFolder()


End Function

Private Sub cboAddress_Click()
        '<EhHeader>
        On Error GoTo cboAddress_Click_Err
        '</EhHeader>

        ' This calls GetRSS which gets the feed from the internet.
100     Call GetRSS

        '<EhFooter>
        Exit Sub

cboAddress_Click_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.cboAddress_Click", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Sub cboAddress_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo cboAddress_KeyPress_Err
        '</EhHeader>
    
        ' Call GetRSS when return is pressed.
        On Error Resume Next

100     If KeyAscii = vbKeyReturn Then
102         Call GetRSS
        End If
    
        '<EhFooter>
        Exit Sub

cboAddress_KeyPress_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.cboAddress_KeyPress", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Public Sub GetGEOPointsEX(oItems As MSXML2.IXMLDOMNodeList)
        '<EhHeader>
        On Error GoTo GetGEOPoints_Err
        '</EhHeader>
        Dim i As Integer
        Dim oNode As MSXML2.IXMLDOMNode
        
        Dim ptg As TatukGIS_XDK10.XGIS_Point
        Dim oSHP As TatukGIS_XDK10.XGIS_Shape
        
        Dim sFont As String
        Dim SymbolList As TatukGIS_XDK10.XGIS_SymbolList
        Dim ostr As New XStringList
                
100     GIS_RSS.Mode = TatukGIS_XDK10.XgisZoomEx

114     For Each oNode In oItems

116         strHeadlines = oNode.selectSingleNode("title").Text
118         strURL = oNode.selectSingleNode("link").Text
120         strFeed = oNode.selectSingleNode("description").Text

122         Set ptg = New TatukGIS_XDK10.XGIS_Point
124         ptg.Prepare CDbl(oNode.selectSingleNode("geo:long").Text), CDbl(oNode.selectSingleNode("geo:lat").Text)
126         Set oSHP = oLyr.CreateShape(XgisShapeTypePoint)
            
128         With oSHP
130             .Lock TatukGIS_XDK10.XgisLockExtent
132             .AddPart
134             .AddPoint ptg

136             With .Params.labels
138                 .Alignment = TatukGIS_XDK10.XgisLabelAlignmentLeftJustify
140                 .color = vbRed
142                 .FontColor = vbWhite
144                 .Allocator = False
'146                 .Duplicates = True
148                 .Font.Size = 8
150                 .OutlineWidth = 0
152                 .Pattern = XbsClear
154                 .Position = TatukGIS_XDK10.XgisLabelPositionDownRight
156                 .value = strHeadlines
                End With
                
158             sFont = "MapInfo Miscellaneous"
160             sFont = sFont & ":" & Asc(",") & ":NORMAL"

162             Set SymbolList = New TatukGIS_XDK10.XGIS_SymbolList

164             .Params.Marker.color = vbBlue
166             .Params.Marker.OutlineColor = 16711680
168             .Params.Marker.Symbol = SymbolList.Prepare(sFont)
170             .Params.Marker.Size = 250
                '172                     .Params.Marker.OutlineStyle = XpsSolid
                '174                     .Params.Marker.OutlineWidth = 20
                '176                     .Params.Marker.OutlineColor = vbRed
                '178                     .Params.Marker.Style = TatukGIS_XDK10.XGISMarkerStyleTriangleDown
                '180                     .Params.Marker.Color = vbBlue
                '182                     .Params.Marker.Size = 200
                        
                '.Params.Marker.SaveToStrings ostr
                        
                'On Error Resume Next
                        
172             .SetField "Title", strHeadlines
174             .SetField "Link", strURL
176             .SetField "Description", strFeed
                
178             .Unlock
            End With
            
'180         GIS_RSS.VisibleExtent = oLyr.Extent
            
182         DebugPrint oNode.selectSingleNode("geo:lat").Text
184         DebugPrint oNode.selectSingleNode("geo:long").Text
        
186     Next oNode
        
188     GIS_RSS.UpDate
        
        Exit Sub
        
        Dim oTestNode As MSXML2.IXMLDOMNode
190     Set oTestNode = oNode.selectSingleNode("pubDate|dc:date")

192     If oTestNode Is Nothing Then
    
            ' If the Node does not exist we simply display No Information.
194         strPubDate = "No Information"
        
        Else
    
            ' Next statement finds the first node matching any of the tags in the list
196         strPubDate = oNode.selectSingleNode("pubDate|dc:date").Text
        
        End If
       
        ' Now we write the info to the HTML page.
198     Call WriteFeeds

        ' Now we display the web page to the user.
200     webFeeds.Navigate g_sAppPath & "\data\user\Exports\Feeds\Feeds.html"  'App.Path & "\HTML\Feeds.html"
202     cboAddress.Text = strURL
        '<EhFooter>
        Exit Sub

GetGEOPoints_Err:
        MsgBox Erl & Err.Description & "OASISClient.OASISRSSBrowser.GetGEOPoints", _
                  "OASISRSSBrowser component failure"
        '</EhFooter>
End Sub

Public Sub GetGEOPoints(SFeed As String)
        '<EhHeader>
        On Error GoTo GetGEOPoints_Err
        '</EhHeader>
        Dim oRSS As New MSXML2.DOMDocument
        Dim i As Integer
        Dim oNode As MSXML2.IXMLDOMNode
        Dim oItems As MSXML2.IXMLDOMNodeList
        Dim ptg As TatukGIS_XDK10.XGIS_Point
        Dim oSHP As TatukGIS_XDK10.XGIS_Shape
        
        Dim sFont As String
        Dim SymbolList As TatukGIS_XDK10.XGIS_SymbolList
        Dim ostr As New XStringList
                
100     GIS_RSS.Mode = TatukGIS_XDK10.XgisZoomEx
        oRSS.async = False
        oRSS.Load SFeed
'110     oRSS.loadXML SFeed
    
112     Set oItems = oRSS.selectNodes("rss/channel/item")
    
114     For Each oNode In oItems

116         strHeadlines = oNode.selectSingleNode("title").Text
118         strURL = oNode.selectSingleNode("link").Text
120         strFeed = oNode.selectSingleNode("description").Text

122         Set ptg = New TatukGIS_XDK10.XGIS_Point
124         ptg.Prepare CDbl(oNode.selectSingleNode("geo:long").Text), CDbl(oNode.selectSingleNode("geo:lat").Text)
126         Set oSHP = oLyr.CreateShape(XgisShapeTypePoint)
            
128         With oSHP
130             .Lock TatukGIS_XDK10.XgisLockExtent
132             .AddPart
134             .AddPoint ptg

136             With .Params.labels
138                 .Alignment = TatukGIS_XDK10.XgisLabelAlignmentLeftJustify
140                 .color = vbRed
142                 .FontColor = vbWhite
144                 .Allocator = False
'146                 .Duplicates = True
148                 .Font.Size = 8
150                 .OutlineWidth = 0
152                 .Pattern = XbsClear
154                 .Position = TatukGIS_XDK10.XgisLabelPositionDownRight
156                 .value = strHeadlines
                End With
                
158             sFont = "MapInfo Miscellaneous"
160             sFont = sFont & ":" & Asc(",") & ":NORMAL"

162             Set SymbolList = New TatukGIS_XDK10.XGIS_SymbolList

164             .Params.Marker.color = vbBlue
166             .Params.Marker.OutlineColor = 16711680
168             .Params.Marker.Symbol = SymbolList.Prepare(sFont)
170             .Params.Marker.Size = 250
                '172                     .Params.Marker.OutlineStyle = XpsSolid
                '174                     .Params.Marker.OutlineWidth = 20
                '176                     .Params.Marker.OutlineColor = vbRed
                '178                     .Params.Marker.Style = TatukGIS_XDK10.XGISMarkerStyleTriangleDown
                '180                     .Params.Marker.Color = vbBlue
                '182                     .Params.Marker.Size = 200
                        
                '.Params.Marker.SaveToStrings ostr
                        
                'On Error Resume Next
                        
172             .SetField "Title", strHeadlines
174             .SetField "Link", strURL
176             .SetField "Description", strFeed
                
178             .Unlock
            End With
            
'180         GIS_RSS.VisibleExtent = oLyr.Extent
            
182         DebugPrint oNode.selectSingleNode("geo:lat").Text
184         DebugPrint oNode.selectSingleNode("geo:long").Text
        
186     Next oNode
        
188     GIS_RSS.UpDate
        
        Exit Sub
        
        Dim oTestNode As MSXML2.IXMLDOMNode
190     Set oTestNode = oNode.selectSingleNode("pubDate|dc:date")

192     If oTestNode Is Nothing Then
    
            ' If the Node does not exist we simply display No Information.
194         strPubDate = "No Information"
        
        Else
    
            ' Next statement finds the first node matching any of the tags in the list
196         strPubDate = oNode.selectSingleNode("pubDate|dc:date").Text
        
        End If
       
        ' Now we write the info to the HTML page.
198     Call WriteFeeds

        ' Now we display the web page to the user.
200     webFeeds.Navigate g_sAppPath & "\data\user\Exports\Feeds\Feeds.html"  'App.Path & "\HTML\Feeds.html"
202     cboAddress.Text = strURL
        '<EhFooter>
        Exit Sub

GetGEOPoints_Err:
        MsgBox Erl & Err.Description & "OASISClient.OASISRSSBrowser.GetGEOPoints", _
                  "OASISRSSBrowser component failure"
        '</EhFooter>
End Sub

Public Sub LoadHeader(oNode As MSXML2.IXMLDOMNode)
        '<EhHeader>
        On Error GoTo GetHeadlines_Err
        '</EhHeader>


102     strHeadlines = oNode.selectSingleNode("title").Text
104     strURL = oNode.selectSingleNode("link").Text
106     strFeed = oNode.selectSingleNode("description").Text
        
        Dim oTestNode As MSXML2.IXMLDOMNode
108     Set oTestNode = oNode.selectSingleNode("pubDate|dc:date")

110     If oTestNode Is Nothing Then
    
            ' If the Node does not exist we simply display No Information.
112         strPubDate = "No Information"
        
        Else
    
            ' Next statement finds the first node matching any of the tags in the list
114         strPubDate = oNode.selectSingleNode("pubDate|dc:date").Text
        
        End If
       
        ' Now we write the info to the HTML page.
116     Call WriteFeeds

        ' Now we display the web page to the user.
118     webFeeds.Navigate g_sAppPath & "\data\user\Exports\Feeds\Feeds.html"  'App.Path & "\HTML\Feeds.html"
        cboAddress.Text = strURL
        '<EhFooter>
        Exit Sub

GetHeadlines_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.GetHeadlines", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Public Sub UpdateStatusBar(sMess As String)
    StatusBar.Panels(2).Text = sMess
End Sub

Public Sub UpdateHeaderList(sURL As String, bUseGeo As Boolean, oItems As MSXML2.IXMLDOMNodeList, sCountryISO As String)
        '<EhHeader>
        On Error GoTo UpdateHeaderList_Err
        '</EhHeader>
    
100     cboAddress.Text = ""
102     GetGEOPointsEX oItems '"http://ws.geonames.org/rssToGeoRSS?feedUrl=" & sURL
104     webFeeds.Navigate g_sAppPath & "\Data\HTML\RSSIntro.html"

        '<EhFooter>
        Exit Sub

UpdateHeaderList_Err:
        Err.Clear
        Exit Sub
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISRSSBrowser.UpdateHeaderList", _
                  "OASISRSSBrowser component failure"
        '</EhFooter>
End Sub

Public Sub UpdateFeedList()

104     cboAddress.Text = ""
106     webFeeds.Navigate "about:blank"
108     strHeadlines = ""
110     strURL = ""
112     strFeed = ""
114     strPubDate = ""

116     DoEvents
    
132     webFeeds.Navigate g_sAppPath & "\Data\HTML\RSSIntro.html"
        Exit Sub
End Sub


Private Sub cboCategory_Click()
        '<EhHeader>
        On Error GoTo cboCategory_Click_Err

        '</EhHeader>
100     If cboCategory.List(cboCategory.ListIndex) = "---Choose your topic---" Then Exit Sub
        ' Tidy up a bit.
102     lstHeadlines.Clear
104     cboAddress.Text = ""
106     webFeeds.Navigate "about:blank"
108     strHeadlines = ""
110     strURL = ""
112     strFeed = ""
114     strPubDate = ""

116     DoEvents

118     Set m_RSFeed = New ADODB.Recordset
    
120     m_RSFeed.Open "SELECT * FROM Feeds WHERE GroupID =" & cboCategory.ItemData(cboCategory.ListIndex), CN, adOpenDynamic, adLockOptimistic
    
122     lstFeeds.Clear
    
124     Do While Not m_RSFeed.EOF
126         lstFeeds.AddItem m_RSFeed.Fields("FeedName").value
128         lstFeeds.ItemData(lstFeeds.ListCount - 1) = m_RSFeed.Fields("FeedID").value
130         m_RSFeed.MoveNext
        Loop
    
132     webFeeds.Navigate g_sAppPath & "\Data\HTML\RSSIntro.html"
        Exit Sub
     
        '    ' Set the fileFeeds path to that listed in cboCategory.
        '    If FSys.FolderExists(cboCategory) Then
        '        fileFeeds.Path = cboCategory
        '    End If
    
        '<EhFooter>
        Exit Sub

cboCategory_Click_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.cboCategory_Click", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Sub ImgURL_Click()
        '<EhHeader>
        On Error GoTo ImgURL_Click_Err
        '</EhHeader>
    
        ' This is for users who do not want the full article to be opened in IE.
        ' this just opens the URL in the users default browser when the button is clicked.
        Dim retValue As Long
    
100     If strURL <> "" Then
102         retValue = ShellExecute(UserControl.hwnd, "Open", strURL, 0&, 0&, 0&)
        Else
104         MsgBox "No feed to open!", vbInformation, "No feed available"
        End If
    
        '<EhFooter>
        Exit Sub

ImgURL_Click_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.ImgURL_Click", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Function CheckAddress()
        '<EhHeader>
        On Error GoTo CheckAddress_Err
        '</EhHeader>
    
        ' This just checks that the address is not blank when called.
100     If cboAddress.Text = "" Then
102         MsgBox "No URL to open!", vbInformation, "No URL"
            Exit Function
        End If

        '<EhFooter>
        Exit Function

CheckAddress_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.CheckAddress", "RSSBrowser component failure"
        '</EhFooter>
End Function

Private Sub About_Click()
        '<EhHeader>
        On Error GoTo About_Click_Err
        '</EhHeader>
    
        ' Display about information when About is clicked.
100     MsgBox "iMMAPs RSS Reader v" & App.major & "." & App.minor & "." & App.Revision & vbNewLine & vbNewLine & "Written by iMMAP Europe", vbInformation, "iMMAPs RSS Reader"

        '<EhFooter>
        Exit Sub

About_Click_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.About_Click", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo Toolbar1_ButtonClick_Err
        '</EhHeader>
    
        ' This the Toolbar buttons.
        On Error Resume Next
     
100     Select Case Button.Key

            Case "Browser"
                Dim retValue As Long
    
102             If strURL <> "" Then
104                 retValue = ShellExecute(UserControl.hwnd, "Open", strURL, 0&, 0&, 0&)
                Else
106                 MsgBox "No feed to open!", vbInformation, "No feed available"
                End If

108         Case "Save"
                ' Call the save feeds form to save a feed.
110             Call SaveFeed

112         Case "Rename"

114         Case "Delete"
116             Call DeleteFeed

118         Case "Open"
120             Call OpenFeed
122             Call GetRSS

124         Case "Create"
126         Case "DeleteFolder"
128             Call DeleteFolder

130         Case "Exit"

132         Case "About"
134             MsgBox "iMMAPs RSS Reader version:" & App.major & "." & App.minor & "." & App.Revision & vbNewLine & vbNewLine & "Written by iMMAP", vbInformation, "iMMAP RSS Reader"
        End Select
    
        '<EhFooter>
        Exit Sub

Toolbar1_ButtonClick_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.ToolBar1_ButtonClick", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Function SaveFeed()
    
End Function

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo Toolbar2_ButtonClick_Err
        '</EhHeader>
    
        On Error Resume Next
     
100     Select Case Button.Key

            Case "Go"
102             Call GetRSS
        End Select
    
        '<EhFooter>
        Exit Sub

Toolbar2_ButtonClick_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.ToolBar2_ButtonClick", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Sub SystemTray()
    '<EhHeader>
    On Error GoTo SystemTray_Err
    '</EhHeader>


    '<EhFooter>
    Exit Sub

SystemTray_Err:
    Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.SystemTray", "RSSBrowser component failure"
    '</EhFooter>
End Sub

Private Sub webFeeds_StatusTextChange(ByVal Text As String)
        '<EhHeader>
        On Error GoTo webFeeds_StatusTextChange_Err
        '</EhHeader>
    
100     StatusBar.Panels(3).Text = Text

        '<EhFooter>
        Exit Sub

webFeeds_StatusTextChange_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.webFeeds_StatusTextChange", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Function WriteFeeds()
        '<EhHeader>
        On Error GoTo WriteFeeds_Err
        '</EhHeader>
    
        ' This is the HTML that will display the feed.
100     Open g_sAppPath & "\data\user\Exports\Feeds\Feeds.html" For Output As #1
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
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.WriteFeeds", "RSSBrowser component failure"
        '</EhFooter>
End Function

Private Sub WriteHTML2()
        '<EhHeader>
        On Error GoTo WriteHTML2_Err
        '</EhHeader>
        On Error Resume Next

100     Kill g_sAppPath & "\Data\HTML\RSSIntro.html"
102     Open g_sAppPath & "\Data\HTML\RSSIntro.html" For Output As #1
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
148     Print #1, "    <td width=""100%""><center><h4> - information Matters - </h4></center></td>"
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
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.WriteHTML2", "RSSBrowser component failure"
        '</EhFooter>
End Sub

Private Function WriteHTML()
        '<EhHeader>
        On Error GoTo WriteHTML_Err
        '</EhHeader>

        ' This is the HTML you see when you first open the application.
100     Open App.Path & "\HTML\RSS.html" For Output As #1
102     Print #1, "<html>"
104     Print #1, "<head>"
106     Print #1, "<title>iMMAPSs RSS Reader</title>"
108     Print #1, "<style type=""text/css"">"
110     Print #1, "<!--"
112     Print #1, "body {background-color: #F4F5F9;margin-left: 5px;margin-top: 5px;margin-right: 5px;margin-bottom: 5px;font-family: Verdana, Arial, Helvetica, sans-serif;}"
114     Print #1, ".style1 {font-size: xx-large;font-weight: bold;}"
116     Print #1, "-->"
118     Print #1, "</style></head>"
120     Print #1, "<body>"
122     Print #1, "<div align=""center"">"
124     Print #1, "<p class=""style1""><u>iMMAPs RSS Reader</u></p>"
126     Print #1, "<p>Written by iMMAP Europe</p>"
128     Print #1, "<p>v" & App.major & "." & App.minor & "." & App.Revision & "</p>"
130     Print #1, "</div>"
132     Print #1, "</body>"
134     Print #1, "</html>"
136     Close #1

        '<EhFooter>
        Exit Function

WriteHTML_Err:
        Err.Raise vbObjectError + 100, "OASISRssTool.RSSBrowser.WriteHTML", "RSSBrowser component failure"
        '</EhFooter>
End Function



