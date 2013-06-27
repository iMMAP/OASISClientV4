VERSION 5.00
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Begin VB.UserControl OASISDynamReports 
   ClientHeight    =   7320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   LockControls    =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   8115
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7320
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8115
      _cx             =   14314
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
      GridRows        =   2
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"OASISDynamReports.ctx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.PictureBox Picture1 
         Height          =   270
         Left            =   90
         ScaleHeight     =   210
         ScaleWidth      =   2205
         TabIndex        =   8
         Top             =   90
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Frame frameButtons 
         BorderStyle     =   0  'None
         Height          =   6810
         Left            =   7665
         TabIndex        =   5
         Top             =   420
         Visible         =   0   'False
         Width           =   360
         Begin CONTROLSLibCtl.dxPicBtn cmdOrientation 
            Height          =   240
            Left            =   0
            TabIndex        =   6
            ToolTipText     =   "test"
            Top             =   240
            Width           =   240
            _Version        =   65536
            _cx             =   423
            _cy             =   423
            Picture         =   "OASISDynamReports.ctx":005C
            BackColor       =   15790320
            Enabled         =   -1  'True
            Style           =   0
            DitherStyle     =   0
            DitherColor     =   255
            GroupIndex      =   -1
            Stuck           =   0   'False
            Pushed          =   0   'False
         End
         Begin CONTROLSLibCtl.dxPicBtn cmdTitle 
            Height          =   240
            Left            =   0
            TabIndex        =   7
            ToolTipText     =   "test"
            Top             =   600
            Width           =   240
            _Version        =   65536
            _cx             =   423
            _cy             =   423
            Picture         =   "OASISDynamReports.ctx":05AE
            BackColor       =   15790320
            Enabled         =   -1  'True
            Style           =   0
            DitherStyle     =   0
            DitherColor     =   255
            GroupIndex      =   -1
            Stuck           =   0   'False
            Pushed          =   0   'False
         End
      End
      Begin VB.ListBox List1 
         Height          =   6690
         Left            =   90
         TabIndex        =   1
         Top             =   420
         Width           =   2265
      End
      Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
         Height          =   6810
         Left            =   2415
         TabIndex        =   2
         Top             =   420
         Width           =   5190
         _cx             =   9155
         _cy             =   12012
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   0   'False
         MousePointer    =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         AutoRTF         =   -1  'True
         Preview         =   -1  'True
         DefaultDevice   =   0   'False
         PhysicalPage    =   -1  'True
         AbortWindow     =   -1  'True
         AbortWindowPos  =   0
         AbortCaption    =   "Printing..."
         AbortTextButton =   "Cancel"
         AbortTextDevice =   "on the %s on %s"
         AbortTextPage   =   "Now printing Page %d of"
         FileName        =   ""
         MarginLeft      =   1440
         MarginTop       =   1440
         MarginRight     =   1440
         MarginBottom    =   1440
         MarginHeader    =   0
         MarginFooter    =   0
         IndentLeft      =   0
         IndentRight     =   0
         IndentFirst     =   0
         IndentTab       =   720
         SpaceBefore     =   0
         SpaceAfter      =   0
         LineSpacing     =   100
         Columns         =   1
         ColumnSpacing   =   180
         ShowGuides      =   2
         LargeChangeHorz =   300
         LargeChangeVert =   300
         SmallChangeHorz =   30
         SmallChangeVert =   30
         Track           =   0   'False
         ProportionalBars=   -1  'True
         Zoom            =   37.8787878787879
         ZoomMode        =   3
         ZoomMax         =   400
         ZoomMin         =   10
         ZoomStep        =   25
         EmptyColor      =   -2147483636
         TextColor       =   0
         HdrColor        =   0
         BrushColor      =   0
         BrushStyle      =   0
         PenColor        =   0
         PenStyle        =   0
         PenWidth        =   0
         PageBorder      =   0
         Header          =   ""
         Footer          =   ""
         TableSep        =   "|;"
         TableBorder     =   7
         TablePen        =   0
         TablePenLR      =   0
         TablePenTB      =   0
         NavBar          =   3
         NavBarColor     =   -2147483633
         ExportFormat    =   0
         URL             =   ""
         Navigation      =   3
         NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
         AutoLinkNavigate=   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Reports"
         Height          =   270
         Left            =   90
         TabIndex        =   4
         Top             =   90
         Width           =   2265
      End
      Begin VB.Label lblReport 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "No Report Selected"
         Height          =   270
         Left            =   2415
         TabIndex        =   3
         Top             =   90
         Width           =   5610
      End
   End
   Begin VSReport8LibCtl.VSReport VSReport1 
      Left            =   120
      Top             =   2280
      _rv             =   800
      ReportName      =   "Customer Labels"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnOpen          =   ""
      OnClose         =   ""
      OnNoData        =   ""
      OnPage          =   ""
      OnError         =   ""
      MaxPages        =   0
      DoEvents        =   -1  'True
      BeginProperty Layout {D853A4F1-D032-4508-909F-18F074BD547A} 
         Width           =   3895
         MarginLeft      =   360
         MarginTop       =   720
         MarginRight     =   360
         MarginBottom    =   720
         Columns         =   3
         ColumnLayout    =   1
         Orientation     =   0
         PageHeader      =   0
         PageFooter      =   0
         PictureAlign    =   7
         PictureShow     =   1
         PaperSize       =   0
      EndProperty
      BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
         ConnectionString=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Visual Studio\VB98\Nwind.mdb;Persist Security Info=False"
         RecordSource    =   "Customers"
         Filter          =   ""
         MaxRecords      =   0
      EndProperty
      GroupCount      =   3
      BeginProperty Group0 {E862F8BF-E806-4B39-9A11-0BBED515338B} 
         Name            =   "Group0"
         GroupBy         =   "Country"
         Sort            =   1
         KeepTogether    =   0
         Object.Tag             =   ""
      EndProperty
      BeginProperty Group1 {E862F8BF-E806-4B39-9A11-0BBED515338B} 
         Name            =   "Group1"
         GroupBy         =   "PostalCode"
         Sort            =   1
         KeepTogether    =   0
         Object.Tag             =   ""
      EndProperty
      BeginProperty Group2 {E862F8BF-E806-4B39-9A11-0BBED515338B} 
         Name            =   "Group2"
         GroupBy         =   "CompanyName"
         Sort            =   1
         KeepTogether    =   0
         Object.Tag             =   ""
      EndProperty
      SectionCount    =   11
      BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Detail"
         Visible         =   -1  'True
         Height          =   1440
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "ReportHeader"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section2 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "ReportFooter"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section3 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section4 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Page Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section5 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Group 0 Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section6 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Group 0 Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section7 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Group 1 Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section8 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Group 1 Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section9 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Group 2 Header"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Section10 {673CB92F-28D3-421F-86CD-1099425A5037} 
         Name            =   "Group 2 Footer"
         Visible         =   0   'False
         Height          =   0
         CanGrow         =   -1  'True
         CanShrink       =   0   'False
         KeepTogether    =   -1  'True
         ForcePageBreak  =   0
         BackColor       =   16777215
         Repeat          =   0   'False
         OnFormat        =   ""
         OnPrint         =   ""
         Object.Tag             =   ""
      EndProperty
      FieldCount      =   4
      BeginProperty Field0 {6AC1BBA5-107E-4F07-BCF0-DF757735D0A8} 
         Name            =   "CompanyNameLine"
         Text            =   "= Trim([CompanyName])"
         Object.Left            =   288
         Object.Top             =   240
         Width           =   3204
         Height          =   300
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Align           =   -1
         BackStyle       =   1
         BorderColor     =   0
         BorderStyle     =   0
         CanGrow         =   -1  'True
         CanShrink       =   -1  'True
         Visible         =   -1  'True
         HideDuplicates  =   0   'False
         RunningSum      =   0
         Format          =   ""
         LineSlant       =   0
         LineWidth       =   0
         PictureAlign    =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         Section         =   0
         ForcePageBreak  =   0
         Calculated      =   -1  'True
         WordWrap        =   -1  'True
         LineSpacing     =   0
         CheckBox        =   0
         RTF             =   0   'False
         Anchor          =   0
         ZOrder          =   0
         LinkTarget      =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Field1 {6AC1BBA5-107E-4F07-BCF0-DF757735D0A8} 
         Name            =   "Address1Line"
         Text            =   "= Trim( [Address])"
         Object.Left            =   288
         Object.Top             =   540
         Width           =   3204
         Height          =   300
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Align           =   -1
         BackStyle       =   1
         BorderColor     =   0
         BorderStyle     =   0
         CanGrow         =   -1  'True
         CanShrink       =   -1  'True
         Visible         =   -1  'True
         HideDuplicates  =   0   'False
         RunningSum      =   0
         Format          =   ""
         LineSlant       =   0
         LineWidth       =   0
         PictureAlign    =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         Section         =   0
         ForcePageBreak  =   0
         Calculated      =   -1  'True
         WordWrap        =   -1  'True
         LineSpacing     =   0
         CheckBox        =   0
         RTF             =   0   'False
         Anchor          =   0
         ZOrder          =   0
         LinkTarget      =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Field2 {6AC1BBA5-107E-4F07-BCF0-DF757735D0A8} 
         Name            =   "Address2Line"
         Text            =   "= Trim( [City] & "" "" & [Region] & ""  "" & [PostalCode])"
         Object.Left            =   288
         Object.Top             =   840
         Width           =   3204
         Height          =   300
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Align           =   -1
         BackStyle       =   1
         BorderColor     =   0
         BorderStyle     =   0
         CanGrow         =   -1  'True
         CanShrink       =   -1  'True
         Visible         =   -1  'True
         HideDuplicates  =   0   'False
         RunningSum      =   0
         Format          =   ""
         LineSlant       =   0
         LineWidth       =   0
         PictureAlign    =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         Section         =   0
         ForcePageBreak  =   0
         Calculated      =   -1  'True
         WordWrap        =   -1  'True
         LineSpacing     =   0
         CheckBox        =   0
         RTF             =   0   'False
         Anchor          =   0
         ZOrder          =   0
         LinkTarget      =   ""
         Object.Tag             =   ""
      EndProperty
      BeginProperty Field3 {6AC1BBA5-107E-4F07-BCF0-DF757735D0A8} 
         Name            =   "CountryLine"
         Text            =   "= Trim( [Country])"
         Object.Left            =   288
         Object.Top             =   1140
         Width           =   3204
         Height          =   300
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Align           =   -1
         BackStyle       =   1
         BorderColor     =   0
         BorderStyle     =   0
         CanGrow         =   -1  'True
         CanShrink       =   -1  'True
         Visible         =   -1  'True
         HideDuplicates  =   0   'False
         RunningSum      =   0
         Format          =   ""
         LineSlant       =   0
         LineWidth       =   0
         PictureAlign    =   0
         MarginLeft      =   0
         MarginTop       =   0
         MarginRight     =   0
         MarginBottom    =   0
         Section         =   0
         ForcePageBreak  =   0
         Calculated      =   -1  'True
         WordWrap        =   -1  'True
         LineSpacing     =   0
         CheckBox        =   0
         RTF             =   0   'False
         Anchor          =   0
         ZOrder          =   0
         LinkTarget      =   ""
         Object.Tag             =   ""
      EndProperty
   End
   Begin VB.Menu Page 
      Caption         =   "Page"
      Begin VB.Menu Page_Portrait 
         Caption         =   "Portrait"
      End
      Begin VB.Menu Page_Landscape 
         Caption         =   "Landscape"
      End
   End
   Begin VB.Menu Headers 
      Caption         =   "Headers"
      Begin VB.Menu Headers_Title 
         Caption         =   "Title"
      End
   End
End
Attribute VB_Name = "OASISDynamReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim sXMLFilePath As String
Dim bPrintedToPDF As Boolean
Dim bIsWVI As Boolean
Dim mConnectionString As String


Private sNarrative_Social As String
Private sNarrative_Crime As String
Private sNarrative_Conflict As String
Private sNarrative_Terrorism As String
Private sNarrative_Kidnapping As String
Private sNarrative_HumSpace As String
Private sNarrative_Insfrast As String
Private sNarrative_Overall As String

Private sRating_Social As String
Private sRating_Crime As String
Private sRating_Conflict As String
Private sRating_Terrorism As String
Private sRating_Kidnapping As String
Private sRating_HumSpace As String
Private sRating_Insfrast As String
Private sRating_Overall As String

Private sReportFilter As String
Private pReportPic As StdPicture
Private sDateRangeText As String

Public Sub PrintToPDF(sFileName As String, _
                      sDynamicDataPath As String)
        '<EhHeader>
        On Error GoTo PrintToPDF_Err
        '</EhHeader>

        Dim sOldRecordSource As String
        Dim sNewRecordSource As String
        Dim i As Integer

100     If VSReport1.IsBusy Then Exit Sub
     
102     sOldRecordSource = VSReport1.DataSource.ConnectionString
104     i = InStr(sOldRecordSource, "Data Source=")

106     If Not bIsWVI Then

            If (IsNull(i) Or i = 0) Then

108             MsgBox "Error exporting.  Please contact OASIS administrator"
    
            ElseIf Not bPrintedToPDF Then
110             i = i - 1
112             sNewRecordSource = Right(sOldRecordSource, Len(sOldRecordSource) - i - Len("Data Source="))
114             sNewRecordSource = sDynamicDataPath & "\" & sNewRecordSource
116             sNewRecordSource = Left(sOldRecordSource, i + Len("Data Source=")) & sNewRecordSource
                VSReport1.DataSource.ConnectionString = sNewRecordSource
            End If
            
        End If

        VSReport1.RenderToFile sFileName, vsrPDF
        bPrintedToPDF = True
        'VSReport1.DataSource.ConnectionString = sOldRecordSource
        '<EhFooter>
        Exit Sub

PrintToPDF_Err:
        MsgBox "Rendering failure.  Please contact an OASIS administrator." ', "DynamicReportsOCX.DynamReports.PrintToPDF", "DynamReports component failure"
        Stop
        '</EhFooter>
End Sub

Public Sub SetWVIParams(sUserName1 As String, _
                        sNarrative_Social1 As String, _
                        sNarrative_Crime1 As String, _
                        sNarrative_Conflict1 As String, _
                        sNarrative_Terrorism1 As String, _
                        sNarrative_Kidnapping1 As String, _
                        sNarrative_HumSpace1 As String, _
                        sNarrative_Insfrast1 As String, _
                        sNarrative_Overall1 As String, _
                        sRating_Social1 As String, _
                        sRating_Crime1 As String, _
                        sRating_Conflict1 As String, _
                        sRating_Terrorism1 As String, _
                        sRating_Kidnapping1 As String, _
                        sRating_HumSpace1 As String, _
                        sRating_Insfrast1 As String, _
                        sRating_Overall1 As String, _
                        sFilter As String, _
                        pImage As StdPicture)
        '<EhHeader>
        On Error GoTo SetWVIParams_Err
        '</EhHeader>
                        
100     sNarrative_Social = sNarrative_Social1
102     sNarrative_Crime = sNarrative_Crime1
104     sNarrative_Conflict = sNarrative_Conflict1
106     sNarrative_Terrorism = sNarrative_Terrorism1
108     sNarrative_Kidnapping = sNarrative_Kidnapping1
110     sNarrative_HumSpace = sNarrative_HumSpace1
112     sNarrative_Insfrast = sNarrative_Insfrast1
114     sNarrative_Overall = sNarrative_Overall1

116     sRating_Social = sRating_Social1
118     sRating_Crime = sRating_Crime1
120     sRating_Conflict = sRating_Conflict1
122     sRating_Terrorism = sRating_Terrorism1
124     sRating_Kidnapping = sRating_Kidnapping1
126     sRating_HumSpace = sRating_HumSpace1
128     sRating_Insfrast = sRating_Insfrast1
130     sRating_Overall = sRating_Overall1

        sDateRangeText = Replace(sFilter, " OR ([DetailEventDate] = null)", "") & "     (generated by: " & sUserName1 & ")"
        
        Set pReportPic = pImage
        sReportFilter = sFilter
136     bIsWVI = True

        If List1.ListCount > 0 Then
            List1.ListIndex = 0
            Call List1_Click
        End If
                         
        '<EhFooter>
        Exit Sub

SetWVIParams_Err:
        Debug.Print "SetWVIParams error: (" & Erl & ") " & Err.Description
        'MsgBox "SetWVIParams error: (" & Erl & ") " & Err.Description
        Resume Next
        '</EhFooter>
End Sub

Public Sub GenerateChart(sConnString As String, _
                         sRecordSource As String, _
                         iHeight As Integer, _
                         iWidth As Integer, _
                         fField As Field)
        '<EhHeader>
        On Error GoTo GenerateChart_Err
        '</EhHeader>

        Dim NewOASISChart As New OASISCharting.ChartProvider
        Dim oCO As OASISChartObj
        Dim sFilePath As String
    
        Dim RS As New ADODB.Recordset
        Dim cn As New ADODB.Connection
100     cn.Open sConnString
102     RS.Open sRecordSource, cn, adOpenDynamic, adLockBatchOptimistic
104     sFilePath = Left(sXMLFilePath, Len(sXMLFilePath) - 4)
    'CreateDynamDBPath
106     Debug.Print "CH: " & fField.Text
108     Debug.Print "RS: " & RS.Source
110     Debug.Print " "
    
112     With oCO
    
114         .iHeight = iHeight - 50
116         .iWidth = iWidth - 50
118         .iParentHeight = iHeight
120         .iParentWidth = iWidth
   
122         With .udtChartTemplate
124             .enmFormat = tplBin
126             .sDecription = "Some Potato Junkie"
128             .sName = sFilePath & "-" & fField.Text
            End With
        
130         .sSQL = sRecordSource
132         .sConnStr = sConnString
        
        End With
    
134     NewOASISChart.UpdateDataSet RS
136     NewOASISChart.InitChart oCO, True, , , , , fField

138     RS.Close
140     Set RS = Nothing
142     NewOASISChart.CloseObj
144     Set NewOASISChart = Nothing
    
        '<EhFooter>
        Exit Sub

GenerateChart_Err:
        Err.Raise vbObjectError + 100, _
                  "DynamicReportsOCX.DynamReports.GenerateChart", _
                  "DynamReports component failure"
        '</EhFooter>
End Sub

Public Sub ShowReportDetail()
        '<EhHeader>
        On Error GoTo ShowReportDetail_Err
        '</EhHeader>
        
        Dim f As Field ' variable used to hold new fields
        Dim f2 As Field ' variable used to hold new fields
        Dim i As Integer
        Dim iFields As Integer
        Dim sString As String

100     i = VSReport1.Groups.Count

102     Do Until i = 0
104         sString = "Group BY " & VSReport1.Groups.Item(i - 1).GroupBy
            'sString = "Group " & VSReport1.Groups.Item(i - 1).Name & " has " & SReport1.Groups.Item(i - 1).SectionHeader.Fields.Count & " fields"
106         MsgBox sString
108         i = i - 1
        Loop
    
110     i = VSReport1.Sections.Count

112     Do Until i = 0
    
114         sString = "Section '" & VSReport1.Sections.Item(i - 1).Name & "' is of type " & VSReport1.Sections.Item(i - 1).Type & " and has " & VSReport1.Sections.Item(i - 1).Fields.Count & " fields"
        
116         iFields = VSReport1.Sections.Item(i - 1).Fields.Count

118         Do Until iFields = 0
        
120             sString = sString & Chr(13) & " fieldname '" & VSReport1.Sections.Item(i - 1).Fields(iFields - 1).Name & "' with text '" & VSReport1.Sections.Item(i - 1).Fields(iFields - 1).Text & "'"

122             iFields = iFields - 1
            Loop
        
124         MsgBox sString
126         i = i - 1
        Loop

        '<EhFooter>
        Exit Sub

ShowReportDetail_Err:
        Err.Raise vbObjectError + 100, "DynamicReportsOCX.DynamReports.ShowReportDetail", "DynamReports component failure"
        '</EhFooter>
End Sub

Public Sub SetTitle(Optional sText As String = "")
        '<EhHeader>
        On Error GoTo SetTitle_Err
        '</EhHeader>
    
    'On Error Resume Next
    
        Dim fField As Field
100     Set fField = VSReport1.Sections(vsrHeader).Fields("TitleLbl")
    
        If Not Len(sText) > 0 Then
    
102     fField.Text = InputBox("Please enter in a new title", "Change report title", fField.Text)

        Else
        
        fField.Text = sText
        
        End If

104     VSReport1.Render VSPrinter1
    
        '<EhFooter>
        Exit Sub

SetTitle_Err:
        Err.Raise vbObjectError + 100, _
                  "DynamicReportsOCX.DynamReports.SetTitle", _
                  "DynamReports component failure"
        '</EhFooter>
End Sub

Private Sub cmdOrientation_Click()
        '<EhHeader>
        On Error GoTo cmdOrientation_Click_Err
        '</EhHeader>

100     If GetOrientIsPortrait Then
102         Me.SetOrientLandscape
        Else
104         Me.SetOrientPortrait
        End If

        '<EhFooter>
        Exit Sub

cmdOrientation_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "DynamicReportsOCX.DynamReports.cmdOrientation_Click", _
                  "DynamReports component failure"
        '</EhFooter>
End Sub

Private Sub cmdTitle_Click()
        '<EhHeader>
        On Error GoTo cmdTitle_Click_Err
        '</EhHeader>
100     Me.SetTitle
        '<EhFooter>
        Exit Sub

cmdTitle_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "DynamicReportsOCX.DynamReports.cmdTitle_Click", _
                  "DynamReports component failure"
        '</EhFooter>
End Sub

Private Sub List1_Click()
        '<EhHeader>
        On Error GoTo List1_Click_Err
        '</EhHeader>
    
        Dim fField As Field
        Dim sImagePath As String
        Dim sXMLPath As String
        Dim t
    
100     If VSReport1.IsBusy Then Exit Sub
102     List1.Enabled = False
104     VSReport1.Clear
106     VSPrinter1.Clear
    
108     t = Timer
110     MousePointer = 11
112     sXMLPath = sXMLFilePath

114     VSPrinter1.Enabled = True
116     VSReport1.Load sXMLPath, List1
118     VSPrinter1.SetFocus
    
120     lblReport = "Report: " & List1.Text
122     frameButtons.Visible = True

124     If bIsWVI Then
            
126         VSReport1.Fields("txtNarrativeOverall").Text = sNarrative_Overall
128         VSReport1.Fields("txtSecRatingOverall").Text = sRating_Overall
            
130         VSReport1.DataSource.Filter = sReportFilter
            VSReport1.DataSource.RecordSource = "select * from [dd_WVISec_qryIncidents_FEA]"
            VSReport1.DataSource.ConnectionString = mConnectionString

            Picture1.AutoSize = True
            Picture1.Picture = pReportPic

132         VSReport1.Fields("PicScreenshot").Picture = Picture1.Image
134         'VSReport1.Sections(vsrHeader).Fields("PicScreenshot").Picture = Picture1.Image
            VSReport1.Fields("txtDateRange").Text = sDateRangeText

        End If
    
        'On Error Resume Next
136     VSReport1.Render VSPrinter1

138     If Err.Number <> 0 Then
            'MsgBox "Error " & Err.Number & vbCrLf & Err.Description
        End If

140     Clipboard.Clear
142     MousePointer = 0
144     VSPrinter1.NavBarText = "Done in " & Format(Timer - t, "#.##") & " seconds"
146     List1.Enabled = True
148     bPrintedToPDF = False

        '"Data Source="
        'VSReport1.DataSource.ConnectionString = VSReport1.DataSource.ConnectionString

        '<EhFooter>
        Exit Sub

List1_Click_Err:
        Debug.Print "List1_Click_Err: (" & Erl & ") " & Err.Description
        'MsgBox "List1_Click_Err: (" & Erl & ") " & Err.Description
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetOrientLandscape()
        '<EhHeader>
        On Error GoTo SetOrientLandscape_Err
        '</EhHeader>
100     VSReport1.Layout.Orientation = vsrLandscape
102     VSReport1.Render VSPrinter1
        '<EhFooter>
        Exit Sub

SetOrientLandscape_Err:
        Err.Raise vbObjectError + 100, _
                  "DynamicReportsOCX.DynamReports.SetOrientLandscape", _
                  "DynamReports component failure"
        '</EhFooter>
End Sub

Public Sub SetOrientPortrait()
        '<EhHeader>
        On Error GoTo SetOrientPortrait_Err
        '</EhHeader>
100     VSReport1.Layout.Orientation = vsrPortrait
102     VSReport1.Render VSPrinter1
        '<EhFooter>
        Exit Sub

SetOrientPortrait_Err:
        Err.Raise vbObjectError + 100, _
                  "DynamicReportsOCX.DynamReports.SetOrientPortrait", _
                  "DynamReports component failure"
        '</EhFooter>
End Sub

Public Function GetOrientIsPortrait()
        '<EhHeader>
        On Error GoTo GetOrientIsPortrait_Err
        '</EhHeader>

100     If VSReport1.Layout.Orientation = vsrPortrait Then
102         GetOrientIsPortrait = True
        Else
104         GetOrientIsPortrait = False
        End If

        '<EhFooter>
        Exit Function

GetOrientIsPortrait_Err:
        Err.Raise vbObjectError + 100, _
                  "DynamicReportsOCX.DynamReports.GetOrientIsPortrait", _
                  "DynamReports component failure"
        '</EhFooter>
End Function

Public Sub loadXML(sPath As String, _
                   Optional bIsWV As Boolean = False, Optional sCN As String = "")
        '<EhHeader>
        On Error GoTo loadXML_Err
        '</EhHeader>
    
        Dim i%, iCnt%

        ' count how many reports are in our definition file
100     sXMLFilePath = sPath
102     iCnt = VSReport1.GetReportInfo(sXMLFilePath, vsrRICount)
        
        ' populate list box
104     List1.Clear

106     For i = 0 To iCnt - 1
108         List1.AddItem VSReport1.GetReportInfo(sXMLFilePath, vsrRIName, i)
        Next
        
        If bIsWV Then
            mConnectionString = sCN
        End If
    
110     lblReport.Caption = "No Report Selected"
112     VSPrinter1.Clear
114     VSPrinter1.Enabled = False
116     frameButtons.Visible = False
118     Clipboard.Clear
        bIsWVI = bIsWV

        '<EhFooter>
        Exit Sub

loadXML_Err:
        Debug.Print "loadXML_Err: (" & Erl & ") " & Err.Description
        'MsgBox "loadXML_Err: (" & Erl & ") " & Err.Description
        Resume Next
        '</EhFooter>
End Sub

Private Sub VSReport1_OnPrint(ByVal SectionIndex As Long)
        '<EhHeader>
        On Error GoTo VSReport1_OnPrint_Err
        '</EhHeader>
        
        Dim fField As Field
        Dim fField2 As Field
        Dim sSQLString As String
        Dim sField2 As String
        Dim i As Integer
        On Error Resume Next
        
100     If bIsWVI Then
        
102         i = 0
            
104         Select Case VSReport1.Groups(0).SectionFooter.Fields("DataField1").Value
    
                Case "Social & Political"
106                 VSReport1.Groups(0).SectionFooter.Fields("txtNarrative").Text = sNarrative_Social
108                 VSReport1.Groups(0).SectionFooter.Fields("txtSecRating").Text = sRating_Social

110             Case "Crime & Security"
112                 VSReport1.Groups(0).SectionFooter.Fields("txtNarrative").Text = sNarrative_Crime
114                 VSReport1.Groups(0).SectionFooter.Fields("txtSecRating").Text = sRating_Crime

116             Case "Conflict"
118                 VSReport1.Groups(0).SectionFooter.Fields("txtNarrative").Text = sNarrative_Conflict
120                 VSReport1.Groups(0).SectionFooter.Fields("txtSecRating").Text = sRating_Conflict

122             Case "Terrorism"
124                 VSReport1.Groups(0).SectionFooter.Fields("txtNarrative").Text = sNarrative_Terrorism
126                 VSReport1.Groups(0).SectionFooter.Fields("txtSecRating").Text = sRating_Terrorism

128             Case "Kidnapping"
130                 VSReport1.Groups(0).SectionFooter.Fields("txtNarrative").Text = sNarrative_Kidnapping
132                 VSReport1.Groups(0).SectionFooter.Fields("txtSecRating").Text = sRating_Kidnapping
  
134             Case "Humanitarian Space"
136                 VSReport1.Groups(0).SectionFooter.Fields("txtNarrative").Text = sNarrative_HumSpace
138                 VSReport1.Groups(0).SectionFooter.Fields("txtSecRating").Text = sRating_HumSpace

140             Case "Infrastructure"
142                 VSReport1.Groups(0).SectionFooter.Fields("txtNarrative").Text = sNarrative_Insfrast
144                 VSReport1.Groups(0).SectionFooter.Fields("txtSecRating").Text = sRating_Insfrast

146             Case "-- OVERALL --"
148                 VSReport1.Groups(0).SectionFooter.Fields("txtNarrative").Text = sNarrative_Overall
150                 VSReport1.Groups(0).SectionFooter.Fields("txtSecRating").Text = sRating_Overall
            
            End Select
        
        Else
        
160         Clipboard.Clear

162         If VSReport1.Sections(SectionIndex).Type = 5 Or VSReport1.Sections(SectionIndex).Type = 7 Or VSReport1.Sections(SectionIndex).Type = 9 Then
                'Debug.Print VSReport1.Sections(SectionIndex).Type
164             i = 0

166             Do Until i = VSReport1.Groups.Count
                         
168                 If VSReport1.Groups.Item(i).SectionHeader = VSReport1.Sections(SectionIndex) Then

                        'Set fField = VSReport1.Sections(SectionIndex).Fields("picChart")
                        'Set fField2 = VSReport1.Sections(SectionIndex).Fields("Field0")
170                     Set fField = VSReport1.Groups.Item(i).SectionHeader.Fields("picChart")
172                     Set fField2 = VSReport1.Groups.Item(i).SectionHeader.Fields("Field0")

174                     If Not fField Is Nothing Then

176                         sField2 = fField2.Value
178                         sField2 = Trim(sField2)

180                         sSQLString = fField.Tag
182                         sSQLString = Replace$(sSQLString, "XXXX", sField2, 1, , vbTextCompare)
                        
                            'Debug.Print Right$(sSQLString, 1) & "    " & sSQLString
184                         Call GenerateChart(VSReport1.DataSource.ConnectionString, sSQLString, fField.Height, fField.Width, fField)
                        End If
                    End If

186                 Set fField = Nothing
188                 Set fField2 = Nothing
190                 i = i + 1
                Loop
        
            Else

192             Set fField = VSReport1.Sections(SectionIndex).Fields("picChart")
194             Set fField2 = VSReport1.Sections(SectionIndex).Fields("Field0")

196             If Not fField Is Nothing Then

198                 sSQLString = fField.Tag
200                 sSQLString = Replace(sSQLString, "XXXX", fField2.Value, , , vbTextCompare)
202                 Call GenerateChart(VSReport1.DataSource.ConnectionString, sSQLString, fField.Height, fField.Width, fField)
204                 Set fField = Nothing
206                 Set fField2 = Nothing
                End If
            End If
        
        End If
        
        '<EhFooter>
        Exit Sub

VSReport1_OnPrint_Err:
        Debug.Print "VSReport1_OnPrint_Err: (" & Erl & ") " & Err.Description
        'MsgBox "VSReport1_OnPrint_Err: (" & Erl & ") " & Err.Description
        Resume Next
        '</EhFooter>
End Sub

