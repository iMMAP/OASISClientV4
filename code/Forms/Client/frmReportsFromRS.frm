VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{C8CF160E-7278-4354-8071-850013B36892}#1.0#0"; "vsrpt8.ocx"
Object = "{54850C51-14EA-4470-A5E4-8C5DB32DC853}#1.0#0"; "vsprint8.ocx"
Begin VB.Form frmReportsFromRS 
   Caption         =   "Tabular Reports"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   720
   ClientWidth     =   7545
   Icon            =   "frmReportsFromRS.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   7275
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   7545
      _cx             =   13309
      _cy             =   12832
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
      GridRows        =   1
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmReportsFromRS.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VSPrinter8LibCtl.VSPrinter VSPrinter1 
         Height          =   7095
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   7365
         _cx             =   12991
         _cy             =   12515
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
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
         Zoom            =   39.6780303030303
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
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            Height          =   615
            Left            =   6120
            ScaleHeight     =   37
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   45
            TabIndex        =   2
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VSReport8LibCtl.VSReport VSReport1 
         Left            =   10920
         Top             =   0
         _rv             =   800
         ReportName      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
            Width           =   0
            MarginLeft      =   1440
            MarginTop       =   1440
            MarginRight     =   1440
            MarginBottom    =   1440
            Columns         =   1
            ColumnLayout    =   0
            Orientation     =   0
            PageHeader      =   0
            PageFooter      =   0
            PictureAlign    =   7
            PictureShow     =   1
            PaperSize       =   0
         EndProperty
         BeginProperty DataSource {D1359088-0913-44EA-AE50-6A7CD77D4C50} 
            ConnectionString=   ""
            RecordSource    =   ""
            Filter          =   ""
            MaxRecords      =   0
         EndProperty
         GroupCount      =   0
         SectionCount    =   5
         BeginProperty Section0 {673CB92F-28D3-421F-86CD-1099425A5037} 
            Name            =   "Detail"
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
         BeginProperty Section1 {673CB92F-28D3-421F-86CD-1099425A5037} 
            Name            =   "Header"
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
            Name            =   "Footer"
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
         FieldCount      =   0
      End
   End
   Begin VB.Menu Title 
      Caption         =   "Title"
   End
   Begin VB.Menu PaperSize 
      Caption         =   "Paper Size"
      Begin VB.Menu Letter 
         Caption         =   "Letter"
      End
      Begin VB.Menu A4 
         Caption         =   "A4"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Orientation 
      Caption         =   "Orientation"
      Begin VB.Menu Orientation_Portrait 
         Caption         =   "Portrait"
      End
      Begin VB.Menu Orientation_Landscape 
         Caption         =   "Landscape"
      End
   End
   Begin VB.Menu SmartFit 
      Caption         =   "SmartFit"
   End
   Begin VB.Menu SaveToPDF 
      Caption         =   "Save to PDF"
   End
End
Attribute VB_Name = "frmReportsFromRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oLocalRS As ADODB.Recordset
Dim sTitle As String
Dim sSummaryGroups As String
Dim bitmapImage As StdPicture
Dim sPicTitle As String
Dim sPicSubtitle As String
Dim sReportSort As String
Private m_StrPDFPath As String
Dim bSmartFit As Boolean

Public Property Get PDFPath() As String
    PDFPath = m_StrPDFPath
End Property

Private Sub A4_Click()
If Not VSReport1.IsBusy And (VSPrinter1.ReadyState = vpstReady Or VSPrinter1.ReadyState = vpstEmpty) Then
A4.Checked = True
  Letter.Checked = False
 ShowReport
End If
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Me.VSReport1.Clear
        bSmartFit = False
        
102     If Not g_sLanguage = "" Then
104         If Not m_Cnn.State = adStateClosed Then
106             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If
        
        bSmartFit = True
        SmartFit.caption = "Undo SmartFit"
        Me.WindowState = 2
        'ShowReport
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmReportsFromRS.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
        
100     If Me.VSReport1.IsBusy Then
        
            On Error Resume Next
102         VSPrinter1.Clear
104         VSReport1.Clear
106         Unload VSPrinter1
108         Unload VSReport1
    
        End If

110     Set bitmapImage = Nothing

        If Len(Me.Tag) > 3 Then
112         Unload Me
        Else
            Me.Tag = "Irish Chicken"
            Me.Visible = False
            Cancel = 1
        End If
        
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmReportsFromRS.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetReportRS(sPassedTitle As String, _
                       oRS As ADODB.Recordset, _
                       sPassedSummaryGroups As String, _
                       Optional bitmapImagePassed As StdPicture, _
                       Optional sPicTitlePassed As String, _
                       Optional sPicSubtitlePassed As String, Optional sSort As String)
        '<EhHeader>
        On Error GoTo SetReportRS_Err
        '</EhHeader>

100     If Not bitmapImagePassed Is Nothing Then
Set bitmapImage = bitmapImagePassed
Picture1 = bitmapImagePassed

End If
102     sPicTitle = sPicTitlePassed
104     sPicSubtitle = sPicSubtitlePassed

106     Set oLocalRS = oRS
108     sTitle = sPassedTitle
110     sSummaryGroups = sPassedSummaryGroups

112     'If oRS.Fields.Count > 4 Then
114         Me.Orientation_Landscape.Checked = True
116         Me.Orientation_Portrait.Checked = False
        'Else
118         'Me.Orientation_Landscape.Checked = False
120        ' Me.Orientation_Portrait.Checked = True
        'End If
        A4.Checked = True
        Letter.Checked = False
        sReportSort = sSort

        '<EhFooter>
        Exit Sub

SetReportRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmReportsFromRS.SetReportRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ShowReport()
        '<EhHeader>
        On Error GoTo ShowReport_Err
        '</EhHeader>

        Dim i As Integer
        Dim iActualSizeMax As Integer
        Dim iCycle As Integer
        Dim iCycle2 As Integer
        Dim iPrintedFields As Integer
        Dim iLevels As Integer
        Dim iLevelsNotAsGroups As Integer
        Dim iWidthOfEachField As Integer
        Dim iWidthOfReport As Integer
        Dim iTitlesOffset As Integer
        Dim fField1
        Dim fField2
        Dim lLabel1
        Dim sGroups() As String
        Dim grp
        Dim iFieldWidth() As Integer
        Dim iFieldWidthTemp() As Integer
        Dim iCountNumericIsh As Integer
        Dim iCountString As Integer
        Dim iLastLeftPosition As Integer
        Dim bNDMA As Boolean
        
100     bNDMA = False 'IIf(InStr(g_sAppServerPath, "srfpakistan.pk") > 0, True, False)
    
102     If Not VSReport1.IsBusy And (VSPrinter1.ReadyState = vpstReady Or VSPrinter1.ReadyState = vpstEmpty) Then

104         VSReport1.Clear
106         VSPrinter1.Clear
            
108         sGroups = Split(sSummaryGroups, ":::")

110         If Letter.Checked And Me.Orientation_Portrait.Checked Then
            
112             Me.VSReport1.Load g_sAppPath & "\data\templates\FixedTemplates\" & IIf(bNDMA, "NDMA", "OASIS") & ".xml", "DynamicPortraitLetter"
                
114         ElseIf Letter.Checked And Me.Orientation_Landscape.Checked Then
            
116             Me.VSReport1.Load g_sAppPath & "\data\templates\FixedTemplates\" & IIf(bNDMA, "NDMA", "OASIS") & ".xml", "DynamicLandscapeLetter"
                
118         ElseIf A4.Checked And Me.Orientation_Portrait.Checked Then
            
120             Me.VSReport1.Load g_sAppPath & "\data\templates\FixedTemplates\" & IIf(bNDMA, "NDMA", "OASIS") & ".xml", "DynamicPortraitA4"
                
122         ElseIf A4.Checked And Me.Orientation_Landscape.Checked Then
            
124             Me.VSReport1.Load g_sAppPath & "\data\templates\FixedTemplates\" & IIf(bNDMA, "NDMA", "OASIS") & ".xml", "DynamicLandscapeA4"
  
            End If

126         VSReport1.Fields("TitleLbl").Text = sTitle
128         iWidthOfReport = VSReport1.Fields("Line1").Width
130         iLevels = oLocalRS.Fields.Count

132         If Not iLevels = 0 Then

134             VSReport1.DataSource.Recordset = oLocalRS.Clone

136             If Not oLocalRS.Filter = 0 Then
138                 VSReport1.DataSource.Filter = oLocalRS.Filter
                Else
140                 VSReport1.DataSource.Filter = ""
                End If
                
142             ReDim iFieldWidth(iLevels)
144             ReDim iFieldWidthTemp(iLevels)
146             iLevelsNotAsGroups = oLocalRS.Fields.Count - IIf(UBound(sGroups) = -1, 0, UBound(sGroups))
                    
148             If Not bSmartFit Then
            
150                 iWidthOfEachField = iWidthOfReport / iLevelsNotAsGroups
152                 iCycle = 0
                    
154                 Do Until iCycle = iLevels

156                     iFieldWidth(iCycle) = iWidthOfReport / iLevelsNotAsGroups
158                     iCycle = iCycle + 1
                    Loop
                    
                Else
                
160                 iCycle = 0

162                 Do Until iCycle = iLevels 'NotAsGroups

164                     iCycle2 = 0
166                     SafeMoveFirst oLocalRS
168                     iActualSizeMax = 0
                         
170                     Do Until oLocalRS.EOF

172                         If oLocalRS(iCycle).ActualSize > iActualSizeMax Then iActualSizeMax = oLocalRS(iCycle).ActualSize
174                         oLocalRS.MoveNext
                            
                        Loop

176                     If iActualSizeMax < 100 Then iActualSizeMax = 100
178                     iFieldWidthTemp(iCycle) = iActualSizeMax

180                     If InStr(sSummaryGroups, oLocalRS.Fields(iCycle).Name) > 0 Then iFieldWidthTemp(iCycle) = 0
182                     iCycle = iCycle + 1
                        
                    Loop
                    
                    'undude
                    
184                 iCycle = 0
186                 iCycle2 = 0
                    
188                 Do Until iCycle = iLevels 'iLevelsNotAsGroups
                    
190                     iCycle2 = iCycle2 + iFieldWidthTemp(iCycle)
192                     iCycle = iCycle + 1
                        
                    Loop
                    
194                 iCycle = 0
                    
196                 Do Until iCycle = iLevels 'iLevelsNotAsGroups
                    
198                     iFieldWidth(iCycle) = iWidthOfReport * (iFieldWidthTemp(iCycle) / iCycle2)
                        
200                     iCycle = iCycle + 1
                    Loop

                End If

202             iTitlesOffset = VSReport1.Fields("TitleLbl").top + VSReport1.Fields("TitleLbl").Height

204             If Not bitmapImage Is Nothing Then

                    'If Orientation_Landscape.Checked Then VSReport1.Sections(vsrHeader).ForcePageBreak = vsrAfter
                
206                 If Not sPicTitle = "" Then

208                     iTitlesOffset = VSReport1.Fields("TitleLbl").top + VSReport1.Fields("TitleLbl").Height
                    
                        'VSReport1.Sections(vsrHeader).Height =  + 400
210                     Set fField1 = VSReport1.Sections(vsrHeader).Fields.Add("sPicTitle", sPicTitle, 0, iTitlesOffset + 200, iWidthOfReport, 200)
212                     fField1.CanGrow = False
214                     fField1.FontBold = True
216                     fField1.FontSize = 22
218                     fField1.ForeColor = vbBlack ' vbWhite
220                     fField1.Align = vsrCenterTop
222                     iTitlesOffset = iTitlesOffset + 600

                        'If A4.Checked And Not sPicTitle = "" Then VSReport1.Fields("TitleLbl").ForcePageBreak = vsrBefore  '= True
                    
                    End If
                
224                 If Not sPicSubtitle = "" And 1 = 3 Then
                    
226                     Set fField2 = VSReport1.Sections(vsrHeader).Fields.Add("sPicSubttle", sPicSubtitle, 0, 600, iWidthOfReport, 200)
228                     fField2.CanGrow = True
230                     fField2.FontBold = True
232                     fField2.FontSize = 12
234                     fField2.ForeColor = vbBlack 'vbWhite
236                     fField2.Align = vsrCenterTop
238                     iTitlesOffset = iTitlesOffset + 200
                    
                    End If
            
240                 Picture1.AutoSize = True
242                 Picture1.Visible = False
244                 Picture1.Height = bitmapImage.Height
246                 Picture1.Width = bitmapImage.Width
                    
248                 Set fField1 = VSReport1.Sections(vsrHeader).Fields.Add("pic", "pic", 0, iTitlesOffset + 200, iWidthOfReport, 50)
250                 fField1.Align = vsrCenterMiddle
252                 fField1.PictureAlign = vsrPAZoom
                    'VSReport1.Sections(vsrHeader).KeepTogether = True
254                 fField1.CanGrow = False
256                 fField1.CanShrink = False
258                 fField1.Picture = Picture1 '.Image  'bitmapImage
260                 fField1.BorderColor = vbBlack
262                 fField1.BorderStyle = vsrBSSolid
                    
264                 If A4.Checked And Orientation_Landscape.Checked Then
266                     fField1.Width = fField1.Width * 0.8
268                     fField1.Height = (bitmapImage.Height / bitmapImage.Width) * fField1.Width
270                     fField1.left = (iWidthOfReport - fField1.Width) / 2
272                 ElseIf Not Me.Orientation_Portrait.Checked Then
274                     fField1.Height = bitmapImage.Height * (fField1.Width / bitmapImage.Width)
                    Else
                        'fField1.Height = bitmapImage.Height * (fField1.Width / bitmapImage.Width)
276                     fField1.Height = (bitmapImage.Height / bitmapImage.Width) * fField1.Width
                    End If
                    
                Else
278                 VSReport1.Sections(vsrPageHeader).ForcePageBreak = vsrBefore
                End If

280             i = 0

282             Do Until i = UBound(sGroups) Or UBound(sGroups) = -1
        
284                 Set grp = VSReport1.Groups.Add(sGroups(i), "[" & sGroups(i) & "]", vsrAscending)
    
286                 With grp.SectionHeader
288                     .BackColor = RGB(224, 237, 194)
290                     .Height = 300
292                     .Visible = True
294                     Set fField1 = .Fields.Add(sGroups(i), "=""" & sGroups(i) & ": "" & [" & sGroups(i) & "]", 0, 50, VSReport1.layout.Width, 100)
296                     fField1.Calculated = True
298                     fField1.Align = vsrLeftMiddle
300                     fField1.FontBold = True
302                     fField1.MarginLeft = 50
                    End With

304                 i = i + 1
                Loop
            
306             i = 0
308             iPrintedFields = 0
310             iLastLeftPosition = 0

312             Do Until i = iLevels
                
314                 If Not InStr(sSummaryGroups, oLocalRS.Fields(i).Name) > 0 Then
                
                        'Set lLabel1 = VSReport1.Sections(vsrPageHeader).Fields.Add("lbl" & i, oLocalRS.Fields(i).Name, iPrintedFields * iWidthOfEachField, 1200, iWidthOfEachField, 300)
                        
316                     Set lLabel1 = VSReport1.Sections(vsrPageHeader).Fields.Add("lbl" & i, oLocalRS.Fields(i).Name, iLastLeftPosition, 100, iFieldWidth(i), 300)
318                     lLabel1.FontBold = True
                        'iFieldWidth
320                     lLabel1.MarginLeft = 50
322                     lLabel1.CanGrow = True
324                     lLabel1.Align = vsrLeftMiddle

326                     If bNDMA Then lLabel1.ForeColor = vbWhite
                        'Set fField1 = VSReport1.Sections(vsrDetail).Fields.Add("txt" & i, "[" & oLocalRS.Fields(i).Name & "]", iLastLeftPosition, 0, iWidthOfEachField, 200)
328                     Set fField1 = VSReport1.Sections(vsrDetail).Fields.Add("txt" & i, "[" & oLocalRS.Fields(i).Name & "]", iLastLeftPosition, 0, iFieldWidth(i), 200)

330                     If oLocalRS.Fields(i).Type = adDate Then
332                         fField1.Format = "Medium Date"
                        End If
                        
334                     fField1.CanGrow = True 'IDEALLY THIS IS SET TO TRUE BUT FOR EVENT INFO THIS SHOWS BAD
336                     fField1.Calculated = True
338                     fField1.Visible = True
340                     fField1.Align = vsrLeftMiddle
342                     fField1.MarginLeft = 50
344                     iPrintedFields = iPrintedFields + 1
346                     Set fField1 = Nothing
348                     iLastLeftPosition = iLastLeftPosition + iFieldWidth(i)
                    End If

350                 i = i + 1
                Loop

                On Error Resume Next
352             VSReport1.Render VSPrinter1
                
            Else
354             MsgBox "The recordset has no columns!"
            End If
        End If

        '<EhFooter>
        Exit Sub

ShowReport_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmReportsFromRS.ShowReport " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Letter_Click()
If Not VSReport1.IsBusy And (VSPrinter1.ReadyState = vpstReady Or VSPrinter1.ReadyState = vpstEmpty) Then

Letter.Checked = True

 A4.Checked = False

 ShowReport
 
 End If
End Sub

Private Sub Orientation_Landscape_Click()
        '<EhHeader>
        On Error GoTo Orientation_Landscape_Click_Err
        '</EhHeader>
If Not VSReport1.IsBusy And (VSPrinter1.ReadyState = vpstReady Or VSPrinter1.ReadyState = vpstEmpty) Then

100     Me.Orientation_Landscape.Checked = True
102     Me.Orientation_Portrait.Checked = False
104     ShowReport

End If
        '<EhFooter>
        Exit Sub

Orientation_Landscape_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmReportsFromRS.Orientation_Landscape_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Orientation_Portrait_Click()
        '<EhHeader>
        On Error GoTo Orientation_Portrait_Click_Err
        '</EhHeader>
If Not VSReport1.IsBusy And (VSPrinter1.ReadyState = vpstReady Or VSPrinter1.ReadyState = vpstEmpty) Then

100     Me.Orientation_Landscape.Checked = False
102     Me.Orientation_Portrait.Checked = True
104     ShowReport
End If
        
        '<EhFooter>
        Exit Sub

Orientation_Portrait_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmReportsFromRS.Orientation_Portrait_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SaveToPDF_Click()
        '<EhHeader>
        On Error GoTo SaveToPDF_Click_Err
        '</EhHeader>

        Dim c As New cCommonDialog
        'On Error Resume Next
        If Not VSReport1.IsBusy And (VSPrinter1.ReadyState = vpstReady Or VSPrinter1.ReadyState = vpstEmpty) Then
100     c.DefaultExt = "*.pdf"
102     c.DialogTitle = "Save to PDF"
104     c.Filter = "PDF File (*.pdf)|*.pdf"
106     c.InitDir = "%userprofile%"
108     c.ShowSave
    
110     If Not c.Filename = "" Then
            m_StrPDFPath = c.Filename
112         Me.VSReport1.RenderToFile c.Filename, vsrPDF
        End If
End If

        '<EhFooter>
        Exit Sub

SaveToPDF_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmReportsFromRS.SaveToPDF_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SmartFit_Click()
        '<EhHeader>
        On Error GoTo SmartFit_Click_Err
        '</EhHeader>

If Not VSReport1.IsBusy And (VSPrinter1.ReadyState = vpstReady Or VSPrinter1.ReadyState = vpstEmpty) Then


100     If bSmartFit Then
102         bSmartFit = False
104         SmartFit.caption = "SmartFit"
        Else
106         bSmartFit = True
108         SmartFit.caption = "Undo SmartFit"
        End If
110     ShowReport
End If


'bSmartFit = True
        '<EhFooter>
        Exit Sub

SmartFit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmReportsFromRS.SmartFit_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Title_Click()
   
    sText = InputBox("Please enter in a new title", "Change report title", VSReport1.Sections(vsrHeader).Fields("TitleLbl").Text)
        
    If Len(sText) > 0 Then
        VSReport1.Sections(vsrHeader).Fields("TitleLbl").Text = sText
        VSReport1.Render VSPrinter1
    End If
    
End Sub
