VERSION 5.00
Object = "{E9DF30CA-4B30-4235-BF0C-7150F6466080}#1.0#0"; "ChartFX.ClientServer.Core.dll"
Begin VB.Form frmChartWin 
   BorderStyle     =   0  'None
   Caption         =   "OASIS Chart preview"
   ClientHeight    =   330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   330
   ScaleMode       =   0  'User
   ScaleWidth      =   960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3690
      Top             =   1530
   End
   Begin Cfx62ClientServerCtl.Chart Chart1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2205
      _Data_          =   "frmChartWin.frx":0000
   End
End
Attribute VB_Name = "frmChartWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

Private Declare Function SetParent _
                Lib "user32" (ByVal hWndChild As Long, _
                              ByVal hWndNewParent As Long) As Long
Public enmRequestType As ChartRequest
Public sChartSQL As String
Public sConnectioString As String
Public FilePAth As String
Public enmFormat As oChartExportFormat
Private m_intWidth As Integer
Private m_intHeight As Integer
Public m_bAutoExport As Boolean
Public oContainer As Object
Private cxAnnotation As New Cfx62ClientServerAnnotation.AnnotationX
Public m_bProcessReady As Boolean
Private m_bAllowAutoResize As Boolean
Private m_bAllowAutoCloseDBLClick As Boolean
Private m_rs As ADODB.Recordset

Public Property Set Dataset(oRS As ADODB.Recordset)
    Set m_rs = oRS
End Property

Public Property Let AllowAutoCloseDBLClick(ByVal bValue As Boolean)
    m_bAllowAutoCloseDBLClick = bValue
End Property

Public Property Let AllowAutoResize(ByVal bValue As Boolean)
    m_bAllowAutoResize = bValue
End Property

Public Property Let iHeight(ByVal intValue As Integer)
    m_intHeight = intValue
    Chart1.Height = m_intHeight
End Property

Public Property Let iWidth(ByVal intValue As Integer)
    m_intWidth = intValue
    Chart1.Width = m_intWidth
End Property

Private Sub Chart1_DoubleClick(ByVal sender As Object, ByVal args As Cfx62ClientServerCtl.MouseEventArgsX)
    If m_bAllowAutoCloseDBLClick Then
        enmRequestType = 666
        Timer1.Enabled = True
    End If
End Sub

Public Property Get ChartWinHwmd() As Long
    ChartWinHwmd = Me.hwnd
End Property

Public Sub cmdExport_Click()
    Chart1.Export FileFormat_Bitmap, Null
    Me.Hide
End Sub

Public Sub TemplateLoad(sPath As String)

    With Chart1
     
        If Len(sPath) > 0 Then
         .Import FileFormat_Binary, sPath
            
        End If
        Chart1.ContextMenus = True
        Chart1.RecalcScale
        
        
        'GetOASISChartObj.iParentHeight = Me.Height
        'GetOASISChartObj.iParentWidth = Me.Width
        m_bProcessReady = True
        
    End With

End Sub

Public Sub TemplateSave(sPath As String)

    With Chart1
        
        'If Len(sPath) > 0 Then
            .FileMask = FileMask_General
            .Export FileFormat_BinaryTemplate, sPath
        
        
        
       ' End If
        
       ' m_bProcessReady = True
        
    End With

End Sub


Public Function ShowTheChart(hwnd As Long) As Long
    
    With Me
    
        .enmRequestType = Init
        .m_bAutoExport = bAutoExport
        .FilePAth = sFilePath
        .Move Screen.Width / 2 + .Width / 2, Screen.Height / 2
        .AllowAutoResize = m_bAllowAutoResize
        .AllowAutoCloseDBLClick = m_bAllowAutoCloseDBLClick
        .BorderStyle = vbSizable
        InitChart = .hwnd
        
        
        '.setOASISChartObj udtOASISChart
    
        If hwnd > 0 Then
            SetParent .hwnd, hwnd
            .Top = xpos
            .Left = ypos
            .cmdExport.Visible = False
            .Show
        Else
            .Show vbModal
        End If
    

    End With

End Function



Public Function GetOASISChartObj(sPath As String) As OASISChartObj
        '<EhHeader>
        On Error GoTo GetOASISChartObj_Err
        '</EhHeader>
        Dim i As Integer
        Dim sFile As String
    

100     With Chart1
        
102         If Len(sPath) > 0 Then
104             .Import FileFormat_BinaryTemplate, sPath
            End If
        
        
106         GetOASISChartObj.bChartTBR = .ToolBar
108         GetOASISChartObj.bMenuBar = .MenuBar
        
            '            If UBound(GetOASISChartObj.sChartTools) > 0 Then
            '
            '                For i = UBound(GetOASISChartObj.sChartTools) - 1 To 0 Step -1
            '                    .ToolBarObj.RemoveAt CLng(GetOASISChartObj.sChartTools(i)), 1
            '                Next
            '
            '            End If
        
            '        If GetOASISChartObj.bAnnoTBR Then
            '            Chart1.Extensions.Add cxAnnotation
            '            cxAnnotation.Enabled = True
            '            cxAnnotation.ToolBar = True
            '            'On Error Resume Next
            '
            '            If UBound(GetOASISChartObj.sAnnoTools) > 0 Then
            '
            '                For i = UBound(GetOASISChartObj.sAnnoTools) - 1 To 0 Step -1
            '                    '.
            '                    '.Extensions.Item(0).ToolBarObj.RemoveAt CLng(sAnnoTools(i)), 1
            '                Next
            '
            '            End If
            '
            '        End If
        
110         GetOASISChartObj.bDataEdtr = .DataEditor
112         GetOASISChartObj.iDataEdtrAlign = .DataEditorObj.Docked
114         GetOASISChartObj.bDataEdtrAllowEdit = .AllowEdit
116         GetOASISChartObj.bDataEdtrAllowDrag = .AllowDrag
118         GetOASISChartObj.bMultipleColors = .MultipleColors
120         GetOASISChartObj.bSeriesLGD = .SerLegBox
122         GetOASISChartObj.iSerLgdAlign = .SerLegBoxObj.Docked
124         GetOASISChartObj.bValueLGD = .LegendBox
126         GetOASISChartObj.iValLgdAlign = .LegendBoxObj.Docked
128         GetOASISChartObj.bDataHigls = .Highlight.Enabled
130         GetOASISChartObj.bHGlsDimmed = .Highlight.Dimmed
132         GetOASISChartObj.bHGlsPointLabel = .Highlight.PointAttributes.PointLabels
134         GetOASISChartObj.bPointLabelsGen = .PointLabels

136         If Len(.Titles(0).Text) > 0 Then

138             With .Titles(0)
140                 GetOASISChartObj.udtTitle.CT_Alignment = .Alignment
142                 GetOASISChartObj.udtTitle.CT_BackColor = .BackColor
144                 GetOASISChartObj.udtTitle.CT_DockArea = .DockArea
146                 GetOASISChartObj.udtTitle.CT_DrawingArea = .DrawingArea
148                 GetOASISChartObj.udtTitle.CT_Flags = .Flags

150                 With .Font
152                     GetOASISChartObj.udtTitle.CT_Font.CF_Bold = .Bold
154                     GetOASISChartObj.udtTitle.CT_Font.CF_Italic = .Italic
156                     GetOASISChartObj.udtTitle.CT_Font.CF_Name = .Name
158                     GetOASISChartObj.udtTitle.CT_Font.CF_Size = .Size
160                     GetOASISChartObj.udtTitle.CT_Font.CF_Strikethrough = .Strikethrough
162                     GetOASISChartObj.udtTitle.CT_Font.CF_Underline = .Underline
164                     GetOASISChartObj.udtTitle.CT_Font.CF_Weight = .Weight
                    End With
                
166                 GetOASISChartObj.udtTitle.CT_Gap = .Gap
168                 GetOASISChartObj.udtTitle.CT_LineAlignment = .LineAlignment
170                 GetOASISChartObj.udtTitle.CT_LineGap = .LineGap
172                 GetOASISChartObj.udtTitle.CT_Link = .Link.Url
174                 GetOASISChartObj.udtTitle.CT_Text = .Text
176                 GetOASISChartObj.udtTitle.CT_TextColor = .TextColor
178                 GetOASISChartObj.udtTitle.CT_Url = .Url
                End With

            End If

180         If Len(.Titles(1).Text) > 0 Then

182             With .Titles(1)
184                 GetOASISChartObj.udtNotes.CT_Alignment = .Alignment
186                 GetOASISChartObj.udtNotes.CT_BackColor = .BackColor
188                 GetOASISChartObj.udtNotes.CT_DockArea = .DockArea
190                 GetOASISChartObj.udtNotes.CT_DrawingArea = .DrawingArea
192                 GetOASISChartObj.udtNotes.CT_Flags = .Flags
    
194                 With .Font
196                     GetOASISChartObj.udtNotes.CT_Font.CF_Bold = .Bold
198                     GetOASISChartObj.udtNotes.CT_Font.CF_Italic = .Italic
200                     GetOASISChartObj.udtNotes.CT_Font.CF_Name = .Name
202                     GetOASISChartObj.udtNotes.CT_Font.CF_Size = .Size
204                     GetOASISChartObj.udtNotes.CT_Font.CF_Strikethrough = .Strikethrough
206                     GetOASISChartObj.udtNotes.CT_Font.CF_Underline = .Underline
208                     GetOASISChartObj.udtNotes.CT_Font.CF_Weight = .Weight
                    End With
                    
210                 GetOASISChartObj.udtNotes.CT_Gap = .Gap
212                 GetOASISChartObj.udtNotes.CT_LineAlignment = .LineAlignment
214                 GetOASISChartObj.udtNotes.CT_LineGap = .LineGap
216                 GetOASISChartObj.udtNotes.CT_Link = .Link.Url
218                 GetOASISChartObj.udtNotes.CT_Text = .Text
220                 GetOASISChartObj.udtNotes.CT_TextColor = .TextColor
222                 GetOASISChartObj.udtNotes.CT_Url = .Url
                End With
                        
            End If
            
224         GetOASISChartObj.sXAxis = .AxisX.Title.Text
226         GetOASISChartObj.XAxisAngle = .AxisX.LabelAngle
228         GetOASISChartObj.XAxisStaggered = .AxisX.Staggered
230         GetOASISChartObj.sYAxis = .AxisY.Title.Text
232         GetOASISChartObj.bContextMenu = .ContextMenus
234         GetOASISChartObj.bScrollable = .Scrollable
            
236         GetOASISChartObj.iWidth = .Width
238         GetOASISChartObj.iHeight = .Height

240         GetOASISChartObj.enmScheme = .Scheme
242         GetOASISChartObj.iAngleX = .AngleX
244         GetOASISChartObj.iAngleY = .AngleY
246         GetOASISChartObj.enmAxesStyle = .AxesStyle
        
248         GetOASISChartObj.bBorder = .Border

250         GetOASISChartObj.enmBorderEffect = .BorderEffect
252         GetOASISChartObj.bCluster = .Cluster
254         GetOASISChartObj.bChart3D = .Chart3D
256         GetOASISChartObj.bCrossHairs = .CrossHairs
258         GetOASISChartObj.sngCylSides = .CylSides
260         GetOASISChartObj.enmGrid = .Grid
            
262         GetOASISChartObj.enmMarkerShape = .MarkerShape
264         GetOASISChartObj.iMarkerSize = .MarkerSize
266         GetOASISChartObj.sngMarkerStep = .MarkerStep
268         GetOASISChartObj.sngPerspective = .Perspective
270         GetOASISChartObj.bShowTips = .ShowTips
272         GetOASISChartObj.enmSmoothFlags = .SmoothFlags
274         GetOASISChartObj.enmStacked = .Stacked
276         GetOASISChartObj.iWallWidth = .WallWidth
278         GetOASISChartObj.bZoom = .Zoom
            
280         Select Case .Palette
                        
                Case "Nature.Sky"
282                 GetOASISChartObj.enmPalette = Sky

284             Case "Default.EarthTones"
286                 GetOASISChartObj.enmPalette = Earth_Tones

288             Case "Default.ModernBusiness"
290                 GetOASISChartObj.enmPalette = Modern_Business

292             Case "DarkPastels.Pastels"
294                 GetOASISChartObj.enmPalette = Pastels

296             Case "Mesa.Mesa"
298                 GetOASISChartObj.enmPalette = Mese

300             Case "Nature.Adventure"
302                 GetOASISChartObj.enmPalette = Adventure

304             Case "ChartFX5.ChartFX5"
306                 GetOASISChartObj.enmPalette = ChartDef5

308             Case "HighContrast.HighContrast"
310                 GetOASISChartObj.enmPalette = High_Contrast

312             Case "Default.ChartFX6"
314                 GetOASISChartObj.enmPalette = ChartDef6

316             Case "Default.Alternate"
318                 GetOASISChartObj.enmPalette = Alternate

320             Case "Vivid"
322                 GetOASISChartObj.enmPalette = Vivid

324             Case "DarkPastels.AltPastels"
326                 GetOASISChartObj.enmPalette = Alt_Pastels

328             Case Else
330                 GetOASISChartObj.enmPalette = Windows
332                 GetOASISChartObj.lngBackColor = .BackColor
334                 GetOASISChartObj.lngBorderColor = .BorderColor
336                 GetOASISChartObj.lngInsideColor = .InsideColor
            
            End Select

338         GetOASISChartObj.enmChartType = .Gallery
            
        End With
                
340     GetOASISChartObj.iParentHeight = Me.Height
342     GetOASISChartObj.iParentWidth = Me.Width
            
344     m_bProcessReady = True

        '<EhFooter>
        Exit Function

GetOASISChartObj_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISCharting.frmChartWin.GetOASISChartObj " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub setOASISChartObj(udtOASISChart As OASISChartObj)
        '<EhHeader>
        On Error GoTo setOASISChartObj_Err
        '</EhHeader>
        Dim i As Integer
        Dim cn As ADODB.Connection
        Dim RS As ADODB.Recordset
        Dim sFile As String

100     With Chart1
                
102         .FileMask = FileMask_All 'Or FileMask_Colors Or FileMask_Data Or FileMask_DataBind Or FileMask_Elements Or FileMask_Extensions Or FileMask_Fonts Or FileMask_General Or FileMask_Internet Or FileMask_Labels Or FileMask_SizeData Or FileMask_Tools Or FileMask_Titles Or FileMask_Template
            '.ClearData ClearDataFlag_Data
104         If Not Len(udtOASISChart.udtChartTemplate.sName) > 0 Then
106             .ToolBar = udtOASISChart.bChartTBR
                
108             If udtOASISChart.bMenuBar Then
110                 .MenuBar = True
112                 .MenuBarObj.RemoveAt .MenuBarObj.Length - 1, 1
                End If
        
114             If 1 = 2 And udtOASISChart.bChartTBR Then
116                 If UBound(udtOASISChart.sChartTools) > 0 Then

118                     For i = UBound(udtOASISChart.sChartTools) - 1 To 0 Step -1
120                         .ToolBarObj.RemoveAt CLng(udtOASISChart.sChartTools(i)), 1
                        Next

                    End If
                End If
    
122             If udtOASISChart.bAnnoTBR Then
124                 Chart1.Extensions.Add cxAnnotation
126                 cxAnnotation.Enabled = True
128                 cxAnnotation.ToolBar = True
                    'On Error Resume Next
                
130                 If UBound(udtOASISChart.sAnnoTools) > 0 Then

132                     For i = UBound(udtOASISChart.sAnnoTools) - 1 To 0 Step -1
                            '.
                            '.Extensions.Item(0).ToolBarObj.RemoveAt CLng(sAnnoTools(i)), 1
                        Next

                    End If
            
                End If
        
134             .DataEditor = udtOASISChart.bDataEdtr
136             .DataEditorObj.Docked = udtOASISChart.iDataEdtrAlign ' 256
138             .AllowEdit = udtOASISChart.bDataEdtrAllowEdit
140             .AllowDrag = udtOASISChart.bDataEdtrAllowDrag
142             .MultipleColors = udtOASISChart.bMultipleColors
144             .SerLegBox = udtOASISChart.bSeriesLGD
146             .SerLegBoxObj.Docked = udtOASISChart.iSerLgdAlign ' 513
148             .LegendBox = udtOASISChart.bValueLGD
150             .LegendBoxObj.Docked = udtOASISChart.iValLgdAlign ' 515
152             .Highlight.Enabled = udtOASISChart.bDataHigls
154             .Highlight.Dimmed = udtOASISChart.bHGlsDimmed
156             .Highlight.PointAttributes.PointLabels = udtOASISChart.bHGlsPointLabel
158             .PointLabels = udtOASISChart.bPointLabelsGen
                ' .Titles(0).Text = udtOASISChart.sTitle
                ' .Titles(1).Text = udtOASISChart.udtNotes

160             If Len(udtOASISChart.udtTitle.CT_Text) > 0 Then

162                 With .Titles(0)
164                     .Alignment = udtOASISChart.udtTitle.CT_Alignment
166                     .BackColor = udtOASISChart.udtTitle.CT_BackColor
168                     .DockArea = udtOASISChart.udtTitle.CT_DockArea
170                     .DrawingArea = udtOASISChart.udtTitle.CT_DrawingArea
172                     .Flags = udtOASISChart.udtTitle.CT_Flags

174                     With .Font
176                         .Bold = udtOASISChart.udtTitle.CT_Font.CF_Bold
178                         .Italic = udtOASISChart.udtTitle.CT_Font.CF_Italic
180                         .Name = udtOASISChart.udtTitle.CT_Font.CF_Name
182                         .Size = udtOASISChart.udtTitle.CT_Font.CF_Size
184                         .Strikethrough = udtOASISChart.udtTitle.CT_Font.CF_Strikethrough
186                         .Underline = udtOASISChart.udtTitle.CT_Font.CF_Underline
188                         .Weight = udtOASISChart.udtTitle.CT_Font.CF_Weight
                        End With
                
190                     .Gap = udtOASISChart.udtTitle.CT_Gap
192                     .LineAlignment = udtOASISChart.udtTitle.CT_LineAlignment
194                     .LineGap = udtOASISChart.udtTitle.CT_LineGap
196                     .Link.Url = udtOASISChart.udtTitle.CT_Link
198                     .Text = udtOASISChart.udtTitle.CT_Text
200                     .TextColor = udtOASISChart.udtTitle.CT_TextColor
202                     .Url = udtOASISChart.udtTitle.CT_Url
                    End With

                End If

204             If Len(udtOASISChart.udtNotes.CT_Text) > 0 Then

206                 With .Titles(1)
208                     .Alignment = udtOASISChart.udtNotes.CT_Alignment
210                     .BackColor = udtOASISChart.udtNotes.CT_BackColor
212                     .DockArea = udtOASISChart.udtNotes.CT_DockArea
214                     .DrawingArea = udtOASISChart.udtNotes.CT_DrawingArea
216                     .Flags = udtOASISChart.udtNotes.CT_Flags
    
218                     With .Font
220                         .Bold = udtOASISChart.udtNotes.CT_Font.CF_Bold
222                         .Italic = udtOASISChart.udtNotes.CT_Font.CF_Italic
224                         .Name = udtOASISChart.udtNotes.CT_Font.CF_Name
226                         .Size = udtOASISChart.udtNotes.CT_Font.CF_Size
228                         .Strikethrough = udtOASISChart.udtNotes.CT_Font.CF_Strikethrough
230                         .Underline = udtOASISChart.udtNotes.CT_Font.CF_Underline
232                         .Weight = udtOASISChart.udtNotes.CT_Font.CF_Weight
                        End With
                    
234                     .Gap = udtOASISChart.udtNotes.CT_Gap
236                     .LineAlignment = udtOASISChart.udtNotes.CT_LineAlignment
238                     .LineGap = udtOASISChart.udtNotes.CT_LineGap
240                     .Link.Url = udtOASISChart.udtNotes.CT_Link
242                     .Text = udtOASISChart.udtNotes.CT_Text
244                     .TextColor = udtOASISChart.udtNotes.CT_TextColor
246                     .Url = udtOASISChart.udtNotes.CT_Url
                    End With
                        
                End If
            
248             .AxisX.Title.Text = udtOASISChart.sXAxis
250             .AxisX.LabelAngle = udtOASISChart.XAxisAngle
252             .AxisX.Staggered = udtOASISChart.XAxisStaggered
254             .AxisY.Title.Text = udtOASISChart.sYAxis
256             .ContextMenus = udtOASISChart.bContextMenu
258             .Scrollable = udtOASISChart.bScrollable

260             .Scheme = udtOASISChart.enmScheme '1
262             .AngleX = udtOASISChart.iAngleX
264             .AngleY = udtOASISChart.iAngleY
266             .AxesStyle = udtOASISChart.enmAxesStyle '3
            
                '.BackgroundImage
268             .Border = udtOASISChart.bBorder

270             .BorderEffect = udtOASISChart.enmBorderEffect
272             .Cluster = udtOASISChart.bCluster
274             .Chart3D = udtOASISChart.bChart3D
276             .CrossHairs = udtOASISChart.bCrossHairs
278             .CylSides = udtOASISChart.sngCylSides
280             .Grid = udtOASISChart.enmGrid
            
282             .MarkerShape = udtOASISChart.enmMarkerShape
284             .MarkerSize = udtOASISChart.iMarkerSize
286             .MarkerStep = udtOASISChart.sngMarkerStep
288             .Perspective = udtOASISChart.sngPerspective
290             .ShowTips = udtOASISChart.bShowTips
292             .SmoothFlags = udtOASISChart.enmSmoothFlags
294             .Stacked = udtOASISChart.enmStacked
296             .WallWidth = udtOASISChart.iWallWidth
298             .Zoom = udtOASISChart.bZoom
            
300             Select Case udtOASISChart.enmScheme
            
                    Case Windows
302                     .Palette = ""
304                     .BackColor = udtOASISChart.lngBackColor
306                     .BorderColor = udtOASISChart.lngBorderColor
308                     .InsideColor = udtOASISChart.lngInsideColor
            
310                 Case Sky
312                     .Palette = "Nature.Sky"

314                 Case Earth_Tones
316                     .Palette = "Default.EarthTones"

318                 Case Modern_Business
320                     .Palette = "Default.ModernBusiness"

322                 Case Pastels
324                     .Palette = "DarkPastels.Pastels"

326                 Case Mese
328                     .Palette = "Mesa.Mesa"

330                 Case Adventure
332                     .Palette = "Nature.Adventure"

334                 Case ChartDef5
336                     .Palette = "ChartFX5.ChartFX5"

338                 Case High_Contrast
340                     .Palette = "HighContrast.HighContrast"

342                 Case ChartDef6
344                     .Palette = "Default.ChartFX6"

346                 Case Alternate
348                     .Palette = "Default.Alternate"

350                 Case Vivid
352                     .Palette = "Vivid"

354                 Case Alt_Pastels
356                     .Palette = "DarkPastels.AltPastels"
            
                End Select

358             .Gallery = udtOASISChart.enmChartType

            Else

360             .Import udtOASISChart.udtChartTemplate.enmFormat, udtOASISChart.udtChartTemplate.sName
                
362             If Len(udtOASISChart.udtTitle.CT_Text) > 0 Then
364                 .Titles(0).Text = udtOASISChart.udtTitle.CT_Text
                End If
                
366             If Len(udtOASISChart.udtNotes.CT_Text) > 0 Then
368                 .Titles(1).Text = udtOASISChart.udtNotes.CT_Text
                End If
                
370             If Len(udtOASISChart.sXAxis) > 0 Then
372                 .AxisX.Title.Text = udtOASISChart.sXAxis
                End If
                
374             If Len(udtOASISChart.sYAxis) > 0 Then
376                 .AxisY.Title.Text = udtOASISChart.sYAxis
                End If
            End If
            
378         .Width = udtOASISChart.iWidth
380         .Height = udtOASISChart.iHeight

            '            Debug.Print .Scheme
            '            Debug.Print .AngleX
            '            Debug.Print .AngleY
            '            Debug.Print .AxesStyle
            '            Debug.Print .BackColor
            '
            '            Debug.Print .Border
            '            Debug.Print .BorderColor
            '            Debug.Print .BorderEffect
            '            Debug.Print .Cluster
            '            Debug.Print .Chart3D
            '            Debug.Print .CrossHairs
            '            Debug.Print .CylSides
            '            Debug.Print .Grid
            '            Debug.Print .InsideColor
            '            Debug.Print .MarkerShape
            '            Debug.Print .MarkerSize
            '            Debug.Print .MarkerStep
            '            Debug.Print .Perspective
            '            Debug.Print .ShowTips
            '            Debug.Print .SmoothFlags
            '            Debug.Print .Stacked
            '            Debug.Print .WallWidth
            '            Debug.Print .Zoom
        End With
                
382     If Len(udtOASISChart.sConnStr) > 5 Then
    
384         Set cn = New ADODB.Connection
        
386         cn.Open udtOASISChart.sConnStr
        
388         If Len(udtOASISChart.sSQL) > 6 Then
390             Set RS = New ADODB.Recordset
392             RS.Open udtOASISChart.sSQL, cn, adOpenDynamic, adLockBatchOptimistic
394             Chart1.DataSource = RS
            End If
        
        End If
        
        
        
        'Chart1.FileMask = FileMask_All Or FileMask_Template Or FileMask_Titles Or FileMask_ReuseExtensions Or F And Not FileMask_Data
        
396     For i = 0 To 4

398         With udtOASISChart.udtExports(i)
            
400             If Len(.sFileName) > 0 Then
402                 sFile = .sPath & .sFileName
                
404                 Select Case .enmExportFormat
                
                        Case dmpData
406                         sFile = sFile & ".bin"

408                     Case imgBMP
410                         sFile = sFile & ".bmp"

412                     Case imgWMF
414                         sFile = sFile & ".wmf"

416                     Case tplBin
418                         sFile = sFile & ".oct"

420                     Case Else
422                         sFile = sFile & ".xml"
                    End Select
                
                    On Error Resume Next
                
424                 If .bForceKill Then Kill sFile
                
426                 Chart1.Export .enmExportFormat, sFile
                
                End If
            
            End With
 
        Next

428     If m_bAutoExport Then
430         enmRequestType = 666
432         m_bProcessReady = True
434         Timer1.Enabled = True
        Else
436         Me.Height = udtOASISChart.iParentHeight
438         Me.Width = udtOASISChart.iParentWidth

440         Chart1.Export FileFormat_Bitmap, Null
            
442         m_bProcessReady = True
            
            On Error Resume Next
              
444         If Not oContainer Is Nothing Then
446             Set oContainer.Picture = Clipboard.GetData(vbCFBitmap)
            End If
              
448         If Not FilePAth = "" Then
450             Kill FilePAth
452             SavePicture Clipboard.GetData(vbCFBitmap), FilePAth
            End If

            'Chart1.UpdateSizeNow
        End If

        On Error Resume Next
        If Not RS Is Nothing Then
            RS.Close
            Set RS = Nothing
        End If
        
        If Not cn Is Nothing Then
            cn.Close
            Set cn = Nothing
        End If
        
        '<EhFooter>
        Exit Sub

setOASISChartObj_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISCharting.frmChartWin.setOASISChartObj " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
    'cmdCommand1_Click
    'Chart1.Export FileFormat_Bitmap, Null
    m_intWidth = Me.Width * 10
    m_intHeight = Me.Height * 10
    
    Chart1.Height = m_intHeight
    Chart1.Width = m_intWidth
End Sub

Private Sub Form_Resize()
    If m_bAllowAutoResize Then
        Chart1.Move 0, 0, Me.Width, Me.Height
    End If
End Sub

Private Sub Timer1_Timer()

Exit Sub
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    If Not m_bProcessReady Then Exit Sub
    
    
    Select Case enmRequestType

        Case ChartRequest.Init
            Timer1.Enabled = False
            Exit Sub

        Case ChartRequest.ExportToFile
            Chart1.Export enmFormat, FilePAth

        Case ChartRequest.ExportExtended
            Dim RS As ADODB.Recordset

            Set RS = LoadRS(sConnectioString, sChartSQL)

            If Not RS Is Nothing Then
                Chart1.DataSource = RS
                Chart1.Export enmFormat, FilePAth
                On Error Resume Next
                RS.Close
                Set RS = Nothing
            End If
            
            
            
            
        Case Else
            Chart1.Export FileFormat_Bitmap, Null
            
            On Error Resume Next
              
            If Not oContainer Is Nothing Then
                Set oContainer.Picture = Clipboard.GetData(vbCFBitmap)
            End If
              
            If Not FilePAth = "" Then
                Kill FilePAth
                SavePicture Clipboard.GetData(vbCFBitmap), FilePAth
            End If
    End Select
        
    Timer1.Enabled = False
     Me.Hide
End Sub

Private Function LoadRS(sCon As String, _
                        sSQL As String) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo LoadRS_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset
        Dim cn As New ADODB.Connection

100     cn.Open sCon
    
102     RS.Open sSQL, cn, adOpenDynamic, adLockBatchOptimistic
    
104     Set LoadRS = RS

        On Error Resume Next
        
        cn.Close
        
        Set cn = Nothing
        
        '<EhFooter>
        Exit Function

LoadRS_Err:
        Err.Raise vbObjectError + 100, "OASISChartEngine.OASISChart.LoadRS", "OASISChart component failure"
        '</EhFooter>
End Function

