VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Begin VB.Form frmIMSMAExporter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "iMMAP IMSMA Utils"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3795
   Icon            =   "frmIMSMAExporter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   3795
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elhold 
      Height          =   6735
      Left            =   0
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   3795
      _cx             =   6694
      _cy             =   11880
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
      BorderWidth     =   2
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
      GridRows        =   3
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmIMSMAExporter.frx":4781A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame2 
         Caption         =   "Export to:"
         Height          =   2040
         Left            =   30
         TabIndex        =   26
         Top             =   2505
         Width           =   3735
         Begin VB.CheckBox chkOASISSynch 
            Caption         =   "OASIS Synch"
            Height          =   315
            Left            =   1740
            TabIndex        =   33
            Top             =   990
            Width           =   1305
         End
         Begin VB.CheckBox chkExpFormat 
            Caption         =   "OGIS GML"
            Height          =   375
            Index           =   4
            Left            =   1725
            TabIndex        =   32
            Top             =   585
            Width           =   1335
         End
         Begin VB.CheckBox chkExpFormat 
            Caption         =   "ESRI SHP"
            Height          =   375
            Index           =   0
            Left            =   165
            TabIndex        =   31
            Top             =   225
            Width           =   1335
         End
         Begin VB.CheckBox chkExpFormat 
            Caption         =   "Mapinfo Mid/Mif"
            Height          =   375
            Index           =   1
            Left            =   165
            TabIndex        =   30
            Top             =   585
            Width           =   1335
         End
         Begin VB.CheckBox chkExpFormat 
            Caption         =   "Google KML"
            Height          =   375
            Index           =   2
            Left            =   150
            TabIndex        =   29
            Top             =   990
            Width           =   1335
         End
         Begin VB.CheckBox chkExpFormat 
            Caption         =   "GPS (gpx)"
            Height          =   375
            Index           =   3
            Left            =   1725
            TabIndex        =   28
            Top             =   225
            Width           =   1455
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Export"
            Height          =   285
            Left            =   2610
            TabIndex        =   27
            Top             =   1590
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Items to Import:"
         Height          =   2475
         Left            =   30
         TabIndex        =   13
         Top             =   30
         Width           =   3735
         Begin VB.CheckBox chkInsertTo 
            Caption         =   "Insert to OASIS"
            Height          =   225
            Left            =   210
            TabIndex        =   34
            Top             =   1860
            Value           =   1  'Checked
            Width           =   1605
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Get Stuff From Imsma"
            Height          =   255
            Left            =   1890
            TabIndex        =   25
            Top             =   1830
            Width           =   1635
         End
         Begin VB.CheckBox chkLocation 
            Caption         =   "Location"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   240
            Width           =   1125
         End
         Begin VB.CheckBox chkAccident 
            Caption         =   "Accident"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   510
            Width           =   975
         End
         Begin VB.CheckBox chkHazard 
            Caption         =   "Hazard"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   780
            Width           =   1215
         End
         Begin VB.CheckBox chkHazardReduction 
            Caption         =   "Hazard Reduction"
            Height          =   195
            Left            =   1680
            TabIndex        =   21
            Top             =   1020
            Width           =   1635
         End
         Begin VB.CheckBox chkQC 
            Caption         =   "QC"
            Height          =   255
            Left            =   1680
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkVictim 
            Caption         =   "Victim"
            Height          =   255
            Left            =   1680
            TabIndex        =   19
            Top             =   510
            Width           =   975
         End
         Begin VB.CheckBox chkMRE 
            Caption         =   "MRE"
            Height          =   195
            Left            =   1680
            TabIndex        =   18
            Top             =   780
            Width           =   975
         End
         Begin VB.CheckBox chkPlace 
            Caption         =   "Place"
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   1020
            Width           =   975
         End
         Begin VB.CheckBox chkOrganisation 
            Caption         =   "Organisation"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1290
            Width           =   1215
         End
         Begin VB.CheckBox chkPreview 
            Caption         =   "Preview"
            Height          =   285
            Left            =   210
            TabIndex        =   15
            Top             =   1590
            Value           =   1  'Checked
            Width           =   915
         End
         Begin VB.CheckBox chkUseLabels 
            Caption         =   "Use Labels"
            Height          =   240
            Left            =   270
            TabIndex        =   14
            Top             =   2130
            Width           =   1185
         End
      End
   End
   Begin VB.CommandButton cmdIMSMA 
      Caption         =   "IMSMA"
      Height          =   285
      Left            =   4890
      TabIndex        =   11
      Top             =   3330
      Width           =   1230
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   1725
      Left            =   8505
      TabIndex        =   9
      Top             =   315
      Width           =   1725
      Begin TatukGIS_DK.XGIS_ViewerWnd GIS 
         Height          =   780
         Left            =   225
         TabIndex        =   10
         Top             =   225
         Width           =   915
         BigExtentMargin =   -10
         RestrictedDrag  =   -1  'True
         CachedPaint     =   -1  'True
         IncrementalPaint=   -1  'True
         FullPaint       =   -1  'True
         CodePage        =   0
         OutCodePage     =   0
         CharSet         =   0
         UseRTree        =   0   'False
         PrinterTileSize =   1024
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
         SelectionPattern=   "frmIMSMAExporter.frx":4786B
         SelectionTransparency=   100
         SelectionWidth  =   100
         SelectionOutlineOnly=   0   'False
         OldCachedPaint  =   0   'False
         PrinterModeDraft=   0   'False
         PrinterModeForceBitmap=   0   'False
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
         ParentColor     =   0   'False
         ParentCtl3D     =   0   'False
         Object.Visible         =   -1  'True
         Cursor          =   16
         DoubleBuffered  =   0   'False
         ModeMouseButton =   0
         CursorForUserDefined=   0
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   6420
      Picture         =   "frmIMSMAExporter.frx":5217D
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   7
      Top             =   5310
      Width           =   720
   End
   Begin VB.CommandButton cmdINFO 
      Caption         =   "INFO"
      Height          =   285
      Left            =   5280
      TabIndex        =   6
      Top             =   6030
      Width           =   1230
   End
   Begin VB.Frame Frame4 
      Caption         =   "IMSMA Variables"
      Height          =   3330
      Left            =   6570
      TabIndex        =   0
      Top             =   3180
      Width           =   4815
      Begin VB.CommandButton Command3 
         Caption         =   "Update"
         Height          =   255
         Left            =   3840
         TabIndex        =   5
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   330
         Left            =   1440
         TabIndex        =   4
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   1095
         Left            =   600
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Arc Engine Home:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   1290
      End
      Begin VB.Label Label1 
         Caption         =   "Path:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label lblSeenSoon 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "IMSMA Utils"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5130
      TabIndex        =   8
      Top             =   6510
      Width           =   2145
   End
End
Attribute VB_Name = "frmIMSMAExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Private sPASS As String
'Private sUser As String
'
'Private m_binfo As Boolean
'Private m_bLyrsLoaded As Boolean
'
'Public HasItExported As Boolean
'Public OASISImported As Boolean
'
'Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
'
'
'Private Sub CreateExportTable()
'        '<EhHeader>
'        On Error GoTo CreateExportTable_Err
'        '</EhHeader>
'
'100     With m_oLyrHazards
'            .caption = "IMSMA Hazards"
'102         .AddField "Name", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'104         .AddField "DAType", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'106         .AddField "UsageType", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'
'        End With
'
'108     Set m_oLyrAccident = New TatukGIS_XDK9.XGIS_LayerVector
'
'110     With m_oLyrAccident
'            .caption = "IMSMA Accidents"
'112         .AddField "source", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'114         .AddField "dateofaccident", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'        End With
'
'116     Set m_oLyrHazRed = New TatukGIS_XDK9.XGIS_LayerVector
'
'118     With m_oLyrHazRed
'            .caption = "IMSMA Hazard Reduction"
'120         .AddField "hazreducname", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'122         .AddField "isactive", TatukGIS_XDK9.XGISFieldTypeNumber, 0, 0
'124         .AddField "startdate", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'126         .AddField "enddate", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'
'        End With
'
'128     Set m_oLyrVictim = New TatukGIS_XDK9.XGIS_LayerVector
'
'130     With m_oLyrVictim
'            .caption = "IMSMA Victims"
'132         .AddField "givenname", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'134         .AddField "familyname", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'136         .AddField "Age", TatukGIS_XDK9.XGISFieldTypeNumber, 0, 0
'
'        End With
'
'138     Set m_oLyrMRE = New TatukGIS_XDK9.XGIS_LayerVector
'
'140     With m_oLyrMRE
'            .caption = "IMSMA MRE"
'142         .AddField "orgname", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'144         .AddField "mrereason", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'146         .AddField "mredate", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'
'        End With
'
'148     Set m_oLyrQC = New TatukGIS_XDK9.XGIS_LayerVector
'
'150     With m_oLyrQC
'            .caption = "IMSMA QC"
'152         .AddField "startdate", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'154         .AddField "qadate", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'156         .AddField "areasize", TatukGIS_XDK9.XGISFieldTypeFloat, 0, 4
'
'        End With
'
'158     Set m_oLyrOrg = New TatukGIS_XDK9.XGIS_LayerVector
'
'160     With m_oLyrOrg
'            .caption = "IMSMA Organisations"
'162         .AddField "orgname", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'164         .AddField "orgstatus", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'166         .AddField "orgaddress", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'
'        End With
'
'168     Set m_oLyrPlace = New TatukGIS_XDK9.XGIS_LayerVector
'
'170     With m_oLyrPlace
'            .caption = "IMSMA Place"
'172         .AddField "placename", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'174         .AddField "placedescription", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'176         .AddField "placeaddress", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'
'        End With
'    '
'    '    Set m_oLyrLocation = New TatukGIS_XDK9.XGIS_LayerVector
'    '
'    '    With m_oLyrLocation
'    '
'    '        .AddField "", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'    '        .AddField "", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'    '        .AddField "", TatukGIS_XDK9.XGISFieldTypeString, 255, 0
'    '
'    '    End With
'    '
'    '    'm_oLyrLocation
'    '    'm_oLyrPlace
'    '    'm_oLyrOrg
'    '    'm_oLyrQC
'    '    'm_oLyrMRE
'    '    'm_oLyrVictim
'    '    'm_oLyrHazRed
'    '    'm_oLyrAccident
'    '
'    '    GIS.Add m_oLyrLocation
'178     GIS.Add m_oLyrPlace
'180     GIS.Add m_oLyrOrg
'182     GIS.Add m_oLyrQC
'184     GIS.Add m_oLyrMRE
'186     GIS.Add m_oLyrVictim
'188     GIS.Add m_oLyrHazRed
'190     GIS.Add m_oLyrAccident
'
'192     GIS.Add m_oLyrHazards
'        '<EhFooter>
'        Exit Sub
'
'CreateExportTable_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.CreateExportTable " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub ImportPlace(x As Double, _
'                           y As Double, _
'                           splacename As String, _
'                           sPlaceDescription As String, sPlaceAddress As String, _
'                           bUseLBL As Boolean)
'        '<EhHeader>
'        On Error GoTo ImportPlace_Err
'        '</EhHeader>
'        Dim oSHP As TatukGIS_XDK9.XGIS_ShapePoint
'        Dim ptg As New TatukGIS_XDK9.XGIS_Point
'
'100     ptg.Prepare x, y
'
'102     Set oSHP = m_oLyrPlace.CreateShape(XgisShapeTypePoint)
'
'104     With oSHP
'106         .Lock TatukGIS_XDK9.XGISLockExtent
'108         .AddPart
'110         .AddPoint ptg
'112         .SetField "placename", splacename
'114         .SetField "placedescription", sPlaceDescription
'116         .SetField "placeaddress", sPlaceAddress
'
'118         .Params.Marker.Color = RGB(123, 123, 123)
'120         If bUseLBL Then
'122             .Params.labels.Value = splacename
'124             .Params.labels.Visible = True
'            End If
'
'126         .Unlock
'        End With
'
'128     GIS.UpDate
'130     GIS.Mode = TatukGIS_XDK9.XGISZoomEx
'132     GIS.FullExtent
'
'        '<EhFooter>
'        Exit Sub
'
'ImportPlace_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.ImportPlace " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub ImportOrg(x As Double, _
'                           y As Double, _
'                           sorgname As String, _
'                           sorgstatus As String, sorgaddress As String, _
'                           bUseLBL As Boolean)
'        '<EhHeader>
'        On Error GoTo ImportOrg_Err
'        '</EhHeader>
'        Dim oSHP As TatukGIS_XDK9.XGIS_ShapePoint
'        Dim ptg As New TatukGIS_XDK9.XGIS_Point
'
'100     ptg.Prepare x, y
'
'102     Set oSHP = m_oLyrOrg.CreateShape(XgisShapeTypePoint)
'
'104     With oSHP
'106         .Lock TatukGIS_XDK9.XGISLockExtent
'108         .AddPart
'110         .AddPoint ptg
'112         .SetField "orgname", sorgname
'114         .SetField "orgstatus", sorgstatus
'116         .SetField "orgaddress", sorgaddress
'
'118         .Params.Marker.Color = RGB(43, 52, 255)
'
'120         If bUseLBL Then
'122             .Params.labels.Value = sorgname
'124             .Params.labels.Visible = True
'            End If
'
'126         .Unlock
'        End With
'
'128     GIS.UpDate
'130     GIS.Mode = TatukGIS_XDK9.XGISZoomEx
'132     GIS.FullExtent
'
'        '<EhFooter>
'        Exit Sub
'
'ImportOrg_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.ImportOrg " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'
'Private Sub ImportQC(x As Double, _
'                           y As Double, _
'                           sstartdate As String, _
'                           sqadate As String, dareasize As Double, _
'                           bUseLBL As Boolean)
'        '<EhHeader>
'        On Error GoTo ImportQC_Err
'        '</EhHeader>
'        Dim oSHP As TatukGIS_XDK9.XGIS_ShapePoint
'        Dim ptg As New TatukGIS_XDK9.XGIS_Point
'
'100     ptg.Prepare x, y
'
'102     Set oSHP = m_oLyrQC.CreateShape(XgisShapeTypePoint)
'
'104     With oSHP
'106         .Lock TatukGIS_XDK9.XGISLockExtent
'108         .AddPart
'110         .AddPoint ptg
'112         .SetField "startdate", sstartdate
'114         .SetField "qadate", sqadate
'116         .SetField "areasize", dareasize
'
'118         .Params.Marker.Color = vbBlack
'
'120         If bUseLBL Then
'122             .Params.labels.Value = "QC"
'124             .Params.labels.Visible = True
'            End If
'
'126         .Unlock
'        End With
'
'128     GIS.UpDate
'130     GIS.Mode = TatukGIS_XDK9.XGISZoomEx
'132     GIS.FullExtent
'
'        '<EhFooter>
'        Exit Sub
'
'ImportQC_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.ImportQC " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'
'Private Sub ImportMRE(x As Double, _
'                           y As Double, _
'                           sorgname As String, _
'                           smrereason As String, smredate As String, _
'                           bUseLBL As Boolean)
'        '<EhHeader>
'        On Error GoTo ImportMRE_Err
'        '</EhHeader>
'        Dim oSHP As TatukGIS_XDK9.XGIS_ShapePoint
'        Dim ptg As New TatukGIS_XDK9.XGIS_Point
'
'100     ptg.Prepare x, y
'
'102     Set oSHP = m_oLyrMRE.CreateShape(XgisShapeTypePoint)
'
'104     With oSHP
'106         .Lock TatukGIS_XDK9.XGISLockExtent
'108         .AddPart
'110         .AddPoint ptg
'112         .SetField "orgname", sorgname
'114         .SetField "mrereason", smrereason
'116         .SetField "mredate", smredate
'
'118         .Params.Marker.Color = vbYellow
'
'120         If bUseLBL Then
'122             .Params.labels.Value = "MRE By:" & sorgname
'124             .Params.labels.Visible = True
'            End If
'
'126         .Unlock
'        End With
'
'128     GIS.UpDate
'130     GIS.Mode = TatukGIS_XDK9.XGISZoomEx
'132     GIS.FullExtent
'
'        '<EhFooter>
'        Exit Sub
'
'ImportMRE_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.ImportMRE " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'Private Sub ImportVictim(x As Double, _
'                       y As Double, _
'                       sgivenname As String, sfamilyname As String, iAge As Integer, bUseLBL As Boolean)
'        '<EhHeader>
'        On Error GoTo ImportVictim_Err
'        '</EhHeader>
'        Dim oSHP As TatukGIS_XDK9.XGIS_ShapePoint
'        Dim ptg As New TatukGIS_XDK9.XGIS_Point
'
'100     ptg.Prepare x, y
'
'102     Set oSHP = m_oLyrVictim.CreateShape(XgisShapeTypePoint)
'
'104     With oSHP
'106         .Lock TatukGIS_XDK9.XGISLockExtent
'108         .AddPart
'110         .AddPoint ptg
'112         .SetField "givenname", sgivenname
'114         .SetField "familyname", sfamilyname
'116         .SetField "Age", iAge
'118         .Params.Marker.Color = vbMagenta
'
'120         If bUseLBL Then
'122             .Params.labels.Value = "Victim Age:" & iAge
'124             .Params.labels.Visible = True
'            End If
'126         .Unlock
'        End With
'
'128     GIS.UpDate
'130     GIS.Mode = TatukGIS_XDK9.XGISZoomEx
'132     GIS.FullExtent
'
'        '<EhFooter>
'        Exit Sub
'
'ImportVictim_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.ImportVictim " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'
'Private Sub ImportHazReduction(x As Double, _
'                       y As Double, _
'                       shazreducname As String, sstartdate As String, senddate As String, iisactive As Integer, bUseLBL As Boolean)
'        '<EhHeader>
'        On Error GoTo ImportHazReduction_Err
'        '</EhHeader>
'        Dim oSHP As TatukGIS_XDK9.XGIS_ShapePoint
'        Dim ptg As New TatukGIS_XDK9.XGIS_Point
'
'100     ptg.Prepare x, y
'
'102     Set oSHP = m_oLyrHazRed.CreateShape(XgisShapeTypePoint)
'
'104     With oSHP
'106         .Lock TatukGIS_XDK9.XGISLockExtent
'108         .AddPart
'110         .AddPoint ptg
'112         .SetField "hazreducname", shazreducname
'114         .SetField "isactive", iisactive
'116         .SetField "startdate", sstartdate
'118         .SetField "enddate", senddate
'120         .Params.Marker.Color = vbBlue
'
'122         If bUseLBL Then
'124             .Params.labels.Value = "Haz Red:" & shazreducname
'126             .Params.labels.Visible = True
'            End If
'
'128         .Unlock
'        End With
'
'130     GIS.UpDate
'132     GIS.Mode = TatukGIS_XDK9.XGISZoomEx
'134     GIS.FullExtent
'
'        '<EhFooter>
'        Exit Sub
'
'ImportHazReduction_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.ImportHazReduction " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub ImportAccident(x As Double, _
'                           y As Double, _
'                           sSource As String, _
'                           sdateofaccident As String, _
'                           bUseLBL As Boolean)
'        '<EhHeader>
'        On Error GoTo ImportAccident_Err
'        '</EhHeader>
'        Dim oSHP As TatukGIS_XDK9.XGIS_ShapePoint
'        Dim ptg As New TatukGIS_XDK9.XGIS_Point
'
'100     ptg.Prepare x, y
'
'102     Set oSHP = m_oLyrAccident.CreateShape(XgisShapeTypePoint)
'
'104     With oSHP
'106         .Lock TatukGIS_XDK9.XGISLockExtent
'108         .AddPart
'110         .AddPoint ptg
'112         .SetField "source", sSource
'114         .SetField "dateofaccident", sdateofaccident
'116         .Params.Marker.Color = vbGreen
'
'118         If bUseLBL Then
'120             .Params.labels.Value = "Accident Source:" & sSource
'122             .Params.labels.Visible = True
'            End If
'
'124         .Unlock
'        End With
'
'126     GIS.UpDate
'128     GIS.Mode = TatukGIS_XDK9.XGISZoomEx
'130     GIS.FullExtent
'
'        '<EhFooter>
'        Exit Sub
'
'ImportAccident_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.ImportAccident " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub Createpoly(x As Double, _
'                       y As Double, _
'                       sName As String, _
'                       sDAType As String, _
'                       sUsageType As String)
'        '<EhHeader>
'        On Error GoTo Createpoly_Err
'        '</EhHeader>
'        Dim oSHP As TatukGIS_XDK9.XGIS_ShapePoint
'        Dim ptg As New TatukGIS_XDK9.XGIS_Point
'
'100     ptg.Prepare x, y
'
'102     Set oSHP = m_oLyrHazards.CreateShape(XgisShapeTypePoint)
'
'104     With oSHP
'106         .Lock TatukGIS_XDK9.XGISLockExtent
'108         .AddPart
'110         .AddPoint ptg
'112         .SetField "Name", sName
'114         .SetField "DAType", sDAType
'116         .SetField "UsageType", sUsageType
'118         .Params.Marker.Color = vbRed 'RGB(CSng(Rnd(255) * 255), CSng(Rnd(255) * 255), CSng(Rnd(Rnd) * 255))
'
'120         If chkUseLabels.Value = vbChecked Then
'122             .Params.labels.Value = sDAType
'124             .Params.labels.Visible = True
'            End If
'
'126         .Unlock
'        End With
'
'128     GIS.UpDate
'130     GIS.Mode = TatukGIS_XDK9.XGISZoomEx
'132     GIS.FullExtent
'        '<EhFooter>
'        Exit Sub
'
'Createpoly_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.Createpoly " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdIMSMA_Click()
'        '<EhHeader>
''        On Error GoTo cmdIMSMA_Click_Err
''        '</EhHeader>
''        Dim WinhWnd As Long
''        Dim oldhwnd As Long  ' receives handle of button's former parent
''100     WinhWnd = FindWindow(vbNullString, "IMSMA Navigation")
''
''102     If WinhWnd = 0 Then
''104         MsgBox "IMSMA Client Must run before.. Start IMSMA first...", vbInformation
''            Exit Sub
''        End If
''
''106     frmIMSMAViewer.Show vbModeless, Me
''108     oldhwnd = SetParent(WinhWnd, frmIMSMAViewer.elMain.hwnd)
''110     MoveWindow WinhWnd, 0, -20, 1100, 750, 1
''        '<EhFooter>
''        Exit Sub
''
''cmdIMSMA_Click_Err:
''        MsgBox Err.Description & vbCrLf & _
''               "in OASISClient.frmIMSMAExporter.cmdIMSMA_Click " & _
''               "at line " & Erl
''        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub cmdINFO_Click()
'        '<EhHeader>
'        On Error GoTo cmdINFO_Click_Err
'        '</EhHeader>
'100     m_binfo = True
'        '<EhFooter>
'        Exit Sub
'
'cmdINFO_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.cmdINFO_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub Command1_Click()
'        '<EhHeader>
'        On Error GoTo Command1_Click_Err
'        '</EhHeader>
'        Dim conn As ADODB.Connection
'100     Set conn = New ADODB.Connection
'        Dim rs As ADODB.Recordset
'        Dim fld As ADODB.Field
'        Dim sql As String
'
'        'SELECT * FROM hazard WHERE hasgeodata = 1
'
'102     frmLoginIMSMA.Show vbModal, Me
'104     sPASS = frmLoginIMSMA.sPASS
'106     sUser = frmLoginIMSMA.sUser
'
'
'108     If chkPreview.Value = vbChecked Then
'110         If Not frmPreView.Visible Then
'112             SetParent GIS.hwnd, frmPreView.elMain.hwnd
'114             frmPreView.Show vbModeless, Me
'116             GIS.Move 0, 0
'
'118             DoEvents
'            End If
'        End If
'
'120     If Not m_bLyrsLoaded Then
'122         CreateExportTable
'124         m_bLyrsLoaded = True
'        End If
'
'126     If chkHazard.Value = vbChecked Then
'128         sql = sql & "SELECT hazard.hazardname AS Hazard_Name, location.locationname AS Location_name, geopoint.latitude,"
'130         sql = sql & "geopoint.longitude, hazard.areasize AS Area, hazard.isactive AS Active,"
'132         sql = sql & "hazard.dataentrydate AS Data_Entry_Date, imsmaenum2.enumvalue AS Type_Of_Hazard,"
'134         sql = sql & "imsmaenum.enumvalue AS Type_Of_Area "
'136         sql = sql & "FROM (((imsma.geopoint geopoint "
'138         sql = sql & "INNER JOIN imsma.geospatialinfo geospatialinfo "
'140         sql = sql & "ON (geopoint.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid)) "
'142         sql = sql & "INNER JOIN imsma.hazard_has_geospatialinfo hazard_has_geospatialinfo "
'144         sql = sql & "ON (hazard_has_geospatialinfo.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid)) "
'146         sql = sql & "INNER JOIN imsma.hazard hazard "
'148         sql = sql & "ON (hazard_has_geospatialinfo.hazard_guid = hazard.hazard_guid)) "
'150         sql = sql & "INNER JOIN imsma.location location "
'152         sql = sql & "ON (hazard.location_guid = location.location_guid) "
'154         sql = sql & "INNER JOIN imsma.imsmaenum imsmaenum "
'156         sql = sql & "ON (hazard.areatype_guid = imsmaenum.imsmaenum_guid) "
'158         sql = sql & "INNER JOIN imsma.imsmaenum imsmaenum2 "
'160         sql = sql & "ON (hazard.dangerousareatype_guid = imsmaenum2.imsmaenum_guid) "
'
'162         Set conn = New ADODB.Connection
'164         conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=localhost;" & " DATABASE=imsma;" & "UID=" & sUser & ";PWD=" & sPASS & ";"
'166         conn.Open
'
'168         Set rs = New ADODB.Recordset
'
'170         rs.Open sql, conn
'
'172         Do While Not rs.EOF
'174             Createpoly rs.Fields("longitude").Value, rs.Fields("latitude").Value, rs.Fields("Hazard_Name").Value, rs.Fields("Type_Of_Hazard").Value, rs.Fields("Type_Of_Area").Value
'176             rs.MoveNext
'            Loop
'
'        End If
'
'178     sql = ""
'
'180     If chkOrganisation.Value = vbChecked Then
'182         sql = sql & " SELECT geopoint.latitude,geopoint.longitude,geopoint.pointno,organisation.orgname,organisation.orgstatus,organisation.orgaddress"
'184         sql = sql & " FROM ((imsma.geopoint geopoint"
'186         sql = sql & " Inner Join imsma.geospatialinfo geospatialinfo"
'188         sql = sql & " ON (geopoint.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid))"
'190         sql = sql & " Inner Join imsma.organisation_has_geospatialinfo organisation_has_geospatialinfo"
'192         sql = sql & " ON (organisation_has_geospatialinfo.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid))"
'194         sql = sql & " Inner Join imsma.organisation organisation"
'196         sql = sql & " ON (organisation_has_geospatialinfo.org_guid = organisation.org_guid)"
'198         sql = sql & " Where (geopoint.pointno <> 666)"
'
'200         Set conn = New ADODB.Connection
'202         conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=localhost;" & " DATABASE=imsma;" & "UID=" & sUser & ";PWD=" & sPASS & ";"
'204         conn.Open
'
'206         Set rs = New ADODB.Recordset
'
'208         rs.Open sql, conn
'
'210         Do While Not rs.EOF
'212             ImportOrg rs.Fields("longitude").Value, rs.Fields("latitude").Value, rs.Fields("orgname").Value, rs.Fields("orgstatus").Value, rs.Fields("orgaddress").Value, IIf(chkUseLabels.Value = vbChecked, True, False)
'214             rs.MoveNext
'            Loop
'
'        End If
'
'216     sql = ""
'
'218     If chkAccident.Value = vbChecked Then
'220         sql = sql & "SELECT accident.dateofaccident,accident.source,geopoint.latitude,geopoint.Longitude "
'222         sql = sql & "FROM ((imsma.accident_has_geospatialinfo accident_has_geospatialinfo "
'224         sql = sql & "Inner Join"
'226         sql = sql & " imsma.geospatialinfo geospatialinfo"
'228         sql = sql & " ON (accident_has_geospatialinfo.geospatialinfo_guid ="
'230         sql = sql & " geospatialinfo.geospatialinfo_guid))"
'232         sql = sql & " Inner Join"
'234         sql = sql & " imsma.accident accident"
'236         sql = sql & " ON (accident_has_geospatialinfo.accident_guid ="
'238         sql = sql & " accident.accident_guid))"
'240         sql = sql & " Inner Join"
'242         sql = sql & " imsma.geopoint geopoint"
'244         sql = sql & " ON (geopoint.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid)"
'
'246         Set conn = New ADODB.Connection
'248         conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=localhost;" & " DATABASE=imsma;" & "UID=" & sUser & ";PWD=" & sPASS & ";"
'250         conn.Open
'
'252         Set rs = New ADODB.Recordset
'
'254         rs.Open sql, conn
'
'256         Do While Not rs.EOF
'258             ImportAccident rs.Fields("longitude").Value, rs.Fields("latitude").Value, IIf(IsNull(rs.Fields("source").Value), "", rs.Fields("source").Value), rs.Fields("dateofaccident").Value, IIf(chkUseLabels.Value = vbChecked, True, False)
'260             rs.MoveNext
'            Loop
'
'        End If
'
'262     sql = ""
'
'264     If chkLocation.Value = vbChecked Then
'            '        sql = sql & " SELECT location.locationname,"
'            '        sql = sql & " location.nearestmedicalfacility,"
'            '        sql = sql & " location.locationdescription"
'            '        sql = sql & " FROM imsma.location location"
'            '
'            '
'            '       Set conn = New ADODB.Connection
'            '        conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=localhost;" & " DATABASE=imsma;" & "UID=" & sUser & ";PWD=" & sPass & ";"
'            '        conn.Open
'            '
'            '        Set rs = New ADODB.Recordset
'            '
'            '        rs.Open sql, conn
'            '
'            '        Do While Not rs.EOF
'            '            Createpoly rs.Fields("longitude").Value, rs.Fields("latitude").Value, "Location", "Location", "Location"
'            '            rs.MoveNext
'            '        Loop
'
'        End If
'
'266     sql = ""
'
'268     If chkQC.Value = vbChecked Then
'
'270         sql = sql & " SELECT geopoint.longitude,geopoint.latitude,qa.qadate,qa.startdate,qa.areasize"
'272         sql = sql & " FROM ((imsma.qa_has_geospatialinfo qa_has_geospatialinfo"
'274         sql = sql & " Inner Join imsma.geospatialinfo geospatialinfo"
'276         sql = sql & " ON (qa_has_geospatialinfo.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid))"
'278         sql = sql & " Inner Join imsma.qa qa"
'280         sql = sql & " ON (qa_has_geospatialinfo.qa_guid = qa.qa_guid))"
'282         sql = sql & " Inner Join imsma.geopoint geopoint"
'284         sql = sql & " ON (geopoint.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid)"
'
'286         Set conn = New ADODB.Connection
'288         conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=localhost;" & " DATABASE=imsma;" & "UID=" & sUser & ";PWD=" & sPASS & ";"
'290         conn.Open
'
'292         Set rs = New ADODB.Recordset
'
'294         rs.Open sql, conn
'
'296         Do While Not rs.EOF
'298             ImportQC rs.Fields("longitude").Value, rs.Fields("latitude").Value, rs.Fields("startdate").Value, rs.Fields("qadate").Value, rs.Fields("areasize").Value, IIf(chkUseLabels.Value = vbChecked, True, False)
'300             rs.MoveNext
'            Loop
'
'        End If
'
'302     sql = ""
'
'304     If chkVictim.Value = vbChecked Then
'306         sql = sql & " SELECT geopoint.latitude,geopoint.longitude,victim.familyname,victim.givenname,victim.Age"
'308         sql = sql & " FROM ((imsma.victim_has_geospatialinfo victim_has_geospatialinfo"
'310         sql = sql & " Inner Join imsma.geospatialinfo geospatialinfo"
'312         sql = sql & " ON (victim_has_geospatialinfo.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid))"
'314         sql = sql & " Inner Join imsma.victim victim"
'316         sql = sql & " ON (victim_has_geospatialinfo.victim_guid = victim.victim_guid))"
'318         sql = sql & " Inner Join imsma.geopoint geopoint"
'320         sql = sql & " ON (geopoint.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid)"
'
'322         Set conn = New ADODB.Connection
'324         conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=localhost;" & " DATABASE=imsma;" & "UID=" & sUser & ";PWD=" & sPASS & ";"
'326         conn.Open
'
'328         Set rs = New ADODB.Recordset
'
'330         rs.Open sql, conn
'
'332         Do While Not rs.EOF
'
'334             ImportVictim rs.Fields("longitude").Value, rs.Fields("latitude").Value, rs.Fields("sgivenname").Value, rs.Fields("familyname").Value, rs.Fields("age").Value, IIf(chkUseLabels.Value = vbChecked, True, False)
'336             Createpoly rs.Fields("longitude").Value, rs.Fields("latitude").Value, "victim", "Victim", "Victim"
'338             rs.MoveNext
'            Loop
'
'        End If
'
'340     sql = ""
'
'342     If chkPlace.Value Then
'344         sql = sql & "  SELECT place.placename,geopoint.latitude,geopoint.longitude,place.placedescription,place.placeaddress"
'346         sql = sql & " FROM ((imsma.place_has_geospatialinfo place_has_geospatialinfo"
'348         sql = sql & " INNER JOIN imsma.geospatialinfo geospatialinfo"
'350         sql = sql & " ON (place_has_geospatialinfo.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid))"
'352         sql = sql & " INNER JOIN imsma.place place"
'354         sql = sql & " ON (place_has_geospatialinfo.place_guid = place.place_guid))"
'356         sql = sql & " INNER JOIN imsma.geopoint geopoint"
'358         sql = sql & " ON (geopoint.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid)"
'
'360         Set conn = New ADODB.Connection
'362         conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=localhost;" & " DATABASE=imsma;" & "UID=" & sUser & ";PWD=" & sPASS & ";"
'364         conn.Open
'
'366         Set rs = New ADODB.Recordset
'
'368         rs.Open sql, conn
'
'370         Do While Not rs.EOF
'                Dim splacename As String
'                Dim sPlaceDescription As String
'                Dim sPlaceAddress As String
'                On Error Resume Next
'
'372             If Not IsNull(rs.Fields("placeaddress").Value) Then
'374                 sPlaceAddress = rs.Fields("placeaddress").Value
'                End If
'
'376             If Not IsNull(rs.Fields("placename").Value) Then
'378                 splacename = rs.Fields("placename").Value
'                End If
'
'380             If Not IsNull(rs.Fields("placedescription").Value) Then
'382                 sPlaceDescription = rs.Fields("placedescription").Value
'                End If
'
'384             ImportPlace rs.Fields("longitude").Value, rs.Fields("latitude").Value, splacename, sPlaceDescription, sPlaceAddress, IIf(chkUseLabels.Value = vbChecked, True, False)
'386             rs.MoveNext
'            Loop
'
'        End If
'
'388     sql = ""
'
'390     If chkMRE.Value = vbChecked Then
'392         sql = sql & " SELECT geopoint.latitude,geopoint.longitude,mre.mredate,organisation.orgname,mre.mrereason"
'394         sql = sql & " FROM(((imsma.geopoint geopoint"
'396         sql = sql & " Inner Join imsma.geospatialinfo geospatialinfo"
'398         sql = sql & " ON (geopoint.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid))"
'400         sql = sql & " Inner Join imsma.mre_has_geospatialinfo mre_has_geospatialinfo"
'402         sql = sql & " ON (mre_has_geospatialinfo.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid))"
'404         sql = sql & " Inner Join imsma.mre mre"
'406         sql = sql & " ON (mre_has_geospatialinfo.mre_guid = mre.mre_guid))"
'408         sql = sql & " Inner Join imsma.organisation organisation"
'410         sql = sql & " ON (mre.org_guid = organisation.org_guid)"
'
'412         Set conn = New ADODB.Connection
'414         conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=localhost;" & " DATABASE=imsma;" & "UID=" & sUser & ";PWD=" & sPASS & ";"
'416         conn.Open
'
'418         Set rs = New ADODB.Recordset
'
'420         rs.Open sql, conn
'
'422         Do While Not rs.EOF
'424             Createpoly rs.Fields("longitude").Value, rs.Fields("latitude").Value, "MRE", "MRE", "MRE"
'426             ImportMRE rs.Fields("longitude").Value, rs.Fields("latitude").Value, rs.Fields("orgname").Value, rs.Fields("mrereason").Value, rs.Fields("mredate").Value, IIf(chkUseLabels.Value = vbChecked, True, False)
'428             rs.MoveNext
'            Loop
'
'        End If
'
'430     If chkOrganisation.Value = vbChecked Then
'
'        End If
'
'432     sql = ""
'
'434     If chkHazardReduction.Value = vbChecked Then
'436         sql = sql & " SELECT geopoint.latitude,geopoint.longitude,hazreduc.hazreducname,hazreduc.startdate,hazreduc.enddate,geopoint.pointno,hazreduc.isactive"
'438         sql = sql & " FROM((imsma.geopoint geopoint"
'440         sql = sql & " Inner Join imsma.geospatialinfo geospatialinfo"
'442         sql = sql & " ON (geopoint.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid))"
'444         sql = sql & " Inner Join imsma.hazreduc_has_geospatialinfo hazreduc_has_geospatialinfo"
'446         sql = sql & " ON (hazreduc_has_geospatialinfo.geospatialinfo_guid = geospatialinfo.geospatialinfo_guid))"
'448         sql = sql & " Inner Join imsma.hazreduc hazreduc"
'450         sql = sql & " ON (hazreduc_has_geospatialinfo.hazreduc_guid = hazreduc.hazreduc_guid)"
'452         sql = sql & " Where (geopoint.pointno = 1)"
'454         Set conn = New ADODB.Connection
'456         conn.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" & "SERVER=localhost;" & " DATABASE=imsma;" & "UID=" & sUser & ";PWD=" & sPASS & ";"
'458         conn.Open
'
'460         Set rs = New ADODB.Recordset
'
'462         rs.Open sql, conn
'
'464         Do While Not rs.EOF
'                '        ImportOrg rs.Fields("longitude").Value, rs.Fields("latitude").Value, rs.Fields("orgname").Value, rs.Fields("orgstatus").Value, rs.Fields("orgaddress").Value, IIf(chkAccident.Value = vbChecked, True, False)
'                Dim hazreducname As String
'                Dim StartDate As String
'                Dim enddate As String
'                Dim isactive As Integer
'
'466             If Not IsNull(rs.Fields("hazreducname").Value) Then
'468                 hazreducname = rs.Fields("hazreducname").Value
'                End If
'
'470             If Not IsNull(rs.Fields("startdate").Value) Then
'472                 StartDate = rs.Fields("startdate").Value
'                End If
'
'474             If Not IsNull(rs.Fields("enddate").Value) Then
'476                 enddate = rs.Fields("enddate").Value
'                End If
'
'478             If Not IsNull(rs.Fields("isactive").Value) Then
'480                 isactive = rs.Fields("isactive").Value
'                End If
'
'482             ImportHazReduction rs.Fields("longitude").Value, rs.Fields("latitude").Value, hazreducname, StartDate, enddate, isactive, IIf(chkUseLabels.Value = vbChecked, True, False)
'
'484             rs.MoveNext
'            Loop
'
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'Command1_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.Command1_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub Command2_Click()
'        '<EhHeader>
'        On Error GoTo Command2_Click_Err
'        '</EhHeader>
'        Dim sPath As String
'        Dim oKMLlyr As New TatukGIS_XDK9.XGIS_LayerKML
'        Dim oGPXLyr As New TatukGIS_XDK9.XGIS_LayerGPX
'        Dim oSHPLyr As New TatukGIS_XDK9.XGIS_LayerSHP
'        Dim oGMLLyr As New TatukGIS_XDK9.XGIS_LayerGML
'        Dim oTABLyr As New TatukGIS_XDK9.XGIS_LayerTAB
'        Dim oMidMif As New TatukGIS_XDK9.XGIS_LayerMIF
'        Dim GisUtils As New TatukGIS_XDK9.XGIS_Utils
'        Dim sMess As String
'
'        'On Error Resume Next
'
'100     sMess = "The Following Formats were Exported: "
'
'102     sPath = g_sAppPath & "\IMSMAExport" & Year(Now) & Month(Now) & Day(Now) & Hour(Now) & Minute(Now) & Second(Now)
'
'104     If chkExpFormat(2).Value = vbChecked Then
'106         Set oKMLlyr = New TatukGIS_XDK9.XGIS_LayerKML
'108         oKMLlyr.Path = sPath & "_Hazards.kml"
'110         m_oLyrHazards.ExportLayer oKMLlyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'112         oKMLlyr.SaveAll
'
'114         Set oKMLlyr = New TatukGIS_XDK9.XGIS_LayerKML
'116         oKMLlyr.Path = sPath & "_Accident.kml"
'118         m_oLyrAccident.ExportLayer oKMLlyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'120         oKMLlyr.SaveAll
'
'122         Set oKMLlyr = New TatukGIS_XDK9.XGIS_LayerKML
'124         oKMLlyr.Path = sPath & "_Hazard_Reduction.kml"
'126         m_oLyrHazRed.ExportLayer oKMLlyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'128         oKMLlyr.SaveAll
'
'130         Set oKMLlyr = New TatukGIS_XDK9.XGIS_LayerKML
'132         oKMLlyr.Path = sPath & "_MRE.kml"
'134         m_oLyrMRE.ExportLayer oKMLlyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'136         oKMLlyr.SaveAll
'
'138         Set oKMLlyr = New TatukGIS_XDK9.XGIS_LayerKML
'140         oKMLlyr.Path = sPath & "_Organisation.kml"
'142         m_oLyrOrg.ExportLayer oKMLlyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'144         oKMLlyr.SaveAll
'
'146         Set oKMLlyr = New TatukGIS_XDK9.XGIS_LayerKML
'148         oKMLlyr.Path = sPath & "_Place.kml"
'150         m_oLyrPlace.ExportLayer oKMLlyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'152         oKMLlyr.SaveAll
'
'154         Set oKMLlyr = New TatukGIS_XDK9.XGIS_LayerKML
'156         oKMLlyr.Path = sPath & "_QC.kml"
'158         m_oLyrQC.ExportLayer oKMLlyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'160         oKMLlyr.SaveAll
'
'162         Set oKMLlyr = New TatukGIS_XDK9.XGIS_LayerKML
'164         oKMLlyr.Path = sPath & "_Victims.kml"
'166         m_oLyrVictim.ExportLayer oKMLlyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'168         oKMLlyr.SaveAll
'
'170         sMess = sMess & vbCrLf & "Google KML format Exported Succesfully..."
'        End If
'
'172     If chkExpFormat(3).Value = vbChecked Then
'
'174         Set oGPXLyr = New TatukGIS_XDK9.XGIS_LayerGPX
'176         oGPXLyr.Path = sPath & "_Hazards.gpx"
'178         m_oLyrHazards.ExportLayer oGPXLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'180         oGPXLyr.SaveAll
'
'182         Set oGPXLyr = New TatukGIS_XDK9.XGIS_LayerGPX
'184         oGPXLyr.Path = sPath & "_Accident.gpx"
'186         m_oLyrAccident.ExportLayer oGPXLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'188         oGPXLyr.SaveAll
'
'190         Set oGPXLyr = New TatukGIS_XDK9.XGIS_LayerGPX
'192         oGPXLyr.Path = sPath & "_Hazard_Reduction.gpx"
'194         m_oLyrHazRed.ExportLayer oGPXLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'196         oGPXLyr.SaveAll
'
'198         Set oGPXLyr = New TatukGIS_XDK9.XGIS_LayerGPX
'200         oGPXLyr.Path = sPath & "_MRE.gpx"
'202         m_oLyrMRE.ExportLayer oGPXLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'204         oGPXLyr.SaveAll
'
'206         Set oGPXLyr = New TatukGIS_XDK9.XGIS_LayerGPX
'208         oGPXLyr.Path = sPath & "_Organisation.gpx"
'210         m_oLyrOrg.ExportLayer oGPXLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'212         oGPXLyr.SaveAll
'
'214         Set oGPXLyr = New TatukGIS_XDK9.XGIS_LayerGPX
'216         oGPXLyr.Path = sPath & "_Place.gpx"
'218         m_oLyrPlace.ExportLayer oGPXLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'220         oGPXLyr.SaveAll
'
'222         Set oGPXLyr = New TatukGIS_XDK9.XGIS_LayerGPX
'224         oGPXLyr.Path = sPath & "_QC.gpx"
'226         m_oLyrQC.ExportLayer oGPXLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'228         oGPXLyr.SaveAll
'
'230         Set oGPXLyr = New TatukGIS_XDK9.XGIS_LayerGPX
'232         oGPXLyr.Path = sPath & "_Victims.gpx"
'234         m_oLyrVictim.ExportLayer oGPXLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'236         oGPXLyr.SaveAll
'
'238         sMess = sMess & vbCrLf & "GPS gpx Format Exported Successfully..."
'
'        End If
'
'240     If chkExpFormat(4).Value = vbChecked Then
'
'242         Set oGMLLyr = New TatukGIS_XDK9.XGIS_LayerGML
'244         oGMLLyr.Path = sPath & "_Hazards.gml"
'246         m_oLyrHazards.ExportLayer oGMLLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'248         oGMLLyr.SaveAll
'
'250         Set oGMLLyr = New TatukGIS_XDK9.XGIS_LayerGML
'252         oGMLLyr.Path = sPath & "_Accident.gml"
'254         m_oLyrAccident.ExportLayer oGMLLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'256         oGMLLyr.SaveAll
'
'258         Set oGMLLyr = New TatukGIS_XDK9.XGIS_LayerGML
'260         oGMLLyr.Path = sPath & "_Hazard_Reduction.gml"
'262         m_oLyrHazRed.ExportLayer oGMLLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'264         oGMLLyr.SaveAll
'
'266         Set oGMLLyr = New TatukGIS_XDK9.XGIS_LayerGML
'268         oGMLLyr.Path = sPath & "_MRE.gml"
'270         m_oLyrMRE.ExportLayer oGMLLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'272         oGMLLyr.SaveAll
'
'274         Set oGMLLyr = New TatukGIS_XDK9.XGIS_LayerGML
'276         oGMLLyr.Path = sPath & "_Organisation.gml"
'278         m_oLyrOrg.ExportLayer oGMLLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'280         oGMLLyr.SaveAll
'
'282         Set oGMLLyr = New TatukGIS_XDK9.XGIS_LayerGML
'284         oGMLLyr.Path = sPath & "_Place.gml"
'286         m_oLyrPlace.ExportLayer oGMLLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'288         oGMLLyr.SaveAll
'
'290         Set oGMLLyr = New TatukGIS_XDK9.XGIS_LayerGML
'292         oGMLLyr.Path = sPath & "_QC.gml"
'294         m_oLyrQC.ExportLayer oGMLLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'296         oGMLLyr.SaveAll
'
'298         Set oGMLLyr = New TatukGIS_XDK9.XGIS_LayerGML
'300         oGMLLyr.Path = sPath & "_Victims.gml"
'302         m_oLyrVictim.ExportLayer oGMLLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'304         oGMLLyr.SaveAll
'
'
'306         sMess = sMess & vbCrLf & "Geographic Markup Language - GML Exported successfully..."
'
'        End If
'
'308     If chkExpFormat(0).Value = vbChecked Then
'
'310         Set oSHPLyr = New TatukGIS_XDK9.XGIS_LayerSHP
'312         oSHPLyr.Path = sPath & "_Hazards.shp"
'314         m_oLyrHazards.ExportLayer oSHPLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'316         oSHPLyr.SaveAll
'
'318         Set oSHPLyr = New TatukGIS_XDK9.XGIS_LayerSHP
'320         oSHPLyr.Path = sPath & "_Accident.shp"
'322         m_oLyrAccident.ExportLayer oSHPLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'324         oSHPLyr.SaveAll
'
'326         Set oSHPLyr = New TatukGIS_XDK9.XGIS_LayerSHP
'328         oSHPLyr.Path = sPath & "_Hazard_Reduction.shp"
'330         m_oLyrHazRed.ExportLayer oSHPLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'332         oSHPLyr.SaveAll
'
'334         Set oSHPLyr = New TatukGIS_XDK9.XGIS_LayerSHP
'336         oSHPLyr.Path = sPath & "_MRE.shp"
'338         m_oLyrMRE.ExportLayer oSHPLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'340         oSHPLyr.SaveAll
'
'342         Set oSHPLyr = New TatukGIS_XDK9.XGIS_LayerSHP
'344         oSHPLyr.Path = sPath & "_Organisation.shp"
'346         m_oLyrOrg.ExportLayer oSHPLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'348         oSHPLyr.SaveAll
'
'350         Set oSHPLyr = New TatukGIS_XDK9.XGIS_LayerSHP
'352         oSHPLyr.Path = sPath & "_Place.shp"
'354         m_oLyrPlace.ExportLayer oSHPLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'356         oSHPLyr.SaveAll
'
'358         Set oSHPLyr = New TatukGIS_XDK9.XGIS_LayerSHP
'360         oSHPLyr.Path = sPath & "_QC.shp"
'362         m_oLyrQC.ExportLayer oSHPLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'364         oSHPLyr.SaveAll
'
'366         Set oSHPLyr = New TatukGIS_XDK9.XGIS_LayerSHP
'368         oSHPLyr.Path = sPath & "_Victims.shp"
'370         m_oLyrVictim.ExportLayer oSHPLyr, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'372         oSHPLyr.SaveAll
'
'374         sMess = sMess & vbCrLf & "ESRI Shapefile Exported Successfully.."
'        End If
'
'376     If chkExpFormat(1).Value = vbChecked Then
'
'378         Set oMidMif = New TatukGIS_XDK9.XGIS_LayerMIF
'380         oMidMif.Path = sPath & "_Hazards.mif"
'382         m_oLyrHazards.ExportLayer oMidMif, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'384         oMidMif.SaveAll
'
'386         Set oMidMif = New TatukGIS_XDK9.XGIS_LayerMIF
'388         oMidMif.Path = sPath & "_Accident.mif"
'390         m_oLyrAccident.ExportLayer oMidMif, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'392         oMidMif.SaveAll
'
'394         Set oMidMif = New TatukGIS_XDK9.XGIS_LayerMIF
'396         oMidMif.Path = sPath & "_Hazard_Reduction.mif"
'398         m_oLyrHazRed.ExportLayer oMidMif, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'400         oMidMif.SaveAll
'
'402         Set oMidMif = New TatukGIS_XDK9.XGIS_LayerMIF
'404         oMidMif.Path = sPath & "_MRE.mif"
'406         m_oLyrMRE.ExportLayer oMidMif, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'408         oMidMif.SaveAll
'
'410         Set oMidMif = New TatukGIS_XDK9.XGIS_LayerMIF
'412         oMidMif.Path = sPath & "_Organisation.mif"
'414         m_oLyrOrg.ExportLayer oMidMif, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'416         oMidMif.SaveAll
'
'418         Set oMidMif = New TatukGIS_XDK9.XGIS_LayerMIF
'420         oMidMif.Path = sPath & "_Place.mif"
'422         m_oLyrPlace.ExportLayer oMidMif, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'424         oMidMif.SaveAll
'
'426         Set oMidMif = New TatukGIS_XDK9.XGIS_LayerMIF
'428         oMidMif.Path = sPath & "_QC.mif"
'430         m_oLyrQC.ExportLayer oMidMif, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'432         oMidMif.SaveAll
'
'434         Set oMidMif = New TatukGIS_XDK9.XGIS_LayerMIF
'436         oMidMif.Path = sPath & "_Victims.mif"
'438         m_oLyrVictim.ExportLayer oMidMif, GisUtils.GisWholeWorld, TatukGIS_XDK9.XGISShapeTypeUnknown, "", True
'440         oMidMif.SaveAll
'
'442         sMess = sMess & vbCrLf & "Mapinfo MID/MIF File Exported successfully..."
'        End If
'
'444     MsgBox sMess, vbInformation, "iMMAP Dudes... They Just do it!!!"
'        HasItExported = True
'        '<EhFooter>
'        Exit Sub
'
'Command2_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmIMSMAExporter.Command2_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub Command3_Click()
'        '<EhHeader>
''        On Error GoTo Command3_Click_Err
''        '</EhHeader>
''        Dim SysObj As New SysEnv
''
''100     SysObj.SetEnv "Path", Text1.Text
''102     SysObj.SetEnv "ARCENGINE_HOME", Text2.Text
''104     MsgBox "Done! Proudly provided to you by the fantastic iMMAP Team", vbInformation
''        '<EhFooter>
''        Exit Sub
''
''Command3_Click_Err:
''        MsgBox Err.Description & vbCrLf & _
''               "in OASISClient.frmIMSMAExporter.Command3_Click " & _
''               "at line " & Erl
''        Resume Next
'        '</EhFooter>
'End Sub
'
'Private Sub Form_Load()
'    '<EhHeader>
'    On Error GoTo Form_Load_Err
'    '</EhHeader>
'    Dim i As Integer
'
'     '   frmLogin.Show vbModal, Me
'     '   sPASS = frmLogin.sPASS
'     '   sUser = frmLogin.sUser
'     '
'     '   For i = 1 To 200
'     '       Debug.Print Environ(i)
'     '   Next i
'    '
'     '   Text1.Text = Environ("Path")
'     '   Text2.Text = Environ("ARCENGINE_HOME")
'    '<EhFooter>
'    Exit Sub
'
'Form_Load_Err:
'    MsgBox Err.Description & vbCrLf & _
'           "in OASISClient.frmIMSMAExporter.Form_Load " & _
'           "at line " & Erl
'    Resume Next
'    '</EhFooter>
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'   ' MsgBox "Thanks for using Quality Utilities from iMMAP...", vbInformation
'End Sub
'
'        'sFont = .Item("Font_Name").value
'        'sFont = sFont & ":" & .Item("Ascii").value & ":NORMAL"
'
''146     m_oW3Lyr.Params.Marker.Color = vbWhite
''148     m_oW3Lyr.Params.Marker.OutlineColor = vbBlue
''150     m_oW3Lyr.Params.Marker.Symbol = SymbolList.Prepare(Replace(g_sAppPath, "\", "\\\\") & "\\\\Data\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE") '"..\\\\..\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE") 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
''152     m_oW3Lyr.Params.Marker.Size = 440
'
'    '        m_oW3Lyr.Params.LabelField = "DAType"
'    '        m_oW3Lyr.Params.LabelVisible = True
'
''158     m_oW3Lyr.Paint
