VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Begin VB.Form frmMapPrint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS - Print Template Creator"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11085
   Icon            =   "frmMapPrintAdmin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   11085
   StartUpPosition =   2  'CenterScreen
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   8505
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11085
      _cx             =   19553
      _cy             =   15002
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
      Begin TatukGIS_DK.XGIS_ControlPrintPreviewSimple PrintPreviewSimple 
         Left            =   9690
         Top             =   225
         Caption         =   "Print Preview"
         WindowLeft      =   0
         WindowTop       =   0
         WindowWidth     =   640
         WindowHeight    =   480
         DoubleBuffered  =   0   'False
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "New"
         Enabled         =   0   'False
         Height          =   360
         Left            =   6750
         TabIndex        =   41
         Top             =   7890
         Width           =   720
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "Open"
         Enabled         =   0   'False
         Height          =   360
         Left            =   9840
         TabIndex        =   40
         Top             =   7890
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   360
         Left            =   9030
         TabIndex        =   39
         Top             =   7890
         Width           =   750
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   360
         Left            =   7530
         TabIndex        =   31
         Top             =   7890
         Width           =   720
      End
      Begin C1SizerLibCtl.C1Tab C1TTabMain 
         Height          =   8205
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Width           =   6585
         _cx             =   11615
         _cy             =   14473
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
         Caption         =   "Map|Print Preview"
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
         Picture(0)      =   "frmMapPrintAdmin.frx":6852
         Picture(1)      =   "frmMapPrintAdmin.frx":D0B4
         Begin C1SizerLibCtl.C1Elastic elPreview 
            Height          =   7740
            Left            =   7230
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   420
            Width           =   6495
            _cx             =   11456
            _cy             =   13653
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
            Begin TatukGIS_DK.XGIS_ControlPrintPreview PrintPreview 
               Height          =   7680
               Left            =   150
               TabIndex        =   27
               Top             =   90
               Width           =   6225
               Align           =   0
               BevelInner      =   0
               BevelOuter      =   0
               BorderStyle     =   0
               Ctl3D           =   -1  'True
               Color           =   8421504
               Enabled         =   -1  'True
               ParentColor     =   0   'False
               ParentCtl3D     =   -1  'True
               Object.Visible         =   -1  'True
               ParentBackground=   0   'False
               DoubleBuffered  =   0   'False
               BevelWidth      =   1
               BorderWidth     =   0
               HelpContextId   =   0
               TabOrder        =   -1
               TabStop         =   0   'False
            End
         End
         Begin C1SizerLibCtl.C1Elastic elMap 
            Height          =   7740
            Left            =   45
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   420
            Width           =   6495
            _cx             =   11456
            _cy             =   13653
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
            Begin TatukGIS_DK.XGIS_ViewerWnd GIS 
               Height          =   6945
               Left            =   60
               TabIndex        =   25
               Top             =   90
               Width           =   6375
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
               SelectionPattern=   "frmMapPrintAdmin.frx":13916
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
            Begin TatukGIS_DK.XGIS_ControlScale ControlScale1 
               Height          =   375
               Left            =   3540
               TabIndex        =   108
               Top             =   6600
               Width           =   2835
               Dividers        =   5
               UnitsType       =   2
               UnitsTxt        =   ""
               Align           =   0
               BevelInner      =   0
               BevelOuter      =   2
               BorderStyle     =   1
               Color           =   -16777201
               Ctl3D           =   0   'False
               Enabled         =   -1  'True
               FullRepaint     =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontColor       =   -2147483630
               ParentColor     =   0   'False
               ParentCtl3D     =   0   'False
               ParentFont      =   -1  'True
               Object.Visible         =   -1  'True
               ParentBackground=   -1  'True
               DoubleBuffered  =   0   'False
            End
            Begin TatukGIS_DK.XGIS_ControlLegend Legend1 
               Height          =   2415
               Left            =   60
               TabIndex        =   53
               Top             =   4590
               Width           =   1695
               BorderStyle     =   1
               BeginProperty FontTitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontColorTitle  =   -16777208
               BeginProperty FontSubtitle {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               FontColorSubTitle=   -16777208
               Spacing         =   1
               ReverseOrder    =   -1  'True
               Align           =   0
               Ctl3D           =   0   'False
               Color           =   16777215
               Enabled         =   -1  'True
               ParentColor     =   0   'False
               ParentCtl3D     =   0   'False
               Object.Visible         =   -1  'True
               DoubleBuffered  =   -1  'True
               AllowMove       =   -1  'True
               AllowActive     =   -1  'True
               AllowExpand     =   -1  'True
               AllowParams     =   -1  'True
            End
            Begin VB.Frame FraMapTools 
               Caption         =   "Map Tools:"
               Height          =   675
               Left            =   60
               TabIndex        =   49
               Top             =   7050
               Width           =   6375
               Begin VB.CommandButton cmdMapTools 
                  Height          =   315
                  Index           =   2
                  Left            =   1020
                  Picture         =   "frmMapPrintAdmin.frx":1E228
                  Style           =   1  'Graphical
                  TabIndex        =   52
                  Top             =   240
                  Width           =   375
               End
               Begin VB.CommandButton cmdMapTools 
                  Height          =   315
                  Index           =   1
                  Left            =   600
                  Picture         =   "frmMapPrintAdmin.frx":24A7A
                  Style           =   1  'Graphical
                  TabIndex        =   51
                  Top             =   240
                  Width           =   405
               End
               Begin VB.CommandButton cmdMapTools 
                  Height          =   315
                  Index           =   0
                  Left            =   180
                  Picture         =   "frmMapPrintAdmin.frx":2B2CC
                  Style           =   1  'Graphical
                  TabIndex        =   50
                  Top             =   240
                  Width           =   405
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Tab C1TTabPrint 
         Height          =   7785
         Left            =   6750
         TabIndex        =   8
         Top             =   30
         Width           =   4245
         _cx             =   7488
         _cy             =   13732
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
         Caption         =   "Settings|TPLs|Objects|Text|Editor"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   5
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
         Begin C1SizerLibCtl.C1Elastic elsettings 
            Height          =   7410
            Left            =   45
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   330
            Width           =   4155
            _cx             =   7329
            _cy             =   13070
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
            Begin VB.Frame FraClientSettings 
               Caption         =   "Client Settings"
               Height          =   2505
               Left            =   150
               TabIndex        =   259
               Top             =   4800
               Width           =   3945
               Begin VB.CheckBox chkAllowChange 
                  Caption         =   "Allow Change of Text Objects"
                  Height          =   315
                  Left            =   120
                  TabIndex        =   260
                  Top             =   240
                  Width           =   2475
               End
            End
            Begin VB.Frame FraToolSettings 
               Caption         =   "Tool Settings:"
               Height          =   1455
               Left            =   150
               TabIndex        =   117
               Top             =   3270
               Width           =   3945
               Begin VB.CheckBox chkSaveText 
                  Caption         =   "Save Text In Templates "
                  Height          =   255
                  Left            =   120
                  TabIndex        =   256
                  Top             =   690
                  Width           =   2385
               End
               Begin VB.ComboBox ComUnits 
                  Height          =   315
                  ItemData        =   "frmMapPrintAdmin.frx":31B1E
                  Left            =   540
                  List            =   "frmMapPrintAdmin.frx":31B34
                  Style           =   2  'Dropdown List
                  TabIndex        =   253
                  Top             =   960
                  Width           =   2025
               End
               Begin VB.CheckBox chkAutoSave 
                  Caption         =   "Auto save template"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   119
                  Top             =   450
                  Width           =   2655
               End
               Begin VB.CheckBox chkShowPrinter 
                  Caption         =   "Show Printer Settings On update"
                  Height          =   240
                  Left            =   120
                  TabIndex        =   118
                  Top             =   240
                  Width           =   2715
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Units:"
                  Height          =   195
                  Index           =   47
                  Left            =   120
                  TabIndex        =   254
                  Top             =   1020
                  Width           =   405
               End
            End
            Begin VB.Frame FraMap 
               Caption         =   "Map"
               Height          =   1425
               Left            =   120
               TabIndex        =   38
               Top             =   1800
               Width           =   3975
               Begin VB.CheckBox chkUseLegendInMap 
                  Caption         =   "Use Legend"
                  Height          =   315
                  Left            =   150
                  TabIndex        =   54
                  Top             =   1020
                  Value           =   1  'Checked
                  Width           =   1665
               End
               Begin VB.CheckBox chkUseStrict 
                  Caption         =   "Use Strict Open"
                  Height          =   315
                  Left            =   150
                  TabIndex        =   48
                  Top             =   270
                  Width           =   1425
               End
               Begin VB.CommandButton cmdMapPath 
                  Caption         =   "..."
                  Height          =   285
                  Left            =   3390
                  TabIndex        =   47
                  Top             =   630
                  Width           =   465
               End
               Begin VB.TextBox txtMapPath 
                  Height          =   315
                  Left            =   120
                  TabIndex        =   46
                  Top             =   630
                  Width           =   3225
               End
            End
            Begin VB.Frame FraLocal 
               Caption         =   "Local"
               Height          =   885
               Left            =   120
               TabIndex        =   37
               Top             =   870
               Width           =   3945
               Begin VB.CommandButton cmdlocalDB 
                  Caption         =   "..."
                  Height          =   315
                  Left            =   3450
                  TabIndex        =   45
                  Top             =   450
                  Width           =   435
               End
               Begin VB.TextBox txtLocalDB 
                  Height          =   345
                  Left            =   60
                  TabIndex        =   44
                  Top             =   450
                  Width           =   3345
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "OASIS Client DB:"
                  Height          =   195
                  Index           =   48
                  Left            =   60
                  TabIndex        =   258
                  Top             =   240
                  Width           =   1230
               End
            End
            Begin VB.Frame FraServer 
               Caption         =   "Server:"
               Height          =   855
               Left            =   150
               TabIndex        =   36
               Top             =   930
               Visible         =   0   'False
               Width           =   3945
               Begin VB.TextBox txtServerURL 
                  Height          =   345
                  Left            =   90
                  TabIndex        =   43
                  Top             =   330
                  Width           =   3315
               End
               Begin VB.CommandButton cmdServerConnect 
                  Caption         =   "..."
                  Height          =   315
                  Left            =   3420
                  TabIndex        =   42
                  Top             =   360
                  Width           =   465
               End
            End
            Begin VB.Frame FraWorkMode 
               Caption         =   "Work Mode:"
               Enabled         =   0   'False
               Height          =   705
               Left            =   120
               TabIndex        =   33
               Top             =   120
               Width           =   3945
               Begin VB.OptionButton OptMode 
                  Caption         =   "Disconnected"
                  Height          =   345
                  Index           =   2
                  Left            =   1830
                  TabIndex        =   179
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1305
               End
               Begin VB.OptionButton OptMode 
                  Caption         =   "Local"
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   0
                  Left            =   90
                  TabIndex        =   35
                  Top             =   210
                  Width           =   885
               End
               Begin VB.OptionButton OptMode 
                  Caption         =   "Server"
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   1
                  Left            =   990
                  TabIndex        =   34
                  Top             =   210
                  Width           =   945
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elObjects 
            Height          =   7410
            Left            =   5190
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   330
            Width           =   4155
            _cx             =   7329
            _cy             =   13070
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
            Begin VB.Frame FraGraphicsObjects 
               Caption         =   "Graphics Objects:"
               Height          =   3225
               Left            =   60
               TabIndex        =   107
               Top             =   4140
               Width           =   4065
               Begin VB.CheckBox chkNoGraphics 
                  Caption         =   "No Graphics"
                  Height          =   285
                  Left            =   1830
                  TabIndex        =   257
                  Top             =   2850
                  Value           =   1  'Checked
                  Width           =   1545
               End
               Begin VB.CheckBox chkUseRoot 
                  Caption         =   "Use Root Path"
                  Height          =   225
                  Left            =   120
                  TabIndex        =   156
                  Top             =   2880
                  Width           =   1395
               End
               Begin VB.CommandButton cmdGraphicPath 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   3
                  Left            =   3600
                  TabIndex        =   154
                  Top             =   2520
                  Width           =   405
               End
               Begin VB.TextBox txtgraphPath 
                  Height          =   285
                  Index           =   3
                  Left            =   540
                  TabIndex        =   153
                  Top             =   2520
                  Width           =   3015
               End
               Begin VB.CommandButton cmdGraphicPath 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   2
                  Left            =   3600
                  TabIndex        =   151
                  Top             =   1890
                  Width           =   405
               End
               Begin VB.TextBox txtgraphPath 
                  Height          =   285
                  Index           =   2
                  Left            =   540
                  TabIndex        =   150
                  Top             =   1890
                  Width           =   3015
               End
               Begin VB.CommandButton cmdGraphicPath 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   1
                  Left            =   3600
                  TabIndex        =   148
                  Top             =   1290
                  Width           =   405
               End
               Begin VB.TextBox txtgraphPath 
                  Height          =   285
                  Index           =   1
                  Left            =   540
                  TabIndex        =   147
                  Top             =   1290
                  Width           =   3015
               End
               Begin VB.CommandButton cmdGraphicPath 
                  Caption         =   "..."
                  Height          =   285
                  Index           =   0
                  Left            =   3600
                  TabIndex        =   145
                  Top             =   690
                  Width           =   405
               End
               Begin VB.TextBox txtgraphPath 
                  Height          =   285
                  Index           =   0
                  Left            =   540
                  TabIndex        =   144
                  Top             =   690
                  Width           =   3015
               End
               Begin VB.TextBox txtGraphicsBottom 
                  Height          =   285
                  Index           =   3
                  Left            =   3090
                  TabIndex        =   143
                  Text            =   "2.00"
                  Top             =   2190
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsRight 
                  Height          =   285
                  Index           =   3
                  Left            =   2610
                  TabIndex        =   142
                  Text            =   "8.00"
                  Top             =   2190
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsTop 
                  Height          =   285
                  Index           =   3
                  Left            =   2130
                  TabIndex        =   141
                  Text            =   "0.5"
                  Top             =   2190
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsLeft 
                  Height          =   285
                  Index           =   3
                  Left            =   1650
                  TabIndex        =   140
                  Text            =   "1.00"
                  Top             =   2190
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsBottom 
                  Height          =   285
                  Index           =   2
                  Left            =   3090
                  TabIndex        =   139
                  Text            =   "2.00"
                  Top             =   1590
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsRight 
                  Height          =   285
                  Index           =   2
                  Left            =   2610
                  TabIndex        =   138
                  Text            =   "8.00"
                  Top             =   1590
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsTop 
                  Height          =   285
                  Index           =   2
                  Left            =   2130
                  TabIndex        =   137
                  Text            =   "0.5"
                  Top             =   1590
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsLeft 
                  Height          =   285
                  Index           =   2
                  Left            =   1650
                  TabIndex        =   136
                  Text            =   "1.00"
                  Top             =   1590
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsBottom 
                  Height          =   285
                  Index           =   1
                  Left            =   3090
                  TabIndex        =   135
                  Text            =   "2.00"
                  Top             =   990
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsRight 
                  Height          =   285
                  Index           =   1
                  Left            =   2610
                  TabIndex        =   134
                  Text            =   "8.00"
                  Top             =   990
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsTop 
                  Height          =   285
                  Index           =   1
                  Left            =   2130
                  TabIndex        =   133
                  Text            =   "0.5"
                  Top             =   990
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsLeft 
                  Height          =   285
                  Index           =   1
                  Left            =   1650
                  TabIndex        =   132
                  Text            =   "1.00"
                  Top             =   990
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsBottom 
                  Height          =   285
                  Index           =   0
                  Left            =   3090
                  TabIndex        =   131
                  Text            =   "2.00"
                  Top             =   390
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsRight 
                  Height          =   285
                  Index           =   0
                  Left            =   2610
                  TabIndex        =   130
                  Text            =   "8.00"
                  Top             =   390
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsTop 
                  Height          =   285
                  Index           =   0
                  Left            =   2130
                  TabIndex        =   129
                  Text            =   "0.5"
                  Top             =   390
                  Width           =   465
               End
               Begin VB.TextBox txtGraphicsLeft 
                  Height          =   285
                  Index           =   0
                  Left            =   1650
                  TabIndex        =   128
                  Text            =   "1.00"
                  Top             =   390
                  Width           =   465
               End
               Begin VB.OptionButton optGraphics 
                  Caption         =   "4 Graphics"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   123
                  Top             =   2220
                  Width           =   1095
               End
               Begin VB.OptionButton optGraphics 
                  Caption         =   "3 Graphics"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   122
                  Top             =   1590
                  Width           =   1065
               End
               Begin VB.OptionButton optGraphics 
                  Caption         =   "2 Graphics"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   121
                  Top             =   990
                  Width           =   1065
               End
               Begin VB.OptionButton optGraphics 
                  Caption         =   "1 Graphic"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   120
                  Top             =   390
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Path:"
                  Height          =   195
                  Index           =   22
                  Left            =   120
                  TabIndex        =   155
                  Top             =   2550
                  Width           =   375
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Path:"
                  Height          =   195
                  Index           =   21
                  Left            =   120
                  TabIndex        =   152
                  Top             =   1920
                  Width           =   375
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Path:"
                  Height          =   195
                  Index           =   20
                  Left            =   120
                  TabIndex        =   149
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Path:"
                  Height          =   195
                  Index           =   19
                  Left            =   120
                  TabIndex        =   146
                  Top             =   720
                  Width           =   375
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Bottom"
                  Height          =   195
                  Index           =   18
                  Left            =   3090
                  TabIndex        =   127
                  Top             =   180
                  Width           =   495
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Right"
                  Height          =   195
                  Index           =   17
                  Left            =   2610
                  TabIndex        =   126
                  Top             =   180
                  Width           =   375
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Top"
                  Height          =   195
                  Index           =   16
                  Left            =   2130
                  TabIndex        =   125
                  Top             =   180
                  Width           =   285
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Left"
                  Height          =   195
                  Index           =   15
                  Left            =   1650
                  TabIndex        =   124
                  Top             =   180
                  Width           =   270
               End
            End
            Begin VB.Frame FraScaleBar 
               Caption         =   "Scale Bar Objects:"
               Height          =   705
               Left            =   60
               TabIndex        =   105
               Top             =   3450
               Width           =   4095
               Begin VB.TextBox txtbottomScale 
                  Height          =   285
                  Left            =   3090
                  TabIndex        =   116
                  Text            =   "-1.00"
                  Top             =   330
                  Width           =   465
               End
               Begin VB.TextBox txtRightScale 
                  Height          =   285
                  Left            =   2610
                  TabIndex        =   115
                  Text            =   "-7.00"
                  Top             =   330
                  Width           =   465
               End
               Begin VB.TextBox txtTopScale 
                  Height          =   285
                  Left            =   2130
                  TabIndex        =   114
                  Text            =   "-2"
                  Top             =   330
                  Width           =   465
               End
               Begin VB.TextBox txtLeftScale 
                  Height          =   285
                  Left            =   1650
                  TabIndex        =   113
                  Text            =   "-10.00"
                  Top             =   330
                  Width           =   465
               End
               Begin VB.CheckBox chkUseScalebar 
                  Caption         =   "Use Scalebar"
                  Height          =   225
                  Left            =   90
                  TabIndex        =   106
                  Top             =   360
                  Value           =   1  'Checked
                  Width           =   1335
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Left"
                  Height          =   195
                  Index           =   14
                  Left            =   1710
                  TabIndex        =   112
                  Top             =   120
                  Width           =   270
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Top"
                  Height          =   195
                  Index           =   13
                  Left            =   2190
                  TabIndex        =   111
                  Top             =   120
                  Width           =   285
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Right"
                  Height          =   195
                  Index           =   12
                  Left            =   2670
                  TabIndex        =   110
                  Top             =   120
                  Width           =   375
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Bottom"
                  Height          =   195
                  Index           =   11
                  Left            =   3090
                  TabIndex        =   109
                  Top             =   120
                  Width           =   495
               End
            End
            Begin VB.Frame FraLegendObjects 
               Caption         =   "Legend Objects:"
               Height          =   1815
               Left            =   30
               TabIndex        =   74
               Top             =   1560
               Width           =   4125
               Begin XpressEditorsLibCtl.dxColorEdit dxLgdColor 
                  Height          =   315
                  Index           =   0
                  Left            =   3540
                  OleObjectBlob   =   "frmMapPrintAdmin.frx":31B70
                  TabIndex        =   99
                  Top             =   480
                  Width           =   555
               End
               Begin VB.TextBox txtLegendBottom 
                  Height          =   285
                  Index           =   3
                  Left            =   3060
                  TabIndex        =   95
                  Text            =   "-1.00"
                  Top             =   1440
                  Width           =   465
               End
               Begin VB.TextBox txtLegendRight 
                  Height          =   285
                  Index           =   3
                  Left            =   2565
                  TabIndex        =   94
                  Text            =   "-1.00"
                  Top             =   1440
                  Width           =   465
               End
               Begin VB.TextBox txtLegendTop 
                  Height          =   285
                  Index           =   3
                  Left            =   2085
                  TabIndex        =   93
                  Text            =   "2.30"
                  Top             =   1440
                  Width           =   465
               End
               Begin VB.TextBox txtLegendLeft 
                  Height          =   285
                  Index           =   3
                  Left            =   1620
                  TabIndex        =   92
                  Text            =   "-6.15"
                  Top             =   1440
                  Width           =   465
               End
               Begin VB.TextBox txtLegendBottom 
                  Height          =   285
                  Index           =   2
                  Left            =   3060
                  TabIndex        =   91
                  Text            =   "-1.00"
                  Top             =   1140
                  Width           =   465
               End
               Begin VB.TextBox txtLegendRight 
                  Height          =   285
                  Index           =   2
                  Left            =   2565
                  TabIndex        =   90
                  Text            =   "-1.00"
                  Top             =   1140
                  Width           =   465
               End
               Begin VB.TextBox txtLegendTop 
                  Height          =   285
                  Index           =   2
                  Left            =   2085
                  TabIndex        =   89
                  Text            =   "2.25"
                  Top             =   1140
                  Width           =   465
               End
               Begin VB.TextBox txtLegendLeft 
                  Height          =   285
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   88
                  Text            =   "-6.25"
                  Top             =   1140
                  Width           =   465
               End
               Begin VB.TextBox txtLegendBottom 
                  Height          =   285
                  Index           =   1
                  Left            =   3060
                  TabIndex        =   87
                  Text            =   "-0.85"
                  Top             =   810
                  Width           =   465
               End
               Begin VB.TextBox txtLegendRight 
                  Height          =   285
                  Index           =   1
                  Left            =   2565
                  TabIndex        =   86
                  Text            =   "-0.85"
                  Top             =   810
                  Width           =   465
               End
               Begin VB.TextBox txtLegendTop 
                  Height          =   285
                  Index           =   1
                  Left            =   2085
                  TabIndex        =   85
                  Text            =   "2.10"
                  Top             =   810
                  Width           =   465
               End
               Begin VB.TextBox txtLegendLeft 
                  Height          =   285
                  Index           =   1
                  Left            =   1620
                  TabIndex        =   84
                  Text            =   "-6.40"
                  Top             =   810
                  Width           =   465
               End
               Begin VB.TextBox txtLegendBottom 
                  Height          =   285
                  Index           =   0
                  Left            =   3060
                  TabIndex        =   83
                  Text            =   "-0.75"
                  Top             =   480
                  Width           =   465
               End
               Begin VB.TextBox txtLegendRight 
                  Height          =   285
                  Index           =   0
                  Left            =   2565
                  TabIndex        =   82
                  Text            =   "-0.75"
                  Top             =   480
                  Width           =   465
               End
               Begin VB.TextBox txtLegendTop 
                  Height          =   285
                  Index           =   0
                  Left            =   2085
                  TabIndex        =   81
                  Text            =   "2.00"
                  Top             =   480
                  Width           =   465
               End
               Begin VB.TextBox txtLegendLeft 
                  Height          =   285
                  Index           =   0
                  Left            =   1620
                  TabIndex        =   80
                  Text            =   "-6.50"
                  Top             =   480
                  Width           =   465
               End
               Begin VB.CheckBox chkUseLegend 
                  Caption         =   "Use Legend"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   75
                  Top             =   1410
                  Value           =   1  'Checked
                  Width           =   1245
               End
               Begin XpressEditorsLibCtl.dxColorEdit dxLgdColor 
                  Height          =   315
                  Index           =   1
                  Left            =   3540
                  OleObjectBlob   =   "frmMapPrintAdmin.frx":31C63
                  TabIndex        =   100
                  Top             =   810
                  Width           =   555
               End
               Begin XpressEditorsLibCtl.dxColorEdit dxLgdColor 
                  Height          =   315
                  Index           =   2
                  Left            =   3540
                  OleObjectBlob   =   "frmMapPrintAdmin.frx":31D56
                  TabIndex        =   101
                  Top             =   1140
                  Width           =   555
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Background"
                  Height          =   195
                  Index           =   10
                  Left            =   180
                  TabIndex        =   104
                  Top             =   1170
                  Width           =   870
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Inner Frame"
                  Height          =   195
                  Index           =   9
                  Left            =   180
                  TabIndex        =   103
                  Top             =   840
                  Width           =   840
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Outer Frame"
                  Height          =   195
                  Index           =   8
                  Left            =   180
                  TabIndex        =   102
                  Top             =   480
                  Width           =   870
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Bottom"
                  Height          =   195
                  Index           =   7
                  Left            =   3060
                  TabIndex        =   79
                  Top             =   240
                  Width           =   495
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Right"
                  Height          =   195
                  Index           =   6
                  Left            =   2580
                  TabIndex        =   78
                  Top             =   240
                  Width           =   375
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Top"
                  Height          =   195
                  Index           =   5
                  Left            =   2130
                  TabIndex        =   77
                  Top             =   240
                  Width           =   285
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Left"
                  Height          =   195
                  Index           =   4
                  Left            =   1650
                  TabIndex        =   76
                  Top             =   240
                  Width           =   270
               End
            End
            Begin VB.Frame FraMapObjects 
               Caption         =   "Map Objects:"
               Height          =   1425
               Left            =   30
               TabIndex        =   29
               Top             =   90
               Width           =   4095
               Begin VB.TextBox txtMapBottom 
                  Height          =   285
                  Index           =   2
                  Left            =   3060
                  TabIndex        =   73
                  Text            =   "-1.00"
                  Top             =   1050
                  Width           =   465
               End
               Begin VB.TextBox txtMapRight 
                  Height          =   285
                  Index           =   2
                  Left            =   2550
                  TabIndex        =   72
                  Text            =   "-7.00"
                  Top             =   1050
                  Width           =   465
               End
               Begin VB.TextBox txtMapTop 
                  Height          =   285
                  Index           =   2
                  Left            =   2070
                  TabIndex        =   71
                  Text            =   "2.25"
                  Top             =   1050
                  Width           =   465
               End
               Begin VB.TextBox txtMapBottom 
                  Height          =   285
                  Index           =   1
                  Left            =   3060
                  TabIndex        =   70
                  Text            =   "-0.85"
                  Top             =   750
                  Width           =   465
               End
               Begin VB.TextBox txtMapRight 
                  Height          =   285
                  Index           =   1
                  Left            =   2550
                  TabIndex        =   69
                  Text            =   "-6.75"
                  Top             =   750
                  Width           =   465
               End
               Begin VB.TextBox txtMapTop 
                  Height          =   285
                  Index           =   1
                  Left            =   2070
                  TabIndex        =   68
                  Text            =   "2.10"
                  Top             =   750
                  Width           =   465
               End
               Begin VB.TextBox txtMapBottom 
                  Height          =   285
                  Index           =   0
                  Left            =   3060
                  TabIndex        =   67
                  Text            =   "-0.75"
                  Top             =   450
                  Width           =   465
               End
               Begin VB.TextBox txtMapRight 
                  Height          =   285
                  Index           =   0
                  Left            =   2550
                  TabIndex        =   66
                  Text            =   "-6.75"
                  Top             =   450
                  Width           =   465
               End
               Begin VB.TextBox txtMapTop 
                  Height          =   285
                  Index           =   0
                  Left            =   2055
                  TabIndex        =   65
                  Text            =   "2.00"
                  Top             =   450
                  Width           =   465
               End
               Begin VB.TextBox txtMapLeft 
                  Height          =   285
                  Index           =   2
                  Left            =   1590
                  TabIndex        =   64
                  Text            =   "1.00"
                  Top             =   1050
                  Width           =   465
               End
               Begin VB.TextBox txtMapLeft 
                  Height          =   285
                  Index           =   1
                  Left            =   1590
                  TabIndex        =   63
                  Text            =   "0.85"
                  Top             =   750
                  Width           =   465
               End
               Begin VB.TextBox txtMapLeft 
                  Height          =   285
                  Index           =   0
                  Left            =   1590
                  TabIndex        =   62
                  Text            =   "0.75"
                  Top             =   450
                  Width           =   465
               End
               Begin VB.OptionButton OptMapFrame 
                  Caption         =   "2 Map Frame"
                  Height          =   225
                  Index           =   2
                  Left            =   90
                  TabIndex        =   57
                  Top             =   810
                  Value           =   -1  'True
                  Width           =   1245
               End
               Begin VB.OptionButton OptMapFrame 
                  Caption         =   "1 Map Frame"
                  Height          =   225
                  Index           =   1
                  Left            =   90
                  TabIndex        =   56
                  Top             =   510
                  Width           =   1245
               End
               Begin VB.OptionButton OptMapFrame 
                  Caption         =   "No Frames"
                  Height          =   225
                  Index           =   0
                  Left            =   90
                  TabIndex        =   55
                  Top             =   210
                  Width           =   1245
               End
               Begin VB.CheckBox chkUseMap 
                  Caption         =   "Use Map"
                  Height          =   285
                  Left            =   120
                  TabIndex        =   30
                  Top             =   1080
                  Value           =   1  'Checked
                  Width           =   1035
               End
               Begin XpressEditorsLibCtl.dxColorEdit dxmapColor 
                  Height          =   315
                  Index           =   0
                  Left            =   3510
                  OleObjectBlob   =   "frmMapPrintAdmin.frx":31E49
                  TabIndex        =   96
                  Top             =   450
                  Width           =   555
               End
               Begin XpressEditorsLibCtl.dxColorEdit dxmapColor 
                  Height          =   315
                  Index           =   1
                  Left            =   3510
                  OleObjectBlob   =   "frmMapPrintAdmin.frx":31F3C
                  TabIndex        =   97
                  Top             =   750
                  Width           =   555
               End
               Begin XpressEditorsLibCtl.dxColorEdit dxmapColor 
                  Height          =   315
                  Index           =   2
                  Left            =   3510
                  OleObjectBlob   =   "frmMapPrintAdmin.frx":3202F
                  TabIndex        =   98
                  Top             =   1050
                  Width           =   555
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Bottom"
                  Height          =   195
                  Index           =   3
                  Left            =   3060
                  TabIndex        =   61
                  Top             =   210
                  Width           =   495
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Right"
                  Height          =   195
                  Index           =   2
                  Left            =   2580
                  TabIndex        =   60
                  Top             =   210
                  Width           =   375
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Top"
                  Height          =   195
                  Index           =   1
                  Left            =   2100
                  TabIndex        =   59
                  Top             =   210
                  Width           =   285
               End
               Begin VB.Label lblMap 
                  AutoSize        =   -1  'True
                  Caption         =   "Left"
                  Height          =   195
                  Index           =   0
                  Left            =   1620
                  TabIndex        =   58
                  Top             =   210
                  Width           =   270
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elRaw 
            Height          =   7410
            Left            =   5790
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   330
            Width           =   4155
            _cx             =   7329
            _cy             =   13070
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
            Begin VB.TextBox txtRawEdt 
               Height          =   7155
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   22
               Top             =   90
               Width           =   4005
            End
         End
         Begin VB.Frame FraPrntSettngs 
            BorderStyle     =   0  'None
            Caption         =   "v"
            Height          =   7410
            Left            =   5490
            TabIndex        =   13
            Top             =   330
            Width           =   4155
            Begin VB.TextBox txtMapProdctionDate 
               Height          =   315
               Left            =   780
               TabIndex        =   255
               Text            =   "Production Date:"
               Top             =   6120
               Width           =   1545
            End
            Begin VB.ComboBox ComJustify 
               Height          =   315
               Index           =   5
               ItemData        =   "frmMapPrintAdmin.frx":32122
               Left            =   2580
               List            =   "frmMapPrintAdmin.frx":32132
               Style           =   2  'Dropdown List
               TabIndex        =   252
               Top             =   6810
               Width           =   1545
            End
            Begin VB.CheckBox chkStrikeThrough 
               Caption         =   "Strike Through"
               Height          =   225
               Index           =   5
               Left            =   1140
               TabIndex        =   251
               Top             =   6870
               Width           =   1365
            End
            Begin VB.CheckBox chkUnderline 
               Caption         =   "Underline"
               Height          =   195
               Index           =   5
               Left            =   30
               TabIndex        =   250
               Top             =   6870
               Width           =   1005
            End
            Begin VB.TextBox txtTextBottom 
               Height          =   285
               Index           =   5
               Left            =   3060
               TabIndex        =   248
               Text            =   "-0.02"
               Top             =   6510
               Width           =   465
            End
            Begin VB.TextBox txtTextRight 
               Height          =   285
               Index           =   5
               Left            =   2040
               TabIndex        =   247
               Text            =   "-0.75"
               Top             =   6510
               Width           =   465
            End
            Begin VB.TextBox txtTextLeft 
               Height          =   285
               Index           =   5
               Left            =   330
               TabIndex        =   246
               Text            =   "1.00"
               Top             =   6510
               Width           =   465
            End
            Begin VB.TextBox txtTextTop 
               Height          =   285
               Index           =   5
               Left            =   1140
               TabIndex        =   245
               Text            =   "-0.50"
               Top             =   6510
               Width           =   465
            End
            Begin VB.ComboBox ComJustify 
               Height          =   315
               Index           =   4
               ItemData        =   "frmMapPrintAdmin.frx":3215F
               Left            =   2580
               List            =   "frmMapPrintAdmin.frx":3216F
               Style           =   2  'Dropdown List
               TabIndex        =   239
               Top             =   5730
               Width           =   1545
            End
            Begin VB.CheckBox chkStrikeThrough 
               Caption         =   "Strike Through"
               Height          =   225
               Index           =   4
               Left            =   1140
               TabIndex        =   238
               Top             =   5790
               Width           =   1365
            End
            Begin VB.CheckBox chkUnderline 
               Caption         =   "Underline"
               Height          =   195
               Index           =   4
               Left            =   30
               TabIndex        =   237
               Top             =   5790
               Width           =   1005
            End
            Begin VB.TextBox txtTextBottom 
               Height          =   285
               Index           =   4
               Left            =   3060
               TabIndex        =   235
               Text            =   "-0.02"
               Top             =   5430
               Width           =   465
            End
            Begin VB.TextBox txtTextRight 
               Height          =   285
               Index           =   4
               Left            =   2040
               TabIndex        =   234
               Text            =   "-0.75"
               Top             =   5430
               Width           =   465
            End
            Begin VB.TextBox txtTextLeft 
               Height          =   285
               Index           =   4
               Left            =   330
               TabIndex        =   233
               Text            =   "1.00"
               Top             =   5430
               Width           =   465
            End
            Begin VB.TextBox txtTextTop 
               Height          =   285
               Index           =   4
               Left            =   1140
               TabIndex        =   232
               Text            =   "-0.50"
               Top             =   5430
               Width           =   465
            End
            Begin VB.ComboBox ComJustify 
               Height          =   315
               Index           =   3
               ItemData        =   "frmMapPrintAdmin.frx":3219C
               Left            =   2580
               List            =   "frmMapPrintAdmin.frx":321AC
               Style           =   2  'Dropdown List
               TabIndex        =   227
               Top             =   4590
               Width           =   1545
            End
            Begin VB.CheckBox chkStrikeThrough 
               Caption         =   "Strike Through"
               Height          =   225
               Index           =   3
               Left            =   1140
               TabIndex        =   226
               Top             =   4650
               Width           =   1365
            End
            Begin VB.CheckBox chkUnderline 
               Caption         =   "Underline"
               Height          =   195
               Index           =   3
               Left            =   30
               TabIndex        =   225
               Top             =   4650
               Width           =   1005
            End
            Begin VB.TextBox txtTextBottom 
               Height          =   285
               Index           =   3
               Left            =   3060
               TabIndex        =   223
               Text            =   "-0.02"
               Top             =   4290
               Width           =   465
            End
            Begin VB.TextBox txtTextRight 
               Height          =   285
               Index           =   3
               Left            =   2040
               TabIndex        =   222
               Text            =   "-0.75"
               Top             =   4290
               Width           =   465
            End
            Begin VB.TextBox txtTextLeft 
               Height          =   285
               Index           =   3
               Left            =   330
               TabIndex        =   221
               Text            =   "1.00"
               Top             =   4290
               Width           =   465
            End
            Begin VB.TextBox txtTextTop 
               Height          =   285
               Index           =   3
               Left            =   1140
               TabIndex        =   220
               Text            =   "-0.50"
               Top             =   4290
               Width           =   465
            End
            Begin VB.ComboBox ComJustify 
               Height          =   315
               Index           =   2
               ItemData        =   "frmMapPrintAdmin.frx":321D9
               Left            =   2580
               List            =   "frmMapPrintAdmin.frx":321E9
               Style           =   2  'Dropdown List
               TabIndex        =   215
               Top             =   3270
               Width           =   1545
            End
            Begin VB.CheckBox chkStrikeThrough 
               Caption         =   "Strike Through"
               Height          =   225
               Index           =   2
               Left            =   1140
               TabIndex        =   214
               Top             =   3330
               Width           =   1365
            End
            Begin VB.CheckBox chkUnderline 
               Caption         =   "Underline"
               Height          =   195
               Index           =   2
               Left            =   30
               TabIndex        =   213
               Top             =   3330
               Width           =   1005
            End
            Begin VB.TextBox txtTextBottom 
               Height          =   285
               Index           =   2
               Left            =   3060
               TabIndex        =   211
               Text            =   "-0.02"
               Top             =   2970
               Width           =   465
            End
            Begin VB.TextBox txtTextRight 
               Height          =   285
               Index           =   2
               Left            =   2040
               TabIndex        =   210
               Text            =   "-0.75"
               Top             =   2970
               Width           =   465
            End
            Begin VB.TextBox txtTextLeft 
               Height          =   285
               Index           =   2
               Left            =   330
               TabIndex        =   209
               Text            =   "0.00"
               Top             =   2970
               Width           =   465
            End
            Begin VB.TextBox txtTextTop 
               Height          =   285
               Index           =   2
               Left            =   1140
               TabIndex        =   208
               Text            =   "-0.50"
               Top             =   2970
               Width           =   465
            End
            Begin VB.ComboBox ComJustify 
               Height          =   315
               Index           =   1
               ItemData        =   "frmMapPrintAdmin.frx":32216
               Left            =   2580
               List            =   "frmMapPrintAdmin.frx":32226
               Style           =   2  'Dropdown List
               TabIndex        =   203
               Top             =   1920
               Width           =   1545
            End
            Begin VB.CheckBox chkStrikeThrough 
               Caption         =   "Strike Through"
               Height          =   225
               Index           =   1
               Left            =   1140
               TabIndex        =   202
               Top             =   1980
               Width           =   1365
            End
            Begin VB.CheckBox chkUnderline 
               Caption         =   "Underline"
               Height          =   195
               Index           =   1
               Left            =   30
               TabIndex        =   201
               Top             =   1980
               Width           =   1005
            End
            Begin VB.TextBox txtTextBottom 
               Height          =   285
               Index           =   1
               Left            =   3060
               TabIndex        =   199
               Text            =   "-0.02"
               Top             =   1620
               Width           =   465
            End
            Begin VB.TextBox txtTextRight 
               Height          =   285
               Index           =   1
               Left            =   2040
               TabIndex        =   198
               Text            =   "-0.75"
               Top             =   1620
               Width           =   465
            End
            Begin VB.TextBox txtTextLeft 
               Height          =   285
               Index           =   1
               Left            =   330
               TabIndex        =   197
               Text            =   "0.75"
               Top             =   1620
               Width           =   465
            End
            Begin VB.TextBox txtTextTop 
               Height          =   285
               Index           =   1
               Left            =   1140
               TabIndex        =   196
               Text            =   "-0.50"
               Top             =   1620
               Width           =   465
            End
            Begin VB.ComboBox ComJustify 
               Height          =   315
               Index           =   0
               ItemData        =   "frmMapPrintAdmin.frx":32253
               Left            =   2610
               List            =   "frmMapPrintAdmin.frx":32263
               Style           =   2  'Dropdown List
               TabIndex        =   191
               Top             =   720
               Width           =   1545
            End
            Begin VB.CheckBox chkStrikeThrough 
               Caption         =   "Strike Through"
               Height          =   225
               Index           =   0
               Left            =   1170
               TabIndex        =   190
               Top             =   780
               Width           =   1365
            End
            Begin VB.CheckBox chkUnderline 
               Caption         =   "Underline"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   189
               Top             =   780
               Width           =   1005
            End
            Begin XpressEditorsLibCtl.dxColorEdit dxColorTXT 
               Height          =   315
               Index           =   0
               Left            =   3660
               OleObjectBlob   =   "frmMapPrintAdmin.frx":32290
               TabIndex        =   188
               Top             =   390
               Width           =   525
            End
            Begin VB.TextBox txtTextBottom 
               Height          =   285
               Index           =   0
               Left            =   3090
               TabIndex        =   187
               Text            =   "3.00"
               Top             =   420
               Width           =   465
            End
            Begin VB.TextBox txtTextRight 
               Height          =   285
               Index           =   0
               Left            =   2070
               TabIndex        =   186
               Text            =   "-0.75"
               Top             =   420
               Width           =   465
            End
            Begin VB.TextBox txtTextLeft 
               Height          =   285
               Index           =   0
               Left            =   360
               TabIndex        =   185
               Text            =   "0.00"
               Top             =   420
               Width           =   465
            End
            Begin VB.TextBox txtTextTop 
               Height          =   285
               Index           =   0
               Left            =   1170
               TabIndex        =   184
               Text            =   "0.75"
               Top             =   420
               Width           =   465
            End
            Begin VB.CheckBox chkMapTexts 
               Caption         =   "Date:"
               Height          =   225
               Index           =   5
               Left            =   0
               TabIndex        =   168
               Top             =   6180
               Width           =   705
            End
            Begin VB.CheckBox chkMapTexts 
               Caption         =   "Map ID:"
               Height          =   315
               Index           =   4
               Left            =   0
               TabIndex        =   167
               Top             =   5040
               Width           =   825
            End
            Begin VB.CheckBox chkMapTexts 
               Caption         =   "Notes:"
               Height          =   225
               Index           =   3
               Left            =   0
               TabIndex        =   166
               Top             =   3690
               Value           =   1  'Checked
               Width           =   825
            End
            Begin VB.CheckBox chkMapTexts 
               Caption         =   "Copy right:"
               Height          =   465
               Index           =   2
               Left            =   0
               TabIndex        =   165
               Top             =   2370
               Value           =   1  'Checked
               Width           =   795
            End
            Begin VB.CheckBox chkMapTexts 
               Caption         =   "Sub Title:"
               Height          =   375
               Index           =   1
               Left            =   0
               TabIndex        =   164
               Top             =   1230
               Value           =   1  'Checked
               Width           =   825
            End
            Begin VB.CheckBox chkMapTexts 
               Caption         =   "Title:"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   163
               Top             =   90
               Value           =   1  'Checked
               Width           =   705
            End
            Begin VB.CommandButton cmdFonts 
               Caption         =   "..."
               Height          =   255
               Index           =   5
               Left            =   3750
               TabIndex        =   162
               Top             =   6150
               Width           =   375
            End
            Begin VB.CommandButton cmdFonts 
               Caption         =   "..."
               Height          =   255
               Index           =   4
               Left            =   2250
               TabIndex        =   161
               Top             =   5040
               Width           =   375
            End
            Begin VB.CommandButton cmdFonts 
               Caption         =   "..."
               Height          =   255
               Index           =   3
               Left            =   3750
               TabIndex        =   160
               Top             =   3720
               Width           =   375
            End
            Begin VB.CommandButton cmdFonts 
               Caption         =   "..."
               Height          =   255
               Index           =   2
               Left            =   3750
               TabIndex        =   159
               Top             =   2370
               Width           =   375
            End
            Begin VB.CommandButton cmdFonts 
               Caption         =   "..."
               Height          =   255
               Index           =   1
               Left            =   3750
               TabIndex        =   158
               Top             =   1290
               Width           =   375
            End
            Begin VB.CommandButton cmdFonts 
               Caption         =   "..."
               Height          =   255
               Index           =   0
               Left            =   3660
               TabIndex        =   157
               Top             =   90
               Width           =   465
            End
            Begin XpressEditorsLibCtl.dxDateEdit dxDateEdit1 
               Height          =   315
               Left            =   2400
               OleObjectBlob   =   "frmMapPrintAdmin.frx":32383
               TabIndex        =   19
               Top             =   6120
               Width           =   1335
            End
            Begin VB.TextBox txtTitle 
               Height          =   330
               Left            =   840
               TabIndex        =   18
               Text            =   "OASIS MAP"
               Top             =   45
               Width           =   2730
            End
            Begin VB.TextBox txtCopyRight 
               Height          =   585
               Left            =   810
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   17
               Text            =   "frmMapPrintAdmin.frx":32423
               Top             =   2340
               Width           =   2850
            End
            Begin VB.TextBox txtSubTitle 
               Height          =   330
               Left            =   840
               TabIndex        =   16
               Text            =   "Map title"
               Top             =   1260
               Width           =   2730
            End
            Begin VB.TextBox txtVeiwer 
               Height          =   540
               Left            =   840
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               Text            =   "frmMapPrintAdmin.frx":32441
               Top             =   3690
               Width           =   2850
            End
            Begin VB.TextBox txtMapID 
               Height          =   300
               Left            =   840
               TabIndex        =   14
               Text            =   "Map ID: 1234"
               Top             =   5040
               Width           =   1320
            End
            Begin XpressEditorsLibCtl.dxColorEdit dxColorTXT 
               Height          =   315
               Index           =   1
               Left            =   3630
               OleObjectBlob   =   "frmMapPrintAdmin.frx":32508
               TabIndex        =   200
               Top             =   1590
               Width           =   525
            End
            Begin XpressEditorsLibCtl.dxColorEdit dxColorTXT 
               Height          =   315
               Index           =   2
               Left            =   3630
               OleObjectBlob   =   "frmMapPrintAdmin.frx":325FB
               TabIndex        =   212
               Top             =   2940
               Width           =   525
            End
            Begin XpressEditorsLibCtl.dxColorEdit dxColorTXT 
               Height          =   315
               Index           =   3
               Left            =   3630
               OleObjectBlob   =   "frmMapPrintAdmin.frx":326EE
               TabIndex        =   224
               Top             =   4260
               Width           =   525
            End
            Begin XpressEditorsLibCtl.dxColorEdit dxColorTXT 
               Height          =   315
               Index           =   4
               Left            =   3630
               OleObjectBlob   =   "frmMapPrintAdmin.frx":327E1
               TabIndex        =   236
               Top             =   5400
               Width           =   525
            End
            Begin XpressEditorsLibCtl.dxColorEdit dxColorTXT 
               Height          =   315
               Index           =   5
               Left            =   3630
               OleObjectBlob   =   "frmMapPrintAdmin.frx":328D4
               TabIndex        =   249
               Top             =   6480
               Width           =   525
            End
            Begin VB.Line Line1 
               Index           =   5
               X1              =   0
               X2              =   4050
               Y1              =   7170
               Y2              =   7170
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Left"
               Height          =   195
               Index           =   46
               Left            =   0
               TabIndex        =   244
               Top             =   6570
               Width           =   270
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Top"
               Height          =   195
               Index           =   45
               Left            =   810
               TabIndex        =   243
               Top             =   6570
               Width           =   285
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Right"
               Height          =   195
               Index           =   44
               Left            =   1620
               TabIndex        =   242
               Top             =   6570
               Width           =   375
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Bottom"
               Height          =   195
               Index           =   43
               Left            =   2520
               TabIndex        =   241
               Top             =   6570
               Width           =   495
            End
            Begin VB.Label lblLabel3 
               Height          =   795
               Left            =   60
               TabIndex        =   240
               Top             =   6630
               Width           =   3945
            End
            Begin VB.Line Line1 
               Index           =   4
               X1              =   0
               X2              =   4050
               Y1              =   6090
               Y2              =   6090
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Left"
               Height          =   195
               Index           =   42
               Left            =   0
               TabIndex        =   231
               Top             =   5490
               Width           =   270
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Top"
               Height          =   195
               Index           =   41
               Left            =   810
               TabIndex        =   230
               Top             =   5490
               Width           =   285
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Right"
               Height          =   195
               Index           =   40
               Left            =   1620
               TabIndex        =   229
               Top             =   5490
               Width           =   375
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Bottom"
               Height          =   195
               Index           =   39
               Left            =   2520
               TabIndex        =   228
               Top             =   5490
               Width           =   495
            End
            Begin VB.Line Line1 
               Index           =   3
               X1              =   0
               X2              =   4050
               Y1              =   4950
               Y2              =   4950
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Left"
               Height          =   195
               Index           =   38
               Left            =   0
               TabIndex        =   219
               Top             =   4350
               Width           =   270
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Top"
               Height          =   195
               Index           =   37
               Left            =   810
               TabIndex        =   218
               Top             =   4350
               Width           =   285
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Right"
               Height          =   195
               Index           =   36
               Left            =   1620
               TabIndex        =   217
               Top             =   4350
               Width           =   375
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Bottom"
               Height          =   195
               Index           =   35
               Left            =   2520
               TabIndex        =   216
               Top             =   4350
               Width           =   495
            End
            Begin VB.Line Line1 
               Index           =   2
               X1              =   0
               X2              =   4050
               Y1              =   3630
               Y2              =   3630
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Left"
               Height          =   195
               Index           =   34
               Left            =   0
               TabIndex        =   207
               Top             =   3030
               Width           =   270
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Top"
               Height          =   195
               Index           =   33
               Left            =   810
               TabIndex        =   206
               Top             =   3030
               Width           =   285
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Right"
               Height          =   195
               Index           =   32
               Left            =   1620
               TabIndex        =   205
               Top             =   3030
               Width           =   375
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Bottom"
               Height          =   195
               Index           =   31
               Left            =   2520
               TabIndex        =   204
               Top             =   3030
               Width           =   495
            End
            Begin VB.Line Line1 
               Index           =   1
               X1              =   0
               X2              =   4050
               Y1              =   2280
               Y2              =   2280
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Left"
               Height          =   195
               Index           =   30
               Left            =   0
               TabIndex        =   195
               Top             =   1680
               Width           =   270
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Top"
               Height          =   195
               Index           =   29
               Left            =   810
               TabIndex        =   194
               Top             =   1680
               Width           =   285
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Right"
               Height          =   195
               Index           =   28
               Left            =   1620
               TabIndex        =   193
               Top             =   1680
               Width           =   375
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Bottom"
               Height          =   195
               Index           =   27
               Left            =   2520
               TabIndex        =   192
               Top             =   1680
               Width           =   495
            End
            Begin VB.Line Line1 
               Index           =   0
               X1              =   30
               X2              =   4080
               Y1              =   1080
               Y2              =   1080
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Left"
               Height          =   195
               Index           =   26
               Left            =   30
               TabIndex        =   183
               Top             =   480
               Width           =   270
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Top"
               Height          =   195
               Index           =   25
               Left            =   840
               TabIndex        =   182
               Top             =   480
               Width           =   285
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Right"
               Height          =   195
               Index           =   24
               Left            =   1650
               TabIndex        =   181
               Top             =   480
               Width           =   375
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Bottom"
               Height          =   195
               Index           =   23
               Left            =   2550
               TabIndex        =   180
               Top             =   480
               Width           =   495
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               Height          =   225
               Index           =   9
               Left            =   3120
               TabIndex        =   178
               Top             =   7140
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               Height          =   225
               Index           =   8
               Left            =   2400
               TabIndex        =   177
               Top             =   7140
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               Height          =   225
               Index           =   7
               Left            =   1680
               TabIndex        =   176
               Top             =   7110
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               Height          =   225
               Index           =   6
               Left            =   960
               TabIndex        =   175
               Top             =   7110
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               Height          =   225
               Index           =   5
               Left            =   240
               TabIndex        =   174
               Top             =   7110
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               Height          =   225
               Index           =   4
               Left            =   2940
               TabIndex        =   173
               Top             =   6840
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               Height          =   225
               Index           =   3
               Left            =   2220
               TabIndex        =   172
               Top             =   6840
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               Height          =   225
               Index           =   2
               Left            =   1500
               TabIndex        =   171
               Top             =   6840
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   840
               TabIndex        =   170
               Top             =   6810
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label lblTemplate 
               Caption         =   "Template"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   24
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   180
               TabIndex        =   169
               Top             =   6780
               Visible         =   0   'False
               Width           =   675
            End
         End
         Begin VB.Frame FraAvailablePrint 
            BorderStyle     =   0  'None
            Height          =   7410
            Left            =   4890
            TabIndex        =   9
            Top             =   330
            Width           =   4155
            Begin VB.TextBox txtClientSettings 
               Height          =   315
               Index           =   1
               Left            =   180
               TabIndex        =   262
               Top             =   1230
               Width           =   3855
            End
            Begin VB.TextBox txtClientSettings 
               Height          =   315
               Index           =   0
               Left            =   180
               TabIndex        =   261
               Top             =   690
               Width           =   3855
            End
            Begin VB.TextBox txtTplDesc 
               Height          =   4560
               Left            =   165
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   11
               Top             =   1830
               Width           =   3885
            End
            Begin VB.ComboBox ComTplPrint 
               Height          =   315
               Left            =   1395
               Style           =   2  'Dropdown List
               TabIndex        =   10
               Top             =   90
               Width           =   2625
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Filename:"
               Height          =   195
               Index           =   50
               Left            =   180
               TabIndex        =   264
               Top             =   1020
               Width           =   675
            End
            Begin VB.Label lblMap 
               AutoSize        =   -1  'True
               Caption         =   "Name:"
               Height          =   195
               Index           =   49
               Left            =   180
               TabIndex        =   263
               Top             =   480
               Width           =   465
            End
            Begin VB.Label lblTemplateDescription 
               Caption         =   "Template Description:"
               Height          =   240
               Left            =   165
               TabIndex        =   20
               Top             =   1575
               Width           =   1950
            End
            Begin VB.Label lblSelectTemplate 
               Caption         =   "Select Template:"
               Height          =   285
               Left            =   135
               TabIndex        =   12
               Top             =   180
               Width           =   1275
            End
         End
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   360
         Left            =   8280
         Picture         =   "frmMapPrintAdmin.frx":329C7
         TabIndex        =   7
         Top             =   7890
         Width           =   720
      End
      Begin VB.TextBox edPrintFooter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7830
         TabIndex        =   5
         Text            =   "www.immap.org"
         Top             =   810
         Width           =   3105
      End
      Begin VB.TextBox edPrintTitle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7830
         TabIndex        =   2
         Text            =   "OASIS Map Print"
         Top             =   45
         Width           =   3105
      End
      Begin VB.TextBox edPrintSubTitle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7830
         TabIndex        =   1
         Text            =   "Oasis Map"
         Top             =   405
         Width           =   3105
      End
      Begin VB.Label lblPrintFooter 
         AutoSize        =   -1  'True
         Caption         =   "Print Footer:"
         Height          =   195
         Left            =   6750
         TabIndex        =   6
         Top             =   855
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Print title:"
         Height          =   255
         Left            =   6795
         TabIndex        =   4
         Top             =   90
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Print subtitle:"
         Height          =   255
         Left            =   6750
         TabIndex        =   3
         Top             =   450
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMapPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_rsPrintTpl As ADODB.Recordset
Private m_Legend As XGIS_Legend
Private m_GIS As XGIS_Viewer
Private udtTXTProp(10) As txtFonts

Public Enum GIS_UnitsType
    UnitsTypeUndefined = 0
    UnitsTypeLinear = 1
    UnitsTypeMeter = 2
    UnitsTypeKiloMeter = 3
    UnitsTypeFoot = 4
    UnitsTypeFootUS = 5
    UnitsTypeFootModifiedAmerican = 6
    UnitsTypeFootClarkes = 7
    UnitsTypeFootIndian = 8
    UnitsTypeLink = 9
    UnitsTypeLinkBenoit = 10
    UnitsTypeLinkSears = 11
    UnitsTypeChainBenoit = 12
    UnitsTypeChainSears = 13
    UnitsTypeYardIndian = 14
    UnitsTypeYardSears = 15
    UnitsTypeFathom = 16
    UnitsTypeMileNautical = 17
    UnitsTypeAngular = 18
    UnitsTypeRadian = 19
    UnitsTypeDegreeDecimal = 20
    UnitsTypeMinuteDecimal = 21
    UnitsTypeSecondDecimal = 22
    UnitsTypeGon = 23
    UnitsTypeGrad = 24
End Enum

Private RSLocalUserGroups As ADODB.Recordset
Private RSTemplates As ADODB.Recordset
Private m_bIsInProgress As Boolean
Private sUserGroupPrefix As String

Private Function UnitAsText(oUnit As GIS_UnitsType) As String
        '<EhHeader>
        On Error GoTo UnitAsText_Err
        '</EhHeader>
        
100     Select Case oUnit
    
            Case UnitsTypeUndefined
102             UnitAsText = " UnDef"

104         Case UnitsTypeLinear
106             UnitAsText = " Linear"

108         Case UnitsTypeMeter
110             UnitAsText = " M"

112         Case UnitsTypeKiloMeter
114             UnitAsText = " KM"

116         Case UnitsTypeFoot
118             UnitAsText = " ft"

120         Case UnitsTypeFootUS
122             UnitAsText = " ft us"

124         Case UnitsTypeFootModifiedAmerican
126             UnitAsText = " ft modified american"

128         Case UnitsTypeFootClarkes
130             UnitAsText = " ft Clarkes"

132         Case UnitsTypeFootIndian
134             UnitAsText = " ft Indian"

136         Case UnitsTypeLink
138             UnitAsText = " Link"

140         Case UnitsTypeLinkBenoit
142             UnitAsText = " Link Benoit"

144         Case UnitsTypeLinkSears
146             UnitAsText = " Link Sears"

148         Case UnitsTypeChainBenoit
150             UnitAsText = " Chain Benoit"

152         Case UnitsTypeChainSears
154             UnitAsText = " Chain Sears"

156         Case UnitsTypeYardIndian
158             UnitAsText = " yrd Indian"

160         Case UnitsTypeYardSears
162             UnitAsText = " yrd Sears"

164         Case UnitsTypeFathom
166             UnitAsText = " Fathom"

168         Case UnitsTypeMileNautical
170             UnitAsText = " Nautical mi"

172         Case UnitsTypeAngular
174             UnitAsText = " Angular"

176         Case UnitsTypeRadian
178             UnitAsText = " rd"

180         Case UnitsTypeDegreeDecimal
182             UnitAsText = " DD"

184         Case UnitsTypeMinuteDecimal
186             UnitAsText = " MD"

188         Case UnitsTypeSecondDecimal
190             UnitAsText = " SD"

192         Case UnitsTypeGon
194             UnitAsText = " Gon"

196         Case UnitsTypeGrad
198             UnitAsText = " Grad"
    
        End Select
  
        '<EhFooter>
        Exit Function

UnitAsText_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.UnitAsText " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub Init(sURL As String, _
                sPassedUserGroupPrefix As String, _
                Optional oViewer As XGIS_Viewer, _
                Optional oLgd As XGIS_Legend)
        '<EhHeader>
        On Error GoTo init_Err
        '</EhHeader>
        
        Dim sString As String
        
100     sUserGroupPrefix = sPassedUserGroupPrefix

102     If Not oViewer Is Nothing Then Set m_GIS = oViewer
104     If Not oLgd Is Nothing Then Set m_Legend = oLgd

106     sString = sURL & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & sUserGroupPrefix & "PrintTemplates")
108     Set m_rsPrintTpl = m_frmOASISProgress.OpenHttpCommsRS(sString, True)
        
110     ComTplPrint.Clear
112     ComTplPrint.AddItem "--No Template--"
        
114     If Not m_rsPrintTpl.EOF And Not m_rsPrintTpl.Bof Then
116         m_rsPrintTpl.MoveFirst

118         Do While Not m_rsPrintTpl.EOF
120             ComTplPrint.AddItem m_rsPrintTpl.fields.Item("Name").Value
122             m_rsPrintTpl.MoveNext
            Loop

        End If
    
124     txtTplDesc.Text = ""
126     txtTitle.Text = ""
128     txtSubTitle.Text = ""
130     txtCopyRight.Text = ""
132     txtVeiwer.Text = ""
134     txtMapID.Text = ""

136     If Not oViewer Is Nothing Then
138         PrintPreview.GIS_Viewer = oViewer
140         PrintPreviewSimple.GIS_Viewer = oViewer
142         oViewer.PrintTitle = edPrintTitle.Text
144         oViewer.PrintSubtitle = edPrintSubTitle.Text
146         oViewer.PrintFooter = edPrintFooter.Text
        End If
    
148     If ComTplPrint.ListCount >= 0 Then ComTplPrint.ListIndex = 0
    
        '<EhFooter>
        Exit Sub

init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.Init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkUseLegendInMap_Click()
        '<EhHeader>
        On Error GoTo chkUseLegendInMap_Click_Err
        '</EhHeader>
100     Legend1.Visible = IIf(chkUseLegendInMap.Value = vbChecked, True, False)
        '<EhFooter>
        Exit Sub

chkUseLegendInMap_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.chkUseLegendInMap_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdFonts_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo cmdFonts_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog

100     c.Font = lblTemplate(Index).Font
102     c.ShowFont
    
104     With lblTemplate(Index)
106         .FontBold = c.FontBold
108         .FontItalic = c.FontItalic
110         .FontName = c.FontName
112         .FontSize = c.FontSize
        End With
    
114     With udtTXTProp(Index)
116         .FontBold = c.FontBold
118         .FontItalic = c.FontItalic
120         .FontName = c.FontName
122         .FontSize = c.FontSize
        End With
       
        '<EhFooter>
        Exit Sub

cmdFonts_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.cmdFonts_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGraphicPath_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo cmdGraphicPath_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog
    
        On Error Resume Next
100     c.DefaultExt = "*.bmp"
102     c.DialogTitle = "Open Map Graphics File"
104     c.Filter = "Map Definition Files (*.bmp;*.wmf)|*.bmp;*.wmf"
106     c.ShowOpen
    
108     If chkUseRoot.Value = vbChecked Then
110         txtgraphPath(Index).Text = c.FileTitle
        Else
112         txtgraphPath(Index).Text = c.fileName
        End If

        '<EhFooter>
        Exit Sub

cmdGraphicPath_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.cmdGraphicPath_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdlocalDB_Click()
        '<EhHeader>
        On Error GoTo cmdlocalDB_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog
        Dim cn As ADODB.Connection
    
        On Error Resume Next
100     c.DefaultExt = "*.mdb"
102     c.DialogTitle = "Open OASIS Client DB File"
104     c.Filter = "OASIS Client DB Files (*.mdb)|*.mdb"
106     c.InitDir = CreateAppPath & "\data\db"
108     c.ShowOpen
    
110     txtLocalDB.Text = c.fileName
    
        Dim fs As New FileSystemObject
    
112     If fs.FileExists(txtLocalDB.Text) Then
114         Set cn = New ADODB.Connection
116         cn.open "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & txtLocalDB.Text & ";Uid=Admin;Pwd=;"
118         Init GIS.Viewer, Legend1.Legend, cn
        Else
120         txtLocalDB.Text = ""
122         MsgBox "File Does Not Exists..."
        End If
    
        '<EhFooter>
        Exit Sub

cmdlocalDB_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.cmdlocalDB_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdMapPath_Click()
        '<EhHeader>
        On Error GoTo cmdMapPath_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog
    
        On Error Resume Next
100     c.DefaultExt = "*.ttkgp"
102     c.DialogTitle = "Open Map Definition File"
104     c.Filter = "Map Definition Files (*.ttkgp;*.prj)|*.ttkgp;*.prj"
106     c.InitDir = CreateAppPath & "\data\user\Maps"
108     c.ShowOpen
110     txtMapPath.Text = c.fileName
112     GIS.open txtMapPath.Text, IIf(chkUseStrict.Value = vbChecked, True, False)
114     Legend1.GIS_Viewer = GIS.Viewer
116     Legend1.Update
118     ControlScale1.GIS_Viewer = GIS.Viewer
120     ControlScale1.UnitsType = GIS.Viewer.Units.Units
122     ControlScale1.UnitsTxt = UnitAsText(ControlScale1.UnitsType)
        
124     PrintPreview.GIS_Viewer = GIS.Viewer
126     PrintPreview.Preview 1
128     cmdUpdate.Enabled = True
130     cmdPrint.Enabled = True
132     cmdNew.Enabled = True
        
        'C1TTabPrint.TabEnabled(1) = True
134     C1TTabPrint.TabEnabled(2) = True
136     C1TTabPrint.TabEnabled(3) = True
138     C1TTabPrint.TabEnabled(4) = True
    
140     GIS.Update
        '<EhFooter>
        Exit Sub

cmdMapPath_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.cmdMapPath_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdMapTools_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo cmdMapTools_Click_Err
        '</EhHeader>
    
100     Select Case Index
    
            Case 0
102             GIS.Mode = XgisZoomEx

104         Case 1
106             GIS.Mode = XgisZoom

108         Case 2
110             GIS.Mode = XgisDrag

112         Case 3
        End Select
    
        '<EhFooter>
        Exit Sub

cmdMapTools_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.cmdMapTools_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdNew_Click()
        '<EhHeader>
        On Error GoTo cmdNew_Click_Err
        '</EhHeader>
        Dim sName As String
        Dim i As Integer
    
100     sName = InputBox("Enter the Name For the new template", "Template Name", "My Template " & Now())
 
102     m_rsPrintTpl.AddNew
104     m_rsPrintTpl.fields.Item("Name").Value = sName
    
106     cmdUpdate.Enabled = True
108     cmdPrint.Enabled = True
110     cmdSave.Enabled = True
    
112     ComTplPrint.Clear
    
114     ComTplPrint.AddItem "--No Template--"
        
116     If Not m_rsPrintTpl.EOF And Not m_rsPrintTpl.Bof Then
118         m_rsPrintTpl.MoveFirst

120         Do While Not m_rsPrintTpl.EOF
122             ComTplPrint.AddItem m_rsPrintTpl.fields.Item("Name").Value
124             m_rsPrintTpl.MoveNext
            Loop
        
126         FindIndexStrEx ComTplPrint, sName
128         txtClientSettings(0).Text = sName
        End If
    
        '<EhFooter>
        Exit Sub

cmdNew_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.cmdNew_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdPrint_Click()
        '<EhHeader>
        On Error GoTo cmdPrint_Click_Err
        '</EhHeader>
100     ApplyPrintTpl , , , True
102     cmdSave.Enabled = True
        '<EhFooter>
        Exit Sub

cmdPrint_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.cmdPrint_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>

        Dim fs As New FileSystemObject
        Dim oFile As File
        Dim oTXTStream As TextStream
        Dim c As New cCommonDialog

100     If Len(txtRawEdt.Text) < 100 Then
102         MsgBox "It seems like The edited parameters are not correct."
            Exit Sub
        End If

104     c.DefaultExt = "*.tpl"
106     c.DialogTitle = "Save Map Print Template"
108     c.Filter = "Map Print Template Files (*.tpl)|*.tpl"
110     c.InitDir = CreateAppPath & "\data\templates\printtemplates"
112     c.ShowSave

114     If Len(c.fileName) < 6 Then Exit Sub
116     Set oTXTStream = fs.CreateTextFile(c.fileName)
118     oTXTStream.Write txtRawEdt.Text
120     oTXTStream.Close
    
122     If Not ComTplPrint.ListIndex = 0 Then
124         m_rsPrintTpl.MoveFirst
126         m_rsPrintTpl.Find "Name = '" & ComTplPrint.List(ComTplPrint.ListIndex) & "'"
        Else
            Exit Sub
        End If
    
128     m_rsPrintTpl.fields.Item("description").Value = txtTplDesc.Text
130     m_rsPrintTpl.fields.Item("MapTitle").Value = txtTitle.Text
132     m_rsPrintTpl.fields.Item("MapSubTitle").Value = txtSubTitle.Text
134     m_rsPrintTpl.fields.Item("copyright").Value = txtCopyRight.Text
136     m_rsPrintTpl.fields.Item("note").Value = txtVeiwer.Text
138     m_rsPrintTpl.fields.Item("MapIDPrefix").Value = txtMapID.Text
140     m_rsPrintTpl.fields.Item("Name").Value = txtClientSettings(0).Text
142     txtClientSettings(0).Text = c.FileTitle
144     m_rsPrintTpl.fields.Item("FileName").Value = c.FileTitle

146     SaveFileToTable c.fileName
    
        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.cmdSave_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CheckEdit() As Boolean
        '<EhHeader>
        On Error GoTo CheckEdit_Err
        '</EhHeader>

        Dim RSProfile As ADODB.Recordset
        Dim sSQL As String
        Dim bReturnValue As String

100     With m_rsPrintTpl
102         .Filter = adFilterPendingRecords
    
104         If Not .Bof And Not .EOF Then
106             If MsgBox("Do you wish to save your changes to Server?", vbYesNo, "Confirm Save") = vbYes Then
                
108                 bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(m_rsPrintTpl, WebSite & "Oasis.asp", True)
                
110                 If bReturnValue Then
112                     IncrementProfileSettingVersion WebSite, "SettingValue9", sUserGroupPrefix
114                     MsgBox "Data saved to server"
                    Else
116                     MsgBox "Saving to server failed!"
                    End If

118                 CheckEdit = True
                End If
           
            End If
                
        End With

        '<EhFooter>
        Exit Function

CheckEdit_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.CheckEdit " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdUpdate_Click()
        '<EhHeader>
        On Error GoTo cmdUpdate_Click_Err
        '</EhHeader>
100     ApplyPrintTpl True, IIf(chkShowPrinter.Value = vbChecked, True, False)
102     cmdSave.Enabled = True
        '<EhFooter>
        Exit Sub

cmdUpdate_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.cmdUpdate_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComTplPrint_Click()
        '<EhHeader>
        On Error GoTo ComTplPrint_Click_Err
        '</EhHeader>
100     ApplyPrintTpl
        '<EhFooter>
        Exit Sub

ComTplPrint_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.ComTplPrint_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetPrntGraphStrings(ind As Integer) As String
        '<EhHeader>
        On Error GoTo GetPrntGraphStrings_Err
        '</EhHeader>
        Dim sRet As String
        Dim sU As String

100     sU = GetUnits
    
102     sRet = "GRAPHIC" & ind + 1 & "= "
104     sRet = sRet & "*" & txtGraphicsLeft(ind) & sU & ", *" & txtGraphicsTop(ind) & sU & ", *" & txtGraphicsRight(ind) & sU & ", *" & txtGraphicsBottom(ind) & sU & ",""" & txtgraphPath(ind).Text & """"

106     GetPrntGraphStrings = sRet
        '<EhFooter>
        Exit Function

GetPrntGraphStrings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.GetPrntGraphStrings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function GetPrntTextStrings(ind As Integer) As String
        '<EhHeader>
        On Error GoTo GetPrntTextStrings_Err
        '</EhHeader>
        Dim sString As String
        Dim sU As String

100     sU = GetUnits
102     sString = "TEXT" & (ind + 1) & "="
104     sString = sString & " *" & txtTextLeft(ind) & sU & ",*" & txtTextTop(ind) & sU & ",*" & txtTextRight(ind) & sU & ",*" & txtTextBottom(ind) & sU
106     GetPrntTextStrings = sString & "," & GetJustified(ind) & "," & dxColorTXT(ind).EditValue & "," & udtTXTProp(ind).FontName & "," & lblTemplate(ind).FontSize
        
        '<EhFooter>
        Exit Function

GetPrntTextStrings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.GetPrntTextStrings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function GetDefaultPrintTemplate() As String
        '<EhHeader>
        On Error GoTo GetDefaultPrintTemplate_Err
        '</EhHeader>
        Dim sTxt As String
        Dim sU As String
        Dim i As Integer
    
100     sU = GetUnits
    
102     sTxt = "[TatukGIS PrintTemplate] "
104     sTxt = sTxt & vbCrLf & "; PAGESIZE needs only be set on PDF device"
106     sTxt = sTxt & vbCrLf & "PAGESIZE=297.00mm,210.00mm"

108     sTxt = sTxt & vbCrLf & "; place graphic and text at the top"
    
110     If chkNoGraphics.Value = vbUnchecked Then

            Do
112             sTxt = sTxt & vbCrLf & GetPrntGraphStrings(i) ' "GRAPHIC1= *1.00cm, *0.5cm, *8.00cm, *2cm,""""logo.BMP"""
114             i = i + 1
116         Loop Until optGraphics(i - 1).Value = True

        End If
    
118     sTxt = sTxt & vbCrLf & "; place Print Main Title"
        'm_frmDebug.DebugPrint IIf(chkSaveText.Value = vbChecked, ",""" & txtTitle.Text & """", "")
    
120     sTxt = sTxt & vbCrLf & GetPrntTextStrings(0) & IIf(chkSaveText.Value = vbChecked, ",,""" & txtTitle.Text & "", "")
122     sTxt = sTxt & vbCrLf & "; place Subtitle"
124     sTxt = sTxt & vbCrLf & GetPrntTextStrings(1) & IIf(chkSaveText.Value = vbChecked, ",,""" & txtSubTitle.Text & """", "")
126     sTxt = sTxt & vbCrLf & "; place text at the top"
128     sTxt = sTxt & vbCrLf & GetPrntTextStrings(2) & IIf(chkSaveText.Value = vbChecked, ",,""" & txtCopyRight.Text & """", "")
130     sTxt = sTxt & vbCrLf & "; place text at the top"
132     sTxt = sTxt & vbCrLf & GetPrntTextStrings(3) & IIf(chkSaveText.Value = vbChecked, ",,""" & txtVeiwer.Text & """", "")

134     sTxt = sTxt & vbCrLf & "; place MAP ID"
136     sTxt = sTxt & vbCrLf & GetPrntTextStrings(4) & IIf(chkSaveText.Value = vbChecked, ",,""" & txtMapID.Text & """", "")
    
138     sTxt = sTxt & vbCrLf & "; place Date"
140     sTxt = sTxt & vbCrLf & GetPrntTextStrings(5) & IIf(chkSaveText.Value = vbChecked, ",,""" & txtMapProdctionDate.Text & """", "")
    
142     If chkUseMap.Value = vbChecked Then

            'left,top,right,bottom
144         If Not OptMapFrame(0).Value = True Then
146             sTxt = sTxt & vbCrLf & "; draw backround border for the map & the map itself"
148             sTxt = sTxt & vbCrLf & "BOX1= *" & txtMapLeft(0).Text & sU & ", *" & txtMapTop(0).Text & sU & ",*" & txtMapRight(0).Text & sU & ",*" & txtMapBottom(0).Text & sU & "," & dxmapColor(0).EditValue   'Blue"

150             If OptMapFrame(2).Value Then
152                 sTxt = sTxt & vbCrLf & "BOX2= *" & txtMapLeft(1).Text & sU & ", *" & txtMapTop(1).Text & sU & ",*" & txtMapRight(1).Text & sU & ",*" & txtMapBottom(1).Text & sU & "," & dxmapColor(1).EditValue 'Yellow"
                End If
            End If
    
154         sTxt = sTxt & vbCrLf & "MAP1= *" & txtMapLeft(2).Text & sU & ", *" & txtMapTop(2).Text & sU & ",*" & txtMapRight(2).Text & sU & ",*" & txtMapBottom(2).Text & sU
        End If
    
156     If chkUseScalebar.Value = vbChecked Then
158         sTxt = sTxt & vbCrLf & "; draw the scale"
160         sTxt = sTxt & vbCrLf & "Scale1= *" & txtLeftScale.Text & sU & ", *" & txtTopScale.Text & sU & ",*" & txtRightScale.Text & sU & ",*" & txtbottomScale.Text & sU
        End If
    
162     If chkUseLegend.Value = vbChecked Then
164         sTxt = sTxt & vbCrLf & "; draw background border for the legend & the legend itself"
166         sTxt = sTxt & vbCrLf & "BOX3= *" & txtLegendLeft(0).Text & sU & ", *" & txtLegendTop(0).Text & sU & ",*" & txtLegendRight(0).Text & sU & ",*" & txtLegendBottom(0).Text & sU & "," & dxLgdColor(0).EditValue  'Blue"
168         sTxt = sTxt & vbCrLf & "BOX4= *" & txtLegendLeft(1).Text & sU & ", *" & txtLegendTop(1).Text & sU & ",*" & txtLegendRight(1).Text & sU & ",*" & txtLegendBottom(1).Text & sU & "," & dxLgdColor(1).EditValue  'Yellow"
170         sTxt = sTxt & vbCrLf & ";white background because legend is transparent by default"
172         sTxt = sTxt & vbCrLf & "BOX5= *" & txtLegendLeft(2).Text & sU & ", *" & txtLegendTop(2).Text & sU & ",*" & txtLegendRight(2).Text & sU & ",*" & txtLegendBottom(2).Text & sU & "," & dxLgdColor(2).EditValue  'White"
174         sTxt = sTxt & vbCrLf & "LEGEND1= *" & txtLegendLeft(3).Text & sU & ", *" & txtLegendTop(3).Text & sU & ",*" & txtLegendRight(3).Text & sU & ",*" & txtLegendBottom(3).Text & sU
        End If
    
176     sTxt = sTxt & vbCrLf & "; draw thin line around the map"
178     sTxt = sTxt & vbCrLf & "FRAME1= *0.01cm, *0.01cm,*-0.01cm,*-0.01cm,BLACK,0.01mm"

180     GetDefaultPrintTemplate = sTxt
182     txtRawEdt.Text = sTxt
        '<EhFooter>
        Exit Function

GetDefaultPrintTemplate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.GetDefaultPrintTemplate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub ApplyPrintTpl(Optional bUseIgnoreDbSettings As Boolean = False, _
                          Optional bShowPrinterSettings As Boolean = False, _
                          Optional bSaveToFile As Boolean, _
                          Optional bDOPrint As Boolean)
        '<EhHeader>
        On Error GoTo ApplyPrintTpl_Err
        '</EhHeader>
        Dim tmp As XGIS_TemplatePrint
        Dim g_sAppPath As String
        Dim fs As New FileSystemObject
        Dim oFile As File
        Dim oTXTStream As TextStream
        Dim c As New cCommonDialog

        'c.Font = frmMapPrint.Font
        'c.ShowFont

100     If m_rsPrintTpl Is Nothing Then Exit Sub

102     g_sAppPath = CreateAppPath
    
        'm_frmDebug.DebugPrint dxmapColor.EditValue
    
104     PrintPreview.GIS_Viewer = GIS.Viewer
    
106     If ComTplPrint.ListIndex = 0 Or m_rsPrintTpl Is Nothing Then
        
108         Set tmp = New XGIS_TemplatePrint

110         If bDOPrint Then PrintPreviewSimple.GIS_Viewer = GIS.Viewer
112         tmp.Create_ GIS.Viewer
        
114         If fs.FileExists(App.Path & "\my.tpl") Then
116             fs.DeleteFile App.Path & "\my.tpl", True
            End If
    
118         Set oTXTStream = fs.CreateTextFile(App.Path & "\my.tpl")
        
120         oTXTStream.Write GetDefaultPrintTemplate
        
122         oTXTStream.Close
        
124         tmp.Path = App.Path & "\my.tpl"
126         tmp.GIS_Legend(1) = Legend1.Legend
    
128         tmp.GIS_Scale(1) = ControlScale1.Scale
130         tmp.GIS_ViewerExtent(1) = GIS.VisibleExtent
            
132         If chkSaveText.Value = vbUnchecked Then
134             If chkMapTexts(0).Value = vbChecked Then
136                 tmp.Text(1) = txtTitle.Text
                End If

138             If chkMapTexts(1).Value = vbChecked Then
140                 tmp.Text(2) = txtSubTitle.Text
                End If

142             If chkMapTexts(2).Value = vbChecked Then
144                 tmp.Text(3) = txtCopyRight.Text
                End If

146             If chkMapTexts(3).Value = vbChecked Then
148                 tmp.Text(4) = txtVeiwer.Text
                End If

150             If chkMapTexts(4).Value = vbChecked Then
152                 tmp.Text(5) = txtMapID.Text
                End If

154             If chkMapTexts(5).Value = vbChecked Then
156                 tmp.Text(6) = txtMapProdctionDate.Text & " " & dxDateEdit1.EditValue
                End If

158             GIS.PrintTitle = edPrintTitle.Text
160             GIS.PrintSubtitle = edPrintSubTitle.Text
162             GIS.PrintFooter = edPrintFooter.Text
                'Printer.Orientation = vbPRORPortrait
            End If
            
164         If bShowPrinterSettings Then PrintPreviewSimple.PrinterSetup
        
166         If bSaveToFile Then
168             c.DefaultExt = "*.tpl"
170             c.DialogTitle = "Save Map Print Template"
172             c.Filter = "Map Print Template Files (*.tpl)|*.tpl"
174             c.ShowSave
                
176             Set oTXTStream = fs.CreateTextFile(c.fileName)
178             oTXTStream.Write GetDefaultPrintTemplate
180             oTXTStream.Close
            End If
        
            'PrintPreviewSimple.GIS_Viewer = GIS.Viewer
            
182         If bDOPrint Then
184             PrintPreviewSimple.Preview
            End If
            
186         PrintPreview.Preview (1)
        Else
        
188         m_rsPrintTpl.MoveFirst
190         m_rsPrintTpl.Find "Name = '" & ComTplPrint.List(ComTplPrint.ListIndex) & "'"
            
192         If Not bUseIgnoreDbSettings Then
                
194             If Not IsNull(m_rsPrintTpl.fields.Item("description").Value) Then txtTplDesc.Text = m_rsPrintTpl.fields.Item("description").Value
196             If Not IsNull(m_rsPrintTpl.fields.Item("MapTitle").Value) Then txtTitle.Text = m_rsPrintTpl.fields.Item("MapTitle").Value
198             If Not IsNull(m_rsPrintTpl.fields.Item("MapSubTitle").Value) Then txtSubTitle.Text = m_rsPrintTpl.fields.Item("MapSubTitle").Value
200             If Not IsNull(m_rsPrintTpl.fields.Item("copyright").Value) Then txtCopyRight.Text = m_rsPrintTpl.fields.Item("copyright").Value
202             If Not IsNull(m_rsPrintTpl.fields.Item("note").Value) Then txtVeiwer.Text = m_rsPrintTpl.fields.Item("note").Value
204             If Not IsNull(m_rsPrintTpl.fields.Item("MapIDPrefix").Value) Then txtMapID.Text = m_rsPrintTpl.fields.Item("MapIDPrefix").Value
        
            End If
        
206         Set tmp = New XGIS_TemplatePrint
    
208         PrintPreviewSimple.GIS_Viewer = GIS.Viewer
210         tmp.Create_ GIS.Viewer
            
212         If Not IsNull(m_rsPrintTpl.fields.Item("FileName").Value) Then
214             If Not FILE_EXISTS(g_sAppPath & "\data\gis\other\Graphics\" & m_rsPrintTpl.fields.Item("FileName").Value) Then
216                 SaveFileToDisk g_sAppPath
                End If
    
218             tmp.Path = g_sAppPath & "\data\gis\other\Graphics\" & m_rsPrintTpl.fields.Item("FileName").Value
            
220             tmp.Text(1) = txtTitle.Text
222             tmp.Text(2) = txtCopyRight.Text
224             tmp.Text(3) = txtSubTitle.Text
226             tmp.Text(4) = txtVeiwer.Text
228             tmp.Text(5) = txtMapID.Text
230             tmp.Text(6) = txtMapProdctionDate.Text & " " & dxDateEdit1.EditValue
            
232             If Not m_GIS Is Nothing Then
234                 m_GIS.PrintTitle = edPrintTitle.Text
236                 m_GIS.PrintSubtitle = edPrintSubTitle.Text
238                 m_GIS.PrintFooter = edPrintFooter.Text
                End If

            Else
 
240             If fs.FileExists(App.Path & "\my.tpl") Then
242                 fs.DeleteFile App.Path & "\my.tpl", True
                End If
    
244             Set oTXTStream = fs.CreateTextFile(App.Path & "\my.tpl")
        
246             oTXTStream.Write GetDefaultPrintTemplate
        
248             oTXTStream.Close
        
250             tmp.Path = App.Path & "\my.tpl"
252             tmp.GIS_Legend(1) = Legend1.Legend
    
254             tmp.GIS_Scale(1) = ControlScale1.Scale
256             tmp.GIS_ViewerExtent(1) = GIS.VisibleExtent
 
258             If chkMapTexts(0).Value = vbChecked Then
260                 tmp.Text(1) = txtTitle.Text
                End If

262             If chkMapTexts(1).Value = vbChecked Then
264                 tmp.Text(2) = txtSubTitle.Text
                End If

266             If chkMapTexts(2).Value = vbChecked Then
268                 tmp.Text(3) = txtCopyRight.Text
                End If

270             If chkMapTexts(3).Value = vbChecked Then
272                 tmp.Text(4) = txtVeiwer.Text
                End If

274             If chkMapTexts(4).Value = vbChecked Then
276                 tmp.Text(5) = txtMapID.Text
                End If

278             If chkMapTexts(5).Value = vbChecked Then
280                 tmp.Text(6) = txtMapProdctionDate.Text & " " & dxDateEdit1.EditValue
                End If

282             GIS.PrintTitle = edPrintTitle.Text
284             GIS.PrintSubtitle = edPrintSubTitle.Text
286             GIS.PrintFooter = edPrintFooter.Text
            End If
            
288         tmp.GIS_Legend(1) = m_Legend
    
290         tmp.GIS_Scale(1) = ControlScale1.Scale

292         If Not m_GIS Is Nothing Then
294             tmp.GIS_ViewerExtent(1) = m_GIS.VisibleExtent
            End If

            'Printer.Orientation = vbPRORPortrait
            
296         If bShowPrinterSettings Then PrintPreviewSimple.PrinterSetup
298         PrintPreview.Preview (1)
        
        End If

        '<EhFooter>
        Exit Sub

ApplyPrintTpl_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.ApplyPrintTpl " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub edPrintFooter_Change()
        '<EhHeader>
        On Error GoTo edPrintFooter_Change_Err
        '</EhHeader>
100     PrintPreview.GIS_Viewer.PrintFooter = edPrintFooter.Text
        '<EhFooter>
        Exit Sub

edPrintFooter_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.edPrintFooter_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub edPrintSubTitle_Change()
        '<EhHeader>
        On Error GoTo edPrintSubTitle_Change_Err
        '</EhHeader>
  
100     PrintPreview.GIS_Viewer.PrintSubtitle = edPrintSubTitle.Text
        '<EhFooter>
        Exit Sub

edPrintSubTitle_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.edPrintSubTitle_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub edPrintTitle_Change()
        '<EhHeader>
        On Error GoTo edPrintTitle_Change_Err
        '</EhHeader>
  
100     PrintPreview.GIS_Viewer.PrintTitle = edPrintTitle.Text
        '<EhFooter>
        Exit Sub

edPrintTitle_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.edPrintTitle_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim i As Integer
100     GIS.ZOrder vbSendToBack
        
102     ComUnits.ListIndex = 2
        
104     For i = 0 To ComJustify.UBound
106         ComJustify(i).ListIndex = 2
        Next
        
108     For i = 0 To 9
110         udtTXTProp(i).FontBold = lblTemplate(i).FontBold
112         udtTXTProp(i).FontName = lblTemplate(i).FontName
114         udtTXTProp(i).FontItalic = lblTemplate(i).FontItalic
116         udtTXTProp(i).FontSize = lblTemplate(i).FontSize
        Next
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GetJustified(ind As Integer) As String
        '<EhHeader>
        On Error GoTo GetJustified_Err
        '</EhHeader>
        Dim SReturn As String

100     Select Case ComJustify(ind).ListIndex
    
            Case 1
102             SReturn = "LEFTJUSTIFY"

104         Case 2
106             SReturn = "CENTER"

108         Case Else
110             SReturn = "RIGHTJUSTIFY"
        End Select

112     If udtTXTProp(ind).FontBold Then SReturn = SReturn & IIf(Len(SReturn) > 2, ":BOLD", "BOLD")
114     If udtTXTProp(ind).FontItalic Then SReturn = SReturn & IIf(Len(SReturn) > 2, ":ITALIC", "ITALIC")
116     If chkStrikeThrough(ind).Value = vbChecked Then SReturn = SReturn & IIf(Len(SReturn) > 2, ":STRIKEOUT", "STRIKEOUT")
118     If chkUnderline(ind).Value = vbChecked Then SReturn = SReturn & IIf(Len(SReturn) > 2, ":UNDERLINE", "UNDERLINE")
    
120     GetJustified = SReturn

        '<EhFooter>
        Exit Function

GetJustified_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.GetJustified " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function GetUnits() As String
        '<EhHeader>
        On Error GoTo GetUnits_Err
        '</EhHeader>
    
100     Select Case ComUnits.ListIndex

            Case 0
102             GetUnits = "" 'Twips

104         Case 1
106             GetUnits = "px" 'Pixels

108         Case 2
110             GetUnits = "cm" 'Centimeters

112         Case 3
114             GetUnits = "mm" 'Milimeters

116         Case 4
118             GetUnits = "pt" 'Points

120         Case 5
122             GetUnits = "in" 'Inches
        End Select
    
        '<EhFooter>
        Exit Function

GetUnits_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.GetUnits " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     CheckEdit
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GIS_OnZoomChange(translated As Boolean)
    '    translated = True
    '    PrintPreview.Preview (1)
    
End Sub

Private Function SaveFileToDisk(sPath As String) As String
        '<EhHeader>
        On Error GoTo SaveFileToDisk_Err
        '</EhHeader>
        Dim FileStream As New ADODB.Stream
        Dim fileName As String
                
100     If Not m_rsPrintTpl.EOF Then
102         fileName = sPath & "\data\gis\other\Graphics\" & m_rsPrintTpl.fields.Item("FileName").Value
104         FileStream.Type = adTypeBinary
106         FileStream.open
108         FileStream.Write m_rsPrintTpl.fields("blob_tpl").Value
110         FileStream.SaveToFile fileName, adSaveCreateOverWrite
        End If
    
112     SaveFileToDisk = fileName

        '<EhFooter>
        Exit Function

SaveFileToDisk_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.SaveFileToDisk " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub SaveFileToTable(sfile As String)
        '<EhHeader>
        On Error GoTo SaveFileToTable_Err
        '</EhHeader>
        Dim FileStream As New ADODB.Stream
        
100     If Not m_rsPrintTpl.EOF Then

102         FileStream.Type = adTypeBinary
104         FileStream.open
106         FileStream.LoadFromFile sfile

108         m_rsPrintTpl.fields("blob_tpl") = FileStream.Read
110         m_rsPrintTpl.Update
        
        End If
    
        Exit Sub

        '<EhFooter>
        Exit Sub

SaveFileToTable_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapPrint.SaveFileToTable " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

