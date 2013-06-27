VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.UserControl ctrWhere 
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11775
   ScaleHeight     =   7125
   ScaleWidth      =   11775
   Begin XpressEditorsLibCtl.dxPickEdit dxPickMoveData 
      Height          =   315
      Left            =   90
      OleObjectBlob   =   "ctrWhere.ctx":0000
      TabIndex        =   41
      Top             =   6660
      Width           =   1950
   End
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   7125
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11775
      _cx             =   20770
      _cy             =   12568
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
      GridRows        =   1
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"ctrWhere.ctx":0411
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab tabWhere 
         Height          =   7065
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   11715
         _cx             =   20664
         _cy             =   12462
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
         Caption         =   "General|Location"
         Align           =   0
         CurrTab         =   1
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
         Begin C1SizerLibCtl.C1Elastic elGeneral 
            Height          =   6690
            Left            =   -12270
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   11625
            _cx             =   20505
            _cy             =   11800
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
            Begin XpressEditorsLibCtl.dxMemoEdit dxMemDescription 
               DataField       =   "Description"
               DataSource      =   "AdodcWhere"
               Height          =   2760
               Left            =   1095
               OleObjectBlob   =   "ctrWhere.ctx":0449
               TabIndex        =   26
               Top             =   1455
               Width           =   4455
            End
            Begin VB.Frame FraPrivacy 
               Caption         =   "Privacy:"
               Height          =   900
               Left            =   5715
               TabIndex        =   49
               Top             =   360
               Width           =   5340
               Begin VB.OptionButton OptPrivacy 
                  Caption         =   "Public"
                  Height          =   330
                  Index           =   2
                  Left            =   2250
                  TabIndex        =   50
                  Top             =   270
                  Width           =   1050
               End
               Begin VB.OptionButton OptPrivacy 
                  Caption         =   "User group"
                  Height          =   330
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   51
                  Top             =   270
                  Width           =   1230
               End
               Begin VB.OptionButton OptPrivacy 
                  Caption         =   "Private"
                  Height          =   330
                  Index           =   0
                  Left            =   135
                  TabIndex        =   52
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   1050
               End
            End
            Begin VB.Frame FraLocationVisibility 
               Caption         =   "Location visibility level:"
               Height          =   2685
               Left            =   5715
               TabIndex        =   42
               Top             =   1440
               Width           =   5340
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "National"
                  Height          =   375
                  Index           =   0
                  Left            =   270
                  TabIndex        =   48
                  Top             =   225
                  Value           =   1  'Checked
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "Province"
                  Height          =   375
                  Index           =   1
                  Left            =   270
                  TabIndex        =   47
                  Top             =   615
                  Value           =   1  'Checked
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "District"
                  Height          =   375
                  Index           =   2
                  Left            =   270
                  TabIndex        =   46
                  Top             =   1005
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "Sub District"
                  Height          =   375
                  Index           =   3
                  Left            =   270
                  TabIndex        =   45
                  Top             =   1380
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "Town"
                  Height          =   375
                  Index           =   4
                  Left            =   270
                  TabIndex        =   44
                  Top             =   1770
                  Width           =   1680
               End
               Begin VB.CheckBox chkVisbilityLevel 
                  Caption         =   "Actual Location"
                  Height          =   375
                  Index           =   5
                  Left            =   270
                  TabIndex        =   43
                  Top             =   2160
                  Width           =   1680
               End
            End
            Begin MSAdodcLib.Adodc AdodcLookUp 
               Height          =   420
               Left            =   180
               Top             =   5715
               Visible         =   0   'False
               Width           =   2445
               _ExtentX        =   4313
               _ExtentY        =   741
               ConnectMode     =   0
               CursorLocation  =   3
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   3
               LockType        =   3
               CommandType     =   8
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "LookUp"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin MSAdodcLib.Adodc AdodcWhere 
               Height          =   330
               Left            =   1125
               Top             =   5175
               Visible         =   0   'False
               Width           =   3465
               _ExtentX        =   6112
               _ExtentY        =   582
               ConnectMode     =   0
               CursorLocation  =   2
               IsolationLevel  =   -1
               ConnectionTimeout=   15
               CommandTimeout  =   30
               CursorType      =   1
               LockType        =   3
               CommandType     =   2
               CursorOptions   =   0
               CacheSize       =   50
               MaxRecords      =   0
               BOFAction       =   0
               EOFAction       =   0
               ConnectStringType=   1
               Appearance      =   1
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   "Where"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   204
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxedtPlaceName 
               DataField       =   "name"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   1095
               OleObjectBlob   =   "ctrWhere.ctx":0545
               TabIndex        =   27
               Top             =   465
               Width           =   4395
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxTextEdit1 
               DataField       =   "pCode"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Index           =   3
               Left            =   1095
               OleObjectBlob   =   "ctrWhere.ctx":05BB
               TabIndex        =   28
               Top             =   0
               Width           =   1215
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxTextEdit1 
               DataField       =   "id"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Index           =   0
               Left            =   4650
               OleObjectBlob   =   "ctrWhere.ctx":064A
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpPlaceType 
               DataField       =   "placeTypeId"
               DataSource      =   "AdodcWhere"
               Height          =   315
               Left            =   1095
               OleObjectBlob   =   "ctrWhere.ctx":06D7
               TabIndex        =   30
               Top             =   915
               Width           =   4380
            End
            Begin XpressEditorsLibCtl.dxStyleController dxStyleController1 
               Index           =   0
               Left            =   225
               OleObjectBlob   =   "ctrWhere.ctx":08C0
               Top             =   3465
            End
            Begin XpressEditorsLibCtl.dxStyleController dxStyleController1 
               Index           =   1
               Left            =   720
               OleObjectBlob   =   "ctrWhere.ctx":09BA
               Top             =   3960
            End
            Begin XpressEditorsLibCtl.dxStyleController dxStyleController1 
               Index           =   3
               Left            =   1260
               OleObjectBlob   =   "ctrWhere.ctx":0AF0
               Top             =   3780
            End
            Begin XpressEditorsLibCtl.dxStyleController dxStyleController1 
               Index           =   2
               Left            =   1800
               OleObjectBlob   =   "ctrWhere.ctx":0C08
               Top             =   3825
            End
            Begin VB.Label lblPCode 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "ID:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   2
               Left            =   4305
               TabIndex        =   35
               Top             =   60
               Visible         =   0   'False
               Width           =   180
            End
            Begin XpressEditorsLibCtl.dxImageLists dxImageLists1 
               Left            =   5190
               OleObjectBlob   =   "ctrWhere.ctx":0D3E
               Top             =   3750
            End
            Begin VB.Label lbPlaceName 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Place Name:"
               Height          =   195
               Left            =   105
               TabIndex        =   34
               Top             =   495
               Width           =   915
            End
            Begin VB.Label lblPlaceType 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Place Type:"
               Height          =   195
               Left            =   165
               TabIndex        =   33
               Top             =   915
               Width           =   855
            End
            Begin VB.Label lblPCode 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "PCode:"
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   360
               TabIndex        =   32
               Top             =   60
               Width           =   525
            End
            Begin VB.Label lblDescription 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Description:"
               Height          =   195
               Left            =   180
               TabIndex        =   31
               Top             =   1500
               Width           =   840
            End
         End
         Begin C1SizerLibCtl.C1Elastic elLocation 
            Height          =   6690
            Left            =   45
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   11625
            _cx             =   20505
            _cy             =   11800
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
            GridRows        =   2
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"ctrWhere.ctx":1AF6
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic elHolder 
               Height          =   2310
               Left            =   90
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   4290
               Width           =   11445
               _cx             =   20188
               _cy             =   4075
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
               BorderWidth     =   2
               ChildSpacing    =   2
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
               _GridInfo       =   $"ctrWhere.ctx":1B36
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Tab tabPCode 
                  Height          =   2250
                  Left            =   30
                  TabIndex        =   9
                  Top             =   30
                  Width           =   11385
                  _cx             =   20082
                  _cy             =   3969
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
                  Caption         =   "Admin Locator|Geo Locator"
                  Align           =   0
                  CurrTab         =   0
                  FirstTab        =   0
                  Style           =   3
                  Position        =   0
                  AutoSwitch      =   -1  'True
                  AutoScroll      =   -1  'True
                  TabPreview      =   -1  'True
                  ShowFocusRect   =   -1  'True
                  TabsPerPage     =   2
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
                  Begin C1SizerLibCtl.C1Elastic elGeo 
                     Height          =   1875
                     Left            =   12030
                     TabIndex        =   10
                     TabStop         =   0   'False
                     Top             =   330
                     Width           =   11295
                     _cx             =   19923
                     _cy             =   3307
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
                     Begin VB.CommandButton cmdRadius 
                        Height          =   330
                        Left            =   3060
                        Picture         =   "ctrWhere.ctx":1B6D
                        Style           =   1  'Graphical
                        TabIndex        =   14
                        Top             =   135
                        Width           =   330
                     End
                     Begin VB.CommandButton cmdCheckIn 
                        Height          =   330
                        Left            =   3060
                        Picture         =   "ctrWhere.ctx":1F96
                        Style           =   1  'Graphical
                        TabIndex        =   13
                        ToolTipText     =   "Check Coordinate in Map"
                        Top             =   855
                        Width           =   330
                     End
                     Begin VB.CommandButton cmdGetFrom 
                        Height          =   330
                        Left            =   3060
                        Picture         =   "ctrWhere.ctx":23E0
                        Style           =   1  'Graphical
                        TabIndex        =   12
                        ToolTipText     =   "Get Coordinates From Map"
                        Top             =   495
                        Width           =   330
                     End
                     Begin VB.TextBox txtMGRS 
                        Height          =   285
                        Left            =   1035
                        TabIndex        =   11
                        Text            =   "N/A"
                        Top             =   810
                        Width           =   1950
                     End
                     Begin XpressEditorsLibCtl.dxTextEdit dxEdtLat 
                        DataField       =   "latX"
                        DataSource      =   "AdodcWhere"
                        Height          =   315
                        Left            =   1035
                        OleObjectBlob   =   "ctrWhere.ctx":282C
                        TabIndex        =   36
                        Top             =   180
                        Width           =   1950
                     End
                     Begin XpressEditorsLibCtl.dxTextEdit dxedtLong 
                        DataField       =   "longY"
                        DataSource      =   "AdodcWhere"
                        Height          =   315
                        Left            =   1035
                        OleObjectBlob   =   "ctrWhere.ctx":2884
                        TabIndex        =   37
                        Top             =   495
                        Width           =   1950
                     End
                     Begin VB.Label lblAdminlevel 
                        BeginProperty Font 
                           Name            =   "MS Serif"
                           Size            =   6.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   1575
                        Left            =   3645
                        TabIndex        =   25
                        Top             =   135
                        Width           =   2925
                     End
                     Begin VB.Label lblMGRS 
                        AutoSize        =   -1  'True
                        Caption         =   "MGRS:"
                        Height          =   195
                        Left            =   0
                        TabIndex        =   17
                        Top             =   900
                        Width           =   525
                     End
                     Begin VB.Label lblLongitudeX 
                        AutoSize        =   -1  'True
                        Caption         =   "Longitude X:"
                        Height          =   195
                        Left            =   0
                        TabIndex        =   16
                        Top             =   540
                        Width           =   900
                     End
                     Begin VB.Label lblLatitudeX 
                        AutoSize        =   -1  'True
                        Caption         =   "Latitude Y:"
                        Height          =   195
                        Left            =   0
                        TabIndex        =   15
                        Top             =   180
                        Width           =   765
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic elAdmin 
                     Height          =   1875
                     Left            =   45
                     TabIndex        =   18
                     TabStop         =   0   'False
                     Top             =   330
                     Width           =   11295
                     _cx             =   19923
                     _cy             =   3307
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
                     Begin VB.Frame FraPcodes 
                        Caption         =   "Pcodes"
                        Height          =   1725
                        Left            =   90
                        TabIndex        =   53
                        Top             =   -45
                        Width           =   5370
                        Begin VB.Frame FraSource 
                           Caption         =   "Source:"
                           BeginProperty Font 
                              Name            =   "MS Serif"
                              Size            =   6.75
                              Charset         =   0
                              Weight          =   700
                              Underline       =   0   'False
                              Italic          =   0   'False
                              Strikethrough   =   0   'False
                           EndProperty
                           Height          =   600
                           Left            =   45
                           TabIndex        =   59
                           Top             =   180
                           Width           =   1410
                           Begin VB.OptionButton OptDatasource 
                              Caption         =   "LIS"
                              BeginProperty Font 
                                 Name            =   "MS Serif"
                                 Size            =   6.75
                                 Charset         =   0
                                 Weight          =   700
                                 Underline       =   0   'False
                                 Italic          =   0   'False
                                 Strikethrough   =   0   'False
                              EndProperty
                              Height          =   195
                              Index           =   1
                              Left            =   90
                              TabIndex        =   61
                              Top             =   360
                              Width           =   960
                           End
                           Begin VB.OptionButton OptDatasource 
                              Caption         =   "OCHA/HIC"
                              BeginProperty Font 
                                 Name            =   "MS Serif"
                                 Size            =   6.75
                                 Charset         =   0
                                 Weight          =   700
                                 Underline       =   0   'False
                                 Italic          =   0   'False
                                 Strikethrough   =   0   'False
                              EndProperty
                              Height          =   195
                              Index           =   0
                              Left            =   90
                              TabIndex        =   60
                              Top             =   180
                              Value           =   -1  'True
                              Width           =   1275
                           End
                        End
                        Begin VB.OptionButton OptSearchType 
                           Caption         =   "Village"
                           Height          =   195
                           Index           =   1
                           Left            =   1575
                           TabIndex        =   58
                           Top             =   495
                           Width           =   780
                        End
                        Begin VB.OptionButton OptSearchType 
                           Caption         =   "PCode"
                           Height          =   195
                           Index           =   0
                           Left            =   1575
                           TabIndex        =   57
                           Top             =   225
                           Value           =   -1  'True
                           Width           =   780
                        End
                        Begin VB.ListBox lstResult 
                           Height          =   840
                           ItemData        =   "ctrWhere.ctx":28FA
                           Left            =   720
                           List            =   "ctrWhere.ctx":28FC
                           TabIndex        =   56
                           Top             =   810
                           Width           =   4425
                        End
                        Begin VB.CommandButton cmdSearch 
                           Height          =   375
                           Left            =   4680
                           MaskColor       =   &H00FFFFFF&
                           Picture         =   "ctrWhere.ctx":28FE
                           Style           =   1  'Graphical
                           TabIndex        =   55
                           ToolTipText     =   "Search"
                           Top             =   270
                           UseMaskColor    =   -1  'True
                           Width           =   510
                        End
                        Begin VB.TextBox txtPcode 
                           Height          =   330
                           Left            =   2430
                           TabIndex        =   54
                           Top             =   315
                           Width           =   2085
                        End
                        Begin VB.Label lblRecords 
                           AutoSize        =   -1  'True
                           Caption         =   "records"
                           Height          =   195
                           Left            =   45
                           TabIndex        =   63
                           Top             =   1305
                           Width           =   525
                        End
                        Begin VB.Label lblResult 
                           AutoSize        =   -1  'True
                           Caption         =   "Result:"
                           Height          =   195
                           Left            =   90
                           TabIndex        =   62
                           Top             =   810
                           Width           =   495
                        End
                     End
                     Begin VB.ComboBox ComProvince 
                        Height          =   315
                        Left            =   5535
                        Style           =   2  'Dropdown List
                        TabIndex        =   21
                        Top             =   315
                        Width           =   2265
                     End
                     Begin VB.ComboBox ComDistrict 
                        Enabled         =   0   'False
                        Height          =   315
                        Left            =   5535
                        Style           =   2  'Dropdown List
                        TabIndex        =   20
                        Top             =   810
                        Width           =   2265
                     End
                     Begin VB.ComboBox ComPlace 
                        Enabled         =   0   'False
                        Height          =   315
                        Left            =   5535
                        Style           =   2  'Dropdown List
                        TabIndex        =   19
                        Top             =   1305
                        Width           =   2265
                     End
                     Begin VB.Label lblPlaceCode 
                        AutoSize        =   -1  'True
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   195
                        Left            =   7875
                        TabIndex        =   40
                        Top             =   1440
                        Width           =   75
                     End
                     Begin VB.Label lblDistrictPCode 
                        AutoSize        =   -1  'True
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   195
                        Left            =   7875
                        TabIndex        =   39
                        Top             =   900
                        Width           =   75
                     End
                     Begin VB.Label lblProvincePcode 
                        AutoSize        =   -1  'True
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   8.25
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   195
                        Left            =   7875
                        TabIndex        =   38
                        Top             =   360
                        Width           =   75
                     End
                     Begin VB.Label lblDistrict 
                        AutoSize        =   -1  'True
                        Caption         =   "District:"
                        Height          =   195
                        Left            =   5535
                        TabIndex        =   23
                        Top             =   630
                        Width           =   525
                     End
                     Begin VB.Label lblCommunity 
                        AutoSize        =   -1  'True
                        Caption         =   "Place:"
                        Height          =   195
                        Left            =   5535
                        TabIndex        =   22
                        Top             =   1125
                        Width           =   450
                     End
                     Begin VB.Label lblProvince 
                        AutoSize        =   -1  'True
                        Caption         =   "Province:"
                        Height          =   195
                        Left            =   5535
                        TabIndex        =   24
                        Top             =   135
                        Width           =   675
                     End
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic v 
               Height          =   4140
               Left            =   90
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   90
               Width           =   11445
               _cx             =   20188
               _cy             =   7303
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
               ChildSpacing    =   1
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
               GridCols        =   2
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"ctrWhere.ctx":2DB1
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin TatukGIS_DK.XGIS_ViewerWnd GIS 
                  Height          =   3855
                  Left            =   15
                  TabIndex        =   5
                  Top             =   15
                  Width           =   11025
                  BigExtentMargin =   -10
                  RestrictedDrag  =   -1  'True
                  CachedPaint     =   -1  'True
                  IncrementalPaint=   -1  'True
                  FullPaint       =   -1  'True
                  CodePage        =   0
                  OutCodePage     =   0
                  CharSet         =   0
                  UseRTree        =   0   'False
                  PrinterTileSize =   512
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
                  SelectionPattern=   "ctrWhere.ctx":2DFB
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
               Begin ActiveBar3LibraryCtl.ActiveBar3 ActiveBar31 
                  Height          =   3855
                  Left            =   11055
                  TabIndex        =   6
                  Top             =   15
                  Width           =   375
                  _LayoutVersion  =   2
                  _ExtentX        =   661
                  _ExtentY        =   6800
                  _DataPath       =   ""
                  Bands           =   "ctrWhere.ctx":2E5D
               End
               Begin VB.Label lblCoords 
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "X: n/a Y: n/a"
                  BeginProperty Font 
                     Name            =   "MS Serif"
                     Size            =   6
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   15
                  TabIndex        =   7
                  Top             =   3885
                  Width           =   11025
               End
            End
         End
      End
   End
End
Attribute VB_Name = "ctrWhere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
