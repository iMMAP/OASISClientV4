VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5B57CA38-0CAB-48F9-BBAE-8A2342F4A7C2}#8.4#0"; "ttkGISDK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9309EA27-780F-4C3D-84E8-79DEB1CECE69}#5.0#0"; "ThemeRangePicker.ocx"
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS User Group Configuration"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9705
   Icon            =   "frmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   9705
   StartUpPosition =   2  'CenterScreen
   Begin OASISRemoteAdmin.Downloader Downloader1 
      Left            =   6810
      Top             =   2790
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin C1SizerLibCtl.C1Elastic scNavigator 
      Height          =   8400
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9705
      _cx             =   17119
      _cy             =   14817
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
      BorderWidth     =   3
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
      GridRows        =   2
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmConfig.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic111 
         Height          =   450
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7905
         Width           =   9570
         _cx             =   16880
         _cy             =   794
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
         BorderWidth     =   0
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
         Begin VB.CommandButton cmdTools 
            Caption         =   "Admin Tools"
            Height          =   375
            Left            =   6690
            TabIndex        =   366
            Top             =   60
            Width           =   1215
         End
         Begin VB.CommandButton cmdSaveClient 
            Caption         =   "Save"
            Enabled         =   0   'False
            Height          =   375
            Left            =   7980
            TabIndex        =   2
            Top             =   60
            Width           =   1215
         End
      End
      Begin C1SizerLibCtl.C1Tab ciTab 
         Height          =   7830
         Left            =   90
         TabIndex        =   3
         Top             =   45
         Width           =   9570
         _cx             =   16880
         _cy             =   13811
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
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Groups/Users|Geek Data View|Graphical View|Synchronisation|Data Definitions/Access|Add-Ons/Modules"
         Align           =   0
         CurrTab         =   2
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
         TabHeight       =   444
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Picture(0)      =   "frmConfig.frx":689F
         Flags(0)        =   2
         Picture(1)      =   "frmConfig.frx":D101
         Picture(2)      =   "frmConfig.frx":13963
         Picture(3)      =   "frmConfig.frx":1A1C5
         Flags(3)        =   2
         Picture(4)      =   "frmConfig.frx":20A27
         Flags(4)        =   2
         Picture(5)      =   "frmConfig.frx":27289
         Flags(5)        =   2
         Begin C1SizerLibCtl.C1Elastic elTab 
            Height          =   7350
            Index           =   0
            Left            =   -10455
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   465
            Width           =   9540
            _cx             =   16828
            _cy             =   12965
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
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmConfig.frx":2DAEB
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic elUG 
               Height          =   2565
               Left            =   30
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   30
               Width           =   9480
               _cx             =   16722
               _cy             =   4524
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
               BorderWidth     =   0
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
               GridCols        =   2
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmConfig.frx":2DB30
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.Frame FraGroupDetails 
                  Caption         =   "Group Details:"
                  Height          =   2565
                  Left            =   0
                  TabIndex        =   6
                  Top             =   0
                  Width           =   3240
                  Begin VB.CommandButton cmdUsrNew 
                     Caption         =   "New"
                     Enabled         =   0   'False
                     Height          =   285
                     Index           =   1
                     Left            =   3510
                     TabIndex        =   11
                     Top             =   930
                     Width           =   825
                  End
                  Begin VB.CommandButton cmdUsrEdit 
                     Caption         =   "Edit"
                     Enabled         =   0   'False
                     Height          =   285
                     Index           =   1
                     Left            =   3510
                     TabIndex        =   10
                     Top             =   1245
                     Width           =   825
                  End
                  Begin VB.CommandButton cmdUsrDelete 
                     Caption         =   "Delete"
                     Enabled         =   0   'False
                     Height          =   285
                     Index           =   1
                     Left            =   3510
                     TabIndex        =   9
                     Top             =   1545
                     Width           =   825
                  End
                  Begin VB.CommandButton cmdUpdateUser 
                     Caption         =   "Save"
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   3510
                     TabIndex        =   8
                     Top             =   1860
                     Width           =   825
                  End
                  Begin VB.CheckBox chkUpdateUsr 
                     Caption         =   "Auto Update User List"
                     Height          =   375
                     Left            =   2070
                     TabIndex        =   7
                     Top             =   120
                     Width           =   2205
                  End
                  Begin C1SizerLibCtl.C1Tab c1TabGroup 
                     Height          =   1950
                     Left            =   90
                     TabIndex        =   12
                     Top             =   225
                     Width           =   3390
                     _cx             =   5980
                     _cy             =   3440
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
                     Caption         =   "Group|Description|Advanced"
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
                     Begin C1SizerLibCtl.C1Elastic elGroupAdvanced 
                        Height          =   1575
                        Left            =   4335
                        TabIndex        =   13
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   3300
                        _cx             =   5821
                        _cy             =   2778
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
                        BorderWidth     =   2
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
                     End
                     Begin C1SizerLibCtl.C1Elastic elGroupDesc 
                        Height          =   1575
                        Left            =   4035
                        TabIndex        =   14
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   3300
                        _cx             =   5821
                        _cy             =   2778
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
                        BorderWidth     =   1
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
                        Begin VB.TextBox txtUGDesc 
                           DataField       =   "Description"
                           Height          =   1455
                           Left            =   90
                           MultiLine       =   -1  'True
                           ScrollBars      =   2  'Vertical
                           TabIndex        =   15
                           Top             =   45
                           Width           =   3120
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elGroup 
                        Height          =   1575
                        Left            =   45
                        TabIndex        =   16
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   3300
                        _cx             =   5821
                        _cy             =   2778
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
                        Begin VB.TextBox txtUGID 
                           DataField       =   "ID"
                           Height          =   285
                           Left            =   135
                           Locked          =   -1  'True
                           TabIndex        =   19
                           Top             =   1215
                           Width           =   3030
                        End
                        Begin VB.TextBox txtUGTablePrefix 
                           DataField       =   "SettingTablePrefix"
                           Height          =   285
                           Left            =   135
                           TabIndex        =   18
                           Top             =   720
                           Width           =   3030
                        End
                        Begin VB.TextBox txtUGName 
                           DataField       =   "Name"
                           Height          =   285
                           Left            =   135
                           TabIndex        =   17
                           Top             =   225
                           Width           =   3030
                        End
                        Begin VB.Label lblUG 
                           AutoSize        =   -1  'True
                           Caption         =   "Group ID:"
                           Height          =   195
                           Index           =   2
                           Left            =   135
                           TabIndex        =   22
                           Top             =   990
                           Width           =   690
                        End
                        Begin VB.Label lblUG 
                           AutoSize        =   -1  'True
                           Caption         =   "Group Prefix:"
                           Height          =   195
                           Index           =   1
                           Left            =   135
                           TabIndex        =   21
                           Top             =   495
                           Width           =   915
                        End
                        Begin VB.Label lblUG 
                           AutoSize        =   -1  'True
                           Caption         =   "Name:"
                           Height          =   195
                           Index           =   0
                           Left            =   135
                           TabIndex        =   20
                           Top             =   0
                           Width           =   465
                        End
                     End
                  End
               End
               Begin DXDBGRIDLibCtl.dxDBGrid dxUG 
                  Height          =   2565
                  Left            =   3240
                  OleObjectBlob   =   "frmConfig.frx":2DB6B
                  TabIndex        =   23
                  Top             =   0
                  Width           =   5925
               End
            End
            Begin C1SizerLibCtl.C1Elastic elUsers 
               Height          =   4710
               Left            =   30
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   2610
               Width           =   9480
               _cx             =   16722
               _cy             =   8308
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
               AutoSizeChildren=   8
               BorderWidth     =   0
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
               GridCols        =   2
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmConfig.frx":2E813
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.Frame FraUsers 
                  Caption         =   "User Details:"
                  Height          =   4770
                  Index           =   1
                  Left            =   0
                  TabIndex        =   25
                  Top             =   0
                  Width           =   3270
                  Begin VB.CommandButton cmdGetData 
                     Caption         =   "Get users"
                     Height          =   285
                     Left            =   3555
                     TabIndex        =   30
                     Top             =   2640
                     Width           =   810
                  End
                  Begin VB.CommandButton CmdSave 
                     Caption         =   "Save"
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   3555
                     TabIndex        =   29
                     Top             =   2340
                     Width           =   810
                  End
                  Begin VB.CommandButton cmdUsrNew 
                     Caption         =   "New"
                     Enabled         =   0   'False
                     Height          =   285
                     Index           =   0
                     Left            =   3555
                     TabIndex        =   28
                     Top             =   1425
                     Width           =   810
                  End
                  Begin VB.CommandButton cmdUsrEdit 
                     Caption         =   "Edit"
                     Enabled         =   0   'False
                     Height          =   285
                     Index           =   0
                     Left            =   3555
                     TabIndex        =   27
                     Top             =   1730
                     Width           =   810
                  End
                  Begin VB.CommandButton cmdUsrDelete 
                     Caption         =   "Delete"
                     Enabled         =   0   'False
                     Height          =   285
                     Index           =   0
                     Left            =   3555
                     TabIndex        =   26
                     Top             =   2035
                     Width           =   810
                  End
                  Begin C1SizerLibCtl.C1Tab c1UserTab 
                     Height          =   2760
                     Left            =   135
                     TabIndex        =   31
                     Top             =   225
                     Width           =   3390
                     _cx             =   5980
                     _cy             =   4868
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
                     Caption         =   "User|Security|Advanced"
                     Align           =   0
                     CurrTab         =   2
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
                     Begin C1SizerLibCtl.C1Elastic elusrAdv 
                        Height          =   2385
                        Left            =   45
                        TabIndex        =   32
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   3300
                        _cx             =   5821
                        _cy             =   4207
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
                        Begin VB.TextBox txtUsrAdvSettingUrl 
                           DataField       =   "SettingUrl"
                           Height          =   285
                           Left            =   90
                           TabIndex        =   33
                           Top             =   360
                           Width           =   3120
                        End
                        Begin MSComctlLib.ListView lvUGAccess 
                           Height          =   1320
                           Left            =   90
                           TabIndex        =   34
                           Top             =   945
                           Width           =   3075
                           _ExtentX        =   5424
                           _ExtentY        =   2328
                           View            =   3
                           LabelEdit       =   1
                           Sorted          =   -1  'True
                           MultiSelect     =   -1  'True
                           LabelWrap       =   -1  'True
                           HideSelection   =   -1  'True
                           Checkboxes      =   -1  'True
                           _Version        =   393217
                           ForeColor       =   -2147483640
                           BackColor       =   -2147483643
                           BorderStyle     =   1
                           Appearance      =   1
                           NumItems        =   1
                           BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                              Text            =   "Availble User Groups"
                              Object.Width           =   5292
                           EndProperty
                        End
                        Begin VB.Label lblApplicationServer 
                           AutoSize        =   -1  'True
                           Caption         =   "Application Server URL:"
                           Height          =   195
                           Left            =   135
                           TabIndex        =   36
                           Top             =   90
                           Width           =   1710
                        End
                        Begin VB.Label lblUserGroupdataAccess 
                           AutoSize        =   -1  'True
                           Caption         =   "User Group Data Access:"
                           Height          =   195
                           Left            =   90
                           TabIndex        =   35
                           Top             =   720
                           Width           =   1815
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elusrSec 
                        Height          =   2385
                        Left            =   -3945
                        TabIndex        =   37
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   3300
                        _cx             =   5821
                        _cy             =   4207
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
                        Begin VB.Frame FraResetPassword 
                           Caption         =   "Reset Password"
                           Height          =   1770
                           Left            =   90
                           TabIndex        =   39
                           Top             =   585
                           Width           =   3165
                           Begin VB.CommandButton cmdusrPwdSubmit 
                              Caption         =   "Submit"
                              Height          =   240
                              Left            =   1845
                              TabIndex        =   42
                              Top             =   1485
                              Width           =   1140
                           End
                           Begin VB.TextBox txtUsrSecNewPassConfirm 
                              Height          =   330
                              Left            =   135
                              TabIndex        =   41
                              Top             =   1125
                              Width           =   2850
                           End
                           Begin VB.TextBox txtUsrSecNewPass 
                              Height          =   330
                              Left            =   135
                              TabIndex        =   40
                              Top             =   495
                              Width           =   2850
                           End
                           Begin VB.Label lblConfirmPassword 
                              Caption         =   "Confirm Password:"
                              Height          =   240
                              Left            =   135
                              TabIndex        =   44
                              Top             =   900
                              Width           =   1410
                           End
                           Begin VB.Label lblNewPassword 
                              Caption         =   "New Password:"
                              Height          =   195
                              Left            =   135
                              TabIndex        =   43
                              Top             =   270
                              Width           =   1455
                           End
                        End
                        Begin VB.TextBox txtUsrSecPassWrd 
                           DataField       =   "pwd"
                           Height          =   285
                           Left            =   90
                           TabIndex        =   38
                           Top             =   270
                           Width           =   3120
                        End
                        Begin VB.Label lblPassword 
                           Caption         =   "Password"
                           Height          =   285
                           Left            =   90
                           TabIndex        =   45
                           Top             =   45
                           Width           =   1725
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elUsr 
                        Height          =   2385
                        Left            =   -4245
                        TabIndex        =   46
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   3300
                        _cx             =   5821
                        _cy             =   4207
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
                        Begin VB.ComboBox comUserGroup 
                           Height          =   315
                           Left            =   135
                           Style           =   2  'Dropdown List
                           TabIndex        =   51
                           Top             =   2025
                           Width           =   3030
                        End
                        Begin VB.TextBox txtUsrLastName 
                           DataField       =   "Lname"
                           Height          =   315
                           Left            =   135
                           TabIndex        =   50
                           Top             =   1440
                           Width           =   3030
                        End
                        Begin VB.TextBox txtUsrFirstName 
                           DataField       =   "Fname"
                           Height          =   315
                           Left            =   135
                           TabIndex        =   49
                           Top             =   855
                           Width           =   3030
                        End
                        Begin VB.TextBox txtUsrName 
                           DataField       =   "user"
                           Height          =   315
                           Left            =   135
                           TabIndex        =   48
                           Top             =   270
                           Width           =   3030
                        End
                        Begin VB.TextBox txtusrUGID 
                           DataField       =   "UserGroupID"
                           Height          =   285
                           Left            =   2610
                           TabIndex        =   47
                           Top             =   1755
                           Visible         =   0   'False
                           Width           =   555
                        End
                        Begin VB.Label lblusr 
                           AutoSize        =   -1  'True
                           Caption         =   "User Name:"
                           Height          =   195
                           Index           =   0
                           Left            =   180
                           TabIndex        =   55
                           Top             =   45
                           Width           =   840
                        End
                        Begin VB.Label lblusr 
                           AutoSize        =   -1  'True
                           Caption         =   "First Name:"
                           Height          =   195
                           Index           =   1
                           Left            =   135
                           TabIndex        =   54
                           Top             =   630
                           Width           =   795
                        End
                        Begin VB.Label lblusr 
                           AutoSize        =   -1  'True
                           Caption         =   "Last Name:"
                           Height          =   195
                           Index           =   2
                           Left            =   135
                           TabIndex        =   53
                           Top             =   1215
                           Width           =   810
                        End
                        Begin VB.Label lblusr 
                           AutoSize        =   -1  'True
                           Caption         =   "User Group:"
                           Height          =   195
                           Index           =   3
                           Left            =   135
                           TabIndex        =   52
                           Top             =   1800
                           Width           =   855
                        End
                     End
                  End
               End
               Begin DXDBGRIDLibCtl.dxDBGrid dxUsers 
                  Height          =   4770
                  Left            =   3270
                  OleObjectBlob   =   "frmConfig.frx":2E853
                  TabIndex        =   56
                  Top             =   0
                  Width           =   5895
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elTab 
            Height          =   7350
            Index           =   1
            Left            =   -10155
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   465
            Width           =   9540
            _cx             =   16828
            _cy             =   12965
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
            BorderWidth     =   3
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
            GridRows        =   5
            GridCols        =   5
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmConfig.frx":2F4FB
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.TextBox TxtSql 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   7260
               Left            =   45
               TabIndex        =   58
               Text            =   "Select * from Users"
               Top             =   45
               Visible         =   0   'False
               Width           =   9450
            End
            Begin DXDBGRIDLibCtl.dxDBGrid dxAppSettings 
               Height          =   7260
               Left            =   45
               OleObjectBlob   =   "frmConfig.frx":2F58F
               TabIndex        =   59
               Top             =   45
               Width           =   9450
            End
         End
         Begin C1SizerLibCtl.C1Elastic elTab 
            Height          =   7350
            Index           =   2
            Left            =   15
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   465
            Width           =   9540
            _cx             =   16828
            _cy             =   12965
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
            BorderWidth     =   3
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
            GridRows        =   5
            GridCols        =   5
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmConfig.frx":34688
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic eltop 
               Height          =   2250
               Left            =   45
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   45
               Width           =   9450
               _cx             =   16669
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
               BorderWidth     =   0
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
               Begin VB.Frame FraGeneralSettings 
                  Caption         =   "General Settings:"
                  Height          =   2235
                  Left            =   0
                  TabIndex        =   62
                  Top             =   0
                  Width           =   9030
                  Begin C1SizerLibCtl.C1Tab C1TTab1Tab2 
                     Height          =   1875
                     Left            =   120
                     TabIndex        =   63
                     Top             =   240
                     Width           =   8790
                     _cx             =   15505
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
                     Appearance      =   2
                     MousePointer    =   0
                     Version         =   801
                     BackColor       =   -2147483633
                     ForeColor       =   -2147483630
                     FrontTabColor   =   -2147483633
                     BackTabColor    =   -2147483633
                     TabOutlineColor =   -2147483632
                     FrontTabForeColor=   -2147483630
                     Caption         =   "General|Synchronisation|Administrative Location|Hot keys|Misc"
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
                     Begin VB.Frame Frame2 
                        BorderStyle     =   0  'None
                        Height          =   1500
                        Left            =   45
                        TabIndex        =   400
                        Top             =   330
                        Width           =   8700
                        Begin VB.TextBox txtServerConnRetries 
                           Height          =   270
                           Left            =   6870
                           TabIndex        =   448
                           Top             =   990
                           Width           =   1290
                        End
                        Begin VB.TextBox txtServerConnTimeout 
                           Height          =   270
                           Left            =   6870
                           TabIndex        =   447
                           Top             =   660
                           Width           =   1290
                        End
                        Begin VB.TextBox txtAppSetting 
                           Height          =   270
                           Index           =   0
                           Left            =   5400
                           TabIndex        =   409
                           Top             =   150
                           Width           =   2760
                        End
                        Begin VB.Frame FraActiveCore 
                           Caption         =   "Active Core Functions"
                           Height          =   1500
                           Index           =   0
                           Left            =   0
                           TabIndex        =   401
                           Top             =   0
                           Width           =   3780
                           Begin VB.CheckBox chkCoreModule 
                              Caption         =   "Dynamic Reports"
                              Height          =   285
                              Index           =   7
                              Left            =   90
                              TabIndex        =   451
                              Top             =   1160
                              Width           =   1575
                           End
                           Begin VB.CheckBox chkCoreModule 
                              Caption         =   "Dynamic Data"
                              Height          =   285
                              Index           =   6
                              Left            =   90
                              TabIndex        =   450
                              Top             =   920
                              Width           =   1695
                           End
                           Begin VB.Frame FraStartModule 
                              Caption         =   "Start Module"
                              Height          =   690
                              Left            =   1740
                              TabIndex        =   407
                              Top             =   210
                              Width           =   1875
                              Begin VB.ComboBox ComStartModule 
                                 Height          =   315
                                 Left            =   150
                                 Style           =   2  'Dropdown List
                                 TabIndex        =   408
                                 Top             =   240
                                 Width           =   1560
                              End
                           End
                           Begin VB.CheckBox chkCoreModule 
                              Caption         =   "Addons"
                              Enabled         =   0   'False
                              Height          =   285
                              Index           =   4
                              Left            =   2400
                              TabIndex        =   406
                              Top             =   960
                              Visible         =   0   'False
                              Width           =   870
                           End
                           Begin VB.CheckBox chkCoreModule 
                              Caption         =   "RSS Feeds"
                              Height          =   285
                              Index           =   2
                              Left            =   90
                              TabIndex        =   405
                              Top             =   690
                              Width           =   1650
                           End
                           Begin VB.CheckBox chkCoreModule 
                              Caption         =   "Synchronize"
                              Enabled         =   0   'False
                              Height          =   285
                              Index           =   3
                              Left            =   2400
                              TabIndex        =   404
                              Top             =   1080
                              Visible         =   0   'False
                              Width           =   1215
                           End
                           Begin VB.CheckBox chkCoreModule 
                              Caption         =   "Oasis Profile"
                              Height          =   285
                              Index           =   0
                              Left            =   90
                              TabIndex        =   403
                              Top             =   210
                              Width           =   1215
                           End
                           Begin VB.CheckBox chkCoreModule 
                              Caption         =   "Operations"
                              Height          =   285
                              Index           =   1
                              Left            =   90
                              TabIndex        =   402
                              Top             =   450
                              Width           =   1215
                           End
                        End
                        Begin VB.Label Label2 
                           Caption         =   "Server connection retries:"
                           Height          =   255
                           Index           =   1
                           Left            =   4050
                           TabIndex        =   446
                           Top             =   1020
                           Width           =   2865
                        End
                        Begin VB.Label Label2 
                           Caption         =   "Server connection timeout (seconds):"
                           Height          =   255
                           Index           =   0
                           Left            =   4050
                           TabIndex        =   445
                           Top             =   690
                           Width           =   2805
                        End
                        Begin VB.Label lblUG 
                           AutoSize        =   -1  'True
                           Caption         =   "Client Language:"
                           Height          =   270
                           Index           =   3
                           Left            =   4020
                           TabIndex        =   410
                           Top             =   150
                           Width           =   1350
                        End
                     End
                     Begin VB.Frame Frame1 
                        BorderStyle     =   0  'None
                        Height          =   1500
                        Left            =   10335
                        TabIndex        =   373
                        Top             =   330
                        Width           =   8700
                        Begin VB.CheckBox chkShowMATab 
                           Caption         =   "Mine Action Module"
                           Height          =   195
                           Left            =   120
                           TabIndex        =   399
                           Top             =   1170
                           Width           =   2010
                        End
                        Begin VB.CheckBox chkShowCommodityTab 
                           Caption         =   "Commodity/NFI Module"
                           Height          =   195
                           Left            =   120
                           TabIndex        =   398
                           Top             =   840
                           Width           =   1980
                        End
                        Begin VB.CheckBox chkCoreModule 
                           Caption         =   "W3 Module"
                           Height          =   285
                           Index           =   5
                           Left            =   120
                           TabIndex        =   397
                           Top             =   510
                           Width           =   1320
                        End
                        Begin VB.Frame FraLocator 
                           Caption         =   "Locator Settings        (Layer Name)                               (Attribute Name)"
                           Height          =   1305
                           Left            =   2310
                           TabIndex        =   375
                           Top             =   120
                           Width           =   6195
                           Begin VB.TextBox txtLocatorAttName 
                              Height          =   285
                              Index           =   3
                              Left            =   3960
                              TabIndex        =   424
                              Text            =   "Level"
                              Top             =   960
                              Width           =   2145
                           End
                           Begin VB.TextBox txtLocatorAttName 
                              Height          =   285
                              Index           =   2
                              Left            =   3960
                              TabIndex        =   421
                              Text            =   "Level"
                              Top             =   720
                              Width           =   2145
                           End
                           Begin VB.TextBox txtLocatorAttName 
                              Height          =   285
                              Index           =   1
                              Left            =   3960
                              TabIndex        =   422
                              Text            =   "Level"
                              Top             =   480
                              Width           =   2145
                           End
                           Begin VB.TextBox txtLocatorAttName 
                              Height          =   255
                              Index           =   0
                              Left            =   3960
                              TabIndex        =   423
                              Text            =   "Level"
                              Top             =   240
                              Width           =   2145
                           End
                           Begin VB.TextBox txtLocatorLayerName 
                              Height          =   285
                              Index           =   3
                              Left            =   1680
                              TabIndex        =   416
                              Text            =   "Level"
                              Top             =   960
                              Width           =   2265
                           End
                           Begin VB.TextBox txtLocatorLayerName 
                              Height          =   285
                              Index           =   2
                              Left            =   1680
                              TabIndex        =   378
                              Text            =   "Level"
                              Top             =   720
                              Width           =   2265
                           End
                           Begin VB.TextBox txtLocatorLayerName 
                              Height          =   285
                              Index           =   1
                              Left            =   1680
                              TabIndex        =   377
                              Text            =   "Level"
                              Top             =   480
                              Width           =   2265
                           End
                           Begin VB.TextBox txtLocatorLayerName 
                              Height          =   255
                              Index           =   0
                              Left            =   1680
                              TabIndex        =   376
                              Text            =   "Level"
                              Top             =   240
                              Width           =   2265
                           End
                           Begin VB.Label Label1 
                              Caption         =   "Pcode Layer"
                              Height          =   255
                              Index           =   3
                              Left            =   120
                              TabIndex        =   420
                              Top             =   960
                              Width           =   1065
                           End
                           Begin VB.Label Label1 
                              Caption         =   "Adm Level 0"
                              Height          =   255
                              Index           =   2
                              Left            =   120
                              TabIndex        =   419
                              Top             =   255
                              Width           =   1575
                           End
                           Begin VB.Label Label1 
                              Caption         =   "Adm Location"
                              Height          =   255
                              Index           =   1
                              Left            =   120
                              TabIndex        =   418
                              Top             =   720
                              Width           =   1695
                           End
                           Begin VB.Label Label1 
                              Caption         =   "Adm Level 1"
                              Height          =   255
                              Index           =   0
                              Left            =   120
                              TabIndex        =   417
                              Top             =   480
                              Width           =   1455
                           End
                        End
                        Begin VB.CheckBox chkPromptSave 
                           Caption         =   "Prompt Save On Exit"
                           Height          =   255
                           Left            =   120
                           TabIndex        =   374
                           Top             =   210
                           Width           =   1965
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elAdmBoundaries 
                        Height          =   1500
                        Left            =   9735
                        TabIndex        =   64
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   8700
                        _cx             =   15346
                        _cy             =   2646
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
                        Begin VB.TextBox txtAdmLevel 
                           Height          =   285
                           Index           =   0
                           Left            =   720
                           TabIndex        =   84
                           Top             =   45
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmLevel 
                           Height          =   285
                           Index           =   1
                           Left            =   720
                           TabIndex        =   83
                           Top             =   333
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmLevel 
                           Height          =   285
                           Index           =   2
                           Left            =   720
                           TabIndex        =   82
                           Top             =   621
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmLevel 
                           Height          =   285
                           Index           =   3
                           Left            =   720
                           TabIndex        =   81
                           Top             =   909
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmLevel 
                           Height          =   285
                           Index           =   4
                           Left            =   720
                           TabIndex        =   80
                           Top             =   1200
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmField 
                           Height          =   285
                           Index           =   0
                           Left            =   2865
                           TabIndex        =   79
                           Top             =   45
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmPCode 
                           Height          =   285
                           Index           =   0
                           Left            =   5250
                           TabIndex        =   78
                           Top             =   45
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmDisplayName 
                           Height          =   285
                           Index           =   0
                           Left            =   7140
                           TabIndex        =   77
                           Top             =   45
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmField 
                           Height          =   285
                           Index           =   1
                           Left            =   2865
                           TabIndex        =   76
                           Top             =   333
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmPCode 
                           Height          =   285
                           Index           =   1
                           Left            =   5250
                           TabIndex        =   75
                           Top             =   333
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmDisplayName 
                           Height          =   285
                           Index           =   1
                           Left            =   7140
                           TabIndex        =   74
                           Top             =   333
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmField 
                           Height          =   285
                           Index           =   2
                           Left            =   2865
                           TabIndex        =   73
                           Top             =   621
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmPCode 
                           Height          =   285
                           Index           =   2
                           Left            =   5250
                           TabIndex        =   72
                           Top             =   621
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmDisplayName 
                           Height          =   285
                           Index           =   2
                           Left            =   7140
                           TabIndex        =   71
                           Top             =   621
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmField 
                           Height          =   285
                           Index           =   3
                           Left            =   2865
                           TabIndex        =   70
                           Top             =   909
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmPCode 
                           Height          =   285
                           Index           =   3
                           Left            =   5250
                           TabIndex        =   69
                           Top             =   909
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmDisplayName 
                           Height          =   285
                           Index           =   3
                           Left            =   7140
                           TabIndex        =   68
                           Top             =   909
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmField 
                           Height          =   285
                           Index           =   4
                           Left            =   2865
                           TabIndex        =   67
                           Top             =   1200
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmPCode 
                           Height          =   285
                           Index           =   4
                           Left            =   5250
                           TabIndex        =   66
                           Top             =   1200
                           Width           =   1410
                        End
                        Begin VB.TextBox txtAdmDisplayName 
                           Height          =   285
                           Index           =   4
                           Left            =   7140
                           TabIndex        =   65
                           Top             =   1200
                           Width           =   1410
                        End
                        Begin VB.Label lblAdmLevel 
                           AutoSize        =   -1  'True
                           Caption         =   "Level 1:"
                           Height          =   195
                           Index           =   0
                           Left            =   45
                           TabIndex        =   104
                           Top             =   90
                           Width           =   570
                        End
                        Begin VB.Label lblAdmLevel 
                           AutoSize        =   -1  'True
                           Caption         =   "Level 2:"
                           Height          =   195
                           Index           =   1
                           Left            =   30
                           TabIndex        =   103
                           Top             =   360
                           Width           =   570
                        End
                        Begin VB.Label lblAdmLevel 
                           AutoSize        =   -1  'True
                           Caption         =   "Level 3:"
                           Height          =   195
                           Index           =   2
                           Left            =   45
                           TabIndex        =   102
                           Top             =   660
                           Width           =   570
                        End
                        Begin VB.Label lblAdmLevel 
                           AutoSize        =   -1  'True
                           Caption         =   "Level 4:"
                           Height          =   195
                           Index           =   3
                           Left            =   45
                           TabIndex        =   101
                           Top             =   960
                           Width           =   570
                        End
                        Begin VB.Label lblAdmLevel 
                           AutoSize        =   -1  'True
                           Caption         =   "Location:"
                           Height          =   195
                           Index           =   4
                           Left            =   45
                           TabIndex        =   100
                           Top             =   1260
                           Width           =   660
                        End
                        Begin VB.Label lblAdmField 
                           AutoSize        =   -1  'True
                           Caption         =   "Fld Name:"
                           Height          =   195
                           Index           =   0
                           Left            =   2145
                           TabIndex        =   99
                           Top             =   90
                           Width           =   720
                        End
                        Begin VB.Label lblAdmCode 
                           AutoSize        =   -1  'True
                           Caption         =   "Adm Code:"
                           Height          =   195
                           Index           =   0
                           Left            =   4395
                           TabIndex        =   98
                           Top             =   120
                           Width           =   780
                        End
                        Begin VB.Label lblAdmAlias 
                           AutoSize        =   -1  'True
                           Caption         =   "Alias:"
                           Height          =   195
                           Index           =   0
                           Left            =   6735
                           TabIndex        =   97
                           Top             =   135
                           Width           =   375
                        End
                        Begin VB.Label lblAdmField 
                           AutoSize        =   -1  'True
                           Caption         =   "Fld Name:"
                           Height          =   195
                           Index           =   1
                           Left            =   2145
                           TabIndex        =   96
                           Top             =   390
                           Width           =   720
                        End
                        Begin VB.Label lblAdmCode 
                           AutoSize        =   -1  'True
                           Caption         =   "Adm Code:"
                           Height          =   195
                           Index           =   1
                           Left            =   4395
                           TabIndex        =   95
                           Top             =   405
                           Width           =   780
                        End
                        Begin VB.Label lblAdmAlias 
                           AutoSize        =   -1  'True
                           Caption         =   "Alias:"
                           Height          =   195
                           Index           =   1
                           Left            =   6735
                           TabIndex        =   94
                           Top             =   405
                           Width           =   375
                        End
                        Begin VB.Label lblAdmField 
                           AutoSize        =   -1  'True
                           Caption         =   "Fld Name:"
                           Height          =   195
                           Index           =   2
                           Left            =   2145
                           TabIndex        =   93
                           Top             =   660
                           Width           =   720
                        End
                        Begin VB.Label lblAdmCode 
                           AutoSize        =   -1  'True
                           Caption         =   "Adm Code:"
                           Height          =   195
                           Index           =   2
                           Left            =   4395
                           TabIndex        =   92
                           Top             =   660
                           Width           =   780
                        End
                        Begin VB.Label lblAdmAlias 
                           AutoSize        =   -1  'True
                           Caption         =   "Alias:"
                           Height          =   195
                           Index           =   2
                           Left            =   6735
                           TabIndex        =   91
                           Top             =   660
                           Width           =   375
                        End
                        Begin VB.Label lblAdmField 
                           AutoSize        =   -1  'True
                           Caption         =   "Fld Name:"
                           Height          =   195
                           Index           =   3
                           Left            =   2145
                           TabIndex        =   90
                           Top             =   960
                           Width           =   720
                        End
                        Begin VB.Label lblAdmCode 
                           AutoSize        =   -1  'True
                           Caption         =   "Adm Code:"
                           Height          =   195
                           Index           =   3
                           Left            =   4395
                           TabIndex        =   89
                           Top             =   960
                           Width           =   780
                        End
                        Begin VB.Label lblAdmAlias 
                           AutoSize        =   -1  'True
                           Caption         =   "Alias:"
                           Height          =   195
                           Index           =   3
                           Left            =   6735
                           TabIndex        =   88
                           Top             =   960
                           Width           =   375
                        End
                        Begin VB.Label lblAdmField 
                           AutoSize        =   -1  'True
                           Caption         =   "Fld Name:"
                           Height          =   195
                           Index           =   4
                           Left            =   2145
                           TabIndex        =   87
                           Top             =   1260
                           Width           =   720
                        End
                        Begin VB.Label lblAdmCode 
                           AutoSize        =   -1  'True
                           Caption         =   "Adm Code:"
                           Height          =   195
                           Index           =   4
                           Left            =   4395
                           TabIndex        =   86
                           Top             =   1290
                           Width           =   780
                        End
                        Begin VB.Label lblAdmAlias 
                           AutoSize        =   -1  'True
                           Caption         =   "Alias:"
                           Height          =   195
                           Index           =   4
                           Left            =   6735
                           TabIndex        =   85
                           Top             =   1260
                           Width           =   375
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elGenSynch 
                        Height          =   1500
                        Index           =   0
                        Left            =   9435
                        TabIndex        =   105
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   8700
                        _cx             =   15346
                        _cy             =   2646
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
                        Begin VB.Frame fraColor 
                           BackColor       =   &H00C0FFFF&
                           BorderStyle     =   0  'None
                           Height          =   285
                           Index           =   1
                           Left            =   1560
                           TabIndex        =   443
                           Top             =   1110
                           Width           =   735
                        End
                        Begin VB.Frame fraColor 
                           BackColor       =   &H000000FF&
                           BorderStyle     =   0  'None
                           Height          =   285
                           Index           =   0
                           Left            =   1560
                           TabIndex        =   442
                           Top             =   780
                           Width           =   735
                        End
                        Begin VB.TextBox txtAttachmentURL 
                           Height          =   285
                           Left            =   3240
                           TabIndex        =   108
                           Top             =   270
                           Width           =   5460
                        End
                        Begin VB.TextBox txtSyncCheckIntervall 
                           Height          =   285
                           Left            =   720
                           TabIndex        =   107
                           Top             =   330
                           Width           =   1185
                        End
                        Begin VB.CheckBox chkEnableInternet 
                           Caption         =   "Enable Internet Sync"
                           Height          =   330
                           Left            =   45
                           TabIndex        =   106
                           Top             =   0
                           Width           =   1995
                        End
                        Begin VB.Label lblIntervall 
                           AutoSize        =   -1  'True
                           Caption         =   "Notifier Fore Color:"
                           Height          =   195
                           Index           =   2
                           Left            =   120
                           TabIndex        =   441
                           Top             =   1170
                           Width           =   1305
                        End
                        Begin VB.Label lblIntervall 
                           AutoSize        =   -1  'True
                           Caption         =   "Notifier Back Color:"
                           Height          =   195
                           Index           =   1
                           Left            =   120
                           TabIndex        =   440
                           Top             =   840
                           Width           =   1365
                        End
                        Begin VB.Label lblAttachmentURL 
                           Caption         =   "Attachment URL:"
                           Height          =   240
                           Left            =   1950
                           TabIndex        =   110
                           Top             =   330
                           Width           =   1275
                        End
                        Begin VB.Label lblIntervall 
                           Caption         =   "Intervall:"
                           Height          =   285
                           Index           =   0
                           Left            =   45
                           TabIndex        =   109
                           Top             =   360
                           Width           =   915
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elGenSynch 
                        Height          =   1500
                        Index           =   1
                        Left            =   10035
                        TabIndex        =   111
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   8700
                        _cx             =   15346
                        _cy             =   2646
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
                        Begin VB.TextBox txtGISGrid 
                           Height          =   285
                           Index           =   10
                           Left            =   3195
                           TabIndex        =   115
                           Top             =   45
                           Width           =   5100
                        End
                        Begin VB.TextBox txtGISGrid 
                           Height          =   285
                           Index           =   11
                           Left            =   3195
                           TabIndex        =   114
                           Top             =   435
                           Width           =   5100
                        End
                        Begin VB.TextBox txtGISGrid 
                           Height          =   285
                           Index           =   12
                           Left            =   3195
                           TabIndex        =   113
                           Top             =   825
                           Width           =   5100
                        End
                        Begin VB.TextBox txtGISGrid 
                           Height          =   285
                           Index           =   13
                           Left            =   3195
                           TabIndex        =   112
                           Top             =   1200
                           Width           =   5100
                        End
                        Begin VB.Label lblAdmLevel 
                           AutoSize        =   -1  'True
                           Caption         =   "Script Name Hot Key 1: (CTRL + ALT + 1)"
                           Height          =   195
                           Index           =   5
                           Left            =   135
                           TabIndex        =   119
                           Top             =   90
                           Width           =   2970
                        End
                        Begin VB.Label lblAdmLevel 
                           AutoSize        =   -1  'True
                           Caption         =   "Script Name Hot Key 2: (CTRL + ALT + 2)"
                           Height          =   195
                           Index           =   6
                           Left            =   135
                           TabIndex        =   118
                           Top             =   480
                           Width           =   2970
                        End
                        Begin VB.Label lblAdmLevel 
                           AutoSize        =   -1  'True
                           Caption         =   "Script Name Hot Key 3: (CTRL + ALT + 3)"
                           Height          =   195
                           Index           =   7
                           Left            =   135
                           TabIndex        =   117
                           Top             =   870
                           Width           =   2970
                        End
                        Begin VB.Label lblAdmLevel 
                           AutoSize        =   -1  'True
                           Caption         =   "Script Name Hot Key 4: (CTRL + ALT + 4)"
                           Height          =   195
                           Index           =   8
                           Left            =   135
                           TabIndex        =   116
                           Top             =   1260
                           Width           =   2970
                        End
                     End
                  End
               End
            End
            Begin C1SizerLibCtl.C1Tab c1TabCoreModules 
               Height          =   4995
               Left            =   45
               TabIndex        =   120
               Top             =   2310
               Width           =   9450
               _cx             =   16669
               _cy             =   8811
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
               Caption         =   "Oasis Profile|Operations"
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
               Separators      =   -1  'True
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   37
               Begin C1SizerLibCtl.C1Elastic elOperations 
                  Height          =   4620
                  Left            =   45
                  TabIndex        =   121
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   9360
                  _cx             =   16510
                  _cy             =   8149
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
                  Begin VB.Frame FraAvailableMap 
                     Caption         =   "Available Map Tools:"
                     Height          =   960
                     Left            =   60
                     TabIndex        =   179
                     Top             =   390
                     Width           =   9300
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   25
                        Left            =   8880
                        TabIndex        =   449
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   24
                        Left            =   8580
                        TabIndex        =   433
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   23
                        Left            =   2670
                        TabIndex        =   426
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   22
                        Left            =   270
                        TabIndex        =   425
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   21
                        Left            =   8280
                        TabIndex        =   382
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   20
                        Left            =   7950
                        TabIndex        =   370
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   19
                        Left            =   7650
                        TabIndex        =   369
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   18
                        Left            =   7320
                        TabIndex        =   368
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Enabled         =   0   'False
                        Height          =   225
                        Index           =   17
                        Left            =   6960
                        TabIndex        =   367
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.PictureBox pctMaptools 
                        BorderStyle     =   0  'None
                        Height          =   345
                        Index           =   0
                        Left            =   60
                        Picture         =   "frmConfig.frx":34719
                        ScaleHeight     =   345
                        ScaleWidth      =   9045
                        TabIndex        =   197
                        Top             =   270
                        Width           =   9045
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   0
                        Left            =   555
                        TabIndex        =   196
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   1
                        Left            =   915
                        TabIndex        =   195
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   2
                        Left            =   1230
                        TabIndex        =   194
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   3
                        Left            =   1545
                        TabIndex        =   193
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   4
                        Left            =   2010
                        TabIndex        =   192
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   5
                        Left            =   2340
                        TabIndex        =   191
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   6
                        Left            =   3165
                        TabIndex        =   190
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   7
                        Left            =   3480
                        TabIndex        =   189
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   8
                        Left            =   3840
                        TabIndex        =   188
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   9
                        Left            =   4200
                        TabIndex        =   187
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   10
                        Left            =   4650
                        TabIndex        =   186
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   11
                        Left            =   4980
                        TabIndex        =   185
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   12
                        Left            =   5310
                        TabIndex        =   184
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   13
                        Left            =   5640
                        TabIndex        =   183
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   14
                        Left            =   5940
                        TabIndex        =   182
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   15
                        Left            =   6270
                        TabIndex        =   181
                        Top             =   600
                        Width           =   195
                     End
                     Begin VB.CheckBox chkMapTool 
                        Height          =   225
                        Index           =   16
                        Left            =   6630
                        TabIndex        =   180
                        Top             =   600
                        Width           =   195
                     End
                  End
                  Begin VB.Frame FraAvailableTool 
                     Caption         =   "Settings:"
                     Height          =   3360
                     Left            =   60
                     TabIndex        =   123
                     Top             =   1380
                     Width           =   9285
                     Begin VB.Frame FraCommonOperations 
                        Caption         =   "COP Themes: "
                        Height          =   1335
                        Left            =   30
                        TabIndex        =   411
                        Top             =   1080
                        Width           =   1380
                        Begin VB.CheckBox chkCOPThemes 
                           Caption         =   "Operations"
                           Height          =   285
                           Index           =   3
                           Left            =   180
                           TabIndex        =   415
                           Top             =   990
                           Width           =   1140
                        End
                        Begin VB.CheckBox chkCOPThemes 
                           Caption         =   "Who"
                           Height          =   285
                           Index           =   2
                           Left            =   180
                           TabIndex        =   414
                           Top             =   720
                           Width           =   690
                        End
                        Begin VB.CheckBox chkCOPThemes 
                           Caption         =   "W3"
                           Height          =   285
                           Index           =   1
                           Left            =   180
                           TabIndex        =   413
                           Top             =   480
                           Width           =   660
                        End
                        Begin VB.CheckBox chkCOPThemes 
                           Caption         =   "Incidents"
                           Height          =   285
                           Index           =   0
                           Left            =   180
                           TabIndex        =   412
                           Top             =   225
                           Width           =   1020
                        End
                     End
                     Begin VB.CheckBox chkShowlegendTab 
                        Caption         =   "Legend"
                        Height          =   195
                        Left            =   60
                        TabIndex        =   133
                        Top             =   360
                        Width           =   1110
                     End
                     Begin VB.CheckBox chkShowSecurityTab 
                        Caption         =   "Security"
                        Height          =   195
                        Left            =   60
                        TabIndex        =   132
                        Top             =   840
                        Width           =   1080
                     End
                     Begin VB.CheckBox chkShowMapLibraryTab 
                        Caption         =   "Map Library"
                        Height          =   195
                        Left            =   60
                        TabIndex        =   131
                        Top             =   600
                        Width           =   1170
                     End
                     Begin VB.ComboBox ComInitialTab 
                        Height          =   315
                        Left            =   60
                        Style           =   2  'Dropdown List
                        TabIndex        =   130
                        Top             =   2700
                        Width           =   2310
                     End
                     Begin VB.Frame FraDataGrid 
                        Caption         =   "Data Grid Settings:"
                        Height          =   2190
                        Left            =   1410
                        TabIndex        =   124
                        Top             =   225
                        Width           =   1695
                        Begin VB.CheckBox chkGridOption 
                           Caption         =   "Map Select"
                           Height          =   285
                           Index           =   1
                           Left            =   90
                           TabIndex        =   372
                           Top             =   1590
                           Width           =   1185
                        End
                        Begin VB.CheckBox chkGridOption 
                           Caption         =   "Map Filter"
                           Height          =   285
                           Index           =   0
                           Left            =   90
                           TabIndex        =   371
                           Top             =   1320
                           Width           =   1185
                        End
                        Begin VB.CheckBox chkUseMax 
                           Caption         =   "Use Max Records"
                           Height          =   240
                           Left            =   60
                           TabIndex        =   127
                           Top             =   270
                           Width           =   1575
                        End
                        Begin VB.TextBox txtrecMaxLevel 
                           Height          =   285
                           Left            =   765
                           TabIndex        =   126
                           Top             =   585
                           Width           =   840
                        End
                        Begin VB.TextBox txtrecWarningLevel 
                           Height          =   285
                           Left            =   765
                           TabIndex        =   125
                           Top             =   990
                           Width           =   840
                        End
                        Begin VB.Label lblMaxLevel 
                           Caption         =   "Max level:"
                           Height          =   420
                           Left            =   90
                           TabIndex        =   129
                           Top             =   495
                           Width           =   420
                        End
                        Begin VB.Label lblWarningLevel 
                           Caption         =   "Warning Level:"
                           Height          =   465
                           Left            =   90
                           TabIndex        =   128
                           Top             =   900
                           Width           =   645
                        End
                     End
                     Begin C1SizerLibCtl.C1Tab C1TabOPS 
                        Height          =   3060
                        Left            =   3150
                        TabIndex        =   134
                        Top             =   210
                        Width           =   6015
                        _cx             =   10610
                        _cy             =   5397
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
                        Caption         =   "Utility|Legend|Maps Library|Security|Mine Action|Security Themes"
                        Align           =   0
                        CurrTab         =   3
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
                        Begin OASISThemeRange.OASISThemeRangePicker OASISThemeRangePicker1 
                           Height          =   2685
                           Left            =   6960
                           TabIndex        =   444
                           Top             =   330
                           Width           =   5925
                           _ExtentX        =   10451
                           _ExtentY        =   4736
                        End
                        Begin C1SizerLibCtl.C1Elastic elSecurityOps 
                           Height          =   2685
                           Left            =   45
                           TabIndex        =   135
                           TabStop         =   0   'False
                           Top             =   330
                           Width           =   5925
                           _cx             =   10451
                           _cy             =   4736
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
                           BorderWidth     =   2
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
                           Begin VB.Frame FraMisc 
                              Caption         =   "Misc"
                              Height          =   615
                              Left            =   3990
                              TabIndex        =   379
                              Top             =   240
                              Width           =   1845
                              Begin VB.TextBox txtLGDSymFontSize 
                                 Height          =   225
                                 Left            =   150
                                 TabIndex        =   381
                                 Top             =   360
                                 Width           =   1425
                              End
                              Begin VB.Label lblLegendSymbol 
                                 AutoSize        =   -1  'True
                                 Caption         =   "Legend Symbol Size:"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   195
                                 Left            =   180
                                 TabIndex        =   380
                                 Top             =   180
                                 Width           =   1290
                              End
                           End
                           Begin VB.TextBox txtOASISIncLyrName 
                              Height          =   255
                              Left            =   1710
                              TabIndex        =   162
                              Top             =   60
                              Width           =   3960
                           End
                           Begin VB.Frame FraAvailableTools 
                              Caption         =   "Available Tools:"
                              Height          =   615
                              Left            =   90
                              TabIndex        =   157
                              Top             =   240
                              Width           =   3885
                              Begin VB.CheckBox chkSecurityAnalysis 
                                 Caption         =   "Security Analysis"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   360
                                 Left            =   90
                                 TabIndex        =   161
                                 Top             =   180
                                 Width           =   900
                              End
                              Begin VB.CheckBox chkSecurityGraphs 
                                 Caption         =   "Security Graphs"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   375
                                 Left            =   1050
                                 TabIndex        =   160
                                 Top             =   180
                                 Width           =   900
                              End
                              Begin VB.CheckBox chkSecurityTrends 
                                 Caption         =   "Security trends"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   360
                                 Left            =   2010
                                 TabIndex        =   159
                                 Top             =   210
                                 Width           =   915
                              End
                              Begin VB.CheckBox chkAddIncident 
                                 Caption         =   "Add incident"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   375
                                 Left            =   2940
                                 TabIndex        =   158
                                 Top             =   180
                                 Width           =   870
                              End
                           End
                           Begin VB.Frame FraSecurityGrid 
                              Caption         =   "Security Grid Settings:"
                              BeginProperty Font 
                                 Name            =   "MS Serif"
                                 Size            =   8.25
                                 Charset         =   0
                                 Weight          =   400
                                 Underline       =   0   'False
                                 Italic          =   0   'False
                                 Strikethrough   =   0   'False
                              EndProperty
                              Height          =   1845
                              Left            =   60
                              TabIndex        =   136
                              Top             =   870
                              Width           =   5790
                              Begin VB.TextBox txtSecAdmKey 
                                 Height          =   255
                                 Index           =   2
                                 Left            =   4470
                                 TabIndex        =   439
                                 Top             =   1560
                                 Width           =   1215
                              End
                              Begin VB.TextBox txtSecAdmKey 
                                 Height          =   255
                                 Index           =   1
                                 Left            =   2460
                                 TabIndex        =   438
                                 Top             =   1560
                                 Width           =   1215
                              End
                              Begin VB.TextBox txtSecAdmKey 
                                 Height          =   255
                                 Index           =   0
                                 Left            =   630
                                 TabIndex        =   437
                                 Top             =   1560
                                 Width           =   1215
                              End
                              Begin VB.TextBox txtSecAdm 
                                 Height          =   285
                                 Index           =   2
                                 Left            =   4470
                                 TabIndex        =   432
                                 Top             =   1290
                                 Width           =   1215
                              End
                              Begin VB.TextBox txtSecAdm 
                                 Height          =   285
                                 Index           =   1
                                 Left            =   2460
                                 TabIndex        =   431
                                 Top             =   1290
                                 Width           =   1215
                              End
                              Begin VB.TextBox txtSecAdm 
                                 Height          =   255
                                 Index           =   0
                                 Left            =   630
                                 TabIndex        =   430
                                 Top             =   1290
                                 Width           =   1215
                              End
                              Begin VB.Frame FraZoomLevels 
                                 Caption         =   "Zoom Level:"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   1095
                                 Left            =   90
                                 TabIndex        =   150
                                 Top             =   180
                                 Width           =   1140
                                 Begin VB.TextBox txtSecZoomLevels 
                                    Height          =   255
                                    Index           =   0
                                    Left            =   270
                                    TabIndex        =   153
                                    Top             =   240
                                    Width           =   735
                                 End
                                 Begin VB.TextBox txtSecZoomLevels 
                                    Height          =   285
                                    Index           =   1
                                    Left            =   270
                                    TabIndex        =   152
                                    Top             =   480
                                    Width           =   735
                                 End
                                 Begin VB.TextBox txtSecZoomLevels 
                                    Height          =   285
                                    Index           =   2
                                    Left            =   270
                                    TabIndex        =   151
                                    Top             =   750
                                    Width           =   735
                                 End
                                 Begin VB.Label lblzLevel 
                                    AutoSize        =   -1  'True
                                    Caption         =   "1:"
                                    BeginProperty Font 
                                       Name            =   "MS Serif"
                                       Size            =   8.25
                                       Charset         =   0
                                       Weight          =   400
                                       Underline       =   0   'False
                                       Italic          =   0   'False
                                       Strikethrough   =   0   'False
                                    EndProperty
                                    Height          =   195
                                    Index           =   0
                                    Left            =   90
                                    TabIndex        =   156
                                    Top             =   270
                                    Width           =   105
                                 End
                                 Begin VB.Label lblzLevel 
                                    AutoSize        =   -1  'True
                                    Caption         =   "2:"
                                    BeginProperty Font 
                                       Name            =   "MS Serif"
                                       Size            =   8.25
                                       Charset         =   0
                                       Weight          =   400
                                       Underline       =   0   'False
                                       Italic          =   0   'False
                                       Strikethrough   =   0   'False
                                    EndProperty
                                    Height          =   195
                                    Index           =   1
                                    Left            =   90
                                    TabIndex        =   155
                                    Top             =   540
                                    Width           =   105
                                 End
                                 Begin VB.Label lblzLevel 
                                    AutoSize        =   -1  'True
                                    Caption         =   "3:"
                                    BeginProperty Font 
                                       Name            =   "MS Serif"
                                       Size            =   8.25
                                       Charset         =   0
                                       Weight          =   400
                                       Underline       =   0   'False
                                       Italic          =   0   'False
                                       Strikethrough   =   0   'False
                                    EndProperty
                                    Height          =   195
                                    Index           =   2
                                    Left            =   90
                                    TabIndex        =   154
                                    Top             =   810
                                    Width           =   105
                                 End
                              End
                              Begin VB.Frame FraGridLayers 
                                 Caption         =   "Grid Layers:"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   1095
                                 Left            =   1260
                                 TabIndex        =   137
                                 Top             =   180
                                 Width           =   4470
                                 Begin VB.TextBox txtSaGrLyrName 
                                    Height          =   240
                                    Index           =   0
                                    Left            =   540
                                    TabIndex        =   143
                                    Top             =   210
                                    Width           =   1995
                                 End
                                 Begin VB.TextBox txtSecGrLyrKey 
                                    Height          =   240
                                    Index           =   0
                                    Left            =   2925
                                    TabIndex        =   142
                                    Top             =   210
                                    Width           =   1500
                                 End
                                 Begin VB.TextBox txtSaGrLyrName 
                                    Height          =   285
                                    Index           =   1
                                    Left            =   540
                                    TabIndex        =   141
                                    Top             =   480
                                    Width           =   1995
                                 End
                                 Begin VB.TextBox txtSecGrLyrKey 
                                    Height          =   285
                                    Index           =   1
                                    Left            =   2925
                                    TabIndex        =   140
                                    Top             =   480
                                    Width           =   1500
                                 End
                                 Begin VB.TextBox txtSaGrLyrName 
                                    Height          =   285
                                    Index           =   2
                                    Left            =   540
                                    TabIndex        =   139
                                    Top             =   750
                                    Width           =   1995
                                 End
                                 Begin VB.TextBox txtSecGrLyrKey 
                                    Height          =   285
                                    Index           =   2
                                    Left            =   2925
                                    TabIndex        =   138
                                    Top             =   750
                                    Width           =   1500
                                 End
                                 Begin VB.Label lblGRName 
                                    AutoSize        =   -1  'True
                                    Caption         =   "Name:"
                                    BeginProperty Font 
                                       Name            =   "MS Serif"
                                       Size            =   8.25
                                       Charset         =   0
                                       Weight          =   400
                                       Underline       =   0   'False
                                       Italic          =   0   'False
                                       Strikethrough   =   0   'False
                                    EndProperty
                                    Height          =   195
                                    Index           =   0
                                    Left            =   45
                                    TabIndex        =   149
                                    Top             =   255
                                    Width           =   420
                                 End
                                 Begin VB.Label lblGRKey 
                                    AutoSize        =   -1  'True
                                    Caption         =   "Key:"
                                    BeginProperty Font 
                                       Name            =   "MS Serif"
                                       Size            =   8.25
                                       Charset         =   0
                                       Weight          =   400
                                       Underline       =   0   'False
                                       Italic          =   0   'False
                                       Strikethrough   =   0   'False
                                    EndProperty
                                    Height          =   195
                                    Index           =   0
                                    Left            =   2565
                                    TabIndex        =   148
                                    Top             =   210
                                    Width           =   300
                                 End
                                 Begin VB.Label lblGRName 
                                    AutoSize        =   -1  'True
                                    Caption         =   "Name:"
                                    BeginProperty Font 
                                       Name            =   "MS Serif"
                                       Size            =   8.25
                                       Charset         =   0
                                       Weight          =   400
                                       Underline       =   0   'False
                                       Italic          =   0   'False
                                       Strikethrough   =   0   'False
                                    EndProperty
                                    Height          =   195
                                    Index           =   1
                                    Left            =   45
                                    TabIndex        =   147
                                    Top             =   525
                                    Width           =   420
                                 End
                                 Begin VB.Label lblGRKey 
                                    AutoSize        =   -1  'True
                                    Caption         =   "Key:"
                                    BeginProperty Font 
                                       Name            =   "MS Serif"
                                       Size            =   8.25
                                       Charset         =   0
                                       Weight          =   400
                                       Underline       =   0   'False
                                       Italic          =   0   'False
                                       Strikethrough   =   0   'False
                                    EndProperty
                                    Height          =   195
                                    Index           =   1
                                    Left            =   2565
                                    TabIndex        =   146
                                    Top             =   480
                                    Width           =   300
                                 End
                                 Begin VB.Label lblGRName 
                                    AutoSize        =   -1  'True
                                    Caption         =   "Name:"
                                    BeginProperty Font 
                                       Name            =   "MS Serif"
                                       Size            =   8.25
                                       Charset         =   0
                                       Weight          =   400
                                       Underline       =   0   'False
                                       Italic          =   0   'False
                                       Strikethrough   =   0   'False
                                    EndProperty
                                    Height          =   195
                                    Index           =   2
                                    Left            =   45
                                    TabIndex        =   145
                                    Top             =   795
                                    Width           =   420
                                 End
                                 Begin VB.Label lblGRKey 
                                    AutoSize        =   -1  'True
                                    Caption         =   "Key:"
                                    BeginProperty Font 
                                       Name            =   "MS Serif"
                                       Size            =   8.25
                                       Charset         =   0
                                       Weight          =   400
                                       Underline       =   0   'False
                                       Italic          =   0   'False
                                       Strikethrough   =   0   'False
                                    EndProperty
                                    Height          =   195
                                    Index           =   2
                                    Left            =   2565
                                    TabIndex        =   144
                                    Top             =   750
                                    Width           =   300
                                 End
                              End
                              Begin VB.Label lblGRKey 
                                 AutoSize        =   -1  'True
                                 Caption         =   "Key:"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   195
                                 Index           =   5
                                 Left            =   4020
                                 TabIndex        =   436
                                 Top             =   1590
                                 Width           =   300
                              End
                              Begin VB.Label lblGRKey 
                                 AutoSize        =   -1  'True
                                 Caption         =   "Key:"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   195
                                 Index           =   4
                                 Left            =   2070
                                 TabIndex        =   435
                                 Top             =   1590
                                 Width           =   300
                              End
                              Begin VB.Label lblGRKey 
                                 AutoSize        =   -1  'True
                                 Caption         =   "Key:"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   195
                                 Index           =   3
                                 Left            =   240
                                 TabIndex        =   434
                                 Top             =   1590
                                 Width           =   300
                              End
                              Begin VB.Label lblGRName 
                                 AutoSize        =   -1  'True
                                 Caption         =   "Location:"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   195
                                 Index           =   5
                                 Left            =   3750
                                 TabIndex        =   429
                                 Top             =   1320
                                 Width           =   570
                              End
                              Begin VB.Label lblGRName 
                                 AutoSize        =   -1  'True
                                 Caption         =   "Adm 2:"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   195
                                 Index           =   4
                                 Left            =   1920
                                 TabIndex        =   428
                                 Top             =   1320
                                 Width           =   450
                              End
                              Begin VB.Label lblGRName 
                                 AutoSize        =   -1  'True
                                 Caption         =   "Adm 1:"
                                 BeginProperty Font 
                                    Name            =   "MS Serif"
                                    Size            =   8.25
                                    Charset         =   0
                                    Weight          =   400
                                    Underline       =   0   'False
                                    Italic          =   0   'False
                                    Strikethrough   =   0   'False
                                 EndProperty
                                 Height          =   195
                                 Index           =   3
                                 Left            =   90
                                 TabIndex        =   427
                                 Top             =   1320
                                 Width           =   450
                              End
                           End
                           Begin VB.Label lblOASISIncident 
                              AutoSize        =   -1  'True
                              Caption         =   "Incident Layer Name:"
                              Height          =   195
                              Left            =   90
                              TabIndex        =   163
                              Top             =   30
                              Width           =   1515
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elMapsOps 
                           Height          =   2685
                           Left            =   -6870
                           TabIndex        =   164
                           TabStop         =   0   'False
                           Top             =   330
                           Width           =   5925
                           _cx             =   10451
                           _cy             =   4736
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
                           Begin VB.TextBox txtHiddenLayers 
                              Height          =   285
                              Left            =   30
                              TabIndex        =   166
                              Top             =   450
                              Width           =   6225
                           End
                           Begin VB.CheckBox chkUseAuto 
                              Caption         =   "Use Predefined Themes"
                              Height          =   465
                              Left            =   45
                              TabIndex        =   165
                              Top             =   750
                              Width           =   2955
                           End
                           Begin VB.Label lblHiddenLayers 
                              Caption         =   "Hidden Layers:"
                              Height          =   285
                              Left            =   45
                              TabIndex        =   167
                              Top             =   180
                              Width           =   1185
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elOPSMapLibrary 
                           Height          =   2685
                           Left            =   -6570
                           TabIndex        =   168
                           TabStop         =   0   'False
                           Top             =   330
                           Width           =   5925
                           _cx             =   10451
                           _cy             =   4736
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
                           BorderWidth     =   0
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
                           Begin VB.ComboBox ComMaps 
                              Height          =   315
                              Left            =   90
                              Style           =   2  'Dropdown List
                              TabIndex        =   169
                              Top             =   135
                              Width           =   2265
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elOpsLegend 
                           Height          =   2685
                           Left            =   -7170
                           TabIndex        =   170
                           TabStop         =   0   'False
                           Top             =   330
                           Width           =   5925
                           _cx             =   10451
                           _cy             =   4736
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
                           BorderWidth     =   0
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
                           Begin VB.CheckBox chkShowUtility 
                              Caption         =   "Show Utility Toolbar"
                              Height          =   195
                              Left            =   135
                              TabIndex        =   177
                              Top             =   135
                              Width           =   2715
                           End
                           Begin VB.Frame FraAvailbleUtility 
                              Caption         =   "Available Utility Tools:"
                              Height          =   780
                              Left            =   90
                              TabIndex        =   171
                              Top             =   495
                              Width           =   4650
                              Begin VB.CheckBox chkUtils 
                                 Caption         =   "Geo Marks"
                                 Height          =   195
                                 Index           =   0
                                 Left            =   90
                                 TabIndex        =   176
                                 Top             =   270
                                 Width           =   1095
                              End
                              Begin VB.CheckBox chkUtils 
                                 Caption         =   "Coordinate tools"
                                 Height          =   195
                                 Index           =   1
                                 Left            =   1260
                                 TabIndex        =   175
                                 Top             =   270
                                 Width           =   1500
                              End
                              Begin VB.CheckBox chkUtils 
                                 Caption         =   "Goto location"
                                 Height          =   195
                                 Index           =   2
                                 Left            =   2880
                                 TabIndex        =   174
                                 Top             =   270
                                 Width           =   1500
                              End
                              Begin VB.CheckBox chkUtils 
                                 Caption         =   "Settings"
                                 Height          =   195
                                 Index           =   3
                                 Left            =   90
                                 TabIndex        =   173
                                 Top             =   495
                                 Width           =   960
                              End
                              Begin VB.CheckBox chkUtils 
                                 Caption         =   "Magnifier"
                                 Height          =   195
                                 Index           =   4
                                 Left            =   1260
                                 TabIndex        =   172
                                 Top             =   495
                                 Width           =   960
                              End
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elAddMods 
                           Height          =   2685
                           Index           =   2
                           Left            =   6660
                           TabIndex        =   383
                           TabStop         =   0   'False
                           Top             =   330
                           Width           =   5925
                           _cx             =   10451
                           _cy             =   4736
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
                           BorderWidth     =   2
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
                           Begin C1SizerLibCtl.C1Elastic elMAOps 
                              Height          =   2925
                              Left            =   0
                              TabIndex        =   384
                              TabStop         =   0   'False
                              Top             =   0
                              Width           =   9375
                              _cx             =   16536
                              _cy             =   5159
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
                              Begin VB.TextBox txtMAThemeMap 
                                 Height          =   285
                                 Left            =   45
                                 TabIndex        =   391
                                 Top             =   1845
                                 Width           =   4605
                              End
                              Begin VB.ComboBox ComNumOfMARegions 
                                 Height          =   315
                                 Left            =   675
                                 Style           =   2  'Dropdown List
                                 TabIndex        =   390
                                 Top             =   2205
                                 Width           =   645
                              End
                              Begin VB.CheckBox chkMAAllowDataUpdate 
                                 Caption         =   "Allow Data Update"
                                 Height          =   285
                                 Left            =   810
                                 TabIndex        =   389
                                 Top             =   45
                                 Width           =   1725
                              End
                              Begin VB.CheckBox chkShowLISTab 
                                 Caption         =   "LIS"
                                 Height          =   195
                                 Left            =   90
                                 TabIndex        =   388
                                 Top             =   90
                                 Width           =   555
                              End
                              Begin VB.TextBox txtMASQL 
                                 ForeColor       =   &H0000FF00&
                                 Height          =   915
                                 Left            =   90
                                 MultiLine       =   -1  'True
                                 ScrollBars      =   2  'Vertical
                                 TabIndex        =   387
                                 Top             =   630
                                 Width           =   4515
                              End
                              Begin VB.TextBox txtMaKeyField 
                                 Height          =   240
                                 Left            =   2880
                                 TabIndex        =   386
                                 Top             =   360
                                 Width           =   1725
                              End
                              Begin VB.ComboBox ComMaregionNames 
                                 Height          =   315
                                 Left            =   2115
                                 TabIndex        =   385
                                 Top             =   2160
                                 Width           =   2535
                              End
                              Begin VB.Label lblMineAction 
                                 Caption         =   "Mine Action Theme Map:"
                                 Height          =   285
                                 Left            =   90
                                 TabIndex        =   396
                                 Top             =   1620
                                 Width           =   1860
                              End
                              Begin VB.Label lblOfRegions 
                                 Caption         =   "# of regions:"
                                 Height          =   420
                                 Left            =   90
                                 TabIndex        =   395
                                 Top             =   2115
                                 Width           =   645
                              End
                              Begin VB.Label lblMAData 
                                 Caption         =   "MA Data SQL:"
                                 Height          =   240
                                 Left            =   90
                                 TabIndex        =   394
                                 Top             =   405
                                 Width           =   1455
                              End
                              Begin VB.Label lblUniqueData 
                                 AutoSize        =   -1  'True
                                 Caption         =   "Data Key Field:"
                                 Height          =   195
                                 Left            =   1710
                                 TabIndex        =   393
                                 Top             =   405
                                 Width           =   1080
                              End
                              Begin VB.Label lblRegionNames 
                                 Caption         =   "Region Names:"
                                 Height          =   375
                                 Left            =   1440
                                 TabIndex        =   392
                                 Top             =   2115
                                 Width           =   645
                              End
                           End
                        End
                     End
                     Begin VB.Label lblInitialMenu 
                        Caption         =   "Initial Menu:"
                        Height          =   195
                        Left            =   60
                        TabIndex        =   178
                        Top             =   2490
                        Width           =   1590
                     End
                  End
                  Begin VB.TextBox txtinitMap 
                     Height          =   285
                     Left            =   930
                     TabIndex        =   122
                     Top             =   30
                     Width           =   4800
                  End
                  Begin VB.Label lblInitialMap 
                     Caption         =   "Initial Map:"
                     Height          =   330
                     Left            =   90
                     TabIndex        =   198
                     Top             =   90
                     Width           =   2490
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elProfiles 
                  Height          =   4620
                  Left            =   -10005
                  TabIndex        =   199
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   9360
                  _cx             =   16510
                  _cy             =   8149
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
                  Begin VB.Frame FraIntranetSettings 
                     Caption         =   "Intranet Settings"
                     Height          =   2235
                     Left            =   135
                     TabIndex        =   205
                     Top             =   2310
                     Width           =   8880
                     Begin VB.CheckBox chkUseIntranet 
                        Caption         =   "Use Intranet"
                        Height          =   285
                        Left            =   225
                        TabIndex        =   208
                        Top             =   270
                        Width           =   3030
                     End
                     Begin VB.CommandButton cmdIntranetTest 
                        Caption         =   "Intranet Test"
                        Height          =   285
                        Left            =   7350
                        TabIndex        =   207
                        Top             =   930
                        Width           =   1050
                     End
                     Begin VB.TextBox txtIntranetUrl 
                        Height          =   330
                        Left            =   180
                        TabIndex        =   206
                        Top             =   900
                        Width           =   7155
                     End
                     Begin VB.Label lblIntranetPage 
                        Caption         =   "Intranet Page URL:"
                        Height          =   240
                        Left            =   225
                        TabIndex        =   209
                        Top             =   630
                        Width           =   2310
                     End
                  End
                  Begin VB.Frame FraGeneral 
                     Caption         =   "General Profile Settings"
                     Height          =   1995
                     Left            =   135
                     TabIndex        =   200
                     Top             =   180
                     Width           =   8910
                     Begin VB.CheckBox chkUseGeneral 
                        Caption         =   "Use General Profile"
                        Height          =   285
                        Left            =   315
                        TabIndex        =   203
                        Top             =   270
                        Width           =   3030
                     End
                     Begin VB.CommandButton cmdTest 
                        Caption         =   "Test"
                        Height          =   330
                        Left            =   7380
                        TabIndex        =   202
                        Top             =   870
                        Width           =   1050
                     End
                     Begin VB.TextBox txtProfileUrl 
                        Height          =   330
                        Left            =   270
                        TabIndex        =   201
                        Top             =   900
                        Width           =   7065
                     End
                     Begin VB.Label lblProfilePage 
                        Caption         =   "Profile Page URL:"
                        Height          =   240
                        Left            =   315
                        TabIndex        =   204
                        Top             =   630
                        Width           =   2310
                     End
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elDataPackDefinitions 
            Height          =   7350
            Left            =   10485
            TabIndex        =   210
            TabStop         =   0   'False
            Top             =   465
            Width           =   9540
            _cx             =   16828
            _cy             =   12965
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
            GridRows        =   1
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmConfig.frx":3624F
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Tab C1TDataDefsAccess 
               Height          =   7290
               Left            =   30
               TabIndex        =   211
               Top             =   30
               Width           =   9480
               _cx             =   16722
               _cy             =   12859
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
               Caption         =   "Data Pack Definitions|GIS Attribute Access|Geo Marks|Maps Definitions|Encryption"
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
               Begin C1SizerLibCtl.C1Elastic elHolder 
                  Height          =   6975
                  Index           =   15
                  Left            =   3000
                  TabIndex        =   212
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   1365
                  _cx             =   2408
                  _cy             =   12303
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
                  Begin VB.Frame FraSettings 
                     Caption         =   "Hash Settings:"
                     Height          =   3825
                     Index           =   1
                     Left            =   4590
                     TabIndex        =   219
                     Top             =   270
                     Width           =   4425
                     Begin VB.CommandButton cmdGenerateHASH 
                        Caption         =   "Generate HASH"
                        Height          =   435
                        Left            =   2220
                        TabIndex        =   221
                        Top             =   1680
                        Width           =   1515
                     End
                     Begin VB.ComboBox ComAlgorithm 
                        Height          =   315
                        Index           =   1
                        ItemData        =   "frmConfig.frx":36287
                        Left            =   180
                        List            =   "frmConfig.frx":36291
                        Style           =   2  'Dropdown List
                        TabIndex        =   220
                        Top             =   990
                        Width           =   4065
                     End
                  End
                  Begin VB.Frame FraSettings 
                     Caption         =   "Encryption Settings:"
                     Height          =   3915
                     Index           =   0
                     Left            =   150
                     TabIndex        =   213
                     Top             =   210
                     Width           =   4365
                     Begin VB.CommandButton cmdTestEncrypt 
                        Caption         =   "Test DeEncrypt"
                        Height          =   405
                        Index           =   1
                        Left            =   2730
                        TabIndex        =   218
                        Top             =   1410
                        Width           =   1455
                     End
                     Begin VB.CommandButton cmdTestEncrypt 
                        Caption         =   "Test Encrypt"
                        Height          =   405
                        Index           =   0
                        Left            =   1230
                        TabIndex        =   217
                        Top             =   1410
                        Width           =   1455
                     End
                     Begin VB.TextBox txtEncryptInput 
                        Height          =   1695
                        Left            =   180
                        TabIndex        =   216
                        Text            =   "Try what is coming out "
                        Top             =   1980
                        Width           =   3975
                     End
                     Begin VB.ComboBox ComAlgorithm 
                        Height          =   315
                        Index           =   0
                        ItemData        =   "frmConfig.frx":362A0
                        Left            =   150
                        List            =   "frmConfig.frx":362BF
                        Style           =   2  'Dropdown List
                        TabIndex        =   215
                        Top             =   960
                        Width           =   4065
                     End
                     Begin VB.CheckBox chkUseEncryption 
                        Caption         =   "Use Encryption"
                        Height          =   375
                        Index           =   0
                        Left            =   120
                        TabIndex        =   214
                        Top             =   240
                        Width           =   1425
                     End
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elMapDefinitionFiled 
                  Height          =   6975
                  Left            =   2700
                  TabIndex        =   222
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   1365
                  _cx             =   2408
                  _cy             =   12303
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
                  GridRows        =   1
                  GridCols        =   1
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmConfig.frx":36303
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin C1SizerLibCtl.C1Tab tabMapProducts 
                     Height          =   6915
                     Left            =   30
                     TabIndex        =   223
                     Top             =   30
                     Width           =   1305
                     _cx             =   2302
                     _cy             =   12197
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
                     Caption         =   "Map Definition 1|Map Definition 2|Map Definition 3|Map Definition 4|Map Descriptions"
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
                     Begin C1SizerLibCtl.C1Elastic leMapDescriptions 
                        Height          =   6795
                        Left            =   4515
                        TabIndex        =   224
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   2880
                        _cx             =   5080
                        _cy             =   11986
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
                        AutoSizeChildren=   7
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
                        Begin DXDBGRIDLibCtl.dxDBGrid dxDBmapDescriptions 
                           Height          =   6705
                           Left            =   15
                           OleObjectBlob   =   "frmConfig.frx":36338
                           TabIndex        =   225
                           Top             =   60
                           Width           =   2850
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elMapDetails 
                        Height          =   6795
                        Index           =   0
                        Left            =   45
                        TabIndex        =   226
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   2880
                        _cx             =   5080
                        _cy             =   11986
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
                        GridRows        =   1
                        GridCols        =   2
                        Frame           =   3
                        FrameStyle      =   0
                        FrameWidth      =   1
                        FrameColor      =   -2147483628
                        FrameShadow     =   -2147483632
                        FloodStyle      =   1
                        _GridInfo       =   $"frmConfig.frx":36FE0
                        AccessibleName  =   ""
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   9
                        Begin C1SizerLibCtl.C1Elastic elHolder 
                           Height          =   6735
                           Index           =   1
                           Left            =   2175
                           TabIndex        =   227
                           TabStop         =   0   'False
                           Top             =   30
                           Width           =   675
                           _cx             =   1191
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
                           Align           =   0
                           AutoSizeChildren=   0
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
                           Begin VB.Frame FraMapPreview 
                              Caption         =   "Map Preview:"
                              Height          =   3045
                              Index           =   0
                              Left            =   90
                              TabIndex        =   230
                              Top             =   3150
                              Width           =   3540
                              Begin VB.PictureBox mapPreview 
                                 Height          =   2670
                                 Index           =   0
                                 Left            =   120
                                 ScaleHeight     =   2610
                                 ScaleWidth      =   3270
                                 TabIndex        =   231
                                 Top             =   240
                                 Width           =   3330
                              End
                           End
                           Begin VB.Frame FraDescription 
                              Caption         =   "Description:"
                              Height          =   2985
                              Index           =   0
                              Left            =   90
                              TabIndex        =   228
                              Top             =   60
                              Width           =   3540
                              Begin VB.TextBox txtMapDescription 
                                 Height          =   2580
                                 Index           =   0
                                 Left            =   60
                                 MultiLine       =   -1  'True
                                 ScrollBars      =   2  'Vertical
                                 TabIndex        =   229
                                 Top             =   225
                                 Width           =   3420
                              End
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elHolder 
                           Height          =   6735
                           Index           =   0
                           Left            =   30
                           TabIndex        =   232
                           TabStop         =   0   'False
                           Top             =   30
                           Width           =   2820
                           _cx             =   4974
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
                           GridCols        =   1
                           Frame           =   3
                           FrameStyle      =   0
                           FrameWidth      =   1
                           FrameColor      =   -2147483628
                           FrameShadow     =   -2147483632
                           FloodStyle      =   1
                           _GridInfo       =   $"frmConfig.frx":3701F
                           AccessibleName  =   ""
                           AccessibleDescription=   ""
                           AccessibleValue =   ""
                           AccessibleRole  =   9
                           Begin TatukGIS_DK.XGIS_ViewerWnd MapProduct 
                              Height          =   6180
                              Index           =   0
                              Left            =   15
                              TabIndex        =   233
                              Top             =   15
                              Width           =   2790
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
                              SelectionPattern=   "frmConfig.frx":3705C
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
                              ScrollBars      =   3
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
                           Begin C1SizerLibCtl.C1Elastic elHolder 
                              Height          =   510
                              Index           =   2
                              Left            =   15
                              TabIndex        =   234
                              TabStop         =   0   'False
                              Top             =   6210
                              Width           =   2790
                              _cx             =   4921
                              _cy             =   900
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
                              BorderWidth     =   0
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
                              Begin VB.CommandButton cmdCreatePreview 
                                 Caption         =   "Create Preview"
                                 Height          =   405
                                 Index           =   0
                                 Left            =   2640
                                 TabIndex        =   239
                                 Top             =   90
                                 Width           =   1320
                              End
                              Begin VB.CommandButton cmdLoadMap 
                                 Caption         =   "Load Map..."
                                 Height          =   405
                                 Index           =   0
                                 Left            =   1290
                                 TabIndex        =   238
                                 Top             =   90
                                 Width           =   1320
                              End
                              Begin VB.Frame FraMapTools 
                                 Caption         =   "Map Tools:"
                                 Height          =   525
                                 Index           =   0
                                 Left            =   60
                                 TabIndex        =   235
                                 Top             =   -30
                                 Width           =   1140
                                 Begin VB.CommandButton cmdZoomOut 
                                    Caption         =   "-"
                                    Height          =   240
                                    Index           =   0
                                    Left            =   600
                                    TabIndex        =   237
                                    Top             =   210
                                    Width           =   330
                                 End
                                 Begin VB.CommandButton cmdZoomin 
                                    Caption         =   "+"
                                    Height          =   240
                                    Index           =   0
                                    Left            =   135
                                    TabIndex        =   236
                                    Top             =   210
                                    Width           =   330
                                 End
                              End
                           End
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elMapDetails 
                        Height          =   6795
                        Index           =   1
                        Left            =   3615
                        TabIndex        =   240
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   2880
                        _cx             =   5080
                        _cy             =   11986
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
                        GridRows        =   1
                        GridCols        =   2
                        Frame           =   3
                        FrameStyle      =   0
                        FrameWidth      =   1
                        FrameColor      =   -2147483628
                        FrameShadow     =   -2147483632
                        FloodStyle      =   1
                        _GridInfo       =   $"frmConfig.frx":4196E
                        AccessibleName  =   ""
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   9
                        Begin C1SizerLibCtl.C1Elastic elHolder 
                           Height          =   6735
                           Index           =   3
                           Left            =   2175
                           TabIndex        =   241
                           TabStop         =   0   'False
                           Top             =   30
                           Width           =   675
                           _cx             =   1191
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
                           Align           =   0
                           AutoSizeChildren=   0
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
                           Begin VB.Frame FraDescription 
                              Caption         =   "Description:"
                              Height          =   2985
                              Index           =   1
                              Left            =   90
                              TabIndex        =   244
                              Top             =   60
                              Width           =   3540
                              Begin VB.TextBox txtMapDescription 
                                 Height          =   2580
                                 Index           =   1
                                 Left            =   60
                                 MultiLine       =   -1  'True
                                 ScrollBars      =   2  'Vertical
                                 TabIndex        =   245
                                 Top             =   225
                                 Width           =   3420
                              End
                           End
                           Begin VB.Frame FraMapPreview 
                              Caption         =   "Map Preview:"
                              Height          =   3045
                              Index           =   1
                              Left            =   90
                              TabIndex        =   242
                              Top             =   3150
                              Width           =   3540
                              Begin VB.PictureBox mapPreview 
                                 Height          =   2670
                                 Index           =   1
                                 Left            =   120
                                 ScaleHeight     =   2610
                                 ScaleWidth      =   3270
                                 TabIndex        =   243
                                 Top             =   240
                                 Width           =   3330
                              End
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elHolder 
                           Height          =   6735
                           Index           =   4
                           Left            =   30
                           TabIndex        =   246
                           TabStop         =   0   'False
                           Top             =   30
                           Width           =   2820
                           _cx             =   4974
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
                           GridCols        =   1
                           Frame           =   3
                           FrameStyle      =   0
                           FrameWidth      =   1
                           FrameColor      =   -2147483628
                           FrameShadow     =   -2147483632
                           FloodStyle      =   1
                           _GridInfo       =   $"frmConfig.frx":419AD
                           AccessibleName  =   ""
                           AccessibleDescription=   ""
                           AccessibleValue =   ""
                           AccessibleRole  =   9
                           Begin TatukGIS_DK.XGIS_ViewerWnd MapProduct 
                              Height          =   6135
                              Index           =   1
                              Left            =   15
                              TabIndex        =   247
                              Top             =   15
                              Width           =   11310
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
                              SelectionPattern=   "frmConfig.frx":419EA
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
                              ScrollBars      =   3
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
                           Begin C1SizerLibCtl.C1Elastic elHolder 
                              Height          =   525
                              Index           =   5
                              Left            =   15
                              TabIndex        =   248
                              TabStop         =   0   'False
                              Top             =   6165
                              Width           =   11310
                              _cx             =   19950
                              _cy             =   926
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
                              BorderWidth     =   0
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
                              Begin VB.Frame FraMapTools 
                                 Caption         =   "Map Tools:"
                                 Height          =   525
                                 Index           =   1
                                 Left            =   60
                                 TabIndex        =   251
                                 Top             =   -30
                                 Width           =   1140
                                 Begin VB.CommandButton cmdZoomin 
                                    Caption         =   "+"
                                    Height          =   240
                                    Index           =   1
                                    Left            =   135
                                    TabIndex        =   253
                                    Top             =   210
                                    Width           =   330
                                 End
                                 Begin VB.CommandButton cmdZoomOut 
                                    Caption         =   "-"
                                    Height          =   240
                                    Index           =   1
                                    Left            =   600
                                    TabIndex        =   252
                                    Top             =   210
                                    Width           =   330
                                 End
                              End
                              Begin VB.CommandButton cmdLoadMap 
                                 Caption         =   "Load Map..."
                                 Height          =   405
                                 Index           =   1
                                 Left            =   1290
                                 TabIndex        =   250
                                 Top             =   90
                                 Width           =   1320
                              End
                              Begin VB.CommandButton cmdCreatePreview 
                                 Caption         =   "Create Preview"
                                 Height          =   405
                                 Index           =   1
                                 Left            =   2640
                                 TabIndex        =   249
                                 Top             =   90
                                 Width           =   1320
                              End
                           End
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elMapDetails 
                        Height          =   6795
                        Index           =   2
                        Left            =   3915
                        TabIndex        =   254
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   2880
                        _cx             =   5080
                        _cy             =   11986
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
                        GridRows        =   1
                        GridCols        =   2
                        Frame           =   3
                        FrameStyle      =   0
                        FrameWidth      =   1
                        FrameColor      =   -2147483628
                        FrameShadow     =   -2147483632
                        FloodStyle      =   1
                        _GridInfo       =   $"frmConfig.frx":4C2FC
                        AccessibleName  =   ""
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   9
                        Begin C1SizerLibCtl.C1Elastic elHolder 
                           Height          =   6735
                           Index           =   6
                           Left            =   2175
                           TabIndex        =   255
                           TabStop         =   0   'False
                           Top             =   30
                           Width           =   675
                           _cx             =   1191
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
                           Align           =   0
                           AutoSizeChildren=   0
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
                           Begin VB.Frame FraDescription 
                              Caption         =   "Description:"
                              Height          =   2985
                              Index           =   2
                              Left            =   90
                              TabIndex        =   258
                              Top             =   60
                              Width           =   3540
                              Begin VB.TextBox txtMapDescription 
                                 Height          =   2580
                                 Index           =   2
                                 Left            =   60
                                 MultiLine       =   -1  'True
                                 ScrollBars      =   2  'Vertical
                                 TabIndex        =   259
                                 Top             =   225
                                 Width           =   3420
                              End
                           End
                           Begin VB.Frame FraMapPreview 
                              Caption         =   "Map Preview:"
                              Height          =   3045
                              Index           =   2
                              Left            =   90
                              TabIndex        =   256
                              Top             =   3150
                              Width           =   3540
                              Begin VB.PictureBox mapPreview 
                                 Height          =   2670
                                 Index           =   2
                                 Left            =   120
                                 ScaleHeight     =   2610
                                 ScaleWidth      =   3270
                                 TabIndex        =   257
                                 Top             =   240
                                 Width           =   3330
                              End
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elHolder 
                           Height          =   6735
                           Index           =   7
                           Left            =   30
                           TabIndex        =   260
                           TabStop         =   0   'False
                           Top             =   30
                           Width           =   2820
                           _cx             =   4974
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
                           GridCols        =   1
                           Frame           =   3
                           FrameStyle      =   0
                           FrameWidth      =   1
                           FrameColor      =   -2147483628
                           FrameShadow     =   -2147483632
                           FloodStyle      =   1
                           _GridInfo       =   $"frmConfig.frx":4C33B
                           AccessibleName  =   ""
                           AccessibleDescription=   ""
                           AccessibleValue =   ""
                           AccessibleRole  =   9
                           Begin TatukGIS_DK.XGIS_ViewerWnd MapProduct 
                              Height          =   6135
                              Index           =   2
                              Left            =   15
                              TabIndex        =   261
                              Top             =   15
                              Width           =   11325
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
                              SelectionPattern=   "frmConfig.frx":4C378
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
                              ScrollBars      =   3
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
                           Begin C1SizerLibCtl.C1Elastic elHolder 
                              Height          =   525
                              Index           =   8
                              Left            =   15
                              TabIndex        =   262
                              TabStop         =   0   'False
                              Top             =   6165
                              Width           =   11325
                              _cx             =   19976
                              _cy             =   926
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
                              BorderWidth     =   0
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
                              Begin VB.Frame FraMapTools 
                                 Caption         =   "Map Tools:"
                                 Height          =   525
                                 Index           =   2
                                 Left            =   60
                                 TabIndex        =   265
                                 Top             =   -30
                                 Width           =   1140
                                 Begin VB.CommandButton cmdZoomin 
                                    Caption         =   "+"
                                    Height          =   240
                                    Index           =   2
                                    Left            =   135
                                    TabIndex        =   267
                                    Top             =   210
                                    Width           =   330
                                 End
                                 Begin VB.CommandButton cmdZoomOut 
                                    Caption         =   "-"
                                    Height          =   240
                                    Index           =   2
                                    Left            =   600
                                    TabIndex        =   266
                                    Top             =   210
                                    Width           =   330
                                 End
                              End
                              Begin VB.CommandButton cmdLoadMap 
                                 Caption         =   "Load Map..."
                                 Height          =   405
                                 Index           =   2
                                 Left            =   1290
                                 TabIndex        =   264
                                 Top             =   90
                                 Width           =   1320
                              End
                              Begin VB.CommandButton cmdCreatePreview 
                                 Caption         =   "Create Preview"
                                 Height          =   405
                                 Index           =   2
                                 Left            =   2640
                                 TabIndex        =   263
                                 Top             =   90
                                 Width           =   1320
                              End
                           End
                        End
                     End
                     Begin C1SizerLibCtl.C1Elastic elMapDetails 
                        Height          =   6795
                        Index           =   3
                        Left            =   4215
                        TabIndex        =   268
                        TabStop         =   0   'False
                        Top             =   330
                        Width           =   2880
                        _cx             =   5080
                        _cy             =   11986
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
                        GridRows        =   1
                        GridCols        =   2
                        Frame           =   3
                        FrameStyle      =   0
                        FrameWidth      =   1
                        FrameColor      =   -2147483628
                        FrameShadow     =   -2147483632
                        FloodStyle      =   1
                        _GridInfo       =   $"frmConfig.frx":56C8A
                        AccessibleName  =   ""
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   9
                        Begin C1SizerLibCtl.C1Elastic elHolder 
                           Height          =   6735
                           Index           =   9
                           Left            =   2175
                           TabIndex        =   269
                           TabStop         =   0   'False
                           Top             =   30
                           Width           =   675
                           _cx             =   1191
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
                           Align           =   0
                           AutoSizeChildren=   0
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
                           Begin VB.Frame FraDescription 
                              Caption         =   "Description:"
                              Height          =   2985
                              Index           =   3
                              Left            =   90
                              TabIndex        =   272
                              Top             =   60
                              Width           =   3540
                              Begin VB.TextBox txtMapDescription 
                                 Height          =   2580
                                 Index           =   3
                                 Left            =   60
                                 MultiLine       =   -1  'True
                                 ScrollBars      =   2  'Vertical
                                 TabIndex        =   273
                                 Top             =   225
                                 Width           =   3420
                              End
                           End
                           Begin VB.Frame FraMapPreview 
                              Caption         =   "Map Preview:"
                              Height          =   3045
                              Index           =   3
                              Left            =   90
                              TabIndex        =   270
                              Top             =   3150
                              Width           =   3540
                              Begin VB.PictureBox mapPreview 
                                 Height          =   2670
                                 Index           =   3
                                 Left            =   120
                                 ScaleHeight     =   2610
                                 ScaleWidth      =   3270
                                 TabIndex        =   271
                                 Top             =   240
                                 Width           =   3330
                              End
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elHolder 
                           Height          =   6735
                           Index           =   10
                           Left            =   30
                           TabIndex        =   274
                           TabStop         =   0   'False
                           Top             =   30
                           Width           =   2820
                           _cx             =   4974
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
                           GridCols        =   1
                           Frame           =   3
                           FrameStyle      =   0
                           FrameWidth      =   1
                           FrameColor      =   -2147483628
                           FrameShadow     =   -2147483632
                           FloodStyle      =   1
                           _GridInfo       =   $"frmConfig.frx":56CC9
                           AccessibleName  =   ""
                           AccessibleDescription=   ""
                           AccessibleValue =   ""
                           AccessibleRole  =   9
                           Begin TatukGIS_DK.XGIS_ViewerWnd MapProduct 
                              Height          =   6120
                              Index           =   3
                              Left            =   15
                              TabIndex        =   275
                              Top             =   15
                              Width           =   15000
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
                              SelectionPattern=   "frmConfig.frx":56D06
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
                              ScrollBars      =   3
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
                           Begin C1SizerLibCtl.C1Elastic elHolder 
                              Height          =   540
                              Index           =   11
                              Left            =   15
                              TabIndex        =   276
                              TabStop         =   0   'False
                              Top             =   6150
                              Width           =   15000
                              _cx             =   26458
                              _cy             =   953
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
                              BorderWidth     =   0
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
                              Begin VB.Frame FraMapTools 
                                 Caption         =   "Map Tools:"
                                 Height          =   525
                                 Index           =   3
                                 Left            =   60
                                 TabIndex        =   279
                                 Top             =   -30
                                 Width           =   1140
                                 Begin VB.CommandButton cmdZoomin 
                                    Caption         =   "+"
                                    Height          =   240
                                    Index           =   3
                                    Left            =   135
                                    TabIndex        =   281
                                    Top             =   210
                                    Width           =   330
                                 End
                                 Begin VB.CommandButton cmdZoomOut 
                                    Caption         =   "-"
                                    Height          =   240
                                    Index           =   3
                                    Left            =   600
                                    TabIndex        =   280
                                    Top             =   210
                                    Width           =   330
                                 End
                              End
                              Begin VB.CommandButton cmdLoadMap 
                                 Caption         =   "Load Map..."
                                 Height          =   405
                                 Index           =   3
                                 Left            =   1290
                                 TabIndex        =   278
                                 Top             =   90
                                 Width           =   1320
                              End
                              Begin VB.CommandButton cmdCreatePreview 
                                 Caption         =   "Create Preview"
                                 Height          =   405
                                 Index           =   3
                                 Left            =   2640
                                 TabIndex        =   277
                                 Top             =   90
                                 Width           =   1320
                              End
                           End
                        End
                     End
                  End
               End
               Begin C1SizerLibCtl.C1Elastic ElGISURLLayers 
                  Height          =   6975
                  Left            =   2400
                  TabIndex        =   282
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   1365
                  _cx             =   2408
                  _cy             =   12303
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
                  GridCols        =   1
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmConfig.frx":61618
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin DXDBGRIDLibCtl.dxDBGrid dxGeoMarks 
                     Height          =   5175
                     Left            =   30
                     OleObjectBlob   =   "frmConfig.frx":6165A
                     TabIndex        =   283
                     Top             =   30
                     Width           =   1305
                  End
                  Begin C1SizerLibCtl.C1Elastic elAddMods 
                     Height          =   1725
                     Index           =   6
                     Left            =   30
                     TabIndex        =   284
                     TabStop         =   0   'False
                     Top             =   5220
                     Width           =   1305
                     _cx             =   2302
                     _cy             =   3043
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
                     BorderWidth     =   2
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
                     Begin VB.CommandButton cmdUsrNew 
                        Caption         =   "New"
                        Height          =   285
                        Index           =   3
                        Left            =   9840
                        TabIndex        =   291
                        Top             =   360
                        Width           =   735
                     End
                     Begin VB.CommandButton cmdUsrEdit 
                        Caption         =   "Edit"
                        Height          =   285
                        Index           =   3
                        Left            =   9840
                        TabIndex        =   290
                        Top             =   675
                        Width           =   735
                     End
                     Begin VB.CommandButton cmdUsrDelete 
                        Caption         =   "Delete"
                        Height          =   285
                        Index           =   3
                        Left            =   9840
                        TabIndex        =   289
                        Top             =   990
                        Width           =   735
                     End
                     Begin VB.TextBox txtGeomarks 
                        DataField       =   "Y"
                        Height          =   285
                        Index           =   2
                        Left            =   2580
                        TabIndex        =   288
                        Top             =   750
                        Width           =   1680
                     End
                     Begin VB.TextBox txtGeomarks 
                        DataField       =   "X"
                        Height          =   285
                        Index           =   1
                        Left            =   225
                        TabIndex        =   287
                        Top             =   720
                        Width           =   1680
                     End
                     Begin VB.TextBox txtGeomarks 
                        DataField       =   "Name"
                        Height          =   285
                        Index           =   0
                        Left            =   540
                        TabIndex        =   286
                        Top             =   360
                        Width           =   3660
                     End
                     Begin VB.TextBox txtGeomarks 
                        DataField       =   "Description"
                        Height          =   1365
                        Index           =   3
                        Left            =   4275
                        MultiLine       =   -1  'True
                        ScrollBars      =   2  'Vertical
                        TabIndex        =   285
                        Top             =   360
                        Width           =   5415
                     End
                     Begin VB.Label lblLBL 
                        AutoSize        =   -1  'True
                        Caption         =   "Y:"
                        Height          =   195
                        Index           =   9
                        Left            =   2115
                        TabIndex        =   295
                        Tag             =   "Nam"
                        Top             =   765
                        Width           =   150
                     End
                     Begin VB.Label lblLBL 
                        AutoSize        =   -1  'True
                        Caption         =   "X:"
                        Height          =   195
                        Index           =   8
                        Left            =   0
                        TabIndex        =   294
                        Tag             =   "Nam"
                        Top             =   765
                        Width           =   150
                     End
                     Begin VB.Label lblLBL 
                        AutoSize        =   -1  'True
                        Caption         =   "Name:"
                        Height          =   195
                        Index           =   7
                        Left            =   0
                        TabIndex        =   293
                        Tag             =   "Nam"
                        Top             =   405
                        Width           =   465
                     End
                     Begin VB.Label lblLBL 
                        AutoSize        =   -1  'True
                        Caption         =   "Description:"
                        Height          =   195
                        Index           =   6
                        Left            =   4275
                        TabIndex        =   292
                        Top             =   45
                        Width           =   840
                     End
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elGISAttrAccess 
                  Height          =   6975
                  Left            =   2100
                  TabIndex        =   296
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   1365
                  _cx             =   2408
                  _cy             =   12303
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
                  GridCols        =   1
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmConfig.frx":62302
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin C1SizerLibCtl.C1Elastic elHolder 
                     Height          =   1920
                     Index           =   12
                     Left            =   30
                     TabIndex        =   297
                     TabStop         =   0   'False
                     Top             =   5025
                     Width           =   1305
                     _cx             =   2302
                     _cy             =   3387
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
                     Begin VB.CommandButton cmdUsrDelete 
                        Caption         =   "Delete"
                        Height          =   285
                        Index           =   2
                        Left            =   9900
                        TabIndex        =   300
                        Top             =   1485
                        Width           =   735
                     End
                     Begin VB.CommandButton cmdUsrEdit 
                        Caption         =   "Edit"
                        Height          =   285
                        Index           =   2
                        Left            =   9900
                        TabIndex        =   299
                        Top             =   1170
                        Width           =   735
                     End
                     Begin VB.CommandButton cmdUsrNew 
                        Caption         =   "New"
                        Height          =   285
                        Index           =   2
                        Left            =   9900
                        TabIndex        =   298
                        Top             =   855
                        Width           =   735
                     End
                     Begin C1SizerLibCtl.C1Tab ciTabDataAccess 
                        Height          =   1770
                        Index           =   0
                        Left            =   0
                        TabIndex        =   301
                        Top             =   0
                        Width           =   9825
                        _cx             =   17330
                        _cy             =   3122
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
                        Caption         =   "General|Accessibility|Inter Active Layer"
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
                        Begin C1SizerLibCtl.C1Elastic elAddMods 
                           Height          =   1395
                           Index           =   3
                           Left            =   45
                           TabIndex        =   302
                           TabStop         =   0   'False
                           Top             =   330
                           Width           =   9735
                           _cx             =   17171
                           _cy             =   2461
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
                           BorderWidth     =   2
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
                           Begin VB.CheckBox chkGisGrid 
                              Caption         =   "Visible"
                              DataField       =   "visible"
                              Height          =   195
                              Index           =   0
                              Left            =   8505
                              TabIndex        =   305
                              Top             =   360
                              Width           =   825
                           End
                           Begin VB.TextBox txtGISGrid 
                              DataField       =   "alias"
                              Height          =   330
                              Index           =   1
                              Left            =   4275
                              TabIndex        =   304
                              Top             =   360
                              Width           =   4065
                           End
                           Begin VB.TextBox txtGISGrid 
                              DataField       =   "name"
                              Height          =   330
                              Index           =   0
                              Left            =   135
                              TabIndex        =   303
                              Top             =   360
                              Width           =   4065
                           End
                           Begin VB.Label lblLBL 
                              AutoSize        =   -1  'True
                              Caption         =   "Alias:"
                              Height          =   195
                              Index           =   1
                              Left            =   4275
                              TabIndex        =   307
                              Top             =   135
                              Width           =   375
                           End
                           Begin VB.Label lblLBL 
                              AutoSize        =   -1  'True
                              Caption         =   "Name:"
                              Height          =   195
                              Index           =   0
                              Left            =   135
                              TabIndex        =   306
                              Tag             =   "Nam"
                              Top             =   135
                              Width           =   465
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elAddMods 
                           Height          =   1395
                           Index           =   4
                           Left            =   10470
                           TabIndex        =   308
                           TabStop         =   0   'False
                           Top             =   330
                           Width           =   9735
                           _cx             =   17171
                           _cy             =   2461
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
                           BorderWidth     =   2
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
                           Begin VB.TextBox txtGISGrid 
                              DataField       =   "excludedFlds"
                              Height          =   960
                              Index           =   4
                              Left            =   4140
                              MultiLine       =   -1  'True
                              ScrollBars      =   2  'Vertical
                              TabIndex        =   312
                              Top             =   315
                              Width           =   5460
                           End
                           Begin VB.TextBox txtGISGrid 
                              DataField       =   "MaxRec"
                              Height          =   330
                              Index           =   3
                              Left            =   1800
                              TabIndex        =   311
                              Top             =   945
                              Width           =   2175
                           End
                           Begin VB.TextBox txtGISGrid 
                              DataField       =   "warninglevel"
                              Height          =   330
                              Index           =   2
                              Left            =   1800
                              TabIndex        =   310
                              Top             =   315
                              Width           =   2175
                           End
                           Begin VB.CheckBox chkGisGrid 
                              Caption         =   "Use Dataset Browsing Warning"
                              DataField       =   "datasetwarning"
                              Height          =   375
                              Index           =   1
                              Left            =   90
                              TabIndex        =   309
                              Top             =   135
                              Width           =   1680
                           End
                           Begin VB.Label lblLBL 
                              AutoSize        =   -1  'True
                              Caption         =   "Hidden Attribute Fields (comma separated):"
                              Height          =   195
                              Index           =   4
                              Left            =   4140
                              TabIndex        =   315
                              Top             =   45
                              Width           =   3030
                           End
                           Begin VB.Label lblLBL 
                              AutoSize        =   -1  'True
                              Caption         =   "Max Record Display Number:"
                              Height          =   195
                              Index           =   3
                              Left            =   1800
                              TabIndex        =   314
                              Top             =   720
                              Width           =   2070
                           End
                           Begin VB.Label lblLBL 
                              AutoSize        =   -1  'True
                              Caption         =   "Initial Warning level:"
                              Height          =   195
                              Index           =   2
                              Left            =   1800
                              TabIndex        =   313
                              Top             =   45
                              Width           =   1425
                           End
                        End
                        Begin C1SizerLibCtl.C1Elastic elAddMods 
                           Height          =   1395
                           Index           =   5
                           Left            =   10770
                           TabIndex        =   316
                           TabStop         =   0   'False
                           Top             =   330
                           Width           =   9735
                           _cx             =   17171
                           _cy             =   2461
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
                           BorderWidth     =   2
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
                           Begin VB.TextBox txtGISGrid 
                              DataField       =   "URLLayerField"
                              Height          =   330
                              Index           =   5
                              Left            =   0
                              TabIndex        =   319
                              Top             =   990
                              Width           =   5280
                           End
                           Begin VB.CheckBox chkGisGrid 
                              Caption         =   "Use Auto Run when using info tool"
                              DataField       =   "autoRunUrls"
                              Height          =   375
                              Index           =   3
                              Left            =   0
                              TabIndex        =   318
                              Top             =   360
                              Width           =   2805
                           End
                           Begin VB.CheckBox chkGisGrid 
                              Caption         =   "Is Inter Active Layer"
                              DataField       =   "isURLLayer"
                              Height          =   375
                              Index           =   2
                              Left            =   0
                              TabIndex        =   317
                              Top             =   0
                              Width           =   1770
                           End
                           Begin VB.Label lblLBL 
                              AutoSize        =   -1  'True
                              Caption         =   "Url/Inter Active Layer Field Name:"
                              Height          =   195
                              Index           =   5
                              Left            =   0
                              TabIndex        =   320
                              Tag             =   "Nam"
                              Top             =   765
                              Width           =   2400
                           End
                        End
                     End
                  End
                  Begin DXDBGRIDLibCtl.dxDBGrid dxGISAttrAccess 
                     Height          =   4980
                     Left            =   30
                     OleObjectBlob   =   "frmConfig.frx":62344
                     TabIndex        =   321
                     Top             =   30
                     Width           =   1305
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elDPDef 
                  Height          =   6975
                  Left            =   45
                  TabIndex        =   322
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   1365
                  _cx             =   2408
                  _cy             =   12303
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
                  Begin C1SizerLibCtl.C1Elastic c1Main 
                     Height          =   6975
                     Left            =   0
                     TabIndex        =   323
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   1365
                     _cx             =   2408
                     _cy             =   12303
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
                     GridCols        =   1
                     Frame           =   3
                     FrameStyle      =   0
                     FrameWidth      =   1
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   $"frmConfig.frx":62FEC
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin DXDBGRIDLibCtl.dxDBGrid dxDataPacks 
                        Height          =   6315
                        Left            =   30
                        OleObjectBlob   =   "frmConfig.frx":6302D
                        TabIndex        =   324
                        Top             =   30
                        Width           =   1305
                     End
                     Begin C1SizerLibCtl.C1Elastic elNavs 
                        Height          =   585
                        Left            =   30
                        TabIndex        =   325
                        TabStop         =   0   'False
                        Top             =   6360
                        Width           =   1305
                        _cx             =   2302
                        _cy             =   1032
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
                        GridCols        =   4
                        Frame           =   3
                        FrameStyle      =   0
                        FrameWidth      =   1
                        FrameColor      =   -2147483628
                        FrameShadow     =   -2147483632
                        FloodStyle      =   1
                        _GridInfo       =   $"frmConfig.frx":63CD5
                        AccessibleName  =   ""
                        AccessibleDescription=   ""
                        AccessibleValue =   ""
                        AccessibleRole  =   9
                        Begin VB.CommandButton cmdDo 
                           Caption         =   "Dowload Data Packs"
                           Enabled         =   0   'False
                           Height          =   555
                           Left            =   1035
                           TabIndex        =   328
                           Top             =   15
                           Width           =   255
                        End
                        Begin VB.CommandButton cmdSET 
                           Caption         =   "Save Data Pack Definitions"
                           Enabled         =   0   'False
                           Height          =   555
                           Left            =   1035
                           TabIndex        =   327
                           Top             =   15
                           Width           =   255
                        End
                        Begin VB.CommandButton cmdGetDataPckDef 
                           Caption         =   "Get Data Pack Definitions"
                           Height          =   555
                           Left            =   1035
                           TabIndex        =   326
                           Top             =   15
                           Width           =   255
                        End
                     End
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elAddons 
            Height          =   7350
            Left            =   10785
            TabIndex        =   329
            TabStop         =   0   'False
            Top             =   465
            Width           =   9540
            _cx             =   16828
            _cy             =   12965
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
            _GridInfo       =   $"frmConfig.frx":63D23
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame FraActiveCore 
               Caption         =   "Active modules"
               Height          =   1365
               Index           =   1
               Left            =   90
               TabIndex        =   330
               Top             =   90
               Width           =   9360
            End
            Begin C1SizerLibCtl.C1Tab c1TabAdditionalmodules 
               Height          =   5745
               Left            =   90
               TabIndex        =   331
               Top             =   1515
               Width           =   9360
               _cx             =   16510
               _cy             =   10134
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
               Caption         =   "W3|Commodity/NFI|Mine Action"
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
               Begin C1SizerLibCtl.C1Elastic elAddMods 
                  Height          =   5430
                  Index           =   1
                  Left            =   1980
                  TabIndex        =   332
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   1245
                  _cx             =   2196
                  _cy             =   9578
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
                  BorderWidth     =   2
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
               End
               Begin C1SizerLibCtl.C1Elastic elAddMods 
                  Height          =   5430
                  Index           =   0
                  Left            =   45
                  TabIndex        =   333
                  TabStop         =   0   'False
                  Top             =   330
                  Width           =   1245
                  _cx             =   2196
                  _cy             =   9578
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
                  BorderWidth     =   2
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
               End
            End
         End
         Begin C1SizerLibCtl.C1Tab C1TFileSynch 
            Height          =   7350
            Left            =   10185
            TabIndex        =   334
            Top             =   465
            Width           =   9540
            _cx             =   16828
            _cy             =   12965
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
            Caption         =   "File Synch|Data Synch"
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
            Begin C1SizerLibCtl.C1Elastic elDatasynch 
               Height          =   7035
               Left            =   45
               TabIndex        =   335
               TabStop         =   0   'False
               Top             =   330
               Width           =   1425
               _cx             =   2514
               _cy             =   12409
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
               Caption         =   "Available Synch Datasets"
               Align           =   0
               AutoSizeChildren=   8
               BorderWidth     =   0
               ChildSpacing    =   2
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
               GridRows        =   5
               GridCols        =   3
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmConfig.frx":63D66
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.Frame FraDatasetDetails 
                  Caption         =   "Dataset Details:"
                  Height          =   6090
                  Left            =   1320
                  TabIndex        =   339
                  Top             =   0
                  Width           =   105
                  Begin VB.CheckBox chkSynchDSOption 
                     Caption         =   "Append Local Data"
                     Height          =   225
                     Index           =   3
                     Left            =   180
                     TabIndex        =   348
                     Top             =   5160
                     Width           =   2355
                  End
                  Begin VB.CommandButton cmdDBSynchSource 
                     Caption         =   "..."
                     Height          =   285
                     Left            =   2430
                     TabIndex        =   347
                     Top             =   4050
                     Width           =   315
                  End
                  Begin VB.TextBox txtDSSynchDetails 
                     Height          =   315
                     Index           =   3
                     Left            =   150
                     TabIndex        =   346
                     Top             =   4020
                     Width           =   2265
                  End
                  Begin VB.CheckBox chkSynchDSOption 
                     Caption         =   "Use Auto Update"
                     Height          =   225
                     Index           =   2
                     Left            =   180
                     TabIndex        =   345
                     Top             =   4920
                     Width           =   2355
                  End
                  Begin VB.CheckBox chkSynchDSOption 
                     Caption         =   "Allow Write"
                     Height          =   225
                     Index           =   1
                     Left            =   180
                     TabIndex        =   344
                     Top             =   4680
                     Width           =   2355
                  End
                  Begin VB.CheckBox chkSynchDSOption 
                     Caption         =   "Is Geo Table"
                     Height          =   225
                     Index           =   0
                     Left            =   180
                     TabIndex        =   343
                     Top             =   4440
                     Width           =   2355
                  End
                  Begin VB.TextBox txtDSSynchDetails 
                     DataField       =   "sDescription"
                     Height          =   2205
                     Index           =   2
                     Left            =   120
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   342
                     Top             =   1500
                     Width           =   2625
                  End
                  Begin VB.TextBox txtDSSynchDetails 
                     DataField       =   "sTableName"
                     Height          =   315
                     Index           =   1
                     Left            =   120
                     TabIndex        =   341
                     Top             =   960
                     Width           =   2625
                  End
                  Begin VB.TextBox txtDSSynchDetails 
                     DataField       =   "sName"
                     Height          =   315
                     Index           =   0
                     Left            =   120
                     TabIndex        =   340
                     Top             =   420
                     Width           =   2625
                  End
                  Begin VB.Label lblDatasetDetails 
                     AutoSize        =   -1  'True
                     Caption         =   "DB Source:"
                     Height          =   195
                     Index           =   3
                     Left            =   150
                     TabIndex        =   352
                     Top             =   3810
                     Width           =   825
                  End
                  Begin VB.Label lblDatasetDetails 
                     AutoSize        =   -1  'True
                     Caption         =   "Description:"
                     Height          =   195
                     Index           =   2
                     Left            =   180
                     TabIndex        =   351
                     Top             =   1290
                     Width           =   840
                  End
                  Begin VB.Label lblDatasetDetails 
                     AutoSize        =   -1  'True
                     Caption         =   "Table Name:"
                     Height          =   195
                     Index           =   1
                     Left            =   120
                     TabIndex        =   350
                     Top             =   750
                     Width           =   915
                  End
                  Begin VB.Label lblDatasetDetails 
                     AutoSize        =   -1  'True
                     Caption         =   "Name:"
                     Height          =   195
                     Index           =   0
                     Left            =   120
                     TabIndex        =   349
                     Top             =   210
                     Width           =   465
                  End
               End
               Begin VB.CommandButton cmdUsrDelete 
                  Caption         =   "Delete"
                  Height          =   285
                  Index           =   4
                  Left            =   1320
                  TabIndex        =   338
                  Top             =   6750
                  Width           =   105
               End
               Begin VB.CommandButton cmdUsrEdit 
                  Caption         =   "Edit"
                  Height          =   285
                  Index           =   4
                  Left            =   1320
                  TabIndex        =   337
                  Top             =   6435
                  Width           =   105
               End
               Begin VB.CommandButton cmdUsrNew 
                  Caption         =   "New"
                  Height          =   285
                  Index           =   4
                  Left            =   1320
                  TabIndex        =   336
                  Top             =   6120
                  Width           =   105
               End
               Begin DXDBGRIDLibCtl.dxDBGrid dxSynchDatasets 
                  Height          =   6795
                  Left            =   0
                  OleObjectBlob   =   "frmConfig.frx":63DDB
                  TabIndex        =   353
                  Top             =   240
                  Width           =   1425
               End
            End
            Begin C1SizerLibCtl.C1Elastic elflSynch 
               Height          =   7035
               Left            =   -2070
               TabIndex        =   354
               TabStop         =   0   'False
               Top             =   330
               Width           =   1425
               _cx             =   2514
               _cy             =   12409
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
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmConfig.frx":6769C
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.Frame FraFileFolders 
                  Caption         =   "File Folders:"
                  Height          =   3660
                  Left            =   30
                  TabIndex        =   355
                  Top             =   30
                  Width           =   1365
                  Begin VB.CommandButton cmdGetSunchFileFolders 
                     Caption         =   "Get Synch Folders"
                     Height          =   345
                     Left            =   225
                     TabIndex        =   359
                     Top             =   2775
                     Width           =   1635
                  End
                  Begin VB.TextBox txtfileSynchUrl 
                     Height          =   330
                     Left            =   1890
                     TabIndex        =   358
                     Text            =   "http://www.immap.org"
                     Top             =   2790
                     Width           =   6120
                  End
                  Begin VB.CommandButton cmdSetSyncFileFolders 
                     Caption         =   "Save Synch Folders"
                     Height          =   345
                     Left            =   240
                     TabIndex        =   357
                     Top             =   3180
                     Width           =   1635
                  End
                  Begin VB.TextBox txtSetSynchFolders 
                     Height          =   330
                     Left            =   1890
                     TabIndex        =   356
                     Text            =   "http://www.immap.org"
                     Top             =   3180
                     Width           =   6135
                  End
                  Begin DXDBGRIDLibCtl.dxDBGrid dxSynch 
                     Height          =   2475
                     Left            =   150
                     OleObjectBlob   =   "frmConfig.frx":676DE
                     TabIndex        =   360
                     Top             =   240
                     Width           =   13155
                  End
               End
               Begin C1SizerLibCtl.C1Elastic elHolder 
                  Height          =   3300
                  Index           =   13
                  Left            =   30
                  TabIndex        =   361
                  TabStop         =   0   'False
                  Top             =   3705
                  Width           =   1365
                  _cx             =   2408
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
                  AutoSizeChildren=   8
                  BorderWidth     =   2
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
                  GridCols        =   3
                  Frame           =   3
                  FrameStyle      =   0
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   $"frmConfig.frx":68386
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin VB.CommandButton cmdFileSynch 
                     Caption         =   "Check Sync"
                     Height          =   495
                     Index           =   0
                     Left            =   14940
                     TabIndex        =   363
                     Top             =   3000
                     Width           =   225
                  End
                  Begin VB.CommandButton cmdFileSynch 
                     Caption         =   "Get Folder Content"
                     Height          =   495
                     Index           =   1
                     Left            =   14940
                     TabIndex        =   362
                     Top             =   3000
                     Width           =   225
                  End
                  Begin C1SizerLibCtl.C1Elastic elHolder 
                     Height          =   2955
                     Index           =   14
                     Left            =   30
                     TabIndex        =   364
                     TabStop         =   0   'False
                     Top             =   30
                     Width           =   15135
                     _cx             =   26696
                     _cy             =   5212
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
                     Caption         =   "Synch Files:"
                     Align           =   0
                     AutoSizeChildren=   8
                     BorderWidth     =   1
                     ChildSpacing    =   2
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
                     GridRows        =   2
                     GridCols        =   1
                     Frame           =   3
                     FrameStyle      =   0
                     FrameWidth      =   1
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   $"frmConfig.frx":683DA
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin DXDBGRIDLibCtl.dxDBGrid dxSynchFolderURL 
                        Height          =   2640
                        Left            =   15
                        OleObjectBlob   =   "frmConfig.frx":6841A
                        TabIndex        =   365
                        Top             =   300
                        Width           =   15105
                     End
                  End
               End
            End
         End
      End
      Begin VB.Image Image 
         Height          =   450
         Left            =   45
         Picture         =   "frmConfig.frx":690C2
         Stretch         =   -1  'True
         ToolTipText     =   "Visit www.immap.org"
         Top             =   7905
         Width           =   15
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WebSite As String '= "http://www.immap.org/"

Private Declare Function CoCreateGuid _
                Lib "ole32" (id As Any) As Long

Dim WithEvents pGetUsersThread As MThreadVB.Thread
Attribute pGetUsersThread.VB_VarHelpID = -1
Dim WithEvents pLoadUsr As MThreadVB.Thread
Attribute pLoadUsr.VB_VarHelpID = -1

Private oSynchExplorer As Object
Private m_bServerConnection As Boolean
Private m_oAES As New clsAES
Private m_sKey As String
Private RSAppSetting As New ADODB.Recordset
Private sTablePrefix As String

Public Sub SetTablePrefix(sPassedTablePrefix As String)
        '<EhHeader>
        On Error GoTo SetTablePrefix_Err
        '</EhHeader>
100     sTablePrefix = sPassedTablePrefix
        '<EhFooter>
        Exit Sub

SetTablePrefix_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.SetTablePrefix " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function CreateGUID() As String
        '<EhHeader>
        On Error GoTo CreateGUID_Err
        '</EhHeader>
        Dim id(0 To 15) As Byte
        Dim Cnt As Long, GUID As String

100     If CoCreateGuid(id(0)) = 0 Then

102         For Cnt = 0 To 15
104             CreateGUID = CreateGUID + IIf(id(Cnt) < 16, "0", "") + Hex$(id(Cnt))
106         Next Cnt

108         CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
        Else
110         MsgBox "Error while creating GUID!"
        End If

        '<EhFooter>
        Exit Function

CreateGUID_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.CreateGUID " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub chkCoreModule_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo chkCoreModule_Click_Err
        '</EhHeader>

100     Select Case Index
    
            Case 0
102             c1TabCoreModules.TabVisible(Index) = IIf(chkCoreModule(Index).Value = vbChecked, True, False)

104             If c1TabCoreModules.TabVisible(0) Then
106                 c1TabCoreModules.CurrTab = 0
108             ElseIf c1TabCoreModules.TabVisible(1) Then
110                 c1TabCoreModules.CurrTab = 1
                End If

112         Case 1
114             c1TabCoreModules.TabVisible(Index) = IIf(chkCoreModule(Index).Value = vbChecked, True, False)

116             If c1TabCoreModules.TabVisible(1) Then
118                 c1TabCoreModules.CurrTab = 1
120             ElseIf c1TabCoreModules.TabVisible(0) Then
122                 c1TabCoreModules.CurrTab = 0
                End If
                
        End Select
        
124     If c1TabCoreModules.TabVisible(0) = False And c1TabCoreModules.TabVisible(1) = False Then
126         c1TabCoreModules.Visible = False
        Else
128         c1TabCoreModules.Visible = True
        End If

        '<EhFooter>
        Exit Sub

chkCoreModule_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.chkCoreModule_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub chkShowlegendTab_Click()
        '<EhHeader>
        On Error GoTo chkShowlegendTab_Click_Err
        '</EhHeader>
100     C1TabOPS.TabVisible(1) = IIf(chkShowlegendTab.Value = vbChecked, True, False)
      
        '<EhFooter>
        Exit Sub

chkShowlegendTab_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.chkShowlegendTab_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkShowMapLibraryTab_Click()
        '<EhHeader>
        On Error GoTo chkShowMapLibraryTab_Click_Err
        '</EhHeader>
100     C1TabOPS.TabVisible(2) = IIf(chkShowMapLibraryTab.Value = vbChecked, True, False)
        '<EhFooter>
        Exit Sub

chkShowMapLibraryTab_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.chkShowMapLibraryTab_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkShowMATab_Click()
        '<EhHeader>
        On Error GoTo chkShowMATab_Click_Err
        '</EhHeader>
100     c1TabAdditionalmodules.TabVisible(2) = IIf(chkShowMATab.Value = vbChecked, True, False)
        '<EhFooter>
        Exit Sub

chkShowMATab_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.chkShowMATab_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkShowSecurityTab_Click()
        '<EhHeader>
        On Error GoTo chkShowSecurityTab_Click_Err
        '</EhHeader>
100     C1TabOPS.TabVisible(3) = IIf(chkShowSecurityTab.Value = vbChecked, True, False)
102     C1TabOPS.TabVisible(5) = IIf(chkShowSecurityTab.Value = vbChecked, True, False)
        '<EhFooter>
        Exit Sub

chkShowSecurityTab_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.chkShowSecurityTab_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function MoveToFieldAndCreateIfNotThere(sFieldName As String, _
                                               sValue As String, _
                                               oRs As ADODB.Recordset) As Boolean
        '<EhHeader>
        On Error GoTo MoveToFieldAndCreateIfNotThere_Err
        '</EhHeader>

        '' This function only works for searching of strings
        '' Returns false if failure to find and add value
    
        On Error GoTo errorreporting
    
100     MoveToFieldAndCreateIfNotThere = False
    
102     If oRs.State = adStateOpen Then
    
104         If (oRs.EOF) And (oRs.Bof) Then
        
106             MoveToFieldAndCreateIfNotThere = False
        
            Else
        
108             MoveToFieldAndCreateIfNotThere = True
110             oRs.MoveFirst
112             oRs.Find sFieldName & " = '" & sValue & "'"
        
114             If oRs.EOF Or oRs.Bof Then
            
                    'Create field
116                 oRs.AddNew
118                 oRs.fields.Item("SettingName").Value = sValue
            
                End If
        
            End If
    
        End If
    
        Exit Function
    
errorreporting:
120     MoveToFieldAndCreateIfNotThere = False

        '<EhFooter>
        Exit Function

MoveToFieldAndCreateIfNotThere_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.MoveToFieldAndCreateIfNotThere " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub SaveAllSettings()
        '<EhHeader>
        On Error GoTo SaveAllSettings_Err
        '</EhHeader>
        Dim sModMenus As String
        Dim i As Integer
        Dim j As Integer
        Dim K As Integer
        
100     SetStatus "Beginning saving settings..."
    
102     With RSAppSetting
            
104         For i = 0 To chkCoreModule.UBound
            
106             Select Case i
        
                    Case 0
                    
                        ''''''''''''''''''''''''''''''''''''''''''''''''
                        ' LOCATOR TOOL
                    
108                     MoveToFieldAndCreateIfNotThere "SettingName", "PCodeAdminLevel0", RSAppSetting
110                     .fields.Item("SettingValue1").Value = Me.txtLocatorLayerName(0) 'Layer Name
112                     .fields.Item("SettingValue2").Value = Me.txtLocatorAttName(0) 'Attribute Field Name
                    
114                     MoveToFieldAndCreateIfNotThere "SettingName", "PCodeAdminLevel1", RSAppSetting
116                     .fields.Item("SettingValue1").Value = Me.txtLocatorLayerName(1) 'Layer Name
118                     .fields.Item("SettingValue2").Value = Me.txtLocatorAttName(1) 'Attribute Field Name
                     
120                     MoveToFieldAndCreateIfNotThere "SettingName", "PCodeAdminLocation", RSAppSetting
122                     .fields.Item("SettingValue1").Value = Me.txtLocatorLayerName(2) 'Layer Name
124                     .fields.Item("SettingValue2").Value = Me.txtLocatorAttName(2) 'Attribute Field Name
                     
126                     MoveToFieldAndCreateIfNotThere "SettingName", "HICPcodeLayer", RSAppSetting
128                     .fields.Item("SettingValue1").Value = Me.txtLocatorLayerName(3) 'Layer Name
130                     .fields.Item("SettingValue2").Value = Me.txtLocatorAttName(3) 'Attribute Field Name
                    
                        ''''''''''''''''''''''''''''''''''''''''''''''''
132                     MoveToFieldAndCreateIfNotThere "SettingName", "InitURL", RSAppSetting
134                     .fields.Item("SettingValue1").Value = txtProfileUrl.Text
                    
136                     MoveToFieldAndCreateIfNotThere "SettingName", "UseOASIProfilePage", RSAppSetting
138                     .fields.Item("SettingValue1").Value = IIf(chkUseGeneral.Value = vbChecked, "1", "0")
                      
140                     MoveToFieldAndCreateIfNotThere "SettingName", "UseIntranet", RSAppSetting
142                     .fields.Item("SettingValue1").Value = IIf(chkUseIntranet.Value = vbChecked, "1", "0")
                         
144                     MoveToFieldAndCreateIfNotThere "SettingName", "IntranetURL", RSAppSetting
146                     .fields.Item("SettingValue1").Value = txtIntranetUrl.Text
                                                 
148                     If chkCoreModule(0).Value = vbChecked Then sModMenus = "cbProfile"
                    
150                 Case 1

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Dim iCount As Integer
                        Dim iCountOfIntervals As Integer
                        
152                     MoveToFieldAndCreateIfNotThere "SettingName", "RangeSettingsColours", RSAppSetting
154                     iCount = 1
156                     iCountOfIntervals = OASISThemeRangePicker1.GetNumberOfIntervals
                        
158                     Do Until iCount = 10
                        
160                         If iCount <= iCountOfIntervals Then
162                             .fields.Item("SettingValue" & iCount).Value = OASISThemeRangePicker1.GetIntervalEndColor(iCount)
                            Else
164                             .fields.Item("SettingValue" & iCount).Value = Null
                            End If

166                         iCount = iCount + 1
                        Loop
                        
168                     .fields.Item("SettingValue10").Value = OASISThemeRangePicker1.GetThemeStartColor
                    
170                     MoveToFieldAndCreateIfNotThere "SettingName", "RangeSettingsMaxVal", RSAppSetting
172                     iCount = 1
                        
174                     Do Until iCount = 10
                        
176                         If iCount <= iCountOfIntervals Then
178                             .fields.Item("SettingValue" & iCount).Value = OASISThemeRangePicker1.GetIntervalRatioValue(iCount)
                            Else
180                             .fields.Item("SettingValue" & iCount).Value = Null
                            End If

182                         iCount = iCount + 1
                        Loop
                        
184                     .fields.Item("SettingValue10").Value = 0

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    
186                     MoveToFieldAndCreateIfNotThere "SettingName", "ShowCommodityTab", RSAppSetting
                        'Disabled functionality
188                     chkShowCommodityTab.Value = vbUnchecked
190                     .fields.Item("SettingValue1").Value = IIf(chkShowCommodityTab.Value = vbChecked, "1", "0")
                    
192                     MoveToFieldAndCreateIfNotThere "SettingName", "ShowMapLibraryTab", RSAppSetting
194                     .fields.Item("SettingValue1").Value = IIf(chkShowMapLibraryTab.Value = vbChecked, "1", "0")

196                     MoveToFieldAndCreateIfNotThere "SettingName", "ShowLISTab", RSAppSetting
198                     .fields.Item("SettingValue1").Value = IIf(chkShowLISTab.Value = vbChecked, "1", "0")

200                     MoveToFieldAndCreateIfNotThere "SettingName", "ShowSecurityTab", RSAppSetting
202                     .fields.Item("SettingValue1").Value = IIf(chkShowSecurityTab.Value = vbChecked, "1", "0")
204                     .fields.Item("SettingValue2").Value = txtLGDSymFontSize.Text
                        
206                     MoveToFieldAndCreateIfNotThere "SettingName", "ShowlegendTab", RSAppSetting
208                     .fields.Item("SettingValue1").Value = IIf(chkShowlegendTab.Value = vbChecked, "1", "0")
                    
210                     MoveToFieldAndCreateIfNotThere "SettingName", "ThemeTool", RSAppSetting
212                     .fields.Item("SettingValue1").Value = IIf(chkUseAuto.Value = vbChecked, "1", "0")
                    
214                     MoveToFieldAndCreateIfNotThere "SettingName", "stdLyrs", RSAppSetting
                        'Disabled functionality
216                     chkCOPThemes(1).Value = vbUnchecked
218                     chkCOPThemes(2).Value = vbUnchecked
220                     .fields.Item("SettingValue1").Value = IIf(chkCOPThemes(0).Value = vbChecked, "1", "0")
222                     .fields.Item("SettingValue2").Value = IIf(chkCOPThemes(1).Value = vbChecked, "1", "0")
224                     .fields.Item("SettingValue3").Value = IIf(chkCOPThemes(2).Value = vbChecked, "1", "0")
226                     .fields.Item("SettingValue4").Value = IIf(chkCOPThemes(3).Value = vbChecked, "1", "0")
                        
228                     MoveToFieldAndCreateIfNotThere "SettingName", "ShowMATab", RSAppSetting
230                     .fields.Item("SettingValue1").Value = IIf(chkShowMATab.Value = vbChecked, "1", "0")
                    
232                     MoveToFieldAndCreateIfNotThere "SettingName", "InitialOperationsTabNumber", RSAppSetting
234                     .fields.Item("SettingValue1").Value = ComInitialTab.ListIndex
236                     .fields.Item("SettingValue2").Value = IIf(chkPromptSave.Value = vbChecked, "1", "0")

                        Dim mapDefFile As String
                        Dim sTmp() As String
                        
238                     MoveToFieldAndCreateIfNotThere "SettingName", "MapLibrarySettings", RSAppSetting

240                     For K = 0 To 3

242                         If Len(MapProduct(K).ProjectName) > 4 Then
244                             sTmp = Split(MapProduct(K).ProjectName, "\")
                                'm_frmDebug.DebugPrint MapProduct(k).ProjectName
246                             sTmp(UBound(sTmp)) = Replace$(sTmp(UBound(sTmp)), ".TTKGP", "", , , vbTextCompare)
                                
248                             If Len(mapDefFile) > 1 Then
250                                 mapDefFile = mapDefFile & ","
                                End If
                                
252                             mapDefFile = mapDefFile & sTmp(UBound(sTmp))
                                
                            End If

                        Next

254                     .fields.Item("SettingValue1").Value = mapDefFile
                    
256                     For K = 0 To 3
258                         MoveToFieldAndCreateIfNotThere "SettingName", "AdminLevel" & K, RSAppSetting
260                         .fields.Item("SettingValue1").Value = txtAdmLevel(K).Text
262                         .fields.Item("SettingValue2").Value = txtAdmField(K).Text
264                         .fields.Item("SettingValue3").Value = txtAdmPCode(K).Text
266                         .fields.Item("SettingValue4").Value = txtAdmDisplayName(K).Text
                        Next

268                     MoveToFieldAndCreateIfNotThere "SettingName", "AdminLocation", RSAppSetting
270                     .fields.Item("SettingValue1").Value = txtAdmLevel(4).Text
272                     .fields.Item("SettingValue2").Value = txtAdmField(4).Text
274                     .fields.Item("SettingValue3").Value = txtAdmPCode(4).Text
276                     .fields.Item("SettingValue4").Value = txtAdmDisplayName(4).Text

278                     MoveToFieldAndCreateIfNotThere "SettingName", "OASIS_Incident_Layer_Name", RSAppSetting
280                     .fields.Item("SettingValue1").Value = txtOASISIncLyrName.Text
                                        
282                     MoveToFieldAndCreateIfNotThere "SettingName", "MineActionThemeMap", RSAppSetting
284                     .fields.Item("SettingValue1").Value = txtMAThemeMap.Text
                    
286                     MoveToFieldAndCreateIfNotThere "SettingName", "HiddenLayers", RSAppSetting
288                     .fields.Item("SettingValue1").Value = txtHiddenLayers.Text
                        
290                     MoveToFieldAndCreateIfNotThere "SettingName", "AdmProvSec", RSAppSetting
292                     .fields.Item("SettingValue1").Value = txtSecAdm(0).Text
294                     .fields.Item("SettingValue3").Value = txtSecAdmKey(0).Text
                        
296                     MoveToFieldAndCreateIfNotThere "SettingName", "AdmDistSec", RSAppSetting
298                     .fields.Item("SettingValue1").Value = txtSecAdm(1).Text
300                     .fields.Item("SettingValue3").Value = txtSecAdmKey(1).Text
                        
302                     MoveToFieldAndCreateIfNotThere "SettingName", "AdmLocalSec", RSAppSetting
304                     .fields.Item("SettingValue1").Value = txtSecAdm(2).Text
306                     .fields.Item("SettingValue3").Value = txtSecAdmKey(2).Text
                        
                        'MoveToFieldAndCreateIfNotThere "SettingName", "HICPcodeLayer", RSAppSetting
                    
                        'MoveToFieldAndCreateIfNotThere "SettingName", "MAPcodeLayer", RSAppSetting
                    
                        'MoveToFieldAndCreateIfNotThere "SettingName", "MADataSetFilesUpdate", RSAppSetting
                     
                        'MoveToFieldAndCreateIfNotThere "SettingName", "MASynchURLs", RSAppSetting
                    
                        'MoveToFieldAndCreateIfNotThere "SettingName", "MADefaultRegion", RSAppSetting
                    
                        '''''''''''''''''''''''' THIS LOOKS WRONG - NEEDS TO BE CHECKED ''''''''''''''''''''''''

308                     MoveToFieldAndCreateIfNotThere "SettingName", "ShowExpandedToolBoxInGISWin", RSAppSetting
310                     .fields.Item("SettingValue1").Value = IIf(chkShowUtility.Value = vbChecked, "1", "0")
                   
312                     .fields.Item("SettingValue3").Value = IIf(chkUtils(0).Value = vbChecked, "1", "0")
314                     .fields.Item("SettingValue4").Value = IIf(chkUtils(1).Value = vbChecked, "1", "0")
316                     .fields.Item("SettingValue5").Value = IIf(chkUtils(2).Value = vbChecked, "1", "0")
318                     .fields.Item("SettingValue7").Value = IIf(chkUtils(3).Value = vbChecked, "1", "0")
320                     .fields.Item("SettingValue7").Value = IIf(chkUtils(4).Value = vbChecked, "1", "0")
                   
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

322                     MoveToFieldAndCreateIfNotThere "SettingName", "MainMapTools", RSAppSetting
324                     .fields("SettingValue1").Value = "1"
326                     .fields("SettingValue2").Value = "0"
328                     .fields("SettingValue3").Value = "1"
330                     .fields("SettingValue4").Value = "1"
332                     .fields("SettingValue5").Value = "1"
                   
                        'AvailableMapTools
                    
                        Dim sTools As String
334                     MoveToFieldAndCreateIfNotThere "SettingName", "AvailableMapTools", RSAppSetting
                    
336                     sTools = IIf(chkMapTool(6).Value = vbChecked, "btnAddLyr", "")
338                     sTools = sTools & IIf(chkMapTool(7).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnRemoveLyr", "")
340                     sTools = sTools & IIf(chkMapTool(8).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnPrint", "")
342                     sTools = sTools & IIf(chkMapTool(9).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnAddClippBoard", "")
344                     .fields.Item("SettingValue1").Value = sTools
                    
346                     sTools = ""
348                     sTools = sTools & IIf(chkMapTool(4).Value = vbChecked, "btnInfo", "")
350                     sTools = sTools & IIf(chkMapTool(5).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnFullExtent", "")
352                     sTools = sTools & IIf(chkMapTool(23).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnLayerExtent", "")
354                     .fields.Item("SettingValue2").Value = sTools
                        
356                     sTools = ""
358                     sTools = sTools & IIf(chkMapTool(0).Value = vbChecked, "btnZoomin", "")
360                     sTools = sTools & IIf(chkMapTool(1).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnZoomout", "")
362                     sTools = sTools & IIf(chkMapTool(2).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnZoom", "")
364                     sTools = sTools & IIf(chkMapTool(3).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnPan", "")
366                     sTools = sTools & IIf(chkMapTool(4).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnInfo", "")
368                     sTools = sTools & IIf(chkMapTool(5).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnFullExtent", "")
370                     sTools = sTools & IIf(chkMapTool(23).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnLayerExtent", "")
372                     sTools = sTools & IIf(chkMapTool(22).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnZoomRect", "")
374                     sTools = sTools & ",mnuSeparator"
376                     .fields.Item("SettingValue3").Value = sTools
                        
378                     sTools = ""
380                     sTools = sTools & IIf(chkMapTool(10).Value = vbChecked, "btnOpenMap", "")
382                     sTools = sTools & IIf(chkMapTool(11).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnRepairDB", "")
384                     sTools = sTools & IIf(chkMapTool(12).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnCreateDBLyr", "")
386                     sTools = sTools & IIf(chkMapTool(13).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnExportMapDefFile", "")
388                     sTools = sTools & IIf(chkMapTool(14).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnExportToShape", "")
390                     sTools = sTools & IIf(chkMapTool(15).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnLoadSQLLyr", "")
392                     sTools = sTools & IIf(chkMapTool(16).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnAdvancedDBManagement", "")
394                     'sTools = sTools & IIf(chkMapTool(17).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnDynamicDataEntry", "")
396                     sTools = sTools & IIf(chkMapTool(18).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnAdminLocator", "")
398                     sTools = sTools & IIf(chkMapTool(19).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnCharting", "")
400                     sTools = sTools & IIf(chkMapTool(20).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnSpatialAnalysis", "")
402                     sTools = sTools & IIf(chkMapTool(21).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnCannedReports", "")
404                     sTools = sTools & IIf(chkMapTool(24).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnAddAnnotation", "")
406                     sTools = sTools & IIf(chkMapTool(25).Value = vbChecked, IIf(Len(sTools) > 0, ",", "") & "btnOASISv1Charts", "")
408                     .fields.Item("SettingValue4").Value = sTools
                    
410                     MoveToFieldAndCreateIfNotThere "SettingName", "SecurityTools", RSAppSetting
412                     .fields.Item("SettingValue1").Value = IIf(chkSecurityAnalysis.Value = vbChecked, "1", "0")
414                     .fields.Item("SettingValue2").Value = IIf(chkSecurityGraphs.Value = vbChecked, "1", "0")
416                     .fields.Item("SettingValue4").Value = IIf(chkSecurityTrends.Value = vbChecked, "1", "0")
418                     .fields.Item("SettingValue3").Value = IIf(chkAddIncident.Value = vbChecked, "1", "0")
                    
420                     MoveToFieldAndCreateIfNotThere "SettingName", "MineActionTools", RSAppSetting
422                     .fields.Item("SettingValue1").Value = IIf(chkMAAllowDataUpdate.Value = vbChecked, "1", "0")
                    
424                     MoveToFieldAndCreateIfNotThere "SettingName", "MAAddonDataSQL", RSAppSetting
426                     .fields.Item("SettingValue1").Value = txtMASQL.Text
                    
428                     MoveToFieldAndCreateIfNotThere "SettingName", "MAAddonDataKeyField", RSAppSetting
430                     .fields.Item("SettingValue1").Value = txtMaKeyField.Text
                    
                        '                    Dim sAddonRegions() As String
                        '                    MoveToFieldAndCreateIfNotThere "SettingName", "MAAddonRegions", RSAppSetting
                        '                    sAddonRegions = Split(.fields.Item("SettingValue1").Value, ",")
                        '                    ComMaregionNames.Clear
                        '                    ComMaregionNames.Text = ""
                        '                    For j = LBound(sAddonRegions) To UBound(sAddonRegions)
                        '                        ComMaregionNames.AddItem sAddonRegions(j)
                        '                    Next
                        '                    If ComMaregionNames.ListCount > 0 Then ComMaregionNames.ListIndex = 0
                        '                    MoveToFieldAndCreateIfNotThere "SettingName", "CommodityAddonDataSQL", RSAppSetting
                    
432                     MoveToFieldAndCreateIfNotThere "SettingName", "SecurityGridZoomLevels", RSAppSetting
434                     .fields.Item("SettingValue1").Value = txtSecZoomLevels(0).Text
436                     .fields.Item("SettingValue2").Value = txtSecZoomLevels(1).Text
438                     .fields.Item("SettingValue3").Value = txtSecZoomLevels(2).Text
                    
440                     MoveToFieldAndCreateIfNotThere "SettingName", "SecGrid1", RSAppSetting
442                     .fields.Item("SettingValue3").Value = txtSecGrLyrKey(0).Text
444                     .fields.Item("SettingValue1").Value = txtSaGrLyrName(0).Text
                    
446                     MoveToFieldAndCreateIfNotThere "SettingName", "SecGrid2", RSAppSetting
448                     .fields.Item("SettingValue3").Value = txtSecGrLyrKey(1).Text
450                     .fields.Item("SettingValue1").Value = txtSaGrLyrName(1).Text
                    
452                     MoveToFieldAndCreateIfNotThere "SettingName", "SecGrid3", RSAppSetting
454                     .fields.Item("SettingValue3").Value = txtSecGrLyrKey(2).Text
456                     .fields.Item("SettingValue1").Value = txtSaGrLyrName(2).Text
                                         
458                     MoveToFieldAndCreateIfNotThere "SettingName", "UserDefGridRecords", RSAppSetting
460                     .fields.Item("SettingValue1").Value = IIf(chkUseMax.Value = vbChecked, "1", "0")
462                     .fields.Item("SettingValue2").Value = txtrecMaxLevel.Text
464                     .fields.Item("SettingValue3").Value = txtrecWarningLevel.Text
466                     .fields.Item("SettingValue4").Value = IIf(chkGridOption(0).Value = vbChecked, "1", "0") & "," & IIf(chkGridOption(1).Value = vbChecked, "1", "0")
            
468                     MoveToFieldAndCreateIfNotThere "SettingName", "Notifier", RSAppSetting
470                     .fields.Item("SettingValue1").Value = fraColor(0).BackColor
472                     .fields.Item("SettingValue2").Value = fraColor(1).BackColor
            
474                     MoveToFieldAndCreateIfNotThere "SettingName", "InitMap", RSAppSetting
476                     .fields.Item("SettingValue1").Value = txtinitMap.Text
                    
478                     If chkCoreModule(1).Value = vbChecked Then
480                         If Len(sModMenus) > 2 Then sModMenus = sModMenus & ","
482                         sModMenus = sModMenus & "cbOperations"
                        End If

484                 Case 2

486                     If chkCoreModule(2).Value = vbChecked Then
488                         If Len(sModMenus) > 2 Then sModMenus = sModMenus & ","
490                         sModMenus = sModMenus & "cbContent"
                        End If
                        
492                 Case 6

494                     If chkCoreModule(6).Value = vbChecked Then
496                         If Len(sModMenus) > 2 Then sModMenus = sModMenus & ","
498                         sModMenus = sModMenus & "cbDynamicData"
                        End If
                        
500                 Case 7

502                     If chkCoreModule(7).Value = vbChecked Then
504                         If Len(sModMenus) > 2 Then sModMenus = sModMenus & ","
506                         sModMenus = sModMenus & "cbReports"
                        End If

                End Select
            
508         Next i

510         MoveToFieldAndCreateIfNotThere "SettingName", "ServerConnectionParameters", RSAppSetting
512         .fields.Item("SettingValue1").Value = txtServerConnTimeout
514         .fields.Item("SettingValue2").Value = txtServerConnRetries

516         m_frmOASISProgress.SetClientDBTimeoutSettings txtServerConnTimeout, txtServerConnRetries, CreateAppPath & "\data\db\Oasisclient.mdb"
                
            '''''''''''''''''''''''' THIS LOOKS WRONG - NEEDS TO BE CHECKED ''''''''''''''''''''''''
518         MoveToFieldAndCreateIfNotThere "SettingName", "CurrentActiveMeny", RSAppSetting

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
520         Select Case ComStartModule.ListIndex
            
                Case 0
522                 .fields.Item("SettingValue1").Value = "cbProfile"

524             Case 1
526                 .fields.Item("SettingValue1").Value = "cbOperations"
        
528             Case Else
530                 .fields.Item("SettingValue1").Value = "cbOperations"
            End Select
                    
532         MoveToFieldAndCreateIfNotThere "SettingName", "InetConnectionSettings", RSAppSetting
534         .fields.Item("SettingValue2").Value = txtSyncCheckIntervall.Text
536         .fields.Item("SettingValue1").Value = IIf(chkEnableInternet.Value = vbChecked, "1", "0")
                    
538         MoveToFieldAndCreateIfNotThere "SettingName", "attachmentsURL", RSAppSetting
540         .fields.Item("SettingValue1").Value = txtAttachmentURL.Text
        
542         MoveToFieldAndCreateIfNotThere "SettingName", "VisibleMainModuleMenus", RSAppSetting
544         .fields.Item("SettingValue1").Value = sModMenus
        
546         MoveToFieldAndCreateIfNotThere "SettingName", "ProfileSettings", RSAppSetting

548         If IsNull(.fields.Item("SettingValue1").Value) Then
550             .fields.Item("SettingValue1").Value = 1
            Else
552             .fields.Item("SettingValue1").Value = CInt(.fields.Item("SettingValue1").Value) + 1
            End If
            
554         .fields.Item("SettingValue2").Value = Now()
556         .fields.Item("SettingValue5").Value = txtAppSetting(0).Text

            Dim sScripts As String
558         MoveToFieldAndCreateIfNotThere "SettingName", "Scripts", RSAppSetting
        
560         If txtGISGrid(10).Text <> "" Then sScripts = txtGISGrid(10).Text

562         If txtGISGrid(11).Text <> "" Then
564             If sScripts <> "" Then
566                 sScripts = sScripts & ","
                End If
                
568             sScripts = sScripts & txtGISGrid(11).Text
            End If
            
570         If txtGISGrid(12).Text <> "" Then
572             If sScripts <> "" Then
574                 sScripts = sScripts & ","
                End If
                
576             sScripts = sScripts & txtGISGrid(12).Text
            End If
            
578         If txtGISGrid(13).Text <> "" Then
580             If sScripts <> "" Then
582                 sScripts = sScripts & ","
                End If
                
584             sScripts = sScripts & txtGISGrid(13).Text
            End If

586         .fields.Item("SettingValue1").Value = sScripts

        End With

        '<EhFooter>
        Exit Sub

SaveAllSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.SaveAllSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

'Private Function CheckIfEncryptedOK() As Boolean
'        'Dim MsXmlHttp As New MSXML2.XMLHTTP
'        'Dim MsXmlDoc As New MSXML2.DOMDocument
'        '<EhHeader>
'        On Error GoTo CheckIfEncryptedOK_Err
'        '</EhHeader>
'        Dim sVariable As String
'        Dim fs As New FileSystemObject
'        Dim oFile As Object
'        Dim oAES As New clsAES
'        Dim sReturnValue As String
'
'100     CheckIfEncryptedOK = False
'
'102     If chkUseEncryption(1).Value = vbChecked Then
'
'104         m_sKey = KeyGen(txtEncPass(1).Text)
'
'106         sVariable = "/oasis.asp?sKey=VALIDATE&str=" & txtEncPass(0).Text
'108         sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(txtServerURL.Text & sVariable, True)
'110         sVariable = oAES.AESEncyptString(txtEncPass(0).Text, m_sKey)
'
'112         If sReturnValue <> sVariable Then
'114             CheckIfEncryptedOK = False
'            Else
'116             CheckIfEncryptedOK = True
'            End If
'
'118         g_bHasEncrypt = True
'
'        Else
'
'120         sReturnValue = m_frmOASISProgress.OpenHttpCommsResponse(WebSite & "oasis.asp?sKey=TEST", True)
'
'122         If sReturnValue <> "0" Then
'124             g_bHasEncrypt = False
'126             CheckIfEncryptedOK = True
'            End If
'
'        End If
'
'        '<EhFooter>
'        Exit Function
'
'CheckIfEncryptedOK_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISRemoteAdmin.frmConfig.CheckIfEncryptedOK " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Function
'
'Private Function CheckEncrypt(sValue As String) As String
'        '<EhHeader>
'        On Error GoTo CheckEncrypt_Err
'        '</EhHeader>
'
'100     If g_bHasEncrypt Then
'102         CheckEncrypt = m_oAES.AESEncyptString(sValue, m_sKey)
'        Else
'104         CheckEncrypt = sValue
'        End If
'
'        '<EhFooter>
'        Exit Function
'
'CheckEncrypt_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISRemoteAdmin.frmConfig.CheckEncrypt " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Function

'Private Sub cmdConnect_Click()
'        '<EhHeader>
'        On Error GoTo cmdConnect_Click_Err
'        '</EhHeader>
'100     m_bServerConnection = True
'
'102     WebSite = txtServerURL.Text
'
'104     If Right$(WebSite, 1) <> "/" Then
'106         WebSite = WebSite & "/"
'        End If
'
'108     SetStatus "Checking Server if encrypted..."
'
'110     If chkUseEncryption(1).Value = vbChecked Then
'112         If Not CheckIfEncryptedOK Then
'114             MsgBox "It seems like the server is encrypted," & vbCrLf & "you will need to provide the correct Encryption details before connecting..."
'                Exit Sub
'            End If
'
'116         SetStatus "ENCRYPTED CONNECTION SUCCESSFUL...."
'        Else
'
'118         If CheckIfEncryptedOK Then
'120             MsgBox "It seems like the server is encrypted, you will need to provide Encryption details before connecting"
'                Exit Sub
'            End If
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'cmdConnect_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISRemoteAdmin.frmConfig.cmdConnect_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub

Private Sub cmdIntranetTest_Click()
100     ShellExecute Me.hwnd, vbNullString, txtIntranetUrl.Text, vbNullString, vbNullString, 1
End Sub

Private Sub cmdSaveClient_Click()
        '<EhHeader>
        On Error GoTo cmdSaveClient_Click_Err
        '</EhHeader>
        
        Dim bReturnValue As Boolean

100     SetStatus "Preparing AppSettings for saving..."
102     SaveAllSettings
       
104     SetStatus "Saving AppSettings..."
106     Me.Visible = False
108     RSAppSetting.Filter = adFilterPendingRecords
110     bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSAppSetting, WebSite & "Oasis.asp", True)
    
112     If bReturnValue Then
114         MsgBox "Data Updated successfully on server: " & WebSite

116         LoadUserData ""
        Else

118         MsgBox "Data Updated failed on server: " & WebSite
        End If

120     Me.Show vbModeless
        '<EhFooter>
        Exit Sub

cmdSaveClient_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmConfig.cmdSaveClient_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTest_Click()
100     ShellExecute Me.hwnd, vbNullString, txtProfileUrl.Text, vbNullString, vbNullString, 1
 
End Sub

Private Sub cmdTools_Click()
        '<EhHeader>
        On Error GoTo cmdTools_Click_Err
        '</EhHeader>
         
        Dim m_frmAdminTools As frmAdminTools
100     Set m_frmAdminTools = New frmAdminTools
102     m_frmAdminTools.Show vbModeless, Me
104     Set m_frmAdminTools = Nothing
         
        '<EhFooter>
        Exit Sub

cmdTools_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.cmdTools_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub fraColor_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo fraColor_Click_Err
        '</EhHeader>
        Dim c As cCommonDialog
100     Set c = New cCommonDialog
102     c.ShowColor

104     fraColor(Index).BackColor = c.Color
    
        '<EhFooter>
        Exit Sub

fraColor_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.fraColor_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetWebsite(sPassedWebsite As String)

    WebSite = sPassedWebsite

    If Right$(WebSite, 1) <> "/" Then
        WebSite = WebSite & "/"
    End If

End Sub

Public Function LoadUserData(vValue As Variant) As Boolean
        '<EhHeader>
        On Error GoTo LoadUserData_Err
        '</EhHeader>
    
        Dim sSQL As String
100     LoadUserData = True
        
102     m_bServerConnection = True
104     ciTab.TabVisible(2) = False
        
106     cmdSaveClient.Enabled = False
        
108     Set RSAppSetting = Nothing
110     Set RSAppSetting = New ADODB.Recordset

112     SetStatus "Beginning to load the user group data..."
    
114     If Not sTablePrefix = "" Then
116         sSQL = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & sTablePrefix & "AppSettings ORDER BY SettingName")
118         Set RSAppSetting = m_frmOASISProgress.OpenHttpCommsRS(sSQL, True)

120         If RSAppSetting.State = adStateClosed Then
            
122             MsgBox "There was a problem loading table: [" & sTablePrefix & "AppSettings] from server address: " & WebSite, vbCritical, "Server database in error"
124             Set RSAppSetting = Nothing
126             LoadUserData = False
                Exit Function
128         ElseIf (RSAppSetting.EOF And RSAppSetting.Bof) Then
130             MsgBox "There was a problem loading table: [" & sTablePrefix & "AppSettings] from server address: " & WebSite, vbCritical, "Server database in error"

132             If RSAppSetting.State = adStateOpen Then RSAppSetting.Close
134             Set RSAppSetting = Nothing
136             LoadUserData = False
                Exit Function
            End If

138         Set dxAppSettings.DataSource = RSAppSetting
140         dxAppSettings.Columns.RetrieveFields
        End If
    
142     cmdSaveClient.Enabled = True
144     ciTab.TabVisible(2) = True

146     SetStatus "Beginning loading application settings..."
148     LoadAppSetts

        '<EhFooter>
        Exit Function

LoadUserData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.LoadUserData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function CheckIfNull(sString As Variant)
        '<EhHeader>
        On Error GoTo CheckIfNull_Err
        '</EhHeader>

100     If IsNull(sString) Then
102         CheckIfNull = ""
        Else
104         CheckIfNull = sString
        End If

        '<EhFooter>
        Exit Function

CheckIfNull_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.CheckIfNull " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function MoveToAndDetectIfEOForBOF(sFieldName As String, _
                                           sValue As String, _
                                           oRs As ADODB.Recordset) As Boolean
        '<EhHeader>
        On Error GoTo MoveToAndDetectIfEOForBOF_Err
        '</EhHeader>

        On Error GoTo errorreporting
100     MoveToAndDetectIfEOForBOF = False
    
102     If oRs.State = adStateOpen Then
    
104         If (oRs.EOF) And (oRs.Bof) Then
        
106             MoveToAndDetectIfEOForBOF = False
        
            Else
        
108             oRs.MoveFirst
110             oRs.Find sFieldName & " = '" & sValue & "'"
        
112             If oRs.EOF Or oRs.Bof Then
            
114                 MoveToAndDetectIfEOForBOF = False
            
                Else
            
116                 MoveToAndDetectIfEOForBOF = True
                End If
        
            End If
    
        End If
    
        Exit Function
    
errorreporting:
118     MoveToAndDetectIfEOForBOF = False

        '<EhFooter>
        Exit Function

MoveToAndDetectIfEOForBOF_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.MoveToAndDetectIfEOForBOF " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub LoadAppSetts()
        '<EhHeader>
        On Error GoTo LoadAppSetts_Err
        '</EhHeader>
        Dim sModMenus() As String
        Dim i As Integer
        Dim j As Integer
        Dim K As Integer
        
100     Me.c1TabCoreModules.TabVisible(0) = False
102     Me.c1TabCoreModules.TabVisible(1) = False
104     c1TabCoreModules.Visible = False
    
106     With RSAppSetting
108         .MoveFirst
110         .Find "SettingName = 'VisibleMainModuleMenus'"
    
112         sModMenus = Split(CheckIfNull(.fields.Item("SettingValue1").Value), ",")
        
114         For i = 0 To UBound(sModMenus)
            
116             Select Case sModMenus(i)

                    Case "cbProfile"

118                     If MoveToAndDetectIfEOForBOF("SettingName", "InitURL", RSAppSetting) Then
120                         txtProfileUrl.Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

                        '''''''''''''''''''' IS THIS CORRECT? ''''''''''''''''''''''''''''''''''''''''''''''
   
122                     If MoveToAndDetectIfEOForBOF("SettingName", "UseOASIProfilePage", RSAppSetting) Then
124                         chkUseGeneral.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                        End If

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

126                     If MoveToAndDetectIfEOForBOF("SettingName", "UseIntranet", RSAppSetting) Then
128                         chkUseIntranet.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                        End If

130                     If MoveToAndDetectIfEOForBOF("SettingName", "IntranetURL", RSAppSetting) Then
132                         txtIntranetUrl.Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

134                     chkCoreModule(0).Value = vbChecked
136                     c1TabCoreModules.TabVisible(0) = True
138                     c1TabCoreModules.CurrTab = 0
140                     c1TabCoreModules.Visible = True

142                 Case "cbOperations"
                    
144                     If MoveToAndDetectIfEOForBOF("SettingName", "ShowCommodityTab", RSAppSetting) Then
                            'Disabled functionality
                            'chkShowCommodityTab.Value = IIf(.fields.Item("SettingValue1").Value = "1", vbChecked, vbUnchecked)
                        End If

146                     chkShowCommodityTab.Value = vbUnchecked
148                     chkShowCommodityTab.Enabled = False
                        
150                     If MoveToAndDetectIfEOForBOF("SettingName", "ShowMapLibraryTab", RSAppSetting) Then
152                         chkShowMapLibraryTab.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                        End If
                                  
154                     If MoveToAndDetectIfEOForBOF("SettingName", "ShowLISTab", RSAppSetting) Then
156                         chkShowLISTab.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                        End If
                                
158                     If MoveToAndDetectIfEOForBOF("SettingName", "ShowSecurityTab", RSAppSetting) Then
160                         chkShowSecurityTab.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                         
162                         If Not IsNull(.fields.Item("SettingValue2").Value) Then
164                             If IsNumeric(.fields.Item("SettingValue2").Value) Then
166                                 txtLGDSymFontSize.Text = .fields.Item("SettingValue2").Value
                                Else
168                                 txtLGDSymFontSize.Text = 500
                                End If

                            Else
170                             txtLGDSymFontSize.Text = 500
                            End If
                
                        Else
                
172                         txtLGDSymFontSize.Text = 500
174                         txtLGDSymFontSize.Text = 500
                        End If
                        
176                     If MoveToAndDetectIfEOForBOF("SettingName", "ShowlegendTab", RSAppSetting) Then
178                         chkShowlegendTab.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                        End If

180                     If MoveToAndDetectIfEOForBOF("SettingName", "ShowMATab", RSAppSetting) Then
182                         chkShowMATab.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                        End If

184                     If MoveToAndDetectIfEOForBOF("SettingName", "ThemeTool", RSAppSetting) Then
186                         chkUseAuto.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                        End If

188                     If MoveToAndDetectIfEOForBOF("SettingName", "stdLyrs", RSAppSetting) Then
190                         chkCOPThemes(0).Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                        
192                         chkCOPThemes(1).Value = IIf(CheckIfNull(.fields.Item("SettingValue2").Value) = "1", vbChecked, vbUnchecked)
194                         chkCOPThemes(2).Value = IIf(CheckIfNull(.fields.Item("SettingValue3").Value) = "1", vbChecked, vbUnchecked)
196                         chkCOPThemes(3).Value = IIf(CheckIfNull(.fields.Item("SettingValue4").Value) = "1", vbChecked, vbUnchecked)
                        End If

                        'Disabled functionality
198                     chkCOPThemes(1).Enabled = False
200                     chkCOPThemes(2).Enabled = False
        
202                     chkCOPThemes(1).Value = vbUnchecked
204                     chkCOPThemes(2).Value = vbUnchecked
                        
206                     If MoveToAndDetectIfEOForBOF("SettingName", "InitialOperationsTabNumber", RSAppSetting) Then
                            'For j = 0 To ComInitialTab.ListCount

208                         If Not IsNull(CInt(.fields.Item("SettingValue1").Value)) Then
210                             ComInitialTab.ListIndex = CInt(.fields.Item("SettingValue1").Value)
                            End If
                        
                            'Next
                        
212                         If Not IsNull(.fields.Item("SettingValue2").Value) Then
214                             If .fields.Item("SettingValue2").Value = "1" Then
216                                 chkPromptSave.Value = vbChecked
                                Else
218                                 chkPromptSave.Value = vbUnchecked
                                End If

                            Else
220                             chkPromptSave.Value = vbUnchecked
                            End If

                        Else
222                         chkPromptSave.Value = vbUnchecked
                        End If

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'Locator Settings
                        
224                     If MoveToAndDetectIfEOForBOF("SettingName", "PCodeAdminLevel0", RSAppSetting) Then
226                         Me.txtLocatorLayerName(0) = CheckIfNull(.fields.Item("SettingValue1").Value) 'Layer Name
228                         Me.txtLocatorAttName(0) = CheckIfNull(.fields.Item("SettingValue2").Value) 'Attribute Field Name
                        End If

230                     If MoveToAndDetectIfEOForBOF("SettingName", "PCodeAdminLevel1", RSAppSetting) Then
232                         Me.txtLocatorLayerName(1) = CheckIfNull(.fields.Item("SettingValue1").Value)  'Layer Name
234                         Me.txtLocatorAttName(1) = CheckIfNull(.fields.Item("SettingValue2").Value) 'Attribute Field Name
                        End If

236                     If MoveToAndDetectIfEOForBOF("SettingName", "PCodeAdminLocation", RSAppSetting) Then
238                         Me.txtLocatorLayerName(2) = CheckIfNull(.fields.Item("SettingValue1").Value) 'Layer Name
240                         Me.txtLocatorAttName(2) = CheckIfNull(.fields.Item("SettingValue2").Value) 'Attribute Field Name
                        End If

242                     If MoveToAndDetectIfEOForBOF("SettingName", "HICPcodeLayer", RSAppSetting) Then
244                         Me.txtLocatorLayerName(3) = CheckIfNull(.fields.Item("SettingValue1").Value) 'Layer Name
246                         Me.txtLocatorAttName(3) = CheckIfNull(.fields.Item("SettingValue2").Value) 'Attribute Field Name
                        End If

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Dim iRangeIntervals As Integer
                        Dim iCount As Integer
                        Dim lColour(11) As Long
                        
248                     If MoveToAndDetectIfEOForBOF("SettingName", "RangeSettingsColours", RSAppSetting) Then
250                         iRangeIntervals = 0
                   
252                         If Not IsNull(.fields.Item("SettingValue9").Value) Then iRangeIntervals = 9
254                         If IsNull(.fields.Item("SettingValue8").Value) Then iRangeIntervals = 7
256                         If IsNull(.fields.Item("SettingValue7").Value) Then iRangeIntervals = 6
258                         If IsNull(.fields.Item("SettingValue6").Value) Then iRangeIntervals = 5
260                         If IsNull(.fields.Item("SettingValue5").Value) Then iRangeIntervals = 4
262                         If IsNull(.fields.Item("SettingValue4").Value) Then iRangeIntervals = 3
264                         If IsNull(.fields.Item("SettingValue3").Value) Then iRangeIntervals = 2
266                         If IsNull(.fields.Item("SettingValue2").Value) Then iRangeIntervals = 1
268                         If IsNull(.fields.Item("SettingValue1").Value) Then iRangeIntervals = 0
                        
270                         OASISThemeRangePicker1.SetNumberOfIntervals iRangeIntervals
272                         iCount = 1
                        
274                         Do Until iRangeIntervals = 0 Or iCount = (iRangeIntervals + 1)
                        
276                             If iCount = 1 Then lColour(10) = IIf(IsNull(.fields.Item("SettingValue10").Value), 255, .fields.Item("SettingValue10").Value)
278                             lColour(iCount) = CheckIfNull(.fields.Item("SettingValue" & iCount).Value)
280                             iCount = iCount + 1

                            Loop

                        End If

282                     If MoveToAndDetectIfEOForBOF("SettingName", "RangeSettingsMaxVal", RSAppSetting) Then
                        
284                         iCount = 1
                        
286                         Do Until iRangeIntervals = 0 Or iCount = (iRangeIntervals + 1)
                            
288                             If iCount = 1 Then OASISThemeRangePicker1.SetThemeStartColor lColour(10)
290                             OASISThemeRangePicker1.SetInterval iCount, lColour(iCount), CheckIfNull(.fields.Item("SettingValue" & iCount).Value)
292                             iCount = iCount + 1
                            Loop

                        End If

294                     OASISThemeRangePicker1.Render

                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        Dim sMaps() As String
                    
296                     If MoveToAndDetectIfEOForBOF("SettingName", "MapLibrarySettings", RSAppSetting) Then
298                         sMaps = Split(CheckIfNull(.fields.Item("SettingValue1").Value), ",")
                        End If

                        '204                     ComMaps.Clear
                        '
                        '206                     For j = 0 To UBound(sMaps)
                        '208                         ComMaps.AddItem sMaps(j)
                        '210                         MapProduct(j).Open App.Path & "\Data\User\Maps\" & sMaps(j) & ".TTKGP"
                        '                        Next
                        '
                        '212                     If ComMaps.ListCount > 0 Then ComMaps.ListIndex = 0

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MapZoomValueOnIncidentListSelection'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'IncidentMapName'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'IncidentMapPath'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'LocationMapName'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'LocationMapPath'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'AdvancedMapTools'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'Maps'"

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MapToolsBand'"

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'SelectTools'"

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'settings'"

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'GISToolBarStyle'"

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'ToolBarFadeAway'"

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'HomeTitle'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'HomeSubTitle'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'HomeTitleToolTipText'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'ShowOperationsAttributesTab'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'ShowIncidentsAttributesTab'"

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'AllRegionChange'"
                        
300                     For K = 0 To 3

302                         If MoveToAndDetectIfEOForBOF("SettingName", "AdminLevel" & K, RSAppSetting) Then
304                             txtAdmLevel(K).Text = CheckIfNull(.fields.Item("SettingValue1").Value)
306                             txtAdmField(K).Text = CheckIfNull(.fields.Item("SettingValue2").Value)
308                             txtAdmPCode(K).Text = CheckIfNull(.fields.Item("SettingValue3").Value)
310                             txtAdmDisplayName(K).Text = CheckIfNull(.fields.Item("SettingValue4").Value)
                            End If

                        Next

312                     If MoveToAndDetectIfEOForBOF("SettingName", "AdminLevel4", RSAppSetting) Then
                        End If

314                     If MoveToAndDetectIfEOForBOF("SettingName", "AdminLevel5", RSAppSetting) Then
                        End If

316                     If MoveToAndDetectIfEOForBOF("SettingName", "AdminLocation", RSAppSetting) Then
318                         txtAdmLevel(4).Text = CheckIfNull(.fields.Item("SettingValue1").Value)
320                         txtAdmField(4).Text = CheckIfNull(.fields.Item("SettingValue2").Value)
322                         txtAdmPCode(4).Text = CheckIfNull(.fields.Item("SettingValue3").Value)
324                         txtAdmDisplayName(4).Text = CheckIfNull(.fields.Item("SettingValue4").Value)
                        End If

326                     If MoveToAndDetectIfEOForBOF("SettingName", "OASIS_Incident_Layer_Name", RSAppSetting) Then
328                         txtOASISIncLyrName.Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

                        '.Find "SettingName = 'MapPreview'"
                    
330                     If MoveToAndDetectIfEOForBOF("SettingName", "MineActionThemeMap", RSAppSetting) Then
332                         txtMAThemeMap.Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MADataSetTheme1'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MADataSetTheme2'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MADataSetTheme3'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MADataSetTheme4'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MADataSetTheme5'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MADataSetFiles'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'securityThemeMap'"
334                     If MoveToAndDetectIfEOForBOF("SettingName", "HiddenLayers", RSAppSetting) Then
336                         txtHiddenLayers.Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

338                     If MoveToAndDetectIfEOForBOF("SettingName", "AdmProvSec", RSAppSetting) Then

340                         If Not IsNull(.fields.Item("SettingValue1").Value) Then
342                             txtSecAdm(0).Text = .fields.Item("SettingValue1").Value
                            End If
                        
344                         If Not IsNull(.fields.Item("SettingValue3").Value) Then
346                             txtSecAdmKey(0).Text = .fields.Item("SettingValue3").Value
                            End If
                        End If

348                     If MoveToAndDetectIfEOForBOF("SettingName", "AdmDistSec", RSAppSetting) Then

350                         If Not IsNull(.fields.Item("SettingValue1").Value) Then
352                             txtSecAdm(1).Text = .fields.Item("SettingValue1").Value
                            End If
                        
354                         If Not IsNull(.fields.Item("SettingValue3").Value) Then
356                             txtSecAdmKey(1).Text = .fields.Item("SettingValue3").Value
                            End If
                        End If

358                     If MoveToAndDetectIfEOForBOF("SettingName", "AdmLocalSec", RSAppSetting) Then

360                         If Not IsNull(.fields.Item("SettingValue1").Value) Then
362                             txtSecAdm(2).Text = .fields.Item("SettingValue1").Value
                            End If
                        
364                         If Not IsNull(.fields.Item("SettingValue3").Value) Then
366                             txtSecAdmKey(2).Text = .fields.Item("SettingValue3").Value
                            End If
                        End If

368                     If MoveToAndDetectIfEOForBOF("SettingName", "HICPcodeLayer", RSAppSetting) Then
                        End If

370                     If MoveToAndDetectIfEOForBOF("SettingName", "MAPcodeLayer", RSAppSetting) Then
                        End If

372                     If MoveToAndDetectIfEOForBOF("SettingName", "MADataSetFilesUpdate", RSAppSetting) Then
                        End If

374                     If MoveToAndDetectIfEOForBOF("SettingName", "MASynchURLs", RSAppSetting) Then
                        End If

376                     If MoveToAndDetectIfEOForBOF("SettingName", "MADefaultRegion", RSAppSetting) Then
                        End If

378                     If MoveToAndDetectIfEOForBOF("SettingName", "ShowExpandedToolBoxInGISWin", RSAppSetting) Then
                    
380                         chkShowUtility.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                   
382                         chkUtils(0).Value = IIf(CheckIfNull(.fields.Item("SettingValue3").Value) = "1", vbChecked, vbUnchecked)
384                         chkUtils(1).Value = IIf(CheckIfNull(.fields.Item("SettingValue4").Value) = "1", vbChecked, vbUnchecked)
386                         chkUtils(2).Value = IIf(CheckIfNull(.fields.Item("SettingValue5").Value) = "1", vbChecked, vbUnchecked)
388                         chkUtils(3).Value = IIf(CheckIfNull(.fields.Item("SettingValue6").Value) = "1", vbChecked, vbUnchecked)
390                         chkUtils(4).Value = IIf(CheckIfNull(.fields.Item("SettingValue7").Value) = "1", vbChecked, vbUnchecked)
                        End If

                        'Number of tabs, 'GeoMarks, Geo Convert, Goto, Settings, Magnifier
                    
392                     If MoveToAndDetectIfEOForBOF("SettingName", "MainMapTools", RSAppSetting) Then
                    
                            '                 chkToolBar.Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue1").Value) = 1, True, False)
                            '             AB.Bands.Item("tbExtents").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue2").Value) = 1, True, False)
                            '             AB.Bands.Item("tbLayer").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue3").Value) = 1, True, False)
                            '             AB.Bands.Item("bToolbarStyle").Visible = IIf(CInt(g_RSAppSettings.Fields("SettingValue4").Value) = 1, True, False)
                        End If

                        '.MoveFirst
                        '.Find "SettingName = 'AvailableMapTools'"
                    
                        'AvailableMapTools
                    
394                     If MoveToAndDetectIfEOForBOF("SettingName", "AvailableMapTools", RSAppSetting) Then
                    
                            Dim sTools() As String

396                         sTools = Split(CheckIfNull(.fields.Item("SettingValue1").Value), ",")

398                         For j = LBound(sTools) To UBound(sTools)

400                             Select Case sTools(j)
                 
                                    Case "btnAddLyr"
402                                     chkMapTool(6).Value = vbChecked

404                                 Case "btnRemoveLyr"
406                                     chkMapTool(7).Value = vbChecked

408                                 Case "btnPrint"
410                                     chkMapTool(8).Value = vbChecked

412                                 Case "btnAddClippBoard"
414                                     chkMapTool(9).Value = vbChecked
                                End Select

                            Next

416                         sTools = Split(CheckIfNull(.fields.Item("SettingValue2").Value), ",")

418                         For j = LBound(sTools) To UBound(sTools)

420                             Select Case sTools(j)
                
                                        '                                Case "btnInfo"
                                        '386                                 chkMapTool(4).Value = vbChecked
                                        '
                                        '388                             Case "btnFullExtent"
                                        '390                                 chkMapTool(5).Value = vbChecked
                                        '                                Case "btnLayerExtent"
                                        '                                    chkMapTool(23).Value = vbChecked
                                End Select

                            Next

                            sTools = Split(CheckIfNull(.fields.Item("SettingValue3").Value), ",")

422                         For j = LBound(sTools) To UBound(sTools)

                                ',,,,btnSelect,btnDeselect,btnInfo
424                             Select Case sTools(j)
                
                                    Case "btnZoomin"
426                                     chkMapTool(0).Value = vbChecked

428                                 Case "btnZoomout"
430                                     chkMapTool(1).Value = vbChecked

432                                 Case "btnZoom"
434                                     chkMapTool(2).Value = vbChecked

436                                 Case "btnPan"
438                                     chkMapTool(3).Value = vbChecked

440                                 Case "btnZoomRect"
442                                     chkMapTool(22).Value = vbChecked

444                                 Case "btnInfo"
446                                     chkMapTool(4).Value = vbChecked

448                                 Case "btnFullExtent"
450                                     chkMapTool(5).Value = vbChecked

452                                 Case "btnLayerExtent"
454                                     chkMapTool(23).Value = vbChecked
                                End Select

                            Next
                    
456                         sTools = Split(CheckIfNull(.fields.Item("SettingValue4").Value), ",")

458                         For j = LBound(sTools) To UBound(sTools)

460                             Select Case sTools(j)
                 
                                    Case "btnOpenMap"
462                                     chkMapTool(10).Value = vbChecked

464                                 Case "btnRepairDB"
466                                     chkMapTool(11).Value = vbChecked

468                                 Case "btnCreateDBLyr"
470                                     chkMapTool(12).Value = vbChecked

472                                 Case "btnExportMapDefFile"
474                                     chkMapTool(13).Value = vbChecked

476                                 Case "btnExportToShape"
478                                     chkMapTool(14).Value = vbChecked

480                                 Case "btnLoadSQLLyr"
482                                     chkMapTool(15).Value = vbChecked

484                                 Case "btnAdvancedDBManagement"
486                                     chkMapTool(16).Value = vbChecked
                            
488                                 Case "btnDynamicDataEntry"
490                                     chkMapTool(17).Value = vbUnchecked ' vbChecked
                                        chkMapTool(17).Enabled = False
492                                 Case "btnAdminLocator"
494                                     chkMapTool(18).Value = vbChecked

496                                 Case "btnCharting"
498                                     chkMapTool(19).Value = vbChecked

500                                 Case "btnSpatialAnalysis"
502                                     chkMapTool(20).Value = vbChecked

504                                 Case "btnCannedReports"
506                                     chkMapTool(21).Value = vbChecked ' vbChecked
                                        'chkMapTool(21).Enabled = False

508                                 Case "btnAddAnnotation"
510                                     chkMapTool(24).Value = vbChecked

512                                 Case "btnOASISv1Charts"
514                                     chkMapTool(25).Value = vbChecked
                                End Select

                            Next

                        End If

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'LgdSecuritySpacing'"
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'LgdMASpacing'"
                        '
516                     If MoveToAndDetectIfEOForBOF("SettingName", "SecurityTools", RSAppSetting) Then
                    
518                         chkSecurityAnalysis.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
520                         chkSecurityGraphs.Value = IIf(CheckIfNull(.fields.Item("SettingValue2").Value) = "1", vbChecked, vbUnchecked)
522                         chkSecurityTrends.Value = IIf(CheckIfNull(.fields.Item("SettingValue4").Value) = "1", vbChecked, vbUnchecked)
524                         chkAddIncident.Value = IIf(CheckIfNull(.fields.Item("SettingValue3").Value) = "1", vbChecked, vbUnchecked)
                        End If

                        'AreaAnalysis,Graphs,AddIncidents, Animation
                    
526                     If MoveToAndDetectIfEOForBOF("SettingName", "MineActionTools", RSAppSetting) Then
528                         chkMAAllowDataUpdate.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
                        End If

530                     If MoveToAndDetectIfEOForBOF("SettingName", "MAAddonDataSQL", RSAppSetting) Then
532                         txtMASQL.Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

534                     If MoveToAndDetectIfEOForBOF("SettingName", "MAAddonDataKeyField", RSAppSetting) Then
536                         txtMaKeyField.Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

538                     ComMaregionNames.Clear
540                     ComMaregionNames.Text = ""
                
542                     If MoveToAndDetectIfEOForBOF("SettingName", "MAAddonRegions", RSAppSetting) Then
                            Dim sAddonRegions() As String
        
544                         sAddonRegions = Split(CheckIfNull(.fields.Item("SettingValue1").Value), ",")
        
546                         For j = LBound(sAddonRegions) To UBound(sAddonRegions)
548                             ComMaregionNames.AddItem sAddonRegions(j)
                            Next

                        End If

550                     If ComMaregionNames.ListCount > 0 Then ComMaregionNames.ListIndex = 0
                    
552                     If MoveToAndDetectIfEOForBOF("SettingName", "CommodityAddonDataSQL", RSAppSetting) Then
                        End If

554                     If MoveToAndDetectIfEOForBOF("SettingName", "SecurityGridZoomLevels", RSAppSetting) Then
                    
556                         txtSecZoomLevels(0).Text = CheckIfNull(.fields.Item("SettingValue1").Value)
558                         txtSecZoomLevels(1).Text = CheckIfNull(.fields.Item("SettingValue2").Value)
560                         txtSecZoomLevels(2).Text = CheckIfNull(.fields.Item("SettingValue3").Value)
                        End If

562                     If MoveToAndDetectIfEOForBOF("SettingName", "SecGrid1", RSAppSetting) Then
564                         txtSecGrLyrKey(0).Text = CheckIfNull(.fields.Item("SettingValue3").Value)
566                         txtSaGrLyrName(0).Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

568                     If MoveToAndDetectIfEOForBOF("SettingName", "SecGrid2", RSAppSetting) Then
570                         txtSecGrLyrKey(1).Text = CheckIfNull(.fields.Item("SettingValue3").Value)
572                         txtSaGrLyrName(1).Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

574                     If MoveToAndDetectIfEOForBOF("SettingName", "SecGrid3", RSAppSetting) Then
576                         txtSecGrLyrKey(2).Text = CheckIfNull(.fields.Item("SettingValue3").Value)
578                         txtSaGrLyrName(2).Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

580                     If MoveToAndDetectIfEOForBOF("SettingName", "UserDefGridRecords", RSAppSetting) Then
582                         chkUseMax.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
            
584                         txtrecMaxLevel.Text = CheckIfNull(.fields.Item("SettingValue2").Value)
586                         txtrecWarningLevel.Text = CheckIfNull(.fields.Item("SettingValue3").Value)

                            Dim sGridOpt() As String
                        
588                         If Not IsNull(.fields.Item("SettingValue4").Value) Then
                        
590                             If Len(.fields.Item("SettingValue4").Value) > 2 Then
592                                 sGridOpt = Split(.fields.Item("SettingValue4").Value, ",")
594                                 chkGridOption(0).Value = IIf(sGridOpt(0) = "1", vbChecked, vbUnchecked)
596                                 chkGridOption(1).Value = IIf(sGridOpt(1) = "1", vbChecked, vbUnchecked)
                                End If
                            End If
                        End If

598                     If MoveToAndDetectIfEOForBOF("SettingName", "ThemeTool", RSAppSetting) Then
                        End If

                        '                    .MoveFirst
                        '                    .Find "SettingName = 'HazardLyr'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'DaLyr'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MFLyr'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'MALyr'"
                        '
                        '                    .MoveFirst
                        '                    .Find "SettingName = 'VictimsLyr'"

600                     If MoveToAndDetectIfEOForBOF("SettingName", "InitMap", RSAppSetting) Then
602                         txtinitMap.Text = CheckIfNull(.fields.Item("SettingValue1").Value)
                        End If

                        '
                        '
                        'UseIntranet
                        'MineActionTools
                        'SecurityTools
                        'LgdMASpacing
                        'LgdSecuritySpacing
                        'AvailableMapTools
                        'MainMapTools
                        'ShowExpandedToolBoxInGISWin
                        'MADefaultRegion
                        'MASynchURLs
                        'MADataSetTheme1
                        'MADataSetTheme2
                        
                        'MADataSetTheme3
                        'MADataSetTheme4
                        'MADataSetTheme5
                        'MADataSetFilesUpdate
                        'MineActionThemeMap

                        'CurrentActiveMeny
                    
604                     If MoveToAndDetectIfEOForBOF("SettingName", "Notifier", RSAppSetting) Then

606                         If Not .EOF And Not .Bof Then
608                             If Not IsNull(.fields.Item("SettingValue1").Value) Then
610                                 fraColor(0).BackColor = CheckIfNull(.fields.Item("SettingValue1").Value)
612                                 fraColor(1).BackColor = CheckIfNull(.fields.Item("SettingValue2").Value)
                                End If
                            End If
                        End If

614                     chkCoreModule(1).Value = vbChecked
616                     c1TabCoreModules.TabVisible(1) = True
618                     c1TabCoreModules.CurrTab = 1

620                     If c1TabCoreModules.TabVisible(0) Then c1TabCoreModules.CurrTab = 0
622                     c1TabCoreModules.Visible = True
                        '596                 Case "cbContent"
                        '598                     chkCoreModule(2).Value = vbChecked
                        '
                        '600                 Case "cbSync"
                        '602                     chkCoreModule(3).Value = vbChecked
                        '
                        '604                 Case "cbAddOns"
                        '606                     chkCoreModule(4).Value = vbChecked
                        '
                        '608                 Case "cbW3"
                        '
                        '610                     If MoveToAndDetectIfEOForBOF("SettingName", "w3OrgID", RSAppSetting) Then
                        '                        End If
                        '
                        '612                     If MoveToAndDetectIfEOForBOF("SettingName", "w3OrgSetup", RSAppSetting) Then
                        '                        End If
                        '
                        '614                     If MoveToAndDetectIfEOForBOF("SettingName", "w3OfficeID", RSAppSetting) Then
                        '                        End If
                        '
                        '616                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHOContact", RSAppSetting) Then
                        '                        End If
                        '
                        '618                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHOStaff", RSAppSetting) Then
                        '                        End If
                        '
                        '620                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHOTransport", RSAppSetting) Then
                        '                        End If
                        '
                        '622                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHOOffice", RSAppSetting) Then
                        '                        End If
                        '
                        '624                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHATLocation", RSAppSetting) Then
                        '                        End If
                        '
                        '626                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHATDetails", RSAppSetting) Then
                        '                        End If
                        '
                        '628                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHATBeneficiaries", RSAppSetting) Then
                        '                        End If
                        '
                        '630                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHATImplementingPartners", RSAppSetting) Then
                        '                        End If
                        '
                        '632                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHATFunding", RSAppSetting) Then
                        '                        End If
                        '
                        '634                     If MoveToAndDetectIfEOForBOF("SettingName", "w3WHERELocator", RSAppSetting) Then
                        '                        End If
                        '
                        '                        'Disabled functionality
                        '                        'chkCoreModule(5).Value = vbChecked
                        
624                 Case "cbContent"
626                     chkCoreModule(2).Value = vbChecked

628                 Case "cbDynamicData"
630                     chkCoreModule(6).Value = vbChecked

632                 Case "cbReports"
634                     chkCoreModule(7).Value = vbChecked
                        
                        'c1TabCoreModules.TabVisible(2) = True
                        'c1TabCoreModules.
                End Select

636             chkCoreModule(5).Enabled = False
638             chkCoreModule(5).Value = vbUnchecked
                'AB.Bands("bNavPane").ChildBands(sModMenus(i)).Visible = True
640         Next i

642         If MoveToAndDetectIfEOForBOF("SettingName", "ServerConnectionParameters", RSAppSetting) Then
                
644             If Not IsNumeric(CheckIfNull(.fields.Item("SettingValue1").Value)) Then
646                 txtServerConnTimeout = 10
                Else
648                 txtServerConnTimeout = CheckIfNull(.fields.Item("SettingValue1").Value)
                End If
                
650             If Not IsNumeric(CheckIfNull(.fields.Item("SettingValue2").Value)) Then
652                 txtServerConnRetries = 3
                Else
654                 txtServerConnRetries = CheckIfNull(.fields.Item("SettingValue2").Value)
                End If
                
            Else
656             txtServerConnTimeout = 10
658             txtServerConnRetries = 3
            End If
        
            '''''''''''''''''''''''' CHECK THIS''''''''''''''''''''''''''''''''''''''''
        
660         If MoveToAndDetectIfEOForBOF("SettingName", "CurrentActiveMeny", RSAppSetting) Then
        
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

662             Select Case CheckIfNull(.fields.Item("SettingValue1").Value)
            
                    Case "cbProfile"
664                     ComStartModule.ListIndex = 0

666                 Case "cbOperations"
668                     ComStartModule.ListIndex = 1

670                 Case "cbContent"
672                     ComStartModule.ListIndex = 2

                        ' Case "cbSync"
                        '  c 'omStartModule.ListIndex = 3

                        'Case "cbAddOns"
                        'ComStartModule.ListIndex = 4

                        'Case "cbW3"
                        'ComStartModule.ListIndex = 5
                End Select

            End If

            '                    OASIS profile
            'Operations
            'Dynamic Content
            'Synchronize
            'OASIS Addons
            'Who, What, Where?
                    
674         If MoveToAndDetectIfEOForBOF("SettingName", "InetConnectionSettings", RSAppSetting) Then
676             txtSyncCheckIntervall.Text = CheckIfNull(.fields.Item("SettingValue2").Value)

678             chkEnableInternet.Value = IIf(CheckIfNull(.fields.Item("SettingValue1").Value) = "1", vbChecked, vbUnchecked)
            End If

680         If MoveToAndDetectIfEOForBOF("SettingName", "attachmentsURL", RSAppSetting) Then
682             txtAttachmentURL.Text = CheckIfNull(.fields.Item("SettingValue1").Value)
            End If

684         If MoveToAndDetectIfEOForBOF("SettingName", "ProfileSettings", RSAppSetting) Then
                Dim sCaption As String
        
686             sCaption = "OASIS Admin Tool Version: " & App.major & "." & App.minor & "." & App.Revision & " © Copyright iMMAP.org 2003-2008"

688             If IsNull(RSAppSetting.fields.Item("SettingValue1").Value) Then
690                 sCaption = sCaption & " Configuration Version: 1"
                Else
692                 sCaption = sCaption & " Configuration Version:" & CInt(CheckIfNull(RSAppSetting.fields.Item("SettingValue1").Value)) + 1
                End If

694             sCaption = sCaption & " Last Update:" & CheckIfNull(.fields.Item("SettingValue2").Value)

696             Me.Caption = sCaption

698             If Not CheckIfNull(.fields.Item("SettingValue5").Value) = vbNull Then
700                 txtAppSetting(0).Text = CheckIfNull(.fields.Item("SettingValue5").Value)
                End If
            End If
     
702         If MoveToAndDetectIfEOForBOF("SettingName", "Scripts", RSAppSetting) Then
            
                Dim sScripts() As String
            
704             If Not IsNull(.fields.Item("SettingValue1").Value) Then
706                 If Len(.fields.Item("SettingValue1").Value) > 4 Then
708                     sScripts = Split(.fields.Item("SettingValue1").Value, ",")

710                     For j = LBound(sScripts) To UBound(sScripts)
712                         txtGISGrid(j + 10).Text = sScripts(j)
                        Next
                        
                    End If
                End If
            End If
        
        End With
        
        '<EhFooter>
        Exit Sub

LoadAppSetts_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmConfig.LoadAppSetts " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ResetStuff()
        '<EhHeader>
        On Error GoTo ResetStuff_Err
        '</EhHeader>
100     c1TabCoreModules.TabVisible(0) = False
102     c1TabCoreModules.TabVisible(1) = False
        '<EhFooter>
        Exit Sub

ResetStuff_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.ResetStuff " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        
100     SetStatus "BEGIN CONFIG TOOL LOG................................"
        Dim mFileSysObj As New FileSystemObject
    
102     Me.Caption = "OASIS Server Administration Toolbox. Version: " & App.major & "." & App.minor & "." & App.Revision & " Developed by iMMAP.org Contact: support@iMMAP.org © Copyright iMMAP.org 2003-2008"
    
104     'CreateThreads
106     ResetStuff
    
108     ComInitialTab.Clear
110     ComInitialTab.AddItem "Legend"
112     ComInitialTab.AddItem "Map Library"
114     ComInitialTab.AddItem "Security"
116     ComInitialTab.AddItem "Commodity/NFI"
118     ComInitialTab.AddItem "Mine Action"
120     ComInitialTab.AddItem "LIS"
122     ComStartModule.AddItem "OASIS Profile"
124     ComStartModule.AddItem "Operations"
126     ComStartModule.AddItem "Dynamic Content"
128     ComStartModule.AddItem "Synchronize"
130     ComStartModule.AddItem "OASIS Addons"
132     ComStartModule.AddItem "Who, What, Where?"

134     Set mFileSysObj = Nothing

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     'TerminateThreads
    
102     'SaveSetting App.EXEName, "Settings", "WebServer", txtServerURL.Text
104     'SaveSetting App.EXEName, "Settings", "FileSynchURL", txtfileSynchUrl.Text
            
        On Error Resume Next

106     Set m_oAES = Nothing
        
108     RSAppSetting.Close
110     Set RSAppSetting = Nothing
    
        Dim i As Integer
    
112     For i = 1 To Forms.Count

114         If Not Forms(i).Name <> Me.Name Then
116             Unload Forms(i)
118             Set Forms(i) = Nothing
            End If

        Next
    
120     SetStatus "END CONFIG TOOL LOG................................"
    
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Image1_Click()
        '<EhHeader>
        On Error GoTo Image1_Click_Err
        '</EhHeader>
100     ShellExecute Me.hwnd, vbNullString, "http://www.immap.org", vbNullString, vbNullString, 1
        '<EhFooter>
        Exit Sub

Image1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.Image1_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CreateNewTableSQL(sNewTable As String, _
                                   db_file As String, _
                                   RSSource As ADODB.Recordset, _
                                   Optional bAppendData As Boolean) As String
        '<EhHeader>
        On Error GoTo CreateNewTableSQL_Err
        '</EhHeader>
        Dim RS As ADODB.Recordset
        Dim num_records As Integer
        Dim oFld As ADODB.Field

        'TODO Drop the table if it already exists.
    
100     For Each oFld In RSSource.fields
        
102         If FieldIsString(oFld) Then
        
104             If oFld.Type = adLongVarChar Or oFld.Type = adLongVarWChar Then
106                 CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " MEMO"
                Else
108                 CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " TEXT(" & oFld.DefinedSize & ")"
                End If
        
110         ElseIf FieldIsBoolean(oFld) Then
112             CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " YESNO"
114         ElseIf FieldIsNumeric(oFld) Then

116             Select Case oFld.Type
        
                    Case adDecimal, adNumeric, adVarNumeric
118                     CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " FLOAT"

120                 Case adDouble
122                     CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " DOUBLE"

124                 Case adInteger, adBigInt
126                     CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " LONG"

128                 Case adSingle
130                     CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " SINGLE"

132                 Case adSmallInt, adTinyInt
134                     CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " SMALLINT"
                    
                End Select
        
136         ElseIf FieldIsTimeDate(oFld) Then

138             Select Case oFld.Type
            
                    Case adDate, adDBDate  'DATE
140                     CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " DATE"

142                 Case adDBTime 'DATETIME
144                     CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " DATETIME"

146                 Case adDBTimeStamp 'TimeStamp
148                     CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " TimeStamp"
                End Select
            
150         ElseIf FieldIsBinary(oFld) Then
152             CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " OLEObject"
154         ElseIf oFld.Type = adGUID Then
156             CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " COUNTER (1,15)" 'GUID" 'ReplicationID"
            Else
158             CreateNewTableSQL = CreateNewTableSQL & IIf(Len(CreateNewTableSQL) = 0, " ", ", ") & oFld.Name & " UNKNOWN"
            End If
        
        Next
    
160     If bAppendData Then
            'TODO Add the data to the Table
        End If
    
162     CreateNewTableSQL = "CREATE TABLE " & sNewTable & " (" & CreateNewTableSQL & ")"
    
        '<EhFooter>
        Exit Function

CreateNewTableSQL_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmConfig.CreateNewTableSQL " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

