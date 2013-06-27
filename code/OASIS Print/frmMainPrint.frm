VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{B6C8B132-5973-4983-AD46-8F3F10B04531}#1.0#0"; "vbalCbEx6.ocx"
Begin VB.Form frmMainPrint 
   Caption         =   "OASIS Print Template Utility"
   ClientHeight    =   9060
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   12840
   Icon            =   "frmMainPrint.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9060
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin OASISClient.OASISDrawObj OASISDrawObj1 
      Height          =   7575
      Left            =   660
      TabIndex        =   60
      Top             =   1080
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   13361
      CanvasWidth     =   800
      CanvasHeight    =   600
      UndoBufferSize  =   0
      ShowCanvasSize  =   0   'False
   End
   Begin ComCtl3.CoolBar CoolBar3 
      Align           =   4  'Align Right
      Height          =   7995
      Left            =   10275
      TabIndex        =   10
      Top             =   750
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   14102
      BandCount       =   2
      Orientation     =   1
      _CBWidth        =   2565
      _CBHeight       =   7995
      _Version        =   "6.7.9782"
      Child1          =   "PicProperty1"
      MinWidth1       =   3495
      MinHeight1      =   2505
      Width1          =   3495
      NewRow1         =   0   'False
      Child2          =   "pctProperties"
      MinWidth2       =   1530
      MinHeight2      =   2400
      Width2          =   8940
      NewRow2         =   0   'False
      Begin VB.PictureBox pctProperties 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4080
         Left            =   75
         ScaleHeight     =   4080
         ScaleWidth      =   2400
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3885
         Width           =   2400
         Begin OASISClient.tipPopup tipPopup1 
            Height          =   795
            Left            =   300
            Top             =   1200
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   1402
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.CommandButton cmdCommand1 
            Caption         =   "Command1"
            Height          =   360
            Left            =   135
            TabIndex        =   59
            Top             =   3510
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.PictureBox ddnCategories 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1005
            Left            =   270
            ScaleHeight     =   945
            ScaleWidth      =   2025
            TabIndex        =   58
            Top             =   2880
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.TextBox txtEdit 
            Height          =   315
            Left            =   225
            TabIndex        =   57
            Top             =   2880
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.PictureBox selCategories 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   180
            ScaleHeight     =   435
            ScaleWidth      =   2205
            TabIndex        =   56
            Top             =   405
            Visible         =   0   'False
            Width           =   2265
         End
         Begin vbalIml6.vbalImageList ilsIcons 
            Left            =   810
            Top             =   3150
            _ExtentX        =   953
            _ExtentY        =   953
            ColourDepth     =   24
            Size            =   13776
            Images          =   "frmMainPrint.frx":1601A
            Version         =   131072
            KeyCount        =   12
            Keys            =   "Big DogÿSmall DogÿBumÿMaggieÿSpace MutantÿRenÿFelixÿHomerÿDraculaÿPirateÿNinjaÿBart"
         End
         Begin vbalComboEx6.vbalCboEx cboIcon 
            Height          =   330
            Left            =   270
            TabIndex        =   55
            Top             =   2385
            Visible         =   0   'False
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ExtendedUI      =   0   'False
            DropDownWidth   =   0
         End
         Begin vbAcceleratorSGrid6.vbalGrid propertyGrid 
            Height          =   4035
            Left            =   35
            TabIndex        =   54
            Top             =   0
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   7117
            GridLines       =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            GridLineColor   =   -2147483632
            GridFillLineColor=   -2147483631
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderButtons   =   0   'False
            HeaderDragReorderColumns=   0   'False
            HeaderHotTrack  =   0   'False
            HeaderFlat      =   -1  'True
            BorderStyle     =   2
            Editable        =   -1  'True
            DisableIcons    =   -1  'True
         End
      End
      Begin TabDlg.SSTab tabMain 
         Height          =   3570
         Left            =   90
         TabIndex        =   11
         Top             =   135
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   6297
         _Version        =   393216
         Style           =   1
         TabHeight       =   485
         TabCaption(0)   =   "Colors"
         TabPicture(0)   =   "frmMainPrint.frx":1960A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "PicProperty1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Style"
         TabPicture(1)   =   "frmMainPrint.frx":19626
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "PicProperty2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Paper"
         TabPicture(2)   =   "frmMainPrint.frx":19642
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "picPaperProp"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.PictureBox picPaperProp 
            BorderStyle     =   0  'None
            Height          =   3120
            Left            =   -74955
            ScaleHeight     =   3120
            ScaleWidth      =   2310
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   315
            Width           =   2310
            Begin OASISClient.ColorPicker ColorPicker1 
               Height          =   315
               Left            =   90
               TabIndex        =   61
               Top             =   2790
               Width           =   2175
               _ExtentX        =   3836
               _ExtentY        =   556
            End
            Begin VB.ComboBox ComPSize 
               Height          =   315
               Left            =   45
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   301
               Width           =   2175
            End
            Begin VB.ComboBox ComPUnits 
               Height          =   315
               Left            =   45
               Style           =   2  'Dropdown List
               TabIndex        =   46
               Top             =   933
               Width           =   2130
            End
            Begin VB.ComboBox ComPOrentation 
               Height          =   315
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   45
               Top             =   1565
               Width           =   2130
            End
            Begin VB.ComboBox ComPQuality 
               Height          =   315
               Left            =   90
               Style           =   2  'Dropdown List
               TabIndex        =   44
               Top             =   2197
               Width           =   2130
            End
            Begin VB.Label lblBackColor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Back Color:"
               Height          =   195
               Left            =   135
               TabIndex        =   52
               Top             =   2580
               Width           =   825
            End
            Begin VB.Label lblSize 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Size:"
               Height          =   195
               Left            =   45
               TabIndex        =   51
               Top             =   45
               Width           =   345
            End
            Begin VB.Label lblUnits 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Trace Units:"
               Height          =   195
               Left            =   45
               TabIndex        =   50
               Top             =   675
               Width           =   870
            End
            Begin VB.Label lblOrientation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Orientation:"
               Height          =   195
               Left            =   45
               TabIndex        =   49
               Top             =   1309
               Width           =   870
            End
            Begin VB.Label lblQuality 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Quality:"
               Height          =   195
               Left            =   90
               TabIndex        =   48
               Top             =   1941
               Width           =   570
            End
         End
         Begin VB.PictureBox PicProperty2 
            BorderStyle     =   0  'None
            Height          =   2850
            Left            =   -74910
            ScaleHeight     =   2850
            ScaleWidth      =   2280
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   540
            Width           =   2280
            Begin VB.VScrollBar VScroll3 
               Height          =   285
               Left            =   1920
               Max             =   1
               Min             =   200
               TabIndex        =   36
               Top             =   2100
               Value           =   25
               Width           =   255
            End
            Begin VB.TextBox TxtRound 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1500
               TabIndex        =   35
               Text            =   "25"
               Top             =   2100
               Width           =   405
            End
            Begin VB.TextBox TxtPoint 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1500
               Locked          =   -1  'True
               TabIndex        =   34
               Text            =   "3"
               ToolTipText     =   "Border Size"
               Top             =   1500
               Width           =   405
            End
            Begin VB.VScrollBar VScroll2 
               Height          =   285
               Left            =   1920
               Max             =   3
               Min             =   30
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   1500
               Value           =   3
               Width           =   255
            End
            Begin VB.ComboBox CboFill 
               Height          =   315
               ItemData        =   "frmMainPrint.frx":1965E
               Left            =   900
               List            =   "frmMainPrint.frx":1967A
               Style           =   2  'Dropdown List
               TabIndex        =   32
               Top             =   450
               Width           =   1275
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   285
               Left            =   1890
               Max             =   0
               Min             =   100
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   60
               Value           =   1
               Width           =   255
            End
            Begin VB.TextBox TxtBorder 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   1470
               Locked          =   -1  'True
               TabIndex        =   30
               Text            =   "1"
               ToolTipText     =   "Border Size"
               Top             =   60
               Width           =   405
            End
            Begin MSComctlLib.Slider Slider1 
               Height          =   345
               Left            =   60
               TabIndex        =   37
               TabStop         =   0   'False
               Top             =   1080
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   609
               _Version        =   393216
               LargeChange     =   1
               Max             =   360
               SelectRange     =   -1  'True
               TickStyle       =   3
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Round Rectangle Size:"
               Height          =   495
               Left            =   90
               TabIndex        =   42
               Top             =   1980
               Width           =   1170
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Points Qty:"
               Height          =   255
               Left            =   90
               TabIndex        =   41
               Top             =   1530
               Width           =   1215
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fill style:"
               Height          =   255
               Left            =   90
               TabIndex        =   40
               Top             =   510
               Width           =   645
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Border size:"
               Height          =   255
               Left            =   90
               TabIndex        =   39
               Top             =   90
               Width           =   915
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rotation: 0°"
               Height          =   195
               Index           =   0
               Left            =   150
               TabIndex        =   38
               Top             =   870
               Width           =   1995
            End
         End
         Begin VB.PictureBox PicProperty1 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   3195
            Left            =   45
            ScaleHeight     =   3195
            ScaleWidth      =   2505
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   315
            Width           =   2505
            Begin OASISClient.ColorPal ColorPal1 
               Height          =   2040
               Left            =   120
               TabIndex        =   62
               Top             =   30
               Width           =   1195
               _ExtentX        =   2117
               _ExtentY        =   3598
               Thumbsize       =   6
            End
            Begin VB.VScrollBar ScrCol 
               Height          =   285
               Index           =   2
               Left            =   2100
               Max             =   0
               Min             =   255
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   2355
               Width           =   225
            End
            Begin VB.VScrollBar ScrCol 
               Height          =   285
               Index           =   1
               Left            =   1350
               Max             =   0
               Min             =   255
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   2355
               Width           =   225
            End
            Begin VB.VScrollBar ScrCol 
               Height          =   285
               Index           =   0
               Left            =   570
               Max             =   0
               Min             =   255
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   2355
               Width           =   225
            End
            Begin VB.TextBox TxtColor 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   2
               Left            =   1650
               Locked          =   -1  'True
               TabIndex        =   18
               Text            =   "0"
               Top             =   2340
               Width           =   465
            End
            Begin VB.TextBox TxtColor 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   1
               Left            =   870
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "0"
               Top             =   2340
               Width           =   465
            End
            Begin VB.TextBox TxtColor 
               Alignment       =   1  'Right Justify
               Height          =   315
               Index           =   0
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   16
               Text            =   "0"
               Top             =   2340
               Width           =   465
            End
            Begin VB.OptionButton OpColor 
               BackColor       =   &H00FFFFFF&
               Height          =   375
               Index           =   2
               Left            =   1350
               MouseIcon       =   "frmMainPrint.frx":196D1
               MousePointer    =   99  'Custom
               TabIndex        =   15
               Top             =   1530
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.OptionButton OpColor 
               BackColor       =   &H00000000&
               Height          =   375
               Index           =   1
               Left            =   1350
               MouseIcon       =   "frmMainPrint.frx":19823
               MousePointer    =   99  'Custom
               TabIndex        =   14
               Top             =   900
               Width           =   945
            End
            Begin VB.OptionButton OpColor 
               BackColor       =   &H00FF0000&
               Height          =   375
               Index           =   0
               Left            =   1365
               MouseIcon       =   "frmMainPrint.frx":19975
               MousePointer    =   99  'Custom
               TabIndex        =   13
               Top             =   270
               Value           =   -1  'True
               Width           =   945
            End
            Begin VB.Label LblCol 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Blue:"
               Height          =   225
               Index           =   2
               Left            =   1650
               TabIndex        =   28
               Top             =   2130
               Width           =   675
            End
            Begin VB.Label LblCol 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Green:"
               Height          =   225
               Index           =   1
               Left            =   870
               TabIndex        =   27
               Top             =   2130
               Width           =   675
            End
            Begin VB.Label LblCol 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Red:"
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   26
               Top             =   2130
               Width           =   675
            End
            Begin VB.Label LblColor 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   465
               Left            =   60
               TabIndex        =   25
               Top             =   2700
               Width           =   2385
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Back Color"
               Height          =   255
               Index           =   2
               Left            =   1260
               TabIndex        =   24
               Top             =   1350
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Border Color"
               Height          =   255
               Index           =   1
               Left            =   1290
               TabIndex        =   23
               Top             =   690
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fill Color"
               Height          =   255
               Index           =   3
               Left            =   1290
               TabIndex        =   22
               Top             =   90
               Width           =   975
            End
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   8745
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7055
            MinWidth        =   7055
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3598
            MinWidth        =   3598
            Object.ToolTipText     =   "Mouse Position"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   6782
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "4:11 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5220
      Top             =   7860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":19AC7
            Key             =   "Select"
            Object.Tag             =   "Select"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":19BD9
            Key             =   "Line"
            Object.Tag             =   "Line"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":19CEB
            Key             =   "Arc"
            Object.Tag             =   "Arc"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":19DFD
            Key             =   "Rectangle"
            Object.Tag             =   "Rectangle"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":19F0F
            Key             =   "RoundRectangle"
            Object.Tag             =   "RoundRectangle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1A021
            Key             =   "Ellipse"
            Object.Tag             =   "Ellipse"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1A133
            Key             =   "Polygon"
            Object.Tag             =   "Polygon"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1A485
            Key             =   "Star"
            Object.Tag             =   "Star"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1A7D7
            Key             =   "Text"
            Object.Tag             =   "Text"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1A8E9
            Key             =   "Picture"
            Object.Tag             =   "Picture"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1AC3B
            Key             =   "map"
            Object.Tag             =   "map"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1AF8D
            Key             =   "legend1"
            Object.Tag             =   "legend1"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1B2DF
            Key             =   "scale1"
            Object.Tag             =   "scale1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1B631
            Key             =   "legend"
            Object.Tag             =   "legend"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1B983
            Key             =   "scale"
            Object.Tag             =   "scale"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1BCD5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar2 
      Align           =   3  'Align Left
      Height          =   7995
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   14102
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   375
      _CBHeight       =   7995
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar4"
      MinHeight1      =   315
      Width1          =   2640
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar4 
         Height          =   3300
         Left            =   30
         TabIndex        =   8
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   5821
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList2"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Select"
               Object.ToolTipText     =   "Select Object"
               Object.Tag             =   "Select"
               ImageIndex      =   1
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Line"
               Object.ToolTipText     =   "Draw Line"
               Object.Tag             =   "Line"
               ImageIndex      =   2
               Style           =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Arc"
               Object.ToolTipText     =   "Draw Arc"
               Object.Tag             =   "Arc"
               ImageIndex      =   3
               Style           =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Rectangle"
               Object.ToolTipText     =   "Draw Rectangle"
               Object.Tag             =   "Rectangle"
               ImageIndex      =   4
               Style           =   2
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "RoundRectangle"
               Object.ToolTipText     =   "Draw Round Rectangle"
               Object.Tag             =   "RoundRectangle"
               ImageIndex      =   5
               Style           =   2
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Ellipse"
               Object.ToolTipText     =   "Draw Ellipse"
               Object.Tag             =   "Ellipse"
               ImageIndex      =   6
               Style           =   2
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Polygon"
               Object.ToolTipText     =   "Draw Polygon"
               Object.Tag             =   "Polygon"
               ImageIndex      =   7
               Style           =   2
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Star"
               Object.ToolTipText     =   "Draw Star"
               Object.Tag             =   "Star"
               ImageIndex      =   8
               Style           =   2
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Text"
               Object.ToolTipText     =   "Draw Text"
               Object.Tag             =   "Text"
               ImageIndex      =   9
               Style           =   2
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Picture"
               Object.ToolTipText     =   "Insert Picture"
               Object.Tag             =   "Picture"
               ImageIndex      =   10
               Style           =   2
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "map"
               Object.ToolTipText     =   "insert map"
               Object.Tag             =   "map"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "legend"
               Object.ToolTipText     =   "insert legend"
               Object.Tag             =   "legend"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "scale"
               Object.ToolTipText     =   "insert scale bar"
               Object.Tag             =   "scale"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "arrow"
               Object.ToolTipText     =   "insert north arrow"
               ImageIndex      =   16
            EndProperty
         EndProperty
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   1323
      _CBWidth        =   12840
      _CBHeight       =   750
      _Version        =   "6.7.9782"
      Child1          =   "Toolbar1"
      MinWidth1       =   9405
      MinHeight1      =   330
      Width1          =   915
      NewRow1         =   0   'False
      Child2          =   "Toolbar3"
      MinWidth2       =   1200
      MinHeight2      =   330
      Width2          =   6225
      NewRow2         =   0   'False
      Child3          =   "Toolbar2"
      MinWidth3       =   5595
      MinHeight3      =   330
      Width3          =   825
      NewRow3         =   -1  'True
      AllowVertical3  =   0   'False
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   330
         Left            =   9795
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UnZoom"
               Object.ToolTipText     =   "UnZoom"
               Object.Tag             =   "UnZoom"
               ImageIndex      =   33
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Zoom-"
               Object.ToolTipText     =   "Zoom -"
               Object.Tag             =   "Zoom-"
               ImageIndex      =   34
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Zoom+"
               Object.ToolTipText     =   "Zoom+"
               Object.Tag             =   "Zoom+"
               ImageIndex      =   35
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   330
         Left            =   165
         TabIndex        =   5
         Top             =   390
         Visible         =   0   'False
         Width           =   12585
         _ExtentX        =   22199
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   18
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SelectAll"
               Object.ToolTipText     =   "Select All"
               Object.Tag             =   "SelectAll"
               ImageIndex      =   18
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "UnselectAll"
               Object.ToolTipText     =   "Unselect All"
               Object.Tag             =   "UnselectAll"
               ImageIndex      =   19
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "AlignLeft"
               Object.ToolTipText     =   "Align Left"
               Object.Tag             =   "AlignLeft"
               ImageIndex      =   20
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "AlignCenterVertical"
               Object.ToolTipText     =   "Align Center Vertical"
               Object.Tag             =   "AlignCenterVertical"
               ImageIndex      =   21
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "AlignRight"
               Object.ToolTipText     =   "Align Right"
               Object.Tag             =   "AlignRight"
               ImageIndex      =   22
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "AlignTop"
               Object.ToolTipText     =   "Align Top"
               Object.Tag             =   "AlignTop"
               ImageIndex      =   23
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "AlignCenterHorizontal"
               Object.ToolTipText     =   "Align Center Horizontal"
               Object.Tag             =   "AlignCenterHorizontal"
               ImageIndex      =   24
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "AlignBottom"
               Object.ToolTipText     =   "Align Bottom"
               Object.Tag             =   "AlignBottom"
               ImageIndex      =   25
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "AlignCenterHorVert"
               Object.ToolTipText     =   "Align Center Horizontal+Vertical"
               Object.Tag             =   "AlignCenterHorVert"
               ImageIndex      =   26
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "BringToFront"
               Object.ToolTipText     =   "Bring to Front"
               Object.Tag             =   "BringToFront"
               ImageIndex      =   27
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SendToBack"
               Object.ToolTipText     =   "Send to Back"
               Object.Tag             =   "SendToBack"
               ImageIndex      =   28
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "BringForward"
               Object.ToolTipText     =   "Bring Forward"
               Object.Tag             =   "BringForward"
               ImageIndex      =   29
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SendBackward"
               Object.ToolTipText     =   "Send Backward"
               Object.Tag             =   "SendBackward"
               ImageIndex      =   30
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Group"
               Object.ToolTipText     =   "Group"
               Object.Tag             =   "Group"
               ImageIndex      =   31
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Ungroup"
               Object.ToolTipText     =   "Ungroup"
               Object.Tag             =   "Ungroup"
               ImageIndex      =   32
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Visible         =   0   'False
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   23
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "New"
               Object.ToolTipText     =   "New"
               Object.Tag             =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Open"
               Object.ToolTipText     =   "Open"
               Object.Tag             =   "Open"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Save"
               Object.ToolTipText     =   "Save"
               Object.Tag             =   "Save"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Export"
               Object.ToolTipText     =   "Export"
               Object.Tag             =   "Export"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
               Object.Tag             =   "Cut"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
               Object.Tag             =   "Copy"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
               Object.Tag             =   "Paste"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo"
               Object.Tag             =   "Undo"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo"
               Object.Tag             =   "Redo"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete"
               Object.Tag             =   "Delete"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TextLeft"
               Object.ToolTipText     =   "Align Text Left"
               Object.Tag             =   "AlignText"
               ImageIndex      =   11
               Style           =   2
               Value           =   1
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TextCenter"
               Object.ToolTipText     =   "Align Text Center"
               Object.Tag             =   "AlignText"
               ImageIndex      =   12
               Style           =   2
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TextRight"
               Object.ToolTipText     =   "Align Text Right"
               Object.Tag             =   "AlignText"
               ImageIndex      =   13
               Style           =   2
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Bold"
               Object.Tag             =   "Bold"
               ImageIndex      =   14
               Style           =   1
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Italic"
               Object.Tag             =   "Italic"
               ImageIndex      =   15
               Style           =   1
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Underline"
               Object.Tag             =   "Underline"
               ImageIndex      =   16
               Style           =   1
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Strikethru"
               Object.ToolTipText     =   "Strikethru"
               Object.Tag             =   "Strikethru"
               ImageIndex      =   17
               Style           =   1
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.ComboBox CboFontName 
            Height          =   315
            Left            =   6540
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Font Name"
            Top             =   0
            Width           =   1905
         End
         Begin VB.ComboBox CboFontSize 
            Height          =   315
            IntegralHeight  =   0   'False
            Left            =   8460
            TabIndex        =   3
            Text            =   "15"
            ToolTipText     =   "Font Size"
            Top             =   0
            Width           =   705
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   7770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   35
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1C027
            Key             =   "New"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1C139
            Key             =   "Open"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1C24B
            Key             =   "Save"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1C35D
            Key             =   "Export"
            Object.Tag             =   "Export"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1C6AF
            Key             =   "Cut"
            Object.Tag             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1C7C1
            Key             =   "Copy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1C8D3
            Key             =   "Paste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1C9E5
            Key             =   "Undo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1CAF7
            Key             =   "Redo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1CC09
            Key             =   "Delete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1CD1B
            Key             =   "TextLeft"
            Object.Tag             =   "TextLeft"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1CE2D
            Key             =   "TextCenter"
            Object.Tag             =   "TextCenter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1CF3F
            Key             =   "TextRight"
            Object.Tag             =   "TextRight"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1D051
            Key             =   "Bold"
            Object.Tag             =   "Bold"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1D163
            Key             =   "Italic"
            Object.Tag             =   "Italic"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1D275
            Key             =   "Underline"
            Object.Tag             =   "Underline"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1D387
            Key             =   "Strikethru"
            Object.Tag             =   "Strikethru"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1D499
            Key             =   "SelectAll"
            Object.Tag             =   "SelectAll"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1D7EB
            Key             =   "UnselectAll"
            Object.Tag             =   "UnselectAll"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1DB3D
            Key             =   "AlignLeft"
            Object.Tag             =   "AlignLeft"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1DE8F
            Key             =   "AlignCenterVertical"
            Object.Tag             =   "AlignCenterVertical"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1E1E1
            Key             =   "AlignRight"
            Object.Tag             =   "AlignRight"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1E533
            Key             =   "AlignTop"
            Object.Tag             =   "AlignTop"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1E885
            Key             =   "AlignCenterHorizontal"
            Object.Tag             =   "AlignCenterHorizontal"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1EBD7
            Key             =   "AlignBottom"
            Object.Tag             =   "AlignBottom"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1EF29
            Key             =   "AlignCenterVerticalHorizontal"
            Object.Tag             =   "AlignCenterVerticalHorizontal"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1F27B
            Key             =   "BringToFront"
            Object.Tag             =   "BringToFront"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1F38D
            Key             =   "SendToBack"
            Object.Tag             =   "SendToBack"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1F49F
            Key             =   "BringForward"
            Object.Tag             =   "BringForward"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1F5B1
            Key             =   "SendBackward"
            Object.Tag             =   "SendBackward"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1F6C3
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1F7D5
            Key             =   "Ungroup"
            Object.Tag             =   "Ungroup"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1F8E7
            Key             =   "Zoom100"
            Object.Tag             =   "Zoom100"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1FC39
            Key             =   "Zoom-"
            Object.Tag             =   "Zoom-"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMainPrint.frx":1FF8B
            Key             =   "Zoom+"
            Object.Tag             =   "Zoom+"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicLoad 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   1350
      ScaleHeight     =   525
      ScaleWidth      =   375
      TabIndex        =   0
      Top             =   5820
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu SmnuFile 
         Caption         =   "New"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Open Print Template"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Open Map File"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Save"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Export to Bitmap"
         Index           =   4
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "-"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Print"
         Index           =   6
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu SmnuFile 
         Caption         =   "Exit"
         Index           =   8
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Visible         =   0   'False
      Begin VB.Menu SmnuEdit 
         Caption         =   "Undo"
         Index           =   0
         Shortcut        =   ^Z
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Redo"
         Index           =   1
         Shortcut        =   ^Y
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Cut"
         Index           =   3
         Shortcut        =   ^X
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Copy"
         Index           =   4
         Shortcut        =   ^C
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Paste"
         Index           =   5
         Shortcut        =   ^V
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Delete"
         Index           =   7
         Shortcut        =   {DEL}
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Select All"
         Index           =   9
         Shortcut        =   ^A
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Group"
         Index           =   11
      End
      Begin VB.Menu SmnuEdit 
         Caption         =   "Ungroup"
         Index           =   12
      End
   End
   Begin VB.Menu mnuZoom 
      Caption         =   "Zoom (100%)"
      Visible         =   0   'False
      Begin VB.Menu SmnuZoom 
         Caption         =   "10%"
         Index           =   0
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "25%"
         Index           =   1
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "50%"
         Index           =   2
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "100%"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "150%"
         Index           =   4
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "200%"
         Index           =   5
      End
      Begin VB.Menu SmnuZoom 
         Caption         =   "400%"
         Index           =   6
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu SmnuOptions 
         Caption         =   "Canvas Size"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMainPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private doNothing    As Boolean
Private Answer       As VbMsgBoxResult
Private Modified     As Boolean

Private mTxtAlign    As AlignmentConstants
Private mBold        As Boolean
Private mItalic      As Boolean
Private mUnderline   As Boolean
Private mStrikethru  As Boolean
Private ColorIndex   As Integer

Private bFillColor   As Long
Private bBorderColor As Long
Private bBackColor   As Long
Private bPtsQty      As Integer
Private bLoading     As Integer
Private bLandscape As Boolean
Private colPaperSizes As New Collection
Private sSize() As String

Private Sub SuperDebug(sText As String)
    DebugPrint "Calling frmMainPrint." & sText, True
End Sub

Private Function FileExist(ByVal MyFile As String) As Boolean
        '<EhHeader>
        On Error GoTo FileExist_Err
        '</EhHeader>
        SuperDebug "sub/fun: FileExist"
100     FileExist = (Dir(MyFile) <> "")
        '<EhFooter>
        Exit Function

FileExist_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.FileExist " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Function

Private Sub CboFill_Click()
        '<EhHeader>
        On Error GoTo CboFill_Click_Err
        '</EhHeader>
SuperDebug "sub/fun: CboFill_Click"
100     If doNothing = True Then Exit Sub
102     If OASISDrawObj1.CurrentObject > -1 Then
104         OASISDrawObj1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
        End If

        '<EhFooter>
        Exit Sub

CboFill_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.CboFill_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub CboFontName_Click()
        '<EhHeader>
        On Error GoTo CboFontName_Click_Err
        '</EhHeader>
        On Error Resume Next
SuperDebug "sub/fun: CboFontName_Click"
100     If doNothing = True Then Exit Sub
102     If OASISDrawObj1.ObjectType = mText And doNothing = False Then
104         OASISDrawObj1.ModifyObject , , , , , , , , , , CboFontName.Text
        End If

106     OASISDrawObj1.SetFocus
        '<EhFooter>
        Exit Sub

CboFontName_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.CboFontName_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub CboFontSize_Change()
        '<EhHeader>
        On Error GoTo CboFontSize_Change_Err
        '</EhHeader>
        On Error Resume Next
SuperDebug "sub/fun: CboFontSize_Change"
100     If doNothing = True Then Exit Sub
102     If OASISDrawObj1.ObjectType = mText And doNothing = False And Len(Trim(CboFontSize.Text)) > 0 Then
104         OASISDrawObj1.ModifyObject , , , , , , , , , , , CboFontSize.Text
        End If

106     OASISDrawObj1.SetFocus
        '<EhFooter>
        Exit Sub

CboFontSize_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.CboFontSize_Change " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub CboFontSize_Click()
        '<EhHeader>
        On Error GoTo CboFontSize_Click_Err
        '</EhHeader>
SuperDebug "sub/fun: CboFontSize_Click"
100     If OASISDrawObj1.ObjectType = mText And doNothing = False Then
102         OASISDrawObj1.ModifyObject , , , , , , , , , , , CboFontSize.Text
        End If

        '<EhFooter>
        Exit Sub

CboFontSize_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.CboFontSize_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCommand1_Click()
        '<EhHeader>
        On Error GoTo cmdCommand1_Click_Err
        '</EhHeader>
100     propertyGrid.Clear
    SuperDebug "sub/fun: cmdCommand1_Click"
102     With OASISDrawObj1
            '.CurrentObject
104         AddProperty "Height", .Height
106         AddProperty "Left", .left
108         AddProperty "Type", .ObjectType
110         AddProperty "Height", .top
112         AddProperty "Height", .Width
        
        End With
        '<EhFooter>
        Exit Sub

cmdCommand1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.cmdCommand1_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ColorPal1_ColorOver(cColor As Long)
        '<EhHeader>
        On Error GoTo ColorPal1_ColorOver_Err
        '</EhHeader>
        SuperDebug "sub/fun: ColorPal1_ColorOver"
        Dim sTmp As String
100     sTmp = right("000000" & Hex(cColor), 6)
102     LblColor.caption = "Hex:" & sTmp & vbCrLf & " Red:" & Int("&H" & right$(sTmp, 2)) & " - Green:" & Int("&H" & Mid$(sTmp, 3, 2)) & " - Blue:" & Int("&H" & left$(sTmp, 2))

        '<EhFooter>
        Exit Sub

ColorPal1_ColorOver_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.ColorPal1_ColorOver " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ColorPal1_ColorSelected(cColor As Long)
        '<EhHeader>
        On Error GoTo ColorPal1_ColorSelected_Err
        '</EhHeader>
        SuperDebug "sub/fun: ColorPal1_ColorSelected"
        Dim sTmp As String
100     sTmp = right("000000" & Hex(cColor), 6)
102     ScrCol(0).value = Int("&H" & right$(sTmp, 2))
104     ScrCol(1).value = Int("&H" & Mid$(sTmp, 3, 2))
106     ScrCol(2).value = Int("&H" & left$(sTmp, 2))

108     OpColor(ColorIndex).BackColor = cColor
110     bFillColor = OpColor(0).BackColor
112     bBorderColor = OpColor(1).BackColor
114     bBackColor = OpColor(2).BackColor

116     If doNothing = True Then Exit Sub

118     Select Case ColorIndex

            Case 0

120             If OASISDrawObj1.CurrentObject > -1 Then
122                 OASISDrawObj1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
                End If

124         Case 1

126             If OASISDrawObj1.CurrentObject > -1 Then
128                 OASISDrawObj1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
                End If

130         Case 2
132             OASISDrawObj1.BackColor = bBackColor
        End Select

134     OASISDrawObj1.SetFocus
        '<EhFooter>
        Exit Sub

ColorPal1_ColorSelected_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.ColorPal1_ColorSelected " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ColorPal1_MouseDown(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
        '<EhHeader>
        On Error GoTo ColorPal1_MouseDown_Err
        '</EhHeader>
        Dim c As New cCommonDialog
SuperDebug "sub/fun: ColorPal1_MouseDown"
100     If Button = 2 Then

102         With c
104             .DialogTitle = "Open Palette"
106             .Filter = "Palette (*.pal)|*.pal"
108             .Filename = ""
110             .ShowOpen
112             .Filename = Trim(.Filename)

114             If Len(.Filename) > 0 And FileExist(.Filename) = True Then
116                 ColorPal1.LoadPalette .Filename
                End If

            End With

        End If

        '<EhFooter>
        Exit Sub

ColorPal1_MouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.ColorPal1_MouseDown " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ColorPicker1_Click()
        '<EhHeader>
        On Error GoTo ColorPicker1_Click_Err
        '</EhHeader>
        SuperDebug "sub/fun: ColorPicker1_Click"
100     OASISDrawObj1.BackColor = ColorPicker1.color
        '<EhFooter>
        Exit Sub

ColorPicker1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.ColorPicker1_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComPOrentation_Click()
        '<EhHeader>
        On Error GoTo ComPOrentation_Click_Err
        '</EhHeader>
    SuperDebug "sub/fun: ComPOrentation_Click"
100     If bLoading Then Exit Sub
 
102     If ComPOrentation.ListIndex = 0 Then
104         bLandscape = False
        Else
106         bLandscape = True
        End If
    
108     SetPaperSize
    
        '<EhFooter>
        Exit Sub

ComPOrentation_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.ComPOrentation_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComPQuality_Click()
        '<EhHeader>
        On Error GoTo ComPQuality_Click_Err
        SuperDebug "sub/fun: ComPQuality_Click"
        '</EhHeader>
100     If bLoading Then Exit Sub
        '<EhFooter>
        Exit Sub

ComPQuality_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.ComPQuality_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComPSize_Click()
        '<EhHeader>
        On Error GoTo ComPSize_Click_Err
        SuperDebug "sub/fun: ComPSize_Click"
        '</EhHeader>
    
100     If bLoading Then Exit Sub
    
102     sSize = Split(colPaperSizes.Item(ComPSize.ListIndex + 1), "x")
               
104     SetPaperSize ScaleX(sSize(0), vbMillimeters, vbPixels), ScaleY(sSize(1), vbMillimeters, vbPixels)
    
        '<EhFooter>
        Exit Sub

ComPSize_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.ComPSize_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetPaperSize(Optional X As Double, Optional Y As Double)
        '<EhHeader>
        On Error GoTo SetPaperSize_Err
        '</EhHeader>
SuperDebug "sub/fun: SetPaperSize"
100     If X = 0 And Y = 0 Then
102         X = OASISDrawObj1.CanvasWidth
104         Y = OASISDrawObj1.CanvasHeight
106         OASISDrawObj1.CanvasWidth = Y
108         OASISDrawObj1.CanvasHeight = X
            Exit Sub
        End If
    
110     If Not bLandscape Then
112         OASISDrawObj1.CanvasWidth = X
114         OASISDrawObj1.CanvasHeight = Y
        Else
116         OASISDrawObj1.CanvasWidth = Y
118         OASISDrawObj1.CanvasHeight = X
        End If

        '<EhFooter>
        Exit Sub

SetPaperSize_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.SetPaperSize " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComPUnits_Click()
        '<EhHeader>
        On Error GoTo ComPUnits_Click_Err
        '</EhHeader>
        SuperDebug "sub/fun: ComPUnits_Click"
100 If bLoading Then Exit Sub

        '<EhFooter>
        Exit Sub

ComPUnits_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.ComPUnits_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub CoolBar1_HeightChanged(ByVal newHeight As Single)
        '<EhHeader>
        On Error GoTo CoolBar1_HeightChanged_Err
        SuperDebug "sub/fun: CoolBar1_HeightChanged"
        '</EhHeader>
100     Form_Resize
        '<EhFooter>
        Exit Sub

CoolBar1_HeightChanged_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.CoolBar1_HeightChanged " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub CoolBar3_HeightChanged(ByVal newHeight As Single)
        '<EhHeader>
        On Error GoTo CoolBar3_HeightChanged_Err
        SuperDebug "sub/fun: CoolBar3_HeightChanged"
        '</EhHeader>
100     Form_Resize
        '<EhFooter>
        Exit Sub

CoolBar3_HeightChanged_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.CoolBar3_HeightChanged " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        SuperDebug "sub/fun: Form_Load"
        Dim n As Integer
100     bLoading = True
        
        sSize = Split("210x297", "x") ' A4 Default
               
        'Property Page Code
        'configureCombo
        'configureCategories
        
        Debug.Print "configureGrid" & Now()
        
102     configureGrid
   
104     'addSampleData
        'propertyGrid.CellSelected(1, 1) = True
   
        'End PropertyGrid
   
106     CboFontName.Clear

108     For n = 1 To Screen.FontCount - 1
110         CboFontName.AddItem Screen.Fonts(n)
112     Next n

114     CboFontName.Text = "Arial"

116     For n = 5 To 100
118         CboFontSize.AddItem n
120     Next n

122     CboFontSize.Text = 15
124     CboFill.ListIndex = 0
126     ColorIndex = 0

128     bFillColor = OpColor(0).BackColor
130     bBorderColor = OpColor(1).BackColor
132     bBackColor = OpColor(2).BackColor

        'CreateDefaultTemplate
    
134     fillOrientation ComPOrentation
136     ComPOrentation.ListIndex = 0
        bLandscape = True
138     fillPaperSizes ComPSize
    
140     ComPSize.ListIndex = 0
142     fillUnits ComPUnits
144     ComPUnits.ListIndex = 0
146     fillQuality ComPQuality
148     ComPQuality.ListIndex = 0

99      SetPaperSize ScaleX(sSize(0), vbMillimeters, vbPixels), ScaleY(sSize(1), vbMillimeters, vbPixels)
        Debug.Print "SetPaperSize" & Now()
150     bLoading = False

        'InitPrint "", 0, 0, 0, 0
        Debug.Print "InitPrint" & Now()
        
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.Form_Load " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub fillQuality(comb As ComboBox)
        '<EhHeader>
        On Error GoTo fillQuality_Err
        SuperDebug "sub/fun: fillQuality"
        '</EhHeader>

100     With comb
102         .Clear
104         .AddItem "96 dpi"
106         .AddItem "150 dpi"
108         .AddItem "300 dpi"
110         .AddItem "500 dpi"
112         .AddItem "1000 dpi"
114         .AddItem "2000 dpi"
116         .AddItem "3000 dpi"
        End With

        '<EhFooter>
        Exit Sub

fillQuality_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.fillQuality " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub
Private Sub fillOrientation(comb As ComboBox)
        '<EhHeader>
        On Error GoTo fillOrientation_Err
        SuperDebug "sub/fun: fillOrientation"
        '</EhHeader>

100     With comb
102         .Clear
104         .AddItem "Landscape"
106         .AddItem "Portrait"
        End With

        '<EhFooter>
        Exit Sub

fillOrientation_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.fillOrientation " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub fillUnits(comb As ComboBox)
        '<EhHeader>
        On Error GoTo fillUnits_Err
        '</EhHeader>
SuperDebug "sub/fun: fillUnits"
100     With comb
102         .Clear
104         .AddItem "mm"
106         .AddItem "inch"
108         .AddItem "pixels"
        End With

        '<EhFooter>
        Exit Sub

fillUnits_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.fillUnits " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub prepareStuff()
    'ComPSize
        '<EhHeader>
        On Error GoTo prepareStuff_Err
        '</EhHeader>
        SuperDebug "sub/fun: prepareStuff"
        Set colPaperSizes = New Collection
        
100     With colPaperSizes
102         .Add ScaleX(600, vbPixels, vbMillimeters) & "x" & ScaleY(800, vbPixels, vbMillimeters), "Pixel 600x800"
104         .Add "1189x841", "A0 (1189x841 mm)"
106         .Add "841x594", "A1 (841x594 mm)"
108         .Add "594x420", "A2 (594x420 mm)"
110         .Add "420x297", "A3 (420x297 mm)"
112         .Add "297x210", "A4 (297x210 mm)"
114         .Add "210x148", "A5 (210x148 mm)"
116         .Add "148x105", "A6 (148x105 mm)"
118         .Add "105x74", "A7 (105x74 mm)"
120         .Add "1414x1000", "B0 (1414x1000 mm)"
122         .Add "1000x707", "B1 (1000x707 mm)"
124         .Add "707x500", "B2 (707x500 mm)"
126         .Add "500x353", "B3 (500x353 mm)"
128         .Add "353x250", "B4 (353x250 mm)"
130         .Add "250x176", "B5 (250x176 mm)"
132         .Add "176x125", "B6 (176x125 mm)"
134         .Add "125x88", "B7 (125x88 mm)"
136         .Add "88x62", "B8 (88x62 mm)"
138         .Add "62x44", "B9 (62x44 mm)"
140         .Add "44x31", "B10 (44x31 mm)"
142         .Add "279.4x215.9", "Letter (ANSI A) (279.4x215.9 mm)"
144         .Add "355.6x215.9", "Legal (355.6x215.9 mm)"
146         .Add "431.8x279.4", "Ledger (ANSI B) (431.8x279.4 mm)"
148         .Add "279.4x431.8", "Tabloid (ANSI B) (279.4x431.8 mm)"
150         .Add "266.7x184.1", "Executive (266.7x184.1 mm)"
152         .Add "432x559", "ANSI C (432x559 mm)"
154         .Add "559x864", "ANSI D (559x864 mm)"
156         .Add "864x1118", "ANSI E (864x1118 mm)"
        End With
        '<EhFooter>
        Exit Sub

prepareStuff_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.prepareStuff " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub fillPaperSizes(comb As ComboBox)
        '<EhHeader>
        On Error GoTo fillPaperSizes_Err
        '</EhHeader>
    SuperDebug "sub/fun: fillPaperSizes"

100     With comb
102         .Clear
104         prepareStuff
106         .AddItem "Pixel (600x800 px)"
108         .AddItem "A0 (1189x841 mm)"
110         .AddItem "A1 (841x594 mm)"
112         .AddItem "A2 (594x420 mm)"
114         .AddItem "A3 (420x297 mm)"
116         .AddItem "A4 (297x210 mm)"
118         .AddItem "A5 (210x148 mm)"
120         .AddItem "A6 (148x105 mm)"
122         .AddItem "A7 (105x74 mm)"
124         .AddItem "B0 (1414x1000 mm)"
126         .AddItem "B1 (1000x707 mm)"
128         .AddItem "B2 (707x500 mm)"
130         .AddItem "B3 (500x353 mm)"
132         .AddItem "B4 (353x250 mm)"
134         .AddItem "B5 (250x176 mm)"
136         .AddItem "B6 (176x125 mm)"
138         .AddItem "B7 (125x88 mm)"
140         .AddItem "B8 (88x62 mm)"
142         .AddItem "B9 (62x44 mm)"
144         .AddItem "B10 (44x31 mm)"
146         .AddItem "Letter (ANSI A) (8.5×11 in)'(279.4x215.9 mm)"
148         .AddItem "Legal (8.5×14 in)" '(355.6x215.9 mm)"
150         .AddItem "Ledger (ANSI B) (17×11 in)" '(431.8x279.4 mm)"
152         .AddItem "Tabloid (ANSI B) (11×17 in)" '(279.4x431.8 mm)"
154         .AddItem "Executive (266.7x184.1 mm)"
156         .AddItem "ANSI C (432x559 mm)"
158         .AddItem "ANSI D (559x864 mm)"
160         .AddItem "ANSI E (864x1118 mm)"
    
        End With

        '<EhFooter>
        Exit Sub

fillPaperSizes_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.fillPaperSizes " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Public Sub InitPrint(sPath As String, xmin As Double, xmax As Double, ymin As Double, ymax As Double, Optional arSelections As Variant)
        '<EhHeader>
        On Error GoTo InitPrint_Err
        '</EhHeader>
        SuperDebug "sub/fun: InitPrint"
        Dim arSelVals() As String
        
        If sPath = "" Then
            sPath = g_sAppPath & "\data\user\Maps\DefaultMap.TTKGP"
            If Not FileExist(sPath) Then
                Err.Raise 666, "OASIS Print module", "File does not exist:" & sPath
            End If
        End If
        
        frmPrintLoader.Show vbModeless, Me

        DoProgress "Loading Print Utilities...", "Preparing map..."
        
        Debug.Print "SetGISComponent" & Now

100     OASISDrawObj1.SetGISComponent sPath, xmin, xmax, ymin, ymax

        If Len(arSelections) > 0 Then
            arSelVals = Split(arSelections, ":::")
            OASISDrawObj1.SetSelection arSelVals(0), arSelVals(1)
        End If
        
        DoProgress "Loading Print Utilities...", "Preparing map ready..."
        
        Debug.Print "After SetGISComponent" & Now
        DoProgress "Loading Print Utilities...", "Preparing Default template..."
102     CreateDefaultTemplate
        Debug.Print "After CreateDefaultTemplate" & Now
        Unload frmPrintLoader
        '<EhFooter>
        Exit Sub

InitPrint_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.InitPrint " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub DoProgress(sCaption As String, sAction As String)
    frmPrintLoader.lblProgress.caption = sCaption
    frmPrintLoader.txtLoading.Text = sAction & vbCrLf & frmPrintLoader.txtLoading.Text
    DoEvents
End Sub
Private Sub Form_Paint()
        '<EhHeader>
        On Error GoTo Form_Paint_Err
        SuperDebug "sub/fun: Form_Paint"
        '</EhHeader>
100     UpdToolBar
102     tabMain.Visible = False
104     tabMain.Visible = True
        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.Form_Paint " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
SuperDebug "sub/fun: Form_Resize"
    If Me.WindowState <> 1 Then
        OASISDrawObj1.Width = Me.Width - CoolBar2.Width - CoolBar3.Width - 175
        OASISDrawObj1.Height = Me.Height - CoolBar1.Height - StatusBar1.Height - 890
        OASISDrawObj1.top = CoolBar1.Height
        OASISDrawObj1.left = CoolBar2.Width
        'pctProperties.Height = ScaleY(CoolBar3.Bands.Item(2).Height, vbPixels, vbTwips)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
        Dim c As New cCommonDialog
        SuperDebug "sub/fun: Form_Unload"
        Exit Sub
100     If Modified = True Then
102         Answer = MsgBox("The map print template has been modified, do you want to save it?", vbDefaultButton1 + vbYesNoCancel, "OASIS Print Template Tool")

104         If Answer = vbYes Then
106             Cancel = True

108             With c
110                 .DialogTitle = "Save Print Template"
112                 .Filter = "OASIS Print Template File (*.ojp)|*.ojp"
114                 .Filename = ""
116                 .ShowSave
118                 .Filename = Trim(.Filename)

120                 If Len(.Filename) > 0 Then
122                     OASISDrawObj1.SaveProjects .Filename
                    End If

                End With

124             End
126         ElseIf Answer = vbCancel Then
128             Cancel = True
                Exit Sub
            End If
        End If


        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.Form_Unload " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub LblColor_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
        '<EhHeader>
        On Error GoTo LblColor_MouseMove_Err
        SuperDebug "sub/fun: LblColor_MouseMove"
        '</EhHeader>
100     LblColor.caption = ""
        '<EhFooter>
        Exit Sub

LblColor_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.LblColor_MouseMove " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuEdit_Click()
        '<EhHeader>
        On Error GoTo mnuEdit_Click_Err
        '</EhHeader>
SuperDebug "sub/fun: mnuEdit_Click"
100     If OASISDrawObj1.CurrentObject > -1 Then
102         SmnuEdit(3).Enabled = True
104         SmnuEdit(4).Enabled = True
106         SmnuEdit(7).Enabled = True
        Else
108         SmnuEdit(3).Enabled = False
110         SmnuEdit(4).Enabled = False
112         SmnuEdit(7).Enabled = False
        End If

114     SmnuEdit(5).Enabled = OASISDrawObj1.ObjectInClipboard
        '<EhFooter>
        Exit Sub

mnuEdit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.mnuEdit_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub mnuZoom_Click()
        '<EhHeader>
        On Error GoTo mnuZoom_Click_Err
        SuperDebug "sub/fun: mnuZoom_Click"
        '</EhHeader>
100     StatusBar1.Panels(1).Text = "You can also change the Zoom Factor with ""+"" & ""-"" on KeyPad"
        '<EhFooter>
        Exit Sub

mnuZoom_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.mnuZoom_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OASISDrawObj1_KeyDown(KeyAscii As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo OASISDrawObj1_KeyDown_Err
        '</EhHeader>
        SuperDebug "sub/fun: OASISDrawObj1_KeyDown"

100     If KeyAscii >= 37 And KeyAscii <= 40 Then
102         StatusBar1.Panels(1).Text = "Press ""Ctrl"" key with arrows keys to switch selection"
104     ElseIf KeyAscii = vbKeyAdd Or KeyAscii = vbKeySubtract Then
106         mnuZoom.caption = "Zoom (" & Round(OASISDrawObj1.ZoomFactor * 100) & "%)"
        End If

        '<EhFooter>
        Exit Sub

OASISDrawObj1_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OASISDrawObj1_KeyDown " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OASISDrawObj1_KeyUp(KeyAscii As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo OASISDrawObj1_KeyUp_Err
        '</EhHeader>
        SuperDebug "sub/fun: OASISDrawObj1_KeyUp"
100     StatusBar1.Panels(1).Text = ""
        '<EhFooter>
        Exit Sub

OASISDrawObj1_KeyUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OASISDrawObj1_KeyUp " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OASISDrawObj1_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
        '<EhHeader>
        On Error GoTo OASISDrawObj1_MouseMove_Err
        SuperDebug "sub/fun: OASISDrawObj1_MouseMove"
        '</EhHeader>
100     Select Case ComPUnits.ListIndex
            Case 0
102             X = Round(ScaleX(X, vbPixels, vbMillimeters), 0)
104             Y = Round(ScaleY(Y, vbPixels, vbMillimeters), 0)
106         Case 1
108             X = Round(ScaleX(X, vbPixels, vbInches), 2)
110             Y = Round(ScaleY(Y, vbPixels, vbInches), 2)
        End Select
    
112     StatusBar1.Panels(2).Text = "X: " & X & " - Y: " & Y
        '<EhFooter>
        Exit Sub

OASISDrawObj1_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OASISDrawObj1_MouseMove " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OASISDrawObj1_NewDrawingEnd()
        '<EhHeader>
        On Error GoTo OASISDrawObj1_NewDrawingEnd_Err
        SuperDebug "sub/fun: OASISDrawObj1_NewDrawingEnd"
        '</EhHeader>
100     Toolbar4.Buttons(1).value = tbrPressed
        '<EhFooter>
        Exit Sub

OASISDrawObj1_NewDrawingEnd_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OASISDrawObj1_NewDrawingEnd " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OASISDrawObj1_ObjectResize(ObjType As myObType, Index As Long, ObjLeft As Single, ObjTop As Single, ObjWidth As Single, ObjHeight As Single, ObjAspect As Single)
        '<EhHeader>
        On Error GoTo OASISDrawObj1_ObjectResize_Err
        SuperDebug "sub/fun: OASISDrawObj1_ObjectResize"
        '</EhHeader>

        Dim tmp As String

100     Select Case ObjType

            Case mline
102             tmp = "Line"

104         Case mArc
106             tmp = "Arc"

108         Case mRectangle

110             If ObjAspect = 0 Then
112                 tmp = "Rectangle"
                Else
114                 tmp = "Square"
                End If

116         Case mEllipse

118             If ObjAspect = 0 Then
120                 tmp = "Ellipse"
                Else
122                 tmp = "Circle"
                End If

124         Case mText
126             tmp = "Text"

128         Case mImage
130             tmp = "Image"
        End Select

132     StatusBar1.Panels(3).Text = tmp & "   Pos. X:" & ObjLeft & "  Y:" & ObjTop & "   Size W:" & ObjWidth & "  H:" & ObjHeight & " "


        '<EhFooter>
        Exit Sub

OASISDrawObj1_ObjectResize_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OASISDrawObj1_ObjectResize " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OASISDrawObj1_ObjSelected(ObjType As myObType, oGISObj As myGISobj, Index As Long, ObjLeft As Single, ObjTop As Single, ObjWidth As Single, ObjHeight As Single, ObjAngle As Single, ObjFillColor As Long, ObjFillStyle As myFill, ObjBorderColor As Long, ObjBorderWidth As Integer, ObjAspect As Single, ObjFontName As String, ObjFontSize As Single, ObjFontBold As Boolean, ObjFontItalic As Boolean, ObjFontUnderline As Boolean, ObjFontStrikethru As Boolean, ObjText As String, ObjTextAlign As AlignmentConstants, ObjPointQty As Integer)
        '<EhHeader>
        On Error GoTo OASISDrawObj1_ObjSelected_Err
        '</EhHeader>
SuperDebug "sub/fun: OASISDrawObj1_ObjSelected"
        Dim tmp As String

100     If ObjType <> -1 Then
        
102         propertyGrid.Clear
        
104         doNothing = True

106         If ObjFillColor > -1 Then bFillColor = ObjFillColor
108         OpColor(0).BackColor = bFillColor

110         If ObjFillStyle > -1 Then CboFill.ListIndex = ObjFillStyle
112         If ObjAngle > -1 Then Slider1.value = ObjAngle
114         Label1(0).caption = "Rotation: " & Slider1.value & "°"

116         If ObjBorderColor > -1 Then bBorderColor = ObjBorderColor
118         OpColor(1).BackColor = bBorderColor

120         If ObjType <> mText Then VScroll1.value = ObjBorderWidth
122         TxtBorder.Text = VScroll1.value

124         If ObjFontName <> "" Then CboFontName.Text = ObjFontName
126         If ObjFontSize > 0 Then CboFontSize.Text = ObjFontSize
128         If ObjType = mPolygon Or ObjType = mStar Then
130             If ObjPointQty > 0 And ObjPointQty <= 30 Then
132                 VScroll2.value = ObjPointQty
134                 TxtPoint.Text = ObjPointQty
                End If

136         ElseIf ObjType = mRoundRectangle Then
138             VScroll3.value = ObjPointQty
140             TxtRound.Text = ObjPointQty
            End If

142         TxtPoint.Text = VScroll2.value
144         mBold = CBool(Int(ObjFontBold))
146         Toolbar1.Buttons(19).value = Abs(Int(ObjFontBold))
148         mItalic = CBool(Int(ObjFontItalic))
150         Toolbar1.Buttons(20).value = Abs(Int(ObjFontItalic))
152         mUnderline = CBool(Int(ObjFontUnderline))
154         Toolbar1.Buttons(21).value = Abs(Int(ObjFontUnderline))
156         mStrikethru = CBool(Int(ObjFontStrikethru))
158         Toolbar1.Buttons(22).value = Abs(Int(ObjFontStrikethru))

160         If ObjTextAlign > -1 Then
162             mTxtAlign = ObjTextAlign

164             Select Case mTxtAlign

                    Case vbLeftJustify
166                     Toolbar1.Buttons(15).value = tbrPressed

168                 Case vbRightJustify
170                     Toolbar1.Buttons(17).value = tbrPressed

172                 Case vbCenter
174                     Toolbar1.Buttons(16).value = tbrPressed
                End Select

            End If

176         doNothing = False

178         Select Case ObjType

                Case mline
180                 tmp = "Line"
                  
182             Case mArc
184                 tmp = "Arc"

186             Case mRectangle

188                 If ObjAspect = 0 Then
190                     tmp = "Rectangle"
                    Else
192                     tmp = "Square"
                    End If

194             Case mEllipse

196                 If ObjAspect = 0 Then
198                     tmp = "Ellipse"
                    Else
200                     tmp = "Circle"
                    End If

202             Case mText
204                 tmp = "Text"

206             Case mImage
208                 tmp = "Image"

210             Case mPolygon
212                 tmp = "Polygon"

214             Case mStar
216                 tmp = "Star"
            End Select

218         propertyGrid.Clear
                
220         Select Case oGISObj
        
                Case 1 'Map
222                 AddProperty "Type", "Map"
224                 AddProperty "Map Color", OASISDrawObj1.MapColor
226                 AddProperty "# of Layers", OASISDrawObj1.MapLayersNumbers
228                 AddProperty "Project Name", OASISDrawObj1.MapProjectName
230                 AddProperty "Rotation Angle", OASISDrawObj1.MapRotationAngle
232                 AddProperty "Rotation X", OASISDrawObj1.MapRotationPointX
234                 AddProperty "Rotation Y", OASISDrawObj1.MapRotationPointY
236                 AddProperty "Scale", OASISDrawObj1.MapScale
238                 AddProperty "Coord Sys Pretty WKT", OASISDrawObj1.MapCoordPrettyWKT
240                 AddProperty "Coord Sys Description", OASISDrawObj1.MapCoordSysDesc
242                 AddProperty "Coord Sys EPSG", OASISDrawObj1.MapCoordSysEPSG
244                 AddProperty "Coord Sys Name", OASISDrawObj1.MapCoordSysName
246                 AddProperty "Coord Sys WKT", OASISDrawObj1.MapCoordSysWKT
248             Case 2 'Legend
250                 AddProperty "Type", "Legend"
252                 AddProperty "Color", OASISDrawObj1.LegendColor
254             Case 3 'North arrow
256                 AddProperty "Type", "North Arrow"
258                 AddProperty "Color", OASISDrawObj1.NAColor
260                 AddProperty "Color1", OASISDrawObj1.NAColor1
262                 AddProperty "Color2", OASISDrawObj1.NAColor2
264                 AddProperty "Font Color", OASISDrawObj1.NAFontColor
266                 AddProperty "Path", OASISDrawObj1.NAPath
268                 AddProperty "Transparent", OASISDrawObj1.NATransparent
270             Case 4 'ScaleBar
272                 AddProperty "Type", "Scale Bar"
274                 AddProperty "Color", OASISDrawObj1.ScaleBarColor
276                 AddProperty "Dividers", OASISDrawObj1.ScaleBarDividers
278                 AddProperty "Font Color", OASISDrawObj1.ScaleBarFontColor
280                 AddProperty "Unit Desc", OASISDrawObj1.ScaleBarUnitsDesc
282                 AddProperty "Unit EPSG", OASISDrawObj1.ScaleBarUnitsEPSG
284                 AddProperty "Unit Name", OASISDrawObj1.ScaleBarUnitsName
286                 AddProperty "Unit Symbol", OASISDrawObj1.ScaleBarUnitsSymbol
288                 AddProperty "Unit Type", OASISDrawObj1.ScaleBarUnitsType
290                 AddProperty "Unit WKT", OASISDrawObj1.ScaleBarUnitsWKT
292             Case 5 'Grid
            
294             Case Else
296                 AddProperty "Type", tmp
            End Select
        
298         AddProperty "Index", Index
300         AddProperty "Left", ObjLeft
302         AddProperty "Top", ObjTop
304         AddProperty "Width", ObjWidth
306         AddProperty "Height", ObjHeight
308         AddProperty "Angle", ObjAngle
            'AddProperty "Fill Color", ObjFillColor
            'AddProperty "Fill Style", ObjFillStyle
            'AddProperty "Border Color", ObjBorderColor
            'AddProperty "Border Width", ObjBorderWidth
            'AddProperty "Aspect", ObjAspect
        
            'AddProperty "Font Name", ObjFontName
            'AddProperty "Font Size", ObjFontSize
            'AddProperty "Font Bold", ObjFontBold
            'AddProperty "Font Italic", ObjFontItalic
            'AddProperty "Font Underline", ObjFontUnderline
            'AddProperty "Font Strikethru", ObjFontStrikethru
310         If tmp = "Text" Then AddProperty "Text", ObjText
            'AddProperty "Align", ObjTextAlign
            'AddProperty "Point Qty", ObjPointQty
312         StatusBar1.Panels(3).Text = tmp & "   Pos. X:" & ObjLeft & "  Y:" & ObjTop & "   Size W:" & ObjWidth & "  H:" & ObjHeight & " "
        Else
314         StatusBar1.Panels(3).Text = ""
        End If

316     UpdToolBar

        '<EhFooter>
        Exit Sub

OASISDrawObj1_ObjSelected_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OASISDrawObj1_ObjSelected " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OASISDrawObj1_Prompt2Save()
        '<EhHeader>
        On Error GoTo OASISDrawObj1_Prompt2Save_Err
        SuperDebug "sub/fun: OASISDrawObj1_Prompt2Save"
        '</EhHeader>
100     Modified = True
        '<EhFooter>
        Exit Sub

OASISDrawObj1_Prompt2Save_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OASISDrawObj1_Prompt2Save " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OASISDrawObj1_UndoRedo(LastUndo As Boolean, LastRedo As Boolean)
        '<EhHeader>
        On Error GoTo OASISDrawObj1_UndoRedo_Err
        SuperDebug "sub/fun: OASISDrawObj1_UndoRedo"
        '</EhHeader>
100     SmnuEdit(0).Enabled = Not LastUndo
102     SmnuEdit(1).Enabled = Not LastRedo
104     Toolbar1.Buttons(10).Enabled = Not LastUndo
106     Toolbar1.Buttons(11).Enabled = Not LastRedo
        '<EhFooter>
        Exit Sub

OASISDrawObj1_UndoRedo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OASISDrawObj1_UndoRedo " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OpColor_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo OpColor_Click_Err
        SuperDebug "sub/fun: OpColor_Click"
        '</EhHeader>
        Dim sTmp As String
100     ColorIndex = Index

102     sTmp = right("000000" & Hex(OpColor(Index).BackColor), 6)
104     ScrCol(0).value = Int("&H" & right$(sTmp, 2))
106     ScrCol(1).value = Int("&H" & Mid$(sTmp, 3, 2))
108     ScrCol(2).value = Int("&H" & left$(sTmp, 2))
        '<EhFooter>
        Exit Sub

OpColor_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OpColor_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub OpColor_MouseMove(Index As Integer, _
                              Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
        '<EhHeader>
        On Error GoTo OpColor_MouseMove_Err
        SuperDebug "sub/fun: OpColor_MouseMove"
        '</EhHeader>
        Dim sTmp As String
100     sTmp = right("000000" & Hex(OpColor(Index).BackColor), 6)
102     LblColor.caption = "Hex:" & sTmp & vbCrLf & " Red:" & Int("&H" & right$(sTmp, 2)) & " - Green:" & Int("&H" & Mid$(sTmp, 3, 2)) & " - Blue:" & Int("&H" & left$(sTmp, 2))
        '<EhFooter>
        Exit Sub

OpColor_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.OpColor_MouseMove " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub pctProperties_Resize()
    '<EhHeader>
    On Error Resume Next
    SuperDebug "sub/fun: pctProperties_Resize"
    '</EhHeader>
    propertyGrid.Height = pctProperties.Height - 10
End Sub

Private Sub PicProperty1_MouseMove(Button As Integer, _
                                   Shift As Integer, _
                                   X As Single, _
                                   Y As Single)
        '<EhHeader>
        On Error GoTo PicProperty1_MouseMove_Err
        SuperDebug "sub/fun: PicProperty1_MouseMove"
        '</EhHeader>
100     LblColor.caption = ""
        '<EhFooter>
        Exit Sub

PicProperty1_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.PicProperty1_MouseMove " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ScrCol_Change(Index As Integer)
        '<EhHeader>
        On Error GoTo ScrCol_Change_Err
        '</EhHeader>
        Dim tColor As Long
SuperDebug "sub/fun: ScrCol_Change"
100     TxtColor(Index).Text = ScrCol(Index).value

102     tColor = RGB(ScrCol(0).value, ScrCol(1).value, ScrCol(2).value)

104     OpColor(ColorIndex).BackColor = tColor
106     bFillColor = OpColor(0).BackColor
108     bBorderColor = OpColor(1).BackColor
110     bBackColor = OpColor(2).BackColor

112     If doNothing = True Then Exit Sub

114     Select Case ColorIndex

            Case 0

116             If OASISDrawObj1.CurrentObject > -1 Then
118                 OASISDrawObj1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
                End If

120         Case 1

122             If OASISDrawObj1.CurrentObject > -1 Then
124                 OASISDrawObj1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor
                End If

126         Case 2
128             OASISDrawObj1.BackColor = bBackColor
        End Select

        '<EhFooter>
        Exit Sub

ScrCol_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.ScrCol_Change " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        '<EhHeader>
        On Error GoTo Slider1_MouseUp_Err
        SuperDebug "sub/fun: Slider1_MouseUp"
        '</EhHeader>

100     If doNothing = True Then Exit Sub
102     OASISDrawObj1.ModifyObject , , , , CSng(Slider1.value)
        '<EhFooter>
        Exit Sub

Slider1_MouseUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.Slider1_MouseUp " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Slider1_Scroll()
        '<EhHeader>
        On Error GoTo Slider1_Scroll_Err
        SuperDebug "sub/fun: Slider1_Scroll"
        '</EhHeader>
100     Label1(0).caption = "Rotation: " & Slider1.value & "°"
        '<EhFooter>
        Exit Sub

Slider1_Scroll_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.Slider1_Scroll " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub SmnuEdit_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo SmnuEdit_Click_Err
        SuperDebug "sub/fun: SmnuEdit_Click"
        '</EhHeader>

100     Select Case Index

            Case 0
102             OASISDrawObj1.Undo

104         Case 1
106             OASISDrawObj1.Redo

108         Case 2

                'Separator
110         Case 3
112             OASISDrawObj1.CopyObject
114             OASISDrawObj1.DeleteObj

116         Case 4
118             OASISDrawObj1.CopyObject

120         Case 5
122             OASISDrawObj1.PasteObject

124         Case 6

                'Separator
126         Case 7
128             OASISDrawObj1.DeleteObj

130         Case 8

                'separator
132         Case 9
134             OASISDrawObj1.SelectAllObjects

136         Case 10

                'separator
138         Case 11
140             OASISDrawObj1.GroupObjects

142         Case 12
144             OASISDrawObj1.UnGroupObjects
        End Select

        '<EhFooter>
        Exit Sub

SmnuEdit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.SmnuEdit_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub SmnuFile_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo SmnuFile_Click_Err
        '</EhHeader>
        Dim c As New cCommonDialog
        
100     Select Case Index

            Case 0
                'OASISDrawObj1.CanvasHeight = 480
                'OASISDrawObj1.CanvasWidth = 640
102             OASISDrawObj1.NewProject
104             Modified = False
                
106         Case 1

108             With c
110                 .DialogTitle = "Open Print Template"
112                 .Filter = "OASIS Print template File (*.ojp)|*.ojp"
114                 .Filename = ""
116                 .ShowOpen
118                 .Filename = Trim(.Filename)

120                 If Len(.Filename) > 0 And FileExist(.Filename) = True Then
122                     OASISDrawObj1.OpenProjects .Filename
124                     bFillColor = OASISDrawObj1.BackColor
                    End If

                End With

126             Modified = False

128         Case 2

130             With c
132                 .DialogTitle = "Open Map Project"
134                 .Filter = "Map Project File (*.ttkgp)|*.ttkgp"
136                 .Filename = ""
138                 .ShowOpen
140                 .Filename = Trim(.Filename)

142                 If Len(.Filename) > 0 And FileExist(.Filename) = True Then
144                     OASISDrawObj1.SetMapProject .Filename
146                     OASISDrawObj1.DrawGIS PicLoad 'bFillColor = OASISDrawObj1.BackColor
                    End If

                End With

                'Modified = False
            
148         Case 3

150             With c
152                 .DialogTitle = "Save Print Template"
154                 .Filter = "OASIS Print template File (*.ojp)|*.ojp"
156                 .Filename = ""
158                 .ShowSave
160                 .Filename = Trim(.Filename)

162                 If Len(.Filename) > 0 Then
164                     OASISDrawObj1.SaveProjects .Filename
                    End If

                End With

166             Modified = False

168         Case 4

170             With c
172                 .DialogTitle = "Export As BitMap"
174                 .Filter = "Bitmap Image File (*.bmp)|*.bmp"
176                 .Filename = ""
178                 .ShowSave
180                 .Filename = Trim(.Filename)

182                 If Len(.Filename) > 0 Then
184                     OASISDrawObj1.Export2BMP .Filename
                    End If

                End With

186         Case 5

                'Separator
188         Case 6
190             OASISDrawObj1.UnSelectAll
            
192             Printer.Orientation = IIf(ComPOrentation.List(ComPOrentation.ListIndex) = "Landscape", vbPRORPortrait, vbPRORLandscape)
          
                Dim oc As cCommonDialog
                Dim sDefPrinter As String
194             Set oc = New cCommonDialog
                
196             Debug.Print Printer.DeviceName
                'debug.Print Printer.
198             Debug.Print oc.PrinterDefault
200             Debug.Print oc.PrinterName
                
                'If oc.ShowPrinter Then
202             oc.ShowPrinter
204             Debug.Print oc.PrinterDefault
206             Debug.Print oc.PrinterName
                ' If oc.flags Then
208             sDefPrinter = GetDefaultPrinter
210             If Not sDefPrinter = oc.PrinterName Then
212                 SetDefaultPrinter oc.PrinterName
                End If
                'Printer.DeviceName = oc.PrinterName
                
214             If ComPOrentation.List(ComPOrentation.ListIndex) = "Landscape" Then
216                 ChngPrinterOrientationLandscape Me
                Else
218                 ChngPrinterOrientationPortrait Me
                End If
                
220             If oc.Copies > 0 Then
222                 Debug.Print Printer.DeviceName
224                 Printer.PaintPicture OASISDrawObj1.Image, 0, 0
226                 Printer.EndDoc
                End If

228             If Not sDefPrinter = oc.PrinterName Then SetDefaultPrinter sDefPrinter

                '  End If
230         Case 7

                'Separator
232         Case 8
234             Unload Me
        End Select

        '<EhFooter>
        Exit Sub

SmnuFile_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMainPrint.SmnuFile_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SmnuOptions_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo SmnuOptions_Click_Err
        SuperDebug "sub/fun: SmnuOptions_Click"
        '</EhHeader>

100     Select Case Index

            Case 0

102             frmCanvasSize.Show vbModal, Me
        End Select

        '<EhFooter>
        Exit Sub

SmnuOptions_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.SmnuOptions_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub SmnuZoom_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo SmnuZoom_Click_Err
        SuperDebug "sub/fun: SmnuZoom_Click"
        '</EhHeader>
        Dim n As Integer

100     Select Case Index

            Case 0
102             OASISDrawObj1.ZoomFactor = 0.1

104         Case 1
106             OASISDrawObj1.ZoomFactor = 0.25

108         Case 2
110             OASISDrawObj1.ZoomFactor = 0.5

112         Case 3
114             OASISDrawObj1.ZoomFactor = 1

116         Case 4
118             OASISDrawObj1.ZoomFactor = 1.5

120         Case 5
122             OASISDrawObj1.ZoomFactor = 2

124         Case 6
126             OASISDrawObj1.ZoomFactor = 4
        End Select

128     For n = 0 To 6
130         SmnuZoom(n).Checked = False
132     Next n

134     SmnuZoom(Index).Checked = True
136     mnuZoom.caption = "Zoom (" & Round(OASISDrawObj1.ZoomFactor * 100) & "%)"
138     StatusBar1.Panels(1).Text = ""
        '<EhFooter>
        Exit Sub

SmnuZoom_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.SmnuZoom_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo Toolbar1_ButtonClick_Err
        SuperDebug "sub/fun: Toolbar1_ButtonClick"
        '</EhHeader>
On Error Resume Next
        Dim c As New cCommonDialog

100     If doNothing = True Then Exit Sub

102     Select Case Button.Index

            Case 1
104             'OASISDrawObj1.CanvasHeight = 480
106             'OASISDrawObj1.CanvasWidth = 640
108             OASISDrawObj1.NewProject
110             Modified = False
112             CreateDefaultTemplate
114         Case 2


346             With c
348                 .DialogTitle = "Open Print Template"
                    '.CancelError = True
350                 .hwnd = Me.hwnd
352                 .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
354                 '.InitDir = g_sAppPath & "\data\user\Maps\"
356                 .Filter = "OASIS Print Template File (*.ojp)|*.ojp)"
358                 .FilterIndex = 1
360                 .ShowOpen
        
                    '.RenewSelection
                End With
                
                'petri: what is this???
362             'If Len(c.FileName) > 0 Then
364                 'InitMap c.FileName
366                 'LoadLayerAttrDataToGridInit
               ' End If
                

116             With c
118                 .DialogTitle = "Open Print Template"
120                 .Filter = "OASIS Print Template File (*.ojp)|*.ojp"
122                 .Filename = ""
124                 .ShowOpen
126                 .Filename = Trim(.Filename)

128                 If Len(.Filename) > 0 And FileExist(.Filename) = True Then
130                     OASISDrawObj1.OpenProjects .Filename
132                     bBackColor = OASISDrawObj1.BackColor
134                     OpColor(2).BackColor = bBackColor
                    End If

                End With

136             Modified = False

138         Case 3

140             With c
142                 .DialogTitle = "Save Print Template"
144                 .Filter = "OASIS Print Template File (*.ojp)|*.ojp"
146                 .Filename = ""
148                 .ShowSave
150                 .Filename = Trim(.Filename)

152                 If Len(.Filename) > 0 Then
154                     OASISDrawObj1.SaveProjects .Filename
                    End If

                End With

156             Modified = False

158         Case 4

160             With c
162                 .DialogTitle = "Export As BitMap"
164                 .Filter = "Bitmap Image File (*.bmp)|*.bmp"
166                 .Filename = ""
168                 .ShowSave
170                 .Filename = Trim(.Filename)

172                 If Len(.Filename) > 0 Then
174                     OASISDrawObj1.Export2BMP .Filename
                    End If

                End With

176         Case 5

                'separator
178         Case 6
180             OASISDrawObj1.CopyObject
182             OASISDrawObj1.DeleteObj

184         Case 7
186             OASISDrawObj1.CopyObject

188         Case 8
190             OASISDrawObj1.PasteObject

192         Case 9

                'separator
194         Case 10
196             OASISDrawObj1.Undo

198         Case 11
200             OASISDrawObj1.Redo

202         Case 12

                'separator
204         Case 13
206             OASISDrawObj1.DeleteObj

208         Case 14

                'separator
210         Case 15
212             mTxtAlign = vbLeftJustify
214             OASISDrawObj1.ModifyObject , , , , , , , , , , , , , , , , , mTxtAlign

216         Case 16
218             mTxtAlign = vbCenter
220             OASISDrawObj1.ModifyObject , , , , , , , , , , , , , , , , , mTxtAlign

222         Case 17
224             mTxtAlign = vbRightJustify
226             OASISDrawObj1.ModifyObject , , , , , , , , , , , , , , , , , mTxtAlign

228         Case 18

                'separator
230         Case 19
232             mBold = Toolbar1.Buttons(19).value
234             OASISDrawObj1.ModifyObject , , , , , , , , , , , , Abs(mBold)

236         Case 20
238             mItalic = Toolbar1.Buttons(20).value
240             OASISDrawObj1.ModifyObject , , , , , , , , , , , , , Abs(mItalic)

242         Case 21
244             mUnderline = Toolbar1.Buttons(21).value
246             OASISDrawObj1.ModifyObject , , , , , , , , , , , , , , Abs(mUnderline)

248         Case 22
250             mStrikethru = Toolbar1.Buttons(22).value
252             OASISDrawObj1.ModifyObject , , , , , , , , , , , , , , , Abs(mStrikethru)

254         Case 23
                'separator
        End Select

256     UpdToolBar
        '<EhFooter>
        Exit Sub

Toolbar1_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.Toolbar1_ButtonClick " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo Toolbar2_ButtonClick_Err
        SuperDebug "sub/fun: Toolbar2_ButtonClick"
        '</EhHeader>

100     Select Case Button.Index

            Case 1
102             OASISDrawObj1.SelectAllObjects

104         Case 2
106             OASISDrawObj1.UnSelectAll

108         Case 3

                'separator
110         Case 4
112             OASISDrawObj1.AlignSelectedObjects mLeft

114         Case 5
116             OASISDrawObj1.AlignSelectedObjects mCenter_V

118         Case 6
120             OASISDrawObj1.AlignSelectedObjects mRight

122         Case 7
124             OASISDrawObj1.AlignSelectedObjects mTop

126         Case 8
128             OASISDrawObj1.AlignSelectedObjects mCenter_H

130         Case 9
132             OASISDrawObj1.AlignSelectedObjects mBottom

134         Case 10
136             OASISDrawObj1.AlignSelectedObjects mCenter_V_H

138         Case 11

                'separator
140         Case 12
142             OASISDrawObj1.SetObjectOrder OASISDrawObj1.CurrentObject, BringToFront

144         Case 13
146             OASISDrawObj1.SetObjectOrder OASISDrawObj1.CurrentObject, SendToBack

148         Case 14
150             OASISDrawObj1.SetObjectOrder OASISDrawObj1.CurrentObject, BringFoward

152         Case 15
154             OASISDrawObj1.SetObjectOrder OASISDrawObj1.CurrentObject, SendBackward

156         Case 16

                'separator
158         Case 17
160             OASISDrawObj1.GroupObjects

162         Case 18
164             OASISDrawObj1.UnGroupObjects
        End Select

166     UpdToolBar
        '<EhFooter>
        Exit Sub

Toolbar2_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.Toolbar2_ButtonClick " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo Toolbar3_ButtonClick_Err
        SuperDebug "sub/fun: Toolbar3_ButtonClick"
        '</EhHeader>

100     Select Case Button.Index

            Case 1
102             OASISDrawObj1.ZoomFactor = 1

104         Case 2
106             OASISDrawObj1.ZoomFactor = OASISDrawObj1.ZoomFactor - 0.1

108         Case 3
110             OASISDrawObj1.ZoomFactor = OASISDrawObj1.ZoomFactor + 0.1
        End Select

112     mnuZoom.caption = "Zoom (" & Round(OASISDrawObj1.ZoomFactor * 100) & "%)"
        '<EhFooter>
        Exit Sub

Toolbar3_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.Toolbar3_ButtonClick " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub Toolbar4_ButtonClick(ByVal Button As MSComctlLib.Button)
        '<EhHeader>
        On Error GoTo Toolbar4_ButtonClick_Err
        SuperDebug "sub/fun: Toolbar4_ButtonClick"
        '</EhHeader>
        Dim tFillColor As Long
        Dim tbSize     As Integer
        Dim c As New cCommonDialog
        
100     tFillColor = bFillColor
102     tbSize = VScroll1.value

104     If tbSize = 0 Then tbSize = 1

106     Select Case Button.Index

            Case 1
108             OASISDrawObj1.UseSelector

110         Case 2
112             OASISDrawObj1.AddObject mline, , , , , , , , bBorderColor, tbSize

114         Case 3
116             OASISDrawObj1.AddObject mArc, , , , , CSng(Slider1.value), , , bBorderColor, tbSize

118         Case 4
120             OASISDrawObj1.AddObject mRectangle, , , , , CSng(Slider1.value), tFillColor, CboFill.ListIndex, bBorderColor, tbSize, PicLoad.Picture

122         Case 5
124             OASISDrawObj1.AddObject mRoundRectangle, , , , , CSng(Slider1.value), tFillColor, CboFill.ListIndex, bBorderColor, tbSize, , , , , , , , , , VScroll3.value

126         Case 6
128             OASISDrawObj1.AddObject mEllipse, , , , , CSng(Slider1.value), tFillColor, CboFill.ListIndex, bBorderColor, tbSize
130             StatusBar1.Panels(1).Text = "Press and Hold ""Ctrl"" Button to make a perfect Circle"

132         Case 7
134             OASISDrawObj1.AddObject mPolygon, , , , , CSng(Slider1.value), tFillColor, CboFill.ListIndex, bBorderColor, VScroll1.value, , , , , , , , , , bPtsQty
136             StatusBar1.Panels(1).Text = "Press and Hold ""Ctrl"" Button to make a perfect Polygon"

138         Case 8
140             OASISDrawObj1.AddObject mStar, , , , , CSng(Slider1.value), tFillColor, CboFill.ListIndex, bBorderColor, VScroll1.value, , , , , , , , , , bPtsQty
142             StatusBar1.Panels(1).Text = "Press and Hold ""Ctrl"" Button to make a perfect Polygon"

144         Case 9
146             OASISDrawObj1.AddObject mText, , , , , CSng(Slider1.value), bFillColor, CboFill.ListIndex, , , , CboFontName.Text, CboFontSize.Text, mBold, mItalic, mUnderline, mStrikethru, , mTxtAlign

148         Case 10

150             With c
152                 .DialogTitle = "Import Image File"
154                 .Filter = "All Picture files|*.jpg;*.bmp;*.gif;*.ico;*.cur;*.dib;*.wmf;*.emf"
156                 .Filename = ""
158                 .ShowOpen

160                 If FileExist(.Filename) = True Then
162                     PicLoad.Picture = LoadPicture(.Filename)
164                     DoEvents
166                     OASISDrawObj1.AddObject mImage, 1, 1, , , , , CboFill.ListIndex, , , PicLoad.Picture
                    End If

                End With

168         Case 11
170             OASISDrawObj1.InitGIS PicLoad
172             PicLoad.Picture = PicLoad.Image
174             DoEvents
176             Debug.Print OASISDrawObj1.CanvasCenterX
178             OASISDrawObj1.AddObject mImage, OASISDrawObj1.CanvasCenterY - ScaleY(PicLoad.Height / 2, vbTwips, vbPixels), OASISDrawObj1.CanvasCenterX - ScaleX(PicLoad.Width / 2, vbTwips, vbPixels), , , , , CboFill.ListIndex, , , PicLoad.Picture, tGISObj:=oMap
180         Case 12
182             OASISDrawObj1.InitLegend PicLoad
184             PicLoad.Picture = PicLoad.Image
186             DoEvents
188             OASISDrawObj1.AddObject mImage, 1, 1, , , , , CboFill.ListIndex, , , PicLoad.Picture, tGISObj:=oLegend
190         Case 13
192             OASISDrawObj1.InitScale PicLoad
194             PicLoad.Picture = PicLoad.Image
196             DoEvents
198             OASISDrawObj1.AddObject mImage, 1, 1, , , , , CboFill.ListIndex, , , PicLoad.Picture, tGISObj:=oScaleBar
200         Case 14
202             OASISDrawObj1.InitNorthArrow PicLoad
204             PicLoad.Picture = PicLoad.Image
206             DoEvents
208             OASISDrawObj1.AddObject mImage, 1, 1, , , , , CboFill.ListIndex, , , PicLoad.Picture, tGISObj:=oNortArrow
        End Select

        '<EhFooter>
        Exit Sub

Toolbar4_ButtonClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.Toolbar4_ButtonClick " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Public Sub CreateDefaultTemplate()
        '<EhHeader>
        On Error GoTo CreateDefaultTemplate_Err
        SuperDebug "sub/fun: CreateDefaultTemplate"
        '</EhHeader>

        Dim sSample As String
        
100     sSample = "Disclaimer" & vbCrLf & "Materials provided on this print are provided “as is”, without warranty of any kind, either express or implied," & vbCrLf & "including, without limitation, warranties of merchantability, fitness for a particular purpose and non-infringement." & vbCrLf & "Mr. specifically does not make any warranties to the accuracy or completeness of any such Materials."
        
''        Debug.Print "CreateDefaultTemplate 1 " & Now
''        DoProgress "Loading Print Utilities...", "Creating map canvas..."
''        'Do the Map Stuff
''102     PicLoad.Move PicLoad.left, PicLoad.top, OASISDrawObj1.CanvasWidth - 100, OASISDrawObj1.CanvasHeight - 100
''        Debug.Print "CreateDefaultTemplate 2 " & Now
''        'Rectangle around map
''        DoProgress "Loading Print Utilities...", "Applying map to canvas..."
''104     OASISDrawObj1.InitGIS PicLoad, OASISDrawObj1.CanvasHeight - 150, OASISDrawObj1.CanvasWidth - 35
''        Debug.Print "CreateDefaultTemplate 3 " & Now
''
''106     PicLoad.Picture = PicLoad.Image
''        Debug.Print "CreateDefaultTemplate 4 " & Now
        
''        DoProgress "Loading Print Utilities...", "Creating map border..."
''1666    OASISDrawObj1.AddObject mRectangle, OASISDrawObj1.CanvasCenterY - ScaleY(PicLoad.Height / 2, vbTwips, vbPixels) - 2, 15, ScaleY(PicLoad.Height, vbTwips, vbPixels) + 10, OASISDrawObj1.CanvasWidth - 30, CSng(Slider1.value), RGB(255, 255, 120), CboFill.ListIndex, RGB(255, 0, 0), 1, PicLoad.Picture
''        DoProgress "Loading Print Utilities...", "Preparing finishing map..."
''110     OASISDrawObj1.AddObject mImage, OASISDrawObj1.CanvasCenterY - ScaleY(PicLoad.Height / 2, vbTwips, vbPixels) + 5, 20, , , , , CboFill.ListIndex, , , PicLoad.Picture, tGISObj:=oMap
''
''        DoProgress "Loading Print Utilities...", "Map finished... Preparing legend border..."
''        'Rectangle around Legend
''112     OASISDrawObj1.AddObject mRectangle, 85, 25, 590, 155, CSng(Slider1.value), RGB(255, 255, 120), CboFill.ListIndex, RGB(255, 0, 0), 1, PicLoad.Picture
''        DoProgress "Loading Print Utilities...", "Created Legend border..."
''
''        Debug.Print "CreateDefaultTemplate 5 " & Now
''
''        DoProgress "Loading Print Utilities...", "Preparing Legend..."
''        'The Legend, No NOT ME!
''114     OASISDrawObj1.InitLegend PicLoad
''116     PicLoad.Picture = PicLoad.Image
''120     OASISDrawObj1.AddObject mImage, 90, 30, 580, 145, , , CboFill.ListIndex, , , PicLoad.Picture, tGISObj:=oLegend
''        DoProgress "Loading Print Utilities...", "Legend finalized..."
''        Debug.Print "CreateDefaultTemplate 6 " & Now
''        DoProgress "Loading Print Utilities...", "Preparing Scalebar..."
''        'Rectangle around scalebar
''122     OASISDrawObj1.AddObject mRectangle, OASISDrawObj1.CanvasHeight - 114, 25, 36, 260, CSng(Slider1.value), RGB(200, 150, 0), CboFill.ListIndex, RGB(0, 0, 0), 1, PicLoad.Picture
''        Debug.Print "CreateDefaultTemplate 7 " & Now
''        'ScaleBar
''124     OASISDrawObj1.InitScale PicLoad
''126     PicLoad.Picture = PicLoad.Image
'''128     DoEvents
''130     OASISDrawObj1.AddObject mImage, OASISDrawObj1.CanvasHeight - 110.5, 30, 30, 250, , , CboFill.ListIndex, , , PicLoad.Picture, tGISObj:=oScaleBar
''
''        DoProgress "Loading Print Utilities...", "Scalebar finished..."
''
''        Debug.Print "CreateDefaultTemplate 8 " & Now
''        'Disclaimer Text Rectangle
''132     OASISDrawObj1.AddObject mRectangle, OASISDrawObj1.CanvasHeight - 60, 15, 51, OASISDrawObj1.CanvasWidth - 30, CSng(Slider1.value), RGB(200, 150, 0), CboFill.ListIndex, RGB(0, 255, 0), 1, PicLoad.Picture
''
''        DoProgress "Loading Print Utilities...", "Preparing disclaimer..."
''
''        Debug.Print "CreateDefaultTemplate 8a " & Now
''        'Disclaimer Text
''134     OASISDrawObj1.AddObject mText, OASISDrawObj1.CanvasHeight - 55, 20, 41, 478, CSng(Slider1.value), RGB(0, 0, 0), CboFill.ListIndex, , , , CboFontName.Text, 6, mBold, mItalic, mUnderline, mStrikethru, sSample, vbLeftJustify
''
''        DoProgress "Loading Print Utilities...", "Disclaimer finished..."
''        Debug.Print "CreateDefaultTemplate 8b " & Now
'''        'North Arrow Rectangle
'''136     OASISDrawObj1.AddObject mRectangle, OASISDrawObj1.CanvasHeight - 57, OASISDrawObj1.CanvasWidth - (OASISDrawObj1.CanvasWidth / 3) - 2, 43, 43, CSng(Slider1.value), RGB(255, 255, 255), CboFill.ListIndex, RGB(0, 0, 0), 1, PicLoad.Picture
'''
'''        DoProgress "Loading Print Utilities...", "Preparing north arrow..."
'''
'''        Debug.Print "CreateDefaultTemplate 9 " & Now
'''        'NorthArrow
'''138     OASISDrawObj1.InitNorthArrow PicLoad
'''140     PicLoad.Picture = PicLoad.Image
''''142     DoEvents
'''144     OASISDrawObj1.AddObject mImage, OASISDrawObj1.CanvasHeight - 55, OASISDrawObj1.CanvasWidth - (OASISDrawObj1.CanvasWidth / 3), 40, 40, , , CboFill.ListIndex, , , PicLoad.Picture, tGISObj:=oNortArrow
'''        Debug.Print "CreateDefaultTemplate 10 " & Now
'''        DoProgress "Loading Print Utilities...", "North arrow finalized..."
'''
'''        'Do iMMAP Logo
''146     PicLoad.Picture = LoadResPicture("LOGO", vbResBitmap)
''        Debug.Print "CreateDefaultTemplate 10a " & Now
''        DoProgress "Loading Print Utilities...", "Preparing Logos..."
''        'Logo Rectangle
''150     OASISDrawObj1.AddObject mRectangle, OASISDrawObj1.CanvasHeight - 56, OASISDrawObj1.CanvasWidth - ScaleX(PicLoad.Width, vbTwips, vbPixels) - 2, 43, 43, CSng(Slider1.value), RGB(255, 255, 255), CboFill.ListIndex, RGB(0, 0, 0), 1, PicLoad.Picture
''        'Draw the Logo rectangle need s to be drawn first otherwise we do not know the position/size of the Logo
''152     OASISDrawObj1.AddObject mImage, OASISDrawObj1.CanvasHeight - 55, OASISDrawObj1.CanvasWidth - ScaleX(PicLoad.Width, vbTwips, vbPixels), 40, 40, , , , , , PicLoad.Picture
''        Debug.Print "CreateDefaultTemplate 10b " & Now
''        DoProgress "Loading Print Utilities...", "1st logo finished..."
''       'Do iMMAP Logo Info Matters
''154     PicLoad.Picture = LoadResPicture("INFO-MATTERS", vbResBitmap)
''        DoProgress "Loading Print Utilities...", "Preparing top Logo..."
''        'Logo Info Matters Rectangle
''        'OASISDrawObj1.AddObject mRectangle, 20, OASISDrawObj1.CanvasWidth - ScaleX(PicLoad.Width, vbTwips, vbMillimeters) - 2, 43, ScaleX(PicLoad.Picture.Width, vbTwips, vbPixels), CSng(Slider1.Value), RGB(255, 255, 255), CboFill.ListIndex, RGB(0, 0, 0), 1, PicLoad.Picture
''        'Draw the Info Matters Logo rectangle needs to be drawn first otherwise we do not know the position/size of the Logo
''158     OASISDrawObj1.AddObject mImage, 15, 15, 40, 160, , , , , , PicLoad.Picture
''        DoProgress "Loading Print Utilities...", "Top logo finished..."
''
'''        Debug.Print "CreateDefaultTemplate 10c " & Now
'''        'Do Map Meta data as text
'''160     OASISDrawObj1.AddObject mText, OASISDrawObj1.CanvasHeight - 55, OASISDrawObj1.CanvasWidth - (OASISDrawObj1.CanvasWidth / 3) - 2 + 46, 41, , , RGB(0, 0, 0), CboFill.ListIndex, , , , CboFontName.Text, 6, mBold, mItalic, mUnderline, mStrikethru, OASISDrawObj1.MapMetaData, vbLeftJustify
'''
'''        DoProgress "Loading Print Utilities...", "Creating meta data text..."
'''        Debug.Print "CreateDefaultTemplate 10c 1 " & Now
'''
''        'Map Title, last object and focus will be set on this
''162     OASISDrawObj1.AddObject mText, 18, 159, 31, 478, CSng(Slider1.value), RGB(0, 0, 0), CboFill.ListIndex, , , , CboFontName.Text, 22, mBold, mItalic, mUnderline, mStrikethru, "OASIS Maps...", vbCenter
        
        DoProgress "Loading Print Utilities...", "Map title added... Refreshing the canvas..."
        Debug.Print "CreateDefaultTemplate 10d " & Now
        'Make sure all objects are redrawn
164     OASISDrawObj1.RefreshCanvas

        DoProgress "Loading Print Utilities...", "Canvas refreshed..."
        Debug.Print "CreateDefaultTemplate 11 " & Now
        '<EhFooter>
        Exit Sub

CreateDefaultTemplate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.CreateDefaultTemplate " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub VScroll1_Change()
        '<EhHeader>
        On Error GoTo VScroll1_Change_Err
        SuperDebug "sub/fun: VScroll1_Change"
        '</EhHeader>

100     If doNothing = True Then Exit Sub
102     TxtBorder.Text = VScroll1.value

104     If OASISDrawObj1.CurrentObject > -1 Then
106         OASISDrawObj1.ModifyObject , , , , , bFillColor, CboFill.ListIndex, bBorderColor, VScroll1.value
        End If

108     OASISDrawObj1.SetFocus
        '<EhFooter>
        Exit Sub

VScroll1_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.VScroll1_Change " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub UpdToolBar()
        '<EhHeader>
        On Error GoTo UpdToolBar_Err
        SuperDebug "sub/fun: UpdToolBar"
        '</EhHeader>
'Exit Sub
100     If OASISDrawObj1.CurrentObject > -1 Then
102         Toolbar1.Buttons(6).Enabled = True
104         Toolbar1.Buttons(7).Enabled = True
106         Toolbar1.Buttons(13).Enabled = True

108         Toolbar2.Buttons(12).Enabled = True
110         Toolbar2.Buttons(13).Enabled = True
112         Toolbar2.Buttons(14).Enabled = True
114         Toolbar2.Buttons(15).Enabled = True
116         Toolbar2.Buttons(17).Enabled = True
118         Toolbar2.Buttons(18).Enabled = True
        Else
120         Toolbar1.Buttons(6).Enabled = False
122         Toolbar1.Buttons(7).Enabled = False
124         Toolbar1.Buttons(13).Enabled = False

126         Toolbar2.Buttons(12).Enabled = False
128         Toolbar2.Buttons(13).Enabled = False
130         Toolbar2.Buttons(14).Enabled = False
132         Toolbar2.Buttons(15).Enabled = False
134         Toolbar2.Buttons(17).Enabled = False
136         Toolbar2.Buttons(18).Enabled = False
        End If

138     If OASISDrawObj1.ObjectType = mText Then
140         Toolbar1.Buttons(15).Enabled = True
142         Toolbar1.Buttons(16).Enabled = True
144         Toolbar1.Buttons(17).Enabled = True
146         Toolbar1.Buttons(19).Enabled = True
148         Toolbar1.Buttons(20).Enabled = True
150         Toolbar1.Buttons(21).Enabled = True
152         Toolbar1.Buttons(22).Enabled = True
154         CboFontName.Enabled = True
156         CboFontSize.Enabled = True
        Else
158         Toolbar1.Buttons(15).Enabled = False
160         Toolbar1.Buttons(16).Enabled = False
162         Toolbar1.Buttons(17).Enabled = False
164         Toolbar1.Buttons(19).Enabled = False
166         Toolbar1.Buttons(20).Enabled = False
168         Toolbar1.Buttons(21).Enabled = False
170         Toolbar1.Buttons(22).Enabled = False
172         CboFontName.Enabled = False
174         CboFontSize.Enabled = False
        End If

176     Toolbar1.Buttons(8).Enabled = OASISDrawObj1.ObjectInClipboard
178     Toolbar2.Buttons(1).Enabled = OASISDrawObj1.ObjectQty
180     Toolbar2.Buttons(2).Enabled = OASISDrawObj1.SelectionQty

182     If OASISDrawObj1.SelectionQty > 1 Then
184         Toolbar2.Buttons(4).Enabled = True
186         Toolbar2.Buttons(5).Enabled = True
188         Toolbar2.Buttons(6).Enabled = True
190         Toolbar2.Buttons(7).Enabled = True
192         Toolbar2.Buttons(8).Enabled = True
194         Toolbar2.Buttons(9).Enabled = True
196         Toolbar2.Buttons(10).Enabled = True
        Else
198         Toolbar2.Buttons(4).Enabled = False
200         Toolbar2.Buttons(5).Enabled = False
202         Toolbar2.Buttons(6).Enabled = False
204         Toolbar2.Buttons(7).Enabled = False
206         Toolbar2.Buttons(8).Enabled = False
208         Toolbar2.Buttons(9).Enabled = False
210         Toolbar2.Buttons(10).Enabled = False
        End If

        '<EhFooter>
        Exit Sub

UpdToolBar_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.UpdToolBar " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub VScroll2_Change()
        '<EhHeader>
        On Error GoTo VScroll2_Change_Err
        SuperDebug "sub/fun: VScroll2_Change"
        '</EhHeader>

100     If doNothing = True Then Exit Sub
102     TxtPoint.Text = VScroll2.value
104     bPtsQty = VScroll2.value

106     If OASISDrawObj1.CurrentObject > -1 And OASISDrawObj1.ObjectType = mPolygon Or OASISDrawObj1.ObjectType = mStar Then
108         OASISDrawObj1.ModifyObject , , , , , , , , , , , , , , , , , , VScroll2.value
        End If

110     OASISDrawObj1.SetFocus
        '<EhFooter>
        Exit Sub

VScroll2_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.VScroll2_Change " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub VScroll3_Change()
        '<EhHeader>
        On Error GoTo VScroll3_Change_Err
        SuperDebug "sub/fun: VScroll3_Change"
        '</EhHeader>

100     If doNothing = True Then Exit Sub
102     TxtRound.Text = VScroll3.value

104     If OASISDrawObj1.CurrentObject > -1 And OASISDrawObj1.ObjectType = mRoundRectangle Then
106         OASISDrawObj1.ModifyObject , , , , , , , , , , , , , , , , , , VScroll3.value
        End If

108     OASISDrawObj1.SetFocus
        '<EhFooter>
        Exit Sub

VScroll3_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.VScroll3_Change " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

'****************************************************************************************
' Proprty Grid Code
'

Private Sub configureCategories()

End Sub

Private Sub configureCombo()
        '<EhHeader>
        SuperDebug "sub/fun: configureCombo"
        On Error GoTo configureCombo_Err
        '</EhHeader>
    Dim i As Long
   
100    With cboIcon
102       .ImageList = ilsIcons
104       For i = 1 To ilsIcons.ImageCount
106          .AddItemAndData ilsIcons.ItemKey(i), i - 1, i - 1
108       Next i
110       .DropDownWidth = 96
       End With
   
        '<EhFooter>
        Exit Sub

configureCombo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.configureCombo " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub configureGrid()
        '<EhHeader>
        On Error GoTo configureGrid_Err
        SuperDebug "sub/fun: configureGrid"
        '</EhHeader>
   
100    With propertyGrid
      
102       .DefaultRowHeight = 330 \ Screen.TwipsPerPixelY
104       .Editable = True
106       .ImageList = ilsIcons
108       .HighlightSelectedIcons = False
110       .AddColumn "hProperty", "Property", , , , , , , , , , CCLSortStringNoCase
112       .AddColumn "hValue", "Value", , , ScaleY(.Width - 23, vbTwips, vbPixels) - .ColumnWidth("hProperty")
            
       End With
   
        '<EhFooter>
        Exit Sub

configureGrid_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.configureGrid " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Function AddProperty( _
    sPropName As String, vPropValue As Variant) As Long
        '<EhHeader>
        On Error GoTo AddProperty_Err
        SuperDebug "sub/fun: AddProperty"
        '</EhHeader>
    
    Dim lRow As Long
   
100    With propertyGrid
102       .AddRow
104       lRow = .Rows

106       .CellText(lRow, 1) = sPropName
108       .CellTextAlign(lRow, 1) = DT_SINGLELINE Or DT_VCENTER Or DT_END_ELLIPSIS
110       .CellText(lRow, 2) = vPropValue
112       .CellTextAlign(lRow, 2) = DT_SINGLELINE Or DT_VCENTER Or DT_END_ELLIPSIS
114       lRow = .ShiftLastRowToSortLocation()
116       AddProperty = lRow
      
       End With
   
        '<EhFooter>
        Exit Function

AddProperty_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.AddProperty " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Function

Private Sub addSampleData()
        '<EhHeader>
        On Error GoTo addSampleData_Err
        SuperDebug "sub/fun: addSampleData"
        '</EhHeader>
'100     AddProperty "Dude", "sdfhushdf"
'102     AddProperty "Dudes", "sdfhushdf"
'104     AddProperty "Dude444", "sdfhushdf"
'106     AddProperty "Dudesjfh", "sdfhushdf"
'108     AddProperty "Dudedhvd", "sdfhushdf"
'110     AddProperty "Dudehdbvh", "sdfhushdf"
   
        Exit Sub
   
112     AddProperty "Dudejdcc", "sdfhushdf"
114     AddProperty "Dudesjcjbc", "sdfhushdf"
116     AddProperty "Dudejbbc", "sdfhushdf"
118     AddProperty "Dudedjvbdbvj", "sdfhushdf"
120     AddProperty "Dudejbvvb", "sdfhushdf"
122     AddProperty "Dudexjbbv", "sdfhushdf"
124     AddProperty "Dudexnbvbm", "sdfhushdf"
   
        Exit Sub
   
        '<EhFooter>
        Exit Sub

addSampleData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.addSampleData " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Function validEditDate(ByVal bHideOnly As Boolean, ByRef vValue As Variant) As Boolean
        '<EhHeader>
        On Error GoTo validEditDate_Err
        SuperDebug "sub/fun: validEditDate"
        '</EhHeader>
    Dim sText As String
    Dim bR As Boolean
100    sText = Trim(txtEdit.Text)
102    If (Len(sText) = 0) Then
104       vValue = Empty
106       bR = True
       Else
108       If (IsDate(sText)) Then
110          vValue = CDate(sText)
112          bR = True
          End If
       End If
114    If Not (bR) Then
116       If Not (bHideOnly) Then
118          If Not (tipPopup1.Showing) Then
120             tipPopup1.Title = "Invalid Date Format"
122             tipPopup1.Text = "Enter a valid date (e.g." & Format(Now, "short date") & "), or blank the text in the cell to remove the date."
124             tipPopup1.Show Me.hwnd, txtEdit.left \ Screen.TwipsPerPixelY, (txtEdit.top + txtEdit.Height) \ Screen.TwipsPerPixelY - 4
             End If
          End If
       Else
126       If (tipPopup1.Showing) Then
128          tipPopup1.Hide
          End If
       End If
130    validEditDate = bR
        '<EhFooter>
        Exit Function

validEditDate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.validEditDate " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Function

Private Sub cboIcon_Click()
        '<EhHeader>
        On Error GoTo cboIcon_Click_Err
        SuperDebug "sub/fun: cboIcon_Click"
        '</EhHeader>
100    If (cboIcon.Tag = "DROPPED") Then
102       Debug.Print "Click whilst dropped"
       Else
104       Debug.Print "Click whilst not dropped"
       End If
        '<EhFooter>
        Exit Sub

cboIcon_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.cboIcon_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub cboIcon_CloseUp()
        '<EhHeader>
        On Error GoTo cboIcon_CloseUp_Err
        SuperDebug "sub/fun: cboIcon_CloseUp"
        '</EhHeader>
100    cboIcon.Tag = ""
        '<EhFooter>
        Exit Sub

cboIcon_CloseUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.cboIcon_CloseUp " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub cboIcon_DropDown()
        '<EhHeader>
        On Error GoTo cboIcon_DropDown_Err
        SuperDebug "sub/fun: cboIcon_DropDown"
        '</EhHeader>
100    cboIcon.Tag = "DROPPED"
        '<EhFooter>
        Exit Sub

cboIcon_DropDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.cboIcon_DropDown " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub cboIcon_KeyDown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo cboIcon_KeyDown_Err
        SuperDebug "sub/fun: cboIcon_KeyDown"
        '</EhHeader>
100    Select Case KeyCode
       Case 9   ' tab
102       propertyGrid.EndEdit
104       KeyCode = 0
106    Case 13  ' return
108       propertyGrid.EndEdit
110       KeyCode = 0
112    Case 27  ' escape
114       propertyGrid.CancelEdit
116       KeyCode = 0
       End Select
        '<EhFooter>
        Exit Sub

cboIcon_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.cboIcon_KeyDown " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub cboIcon_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo cboIcon_KeyPress_Err
        SuperDebug "sub/fun: cboIcon_KeyPress"
        '</EhHeader>
100    Select Case KeyAscii
       Case 9
102       KeyAscii = 0
104    Case 13
106       KeyAscii = 0
108    Case 27
110       KeyAscii = 0
       End Select
        '<EhFooter>
        Exit Sub

cboIcon_KeyPress_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.cboIcon_KeyPress " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub propertyGrid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
        '<EhHeader>
        On Error GoTo propertyGrid_KeyDown_Err
        SuperDebug "sub/fun: propertyGrid_KeyDown"
        '</EhHeader>
100    If (KeyCode = vbKeyDelete) Then
          Dim lCol As Long
          Dim lRow As Long
102       lCol = propertyGrid.SelectedCol
104       lRow = propertyGrid.SelectedRow
106       If (lCol > 0) And (lRow > 0) Then
108          Select Case propertyGrid.ColumnKey(lCol)
             Case "Icon"
110             Beep
112          Case "DisplayName"
114             propertyGrid.CellText(lRow, lCol) = Empty
             End Select
          End If
       End If
        '<EhFooter>
        Exit Sub

propertyGrid_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.propertyGrid_KeyDown " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub selCategories_KeyDown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo selCategories_KeyDown_Err
        SuperDebug "sub/fun: selCategories_KeyDown"
        '</EhHeader>
100    Select Case KeyCode
       Case 9   ' tab
102       propertyGrid.EndEdit
104       KeyCode = 0
106    Case 13  ' return
108       propertyGrid.EndEdit
110       KeyCode = 0
112    Case 27  ' escape
114       propertyGrid.CancelEdit
116       KeyCode = 0
       End Select
        '<EhFooter>
        Exit Sub

selCategories_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.selCategories_KeyDown " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub propertyGrid_CancelEdit()
       '
        '<EhHeader>
        On Error GoTo propertyGrid_CancelEdit_Err
        SuperDebug "sub/fun: propertyGrid_CancelEdit"
        '</EhHeader>
100  '  selCategories.EndEdit
102    selCategories.Visible = False
104    cboIcon.Visible = False
106    txtEdit.Visible = False
108    tipPopup1.Hide
       '
        '<EhFooter>
        Exit Sub

propertyGrid_CancelEdit_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.propertyGrid_CancelEdit " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub propertyGrid_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, NewValue As Variant, bStayInEditMode As Boolean)
        '<EhHeader>
        On Error GoTo propertyGrid_PreCancelEdit_Err
        SuperDebug "sub/fun: propertyGrid_PreCancelEdit"
        '</EhHeader>
    Dim sText As String
       '
100    Select Case propertyGrid.ColumnKey(lCol)
       Case "Icon"
102       propertyGrid.CellIcon(lRow, lCol) = cboIcon.ItemIcon(cboIcon.ListIndex)
104    Case "DisplayName"
106       propertyGrid.CellText(lRow, lCol) = txtEdit.Text
       End Select
       '
        '<EhFooter>
        Exit Sub

propertyGrid_PreCancelEdit_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.propertyGrid_PreCancelEdit " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub selCategories_RequestDropDownInstance(ctl As ddnMultiSelect)
        '<EhHeader>
        On Error GoTo selCategories_RequestDropDownInstance_Err
        SuperDebug "sub/fun: selCategories_RequestDropDownInstance"
        '</EhHeader>
100    Set ctl = ddnCategories
        '<EhFooter>
        Exit Sub

selCategories_RequestDropDownInstance_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.selCategories_RequestDropDownInstance " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub propertyGrid_RequestEdit(ByVal lRow As Long, _
                                     ByVal lCol As Long, _
                                     ByVal iKeyAscii As Integer, _
                                     bCancel As Boolean)
        '<EhHeader>
        On Error GoTo propertyGrid_RequestEdit_Err
        SuperDebug "sub/fun: propertyGrid_RequestEdit"
        '</EhHeader>
        Dim lLeft   As Long
        Dim lTop    As Long
        Dim lWidth  As Long
        Dim lHeight As Long

100     If (lRow > 0) And (lCol > 1) Then
102         propertyGrid.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
104         lLeft = lLeft + propertyGrid.left
106         lTop = lTop + propertyGrid.top + Screen.TwipsPerPixelY
      
            'If (lWidth < 32 * Screen.TwipsPerPixelX) Then
            '   lWidth = 32 * Screen.TwipsPerPixelX
            'End If
      
            Exit Sub
108         Select Case lCol

                Case 1
                    ' Icon
110                 cboIcon.ListIndex = propertyGrid.CellIcon(lRow, lCol)
112                 cboIcon.Move lLeft - 16 * Screen.TwipsPerPixelX, lTop, lWidth + 16 * Screen.TwipsPerPixelX, lHeight
114                 cboIcon.Visible = True
116                 cboIcon.SetFocus

118             Case 7
                    ' Categories
120                 'selCategories.Selection = propertyGrid.CellText(lRow, lCol)
122                 selCategories.Move lLeft, lTop, lWidth, lHeight
124                 selCategories.Visible = True
126                 selCategories.SetFocus

128             Case Else
130                 txtEdit.Text = propertyGrid.CellText(lRow, lCol)
132                 txtEdit.Move lLeft, lTop, lWidth, lHeight
134                 txtEdit.Visible = True
136                 txtEdit.SetFocus
         
            End Select

        End If

        '<EhFooter>
        Exit Sub

propertyGrid_RequestEdit_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.propertyGrid_RequestEdit " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtEdit_Change()
        '<EhHeader>
        On Error GoTo txtEdit_Change_Err
        SuperDebug "sub/fun: txtEdit_Change"
        '</EhHeader>
    Dim lCol As Long
100    If (propertyGrid.InEditMode) Then
102       lCol = propertyGrid.EditCol
104       Select Case propertyGrid.ColumnKey(lCol)
          Case "DisplayName"
             Dim vJunk As Variant
116          validEditDate True, vJunk
          End Select
       End If
        '<EhFooter>
        Exit Sub

txtEdit_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.txtEdit_Change " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
        '<EhHeader>
        On Error GoTo txtEdit_KeyDown_Err
        SuperDebug "sub/fun: txtEdit_KeyDown"
        '</EhHeader>
100    Select Case KeyCode
       Case 9   ' tab
102       propertyGrid.EndEdit
104       KeyCode = 0
106    Case 13  ' return
108       propertyGrid.EndEdit
110       KeyCode = 0
112    Case 27  ' escape
114       propertyGrid.CancelEdit
116       KeyCode = 0
       End Select
        '<EhFooter>
        Exit Sub

txtEdit_KeyDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.txtEdit_KeyDown " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo txtEdit_KeyPress_Err
        SuperDebug "sub/fun: txtEdit_KeyPress"
        '</EhHeader>
100    Select Case KeyAscii
       Case 9
102       KeyAscii = 0
104    Case 13
106       KeyAscii = 0
108    Case 27
110       KeyAscii = 0
       End Select
        '<EhFooter>
        Exit Sub

txtEdit_KeyPress_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISPrint.frmMainPrint.txtEdit_KeyPress " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

