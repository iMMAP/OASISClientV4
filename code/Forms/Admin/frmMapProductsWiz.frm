VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.Form frmMapProductsWiz 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Map Products Wizard"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frmMapProductsWiz.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   7050
      TabIndex        =   104
      Top             =   4620
      Width           =   675
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   375
      Left            =   6360
      TabIndex        =   103
      Top             =   4620
      Width           =   675
   End
   Begin C1SizerLibCtl.C1Tab C1TTab1Tab2 
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7845
      _cx             =   13838
      _cy             =   8440
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
      Caption         =   "Tab&1|New Tab|New Tab|New Tab"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   10
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   1
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin C1SizerLibCtl.C1Elastic elNav 
         Height          =   4695
         Index           =   0
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   7755
         _cx             =   13679
         _cy             =   8281
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
         Begin VB.ComboBox ComMapNum 
            Height          =   315
            ItemData        =   "frmMapProductsWiz.frx":6852
            Left            =   1380
            List            =   "frmMapProductsWiz.frx":6854
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   180
            Width           =   735
         End
         Begin VB.Label lblNumberOf 
            Caption         =   "Number of Maps in Map library:"
            Height          =   435
            Left            =   120
            TabIndex        =   4
            Top             =   120
            Width           =   1185
         End
      End
      Begin C1SizerLibCtl.C1Elastic elNav 
         Height          =   4695
         Index           =   1
         Left            =   8490
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   45
         Width           =   7755
         _cx             =   13679
         _cy             =   8281
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
         Begin VB.Frame FraMap 
            Caption         =   "Map 4:"
            Height          =   1050
            Index           =   3
            Left            =   90
            TabIndex        =   21
            Top             =   3210
            Visible         =   0   'False
            Width           =   6375
            Begin VB.TextBox txtMapAlias 
               DataField       =   "Alias"
               Height          =   345
               Index           =   3
               Left            =   660
               TabIndex        =   23
               Top             =   570
               Width           =   5595
            End
            Begin VB.TextBox txtMapName 
               DataField       =   "Name"
               Height          =   345
               Index           =   3
               Left            =   660
               TabIndex        =   22
               Top             =   180
               Width           =   5595
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Alias:"
               Height          =   195
               Index           =   6
               Left            =   180
               TabIndex        =   25
               Top             =   570
               Width           =   375
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Name:"
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   24
               Top             =   210
               Width           =   465
            End
         End
         Begin VB.Frame FraMap 
            Caption         =   "Map 3:"
            Height          =   1005
            Index           =   2
            Left            =   90
            TabIndex        =   16
            Top             =   2100
            Visible         =   0   'False
            Width           =   6345
            Begin VB.TextBox txtMapName 
               DataField       =   "Name"
               Height          =   345
               Index           =   2
               Left            =   660
               TabIndex        =   18
               Top             =   180
               Width           =   5595
            End
            Begin VB.TextBox txtMapAlias 
               DataField       =   "Alias"
               Height          =   345
               Index           =   2
               Left            =   660
               TabIndex        =   17
               Top             =   570
               Width           =   5595
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Name:"
               Height          =   195
               Index           =   5
               Left            =   120
               TabIndex        =   20
               Top             =   210
               Width           =   465
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Alias:"
               Height          =   195
               Index           =   4
               Left            =   180
               TabIndex        =   19
               Top             =   570
               Width           =   375
            End
         End
         Begin VB.Frame FraMap 
            Caption         =   "Map 2:"
            Height          =   1005
            Index           =   1
            Left            =   90
            TabIndex        =   11
            Top             =   1050
            Visible         =   0   'False
            Width           =   6345
            Begin VB.TextBox txtMapName 
               DataField       =   "Name"
               Height          =   345
               Index           =   1
               Left            =   660
               TabIndex        =   13
               Top             =   180
               Width           =   5595
            End
            Begin VB.TextBox txtMapAlias 
               DataField       =   "Alias"
               Height          =   345
               Index           =   1
               Left            =   660
               TabIndex        =   12
               Top             =   570
               Width           =   5595
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Name:"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   15
               Top             =   210
               Width           =   465
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Alias:"
               Height          =   195
               Index           =   2
               Left            =   180
               TabIndex        =   14
               Top             =   570
               Width           =   375
            End
         End
         Begin VB.Frame FraMap 
            Caption         =   "Map 1:"
            Height          =   1005
            Index           =   0
            Left            =   90
            TabIndex        =   6
            Top             =   30
            Width           =   6345
            Begin VB.TextBox txtMapAlias 
               DataField       =   "Alias"
               Height          =   345
               Index           =   0
               Left            =   660
               TabIndex        =   10
               Top             =   570
               Width           =   5595
            End
            Begin VB.TextBox txtMapName 
               DataField       =   "Name"
               Height          =   345
               Index           =   0
               Left            =   660
               TabIndex        =   9
               Top             =   180
               Width           =   5595
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Alias:"
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   8
               Top             =   570
               Width           =   375
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Name:"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   7
               Top             =   210
               Width           =   465
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic elNav 
         Height          =   4695
         Index           =   2
         Left            =   8790
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   7755
         _cx             =   13679
         _cy             =   8281
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
         Begin VB.Frame FraFileAttr 
            Caption         =   "Map project:"
            Height          =   1050
            Index           =   3
            Left            =   90
            TabIndex        =   44
            Top             =   3300
            Visible         =   0   'False
            Width           =   7605
            Begin VB.TextBox txtMapPrjFile 
               DataField       =   "FileName"
               Height          =   315
               Index           =   3
               Left            =   1830
               TabIndex        =   47
               Top             =   150
               Width           =   5175
            End
            Begin VB.CommandButton cmdOpenMapFile 
               Caption         =   "..."
               Height          =   255
               Index           =   3
               Left            =   7080
               TabIndex        =   46
               Top             =   180
               Width           =   435
            End
            Begin VB.TextBox txtMapDesc 
               DataField       =   "Description"
               Height          =   495
               Index           =   3
               Left            =   960
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   45
               Top             =   480
               Width           =   6045
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Map Project File Name:"
               Height          =   195
               Index           =   15
               Left            =   60
               TabIndex        =   49
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Description:"
               Height          =   195
               Index           =   14
               Left            =   60
               TabIndex        =   48
               Top             =   570
               Width           =   840
            End
         End
         Begin VB.Frame FraFileAttr 
            Caption         =   "Map project:"
            Height          =   1050
            Index           =   2
            Left            =   90
            TabIndex        =   38
            Top             =   2220
            Visible         =   0   'False
            Width           =   7605
            Begin VB.TextBox txtMapPrjFile 
               DataField       =   "FileName"
               Height          =   315
               Index           =   2
               Left            =   1830
               TabIndex        =   41
               Top             =   150
               Width           =   5175
            End
            Begin VB.CommandButton cmdOpenMapFile 
               Caption         =   "..."
               Height          =   255
               Index           =   2
               Left            =   7080
               TabIndex        =   40
               Top             =   180
               Width           =   435
            End
            Begin VB.TextBox txtMapDesc 
               DataField       =   "Description"
               Height          =   495
               Index           =   2
               Left            =   960
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   39
               Top             =   480
               Width           =   6045
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Map Project File Name:"
               Height          =   195
               Index           =   13
               Left            =   60
               TabIndex        =   43
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Description:"
               Height          =   195
               Index           =   12
               Left            =   60
               TabIndex        =   42
               Top             =   570
               Width           =   840
            End
         End
         Begin VB.Frame FraFileAttr 
            Caption         =   "Map project:"
            Height          =   1050
            Index           =   1
            Left            =   90
            TabIndex        =   32
            Top             =   1110
            Visible         =   0   'False
            Width           =   7605
            Begin VB.TextBox txtMapPrjFile 
               DataField       =   "FileName"
               Height          =   315
               Index           =   1
               Left            =   1830
               TabIndex        =   35
               Top             =   150
               Width           =   5175
            End
            Begin VB.CommandButton cmdOpenMapFile 
               Caption         =   "..."
               Height          =   255
               Index           =   1
               Left            =   7080
               TabIndex        =   34
               Top             =   180
               Width           =   435
            End
            Begin VB.TextBox txtMapDesc 
               DataField       =   "Description"
               Height          =   495
               Index           =   1
               Left            =   960
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   33
               Top             =   480
               Width           =   6045
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Map Project File Name:"
               Height          =   195
               Index           =   11
               Left            =   60
               TabIndex        =   37
               Top             =   240
               Width           =   1650
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Description:"
               Height          =   195
               Index           =   10
               Left            =   60
               TabIndex        =   36
               Top             =   570
               Width           =   840
            End
         End
         Begin VB.Frame FraFileAttr 
            Caption         =   "Map project:"
            Height          =   1050
            Index           =   0
            Left            =   90
            TabIndex        =   26
            Top             =   30
            Width           =   7605
            Begin VB.TextBox txtMapDesc 
               DataField       =   "Description"
               Height          =   495
               Index           =   0
               Left            =   960
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   31
               Top             =   480
               Width           =   6045
            End
            Begin VB.CommandButton cmdOpenMapFile 
               Caption         =   "..."
               Height          =   255
               Index           =   0
               Left            =   7080
               TabIndex        =   29
               Top             =   180
               Width           =   435
            End
            Begin VB.TextBox txtMapPrjFile 
               DataField       =   "FileName"
               Height          =   315
               Index           =   0
               Left            =   1830
               TabIndex        =   28
               Top             =   150
               Width           =   5175
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Description:"
               Height          =   195
               Index           =   9
               Left            =   60
               TabIndex        =   30
               Top             =   570
               Width           =   840
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Map Project File Name:"
               Height          =   195
               Index           =   8
               Left            =   60
               TabIndex        =   27
               Top             =   240
               Width           =   1650
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic elNav 
         Height          =   4695
         Index           =   3
         Left            =   9090
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   45
         Width           =   7755
         _cx             =   13679
         _cy             =   8281
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
         Begin VB.Frame FraDetails 
            Caption         =   "Details 1:"
            Height          =   1005
            Index           =   0
            Left            =   90
            TabIndex        =   54
            Top             =   30
            Width           =   7605
            Begin VB.TextBox txtMappreview 
               DataField       =   "Image"
               Height          =   285
               Index           =   0
               Left            =   4170
               TabIndex        =   66
               Top             =   660
               Width           =   3375
            End
            Begin VB.TextBox txtMapthumb 
               DataField       =   "ThumbNail"
               Height          =   285
               Index           =   0
               Left            =   4170
               TabIndex        =   64
               Top             =   390
               Width           =   3375
            End
            Begin VB.TextBox txtMapContact 
               DataField       =   "Contact"
               Height          =   285
               Index           =   0
               Left            =   4170
               TabIndex        =   62
               Top             =   120
               Width           =   3375
            End
            Begin VB.TextBox txtMapCopyright 
               DataField       =   "Copyright"
               Height          =   285
               Index           =   0
               Left            =   960
               TabIndex        =   60
               Top             =   660
               Width           =   2115
            End
            Begin VB.TextBox txtMapCreatedBy 
               DataField       =   "CreatedBy"
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   56
               Top             =   150
               Width           =   2115
            End
            Begin VB.TextBox txtMapDate 
               DataField       =   "CreatedDate"
               Height          =   285
               Index           =   0
               Left            =   960
               TabIndex        =   55
               Top             =   390
               Width           =   2115
            End
            Begin VB.Label lblMappreview 
               AutoSize        =   -1  'True
               Caption         =   "Preview:"
               Height          =   195
               Index           =   27
               Left            =   3360
               TabIndex        =   65
               Top             =   720
               Width           =   615
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Thumb nail:"
               Height          =   195
               Index           =   26
               Left            =   3150
               TabIndex        =   63
               Top             =   450
               Width           =   825
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Contact:"
               Height          =   195
               Index           =   25
               Left            =   3330
               TabIndex        =   61
               Top             =   180
               Width           =   600
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Copyright:"
               Height          =   195
               Index           =   24
               Left            =   150
               TabIndex        =   59
               Top             =   720
               Width           =   705
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Created By:"
               Height          =   195
               Index           =   23
               Left            =   120
               TabIndex        =   58
               Top             =   210
               Width           =   825
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Date:"
               Height          =   195
               Index           =   22
               Left            =   480
               TabIndex        =   57
               Top             =   450
               Width           =   390
            End
         End
         Begin VB.Frame FraDetails 
            Caption         =   "Details 2:"
            Height          =   1005
            Index           =   1
            Left            =   90
            TabIndex        =   53
            Top             =   1050
            Width           =   7605
            Begin VB.TextBox txtMappreview 
               DataField       =   "Image"
               Height          =   285
               Index           =   1
               Left            =   4170
               TabIndex        =   78
               Top             =   690
               Width           =   3375
            End
            Begin VB.TextBox txtMapthumb 
               DataField       =   "ThumbNail"
               Height          =   285
               Index           =   1
               Left            =   4170
               TabIndex        =   76
               Top             =   420
               Width           =   3375
            End
            Begin VB.TextBox txtMapContact 
               DataField       =   "Contact"
               Height          =   285
               Index           =   1
               Left            =   4170
               TabIndex        =   74
               Top             =   150
               Width           =   3375
            End
            Begin VB.TextBox txtMapCopyright 
               DataField       =   "Copyright"
               Height          =   285
               Index           =   1
               Left            =   960
               TabIndex        =   72
               Top             =   690
               Width           =   2115
            End
            Begin VB.TextBox txtMapCreatedBy 
               DataField       =   "CreatedBy"
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   68
               Top             =   180
               Width           =   2115
            End
            Begin VB.TextBox txtMapDate 
               DataField       =   "CreatedDate"
               Height          =   285
               Index           =   1
               Left            =   960
               TabIndex        =   67
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label lblMappreview 
               AutoSize        =   -1  'True
               Caption         =   "Preview:"
               Height          =   195
               Index           =   0
               Left            =   3360
               TabIndex        =   77
               Top             =   750
               Width           =   615
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Thumb nail:"
               Height          =   195
               Index           =   20
               Left            =   3150
               TabIndex        =   75
               Top             =   480
               Width           =   825
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Contact:"
               Height          =   195
               Index           =   19
               Left            =   3330
               TabIndex        =   73
               Top             =   210
               Width           =   600
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Copyright:"
               Height          =   195
               Index           =   18
               Left            =   150
               TabIndex        =   71
               Top             =   750
               Width           =   705
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Created By:"
               Height          =   195
               Index           =   17
               Left            =   120
               TabIndex        =   70
               Top             =   240
               Width           =   825
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Date:"
               Height          =   195
               Index           =   16
               Left            =   480
               TabIndex        =   69
               Top             =   480
               Width           =   390
            End
         End
         Begin VB.Frame FraDetails 
            Caption         =   "Details 3:"
            Height          =   1005
            Index           =   2
            Left            =   90
            TabIndex        =   52
            Top             =   2100
            Width           =   7605
            Begin VB.TextBox txtMappreview 
               DataField       =   "Image"
               Height          =   285
               Index           =   2
               Left            =   4140
               TabIndex        =   90
               Top             =   690
               Width           =   3405
            End
            Begin VB.TextBox txtMapthumb 
               DataField       =   "ThumbNail"
               Height          =   285
               Index           =   2
               Left            =   4140
               TabIndex        =   88
               Top             =   420
               Width           =   3405
            End
            Begin VB.TextBox txtMapContact 
               DataField       =   "Contact"
               Height          =   285
               Index           =   2
               Left            =   4140
               TabIndex        =   86
               Top             =   150
               Width           =   3405
            End
            Begin VB.TextBox txtMapCopyright 
               DataField       =   "Copyright"
               Height          =   285
               Index           =   2
               Left            =   930
               TabIndex        =   84
               Top             =   690
               Width           =   2115
            End
            Begin VB.TextBox txtMapCreatedBy 
               DataField       =   "CreatedBy"
               Height          =   255
               Index           =   2
               Left            =   930
               TabIndex        =   80
               Top             =   180
               Width           =   2115
            End
            Begin VB.TextBox txtMapDate 
               DataField       =   "CreatedDate"
               Height          =   285
               Index           =   2
               Left            =   930
               TabIndex        =   79
               Top             =   420
               Width           =   2115
            End
            Begin VB.Label lblMappreview 
               AutoSize        =   -1  'True
               Caption         =   "Preview:"
               Height          =   195
               Index           =   1
               Left            =   3330
               TabIndex        =   89
               Top             =   750
               Width           =   615
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Thumb nail:"
               Height          =   195
               Index           =   30
               Left            =   3120
               TabIndex        =   87
               Top             =   480
               Width           =   825
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Contact:"
               Height          =   195
               Index           =   29
               Left            =   3300
               TabIndex        =   85
               Top             =   210
               Width           =   600
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Copyright:"
               Height          =   195
               Index           =   28
               Left            =   120
               TabIndex        =   83
               Top             =   750
               Width           =   705
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Created By:"
               Height          =   195
               Index           =   27
               Left            =   90
               TabIndex        =   82
               Top             =   240
               Width           =   825
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Date:"
               Height          =   195
               Index           =   21
               Left            =   450
               TabIndex        =   81
               Top             =   480
               Width           =   390
            End
         End
         Begin VB.Frame FraDetails 
            Caption         =   "Details 4:"
            Height          =   1005
            Index           =   3
            Left            =   90
            TabIndex        =   51
            Top             =   3210
            Width           =   7605
            Begin VB.TextBox txtMappreview 
               DataField       =   "Image"
               Height          =   285
               Index           =   3
               Left            =   4170
               TabIndex        =   102
               Top             =   690
               Width           =   3375
            End
            Begin VB.TextBox txtMapthumb 
               DataField       =   "ThumbNail"
               Height          =   285
               Index           =   3
               Left            =   4170
               TabIndex        =   100
               Top             =   420
               Width           =   3375
            End
            Begin VB.TextBox txtMapContact 
               DataField       =   "Contact"
               Height          =   285
               Index           =   3
               Left            =   4170
               TabIndex        =   98
               Top             =   150
               Width           =   3375
            End
            Begin VB.TextBox txtMapCopyright 
               DataField       =   "Copyright"
               Height          =   285
               Index           =   3
               Left            =   960
               TabIndex        =   96
               Top             =   660
               Width           =   2115
            End
            Begin VB.TextBox txtMapCreatedBy 
               DataField       =   "CreatedBy"
               Height          =   255
               Index           =   3
               Left            =   960
               TabIndex        =   92
               Top             =   150
               Width           =   2115
            End
            Begin VB.TextBox txtMapDate 
               DataField       =   "CreatedDate"
               Height          =   285
               Index           =   3
               Left            =   960
               TabIndex        =   91
               Top             =   390
               Width           =   2115
            End
            Begin VB.Label lblMappreview 
               AutoSize        =   -1  'True
               Caption         =   "Preview:"
               Height          =   195
               Index           =   2
               Left            =   3360
               TabIndex        =   101
               Top             =   750
               Width           =   615
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Thumb nail:"
               Height          =   195
               Index           =   35
               Left            =   3150
               TabIndex        =   99
               Top             =   480
               Width           =   825
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Contact:"
               Height          =   195
               Index           =   34
               Left            =   3330
               TabIndex        =   97
               Top             =   210
               Width           =   600
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Copyright:"
               Height          =   195
               Index           =   33
               Left            =   150
               TabIndex        =   95
               Top             =   750
               Width           =   705
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Created By:"
               Height          =   195
               Index           =   32
               Left            =   120
               TabIndex        =   94
               Top             =   240
               Width           =   825
            End
            Begin VB.Label lblMapAttr 
               AutoSize        =   -1  'True
               Caption         =   "Date:"
               Height          =   195
               Index           =   31
               Left            =   480
               TabIndex        =   93
               Top             =   480
               Width           =   390
            End
         End
      End
   End
End
Attribute VB_Name = "frmMapProductsWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RSMapProducts As ADODB.Recordset
Private RSLocalUserGroups  As ADODB.Recordset

Private Sub cmdBack_Click()
        '<EhHeader>
        On Error GoTo cmdBack_Click_Err
        '</EhHeader>

100     With C1TTab1Tab2
    
102         cmdNext.Enabled = True
104         cmdNext.Caption = "Next"
        
106         If Not .CurrTab = 0 Then
108             .CurrTab = .CurrTab - 1
            
110             If .CurrTab = 0 Then
112                 cmdBack.Enabled = False
                End If
            
            End If
    
        End With

        '<EhFooter>
        Exit Sub

cmdBack_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapProductsWiz.cmdBack_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdNext_Click()
        '<EhHeader>
        On Error GoTo cmdNext_Click_Err
        '</EhHeader>
        Dim i As Integer

100     With C1TTab1Tab2
102         cmdBack.Enabled = True
        
104         If Not .CurrTab = .NumTabs - 1 Then
        
106             Select Case .CurrTab
            
                    Case 1
108                     i = 0

110                     Do While i <> CInt(ComMapNum.List(ComMapNum.ListIndex))

112                         If Len(txtMapName(i).Text) < 1 Or Len(txtMapAlias(i).Text) < 1 Then
114                             MsgBox "Not valid Name or Alias Value", vbInformation, "OASIS Admin"
                                Exit Sub
                            End If

116                         i = i + 1
                        Loop

118                 Case 2
120                     i = 0

122                     Do While i <> CInt(ComMapNum.List(ComMapNum.ListIndex))

124                         If Len(txtMapPrjFile(i).Text) < 1 Or Len(txtMapDesc(i).Text) < 1 Then
126                             MsgBox "Not valid Values", vbInformation, "OASIS Admin"
                                Exit Sub
                            End If

128                         i = i + 1
                        Loop
                
130                 Case 3
                
                End Select
          
132             .CurrTab = .CurrTab + 1
                          
134             If .CurrTab = .NumTabs - 1 Then
136                 cmdNext.Caption = "Finish"
                End If

            Else
            
138             i = 0
140             Do While i <> CInt(ComMapNum.List(ComMapNum.ListIndex))

142                 If Len(txtMapCreatedBy(i).Text) < 1 Or Len(txtMapDate(i).Text) < 1 Or Len(txtMapCopyright(i).Text) < 1 Or Len(txtMapContact(i).Text) < 1 Or Len(txtMapthumb(i).Text) < 1 Or Len(txtMappreview(i).Text) < 1 Then
144                     MsgBox "Please fill in all values!", vbInformation, "OASIS Admin"
                        Exit Sub
                    End If

146                 i = i + 1
                Loop
            
148             SubmitSettings
            End If
    
        End With

        '<EhFooter>
        Exit Sub

cmdNext_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapProductsWiz.cmdNext_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SubmitSettings()
        '<EhHeader>
        On Error GoTo SubmitSettings_Err
        '</EhHeader>
    
        Dim i As Integer
        Dim bReturnValue As Boolean

100     If MsgBox("Do you wish to save your changes?", vbYesNo, "Confirm Save") = vbYes Then
                  
102         If Not RSMapProducts.EOF Or Not RSMapProducts.Bof Then
            
104             RSMapProducts.MoveFirst
                
106             For i = 0 To 3
            
108                 If i > (ComMapNum - 1) Then
                
110                     RSMapProducts.Delete adAffectCurrent
112                     RSMapProducts.MoveLast
                    Else
                
114                     RSMapProducts.fields("Name").Value = txtMapName(i).Text
116                     RSMapProducts.fields("Alias").Value = txtMapAlias(i).Text
118                     RSMapProducts.fields("FileName").Value = txtMapPrjFile(i).Text
120                     RSMapProducts.fields("Description").Value = txtMapDesc(i).Text
122                     RSMapProducts.fields("CreatedBy").Value = txtMapCreatedBy(i).Text
124                     RSMapProducts.fields("CreatedDate").Value = txtMapDate(i).Text
126                     RSMapProducts.fields("Copyright").Value = txtMapCopyright(i).Text
128                     RSMapProducts.fields("Contact").Value = txtMapContact(i).Text
130                     RSMapProducts.fields("ThumbNail").Value = txtMapthumb(i).Text
132                     RSMapProducts.fields("Image").Value = txtMappreview(i).Text
134                     RSMapProducts.MoveNext
                    End If
                
                Next
            
136             RSMapProducts.MoveFirst
138             bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSMapProducts, WebSite & "Oasis.asp", True)
            
140             If bReturnValue Then
            
142                 IncrementProfileSettingVersion WebSite, "SettingValue6", RSLocalUserGroups.fields("Name").Value
144                 MsgBox "Data saved to server"
                
                Else
146                 MsgBox "Saving to server failed!"
                End If
            
            End If
            
        End If

148     Unload Me
    
        '<EhFooter>
        Exit Sub

SubmitSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapProductsWiz.SubmitSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub setUserGroupsRS(ByVal PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserGroupsRS_Err
        '</EhHeader>
100     Set RSLocalUserGroups = PassedRS
    
        '<EhFooter>
        Exit Sub

setUserGroupsRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapProductsWiz.setUserGroupsRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOpenMapFile_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo cmdOpenMapFile_Click_Err
        '</EhHeader>
        Dim i As Long

        Dim c As New cCommonDialog
        Dim SymbolList As New XGIS_SymbolList
        On Error Resume Next
100     c.DefaultExt = "*.ttkgp"
102     c.DialogTitle = "Open Map Definition File"
104     c.Filter = "Map Definition Files (*.ttkgp;*.prj)|*.ttkgp;*.prj"
106     c.ShowOpen
    
108     If Not c.fileName = "" Then
110         txtMapPrjFile(Index).Text = c.FileTitle
        End If

        '<EhFooter>
        Exit Sub

cmdOpenMapFile_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapProductsWiz.cmdOpenMapFile_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

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
               "in OASISRemoteAdmin.frmMapProductsWiz.CheckIfNull " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub LoadExistingMapProducts()
        '<EhHeader>
        On Error GoTo LoadExistingMapProducts_Err
        '</EhHeader>

        Dim i  As Integer
        Dim sString As String
        Dim iRecCount As Integer
    
100     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & "Maps")
102     Set RSMapProducts = m_frmOASISProgress.OpenHttpCommsRS(sString, True)
        
104     If Not RSMapProducts.EOF Or Not RSMapProducts.Bof Then

106         RSMapProducts.MoveFirst
108         iRecCount = RSMapProducts.RecordCount
        Else
        
            iRecCount = 0
        End If

110     ComMapNum.Clear
112     ComMapNum.AddItem "1"
114     ComMapNum.AddItem "2"
116     ComMapNum.AddItem "3"
118     ComMapNum.AddItem "4"

120     If iRecCount > 0 Then
122         ComMapNum.ListIndex = iRecCount - 1
        Else
124         ComMapNum.ListIndex = 0
        End If
            
126     For i = 0 To 3

128         If i > (iRecCount - 1) Then
130             RSMapProducts.AddNew
132             txtMapName(i).Text = "new map name"
134             RSMapProducts.fields("Name").Value = "new map name"
            Else
    
136             FraMap(i).Visible = True
138             FraFileAttr(i).Visible = True
140             FraDetails(i).Visible = True
        
142             txtMapName(i).Text = CheckIfNull(RSMapProducts.fields("Name").Value)
144             txtMapAlias(i).Text = CheckIfNull(RSMapProducts.fields("Alias").Value)
146             txtMapPrjFile(i).Text = CheckIfNull(RSMapProducts.fields("FileName").Value)
148             txtMapDesc(i).Text = CheckIfNull(RSMapProducts.fields("Description").Value)
150             txtMapCreatedBy(i).Text = CheckIfNull(RSMapProducts.fields("CreatedBy").Value)
152             txtMapDate(i).Text = CheckIfNull(RSMapProducts.fields("CreatedDate").Value)
154             txtMapCopyright(i).Text = CheckIfNull(RSMapProducts.fields("Copyright").Value)
156             txtMapContact(i).Text = CheckIfNull(RSMapProducts.fields("Contact").Value)
158             txtMapthumb(i).Text = CheckIfNull(RSMapProducts.fields("ThumbNail").Value)
160             txtMappreview(i).Text = CheckIfNull(RSMapProducts.fields("Image").Value)
    
162             RSMapProducts.MoveNext
            End If

        Next
    
        '<EhFooter>
        Exit Sub

LoadExistingMapProducts_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmMapProductsWiz.LoadExistingMapProducts " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComMapNum_Click()
        '<EhHeader>
        On Error GoTo ComMapNum_Click_Err
        '</EhHeader>
        Dim i  As Integer
        Dim sString As String
        Dim iRecCount As Integer
        
100     For i = 1 To 3
102         FraMap(i).Visible = False
104         FraFileAttr(i).Visible = False
106         FraDetails(i).Visible = False
        Next
                    
        If ComMapNum.List(ComMapNum.ListIndex) <> "" And ComMapNum.List(ComMapNum.ListIndex) <> 0 Then
                    
108         For i = 0 To ComMapNum.List(ComMapNum.ListIndex) - 1
110             FraMap(i).Visible = True
112             FraFileAttr(i).Visible = True
114             FraDetails(i).Visible = True
            Next

        End If

116     Select Case ComMapNum.ListIndex
    
            Case 0
118             C1TTab1Tab2.Height = 4785 - (1050 * 3)

120         Case 1
122             C1TTab1Tab2.Height = 4785 - (1050 * 2)

124         Case 2
126             C1TTab1Tab2.Height = 4785 - 1050

128         Case Else
130             C1TTab1Tab2.Height = 4785
        End Select
    
132     Me.Height = C1TTab1Tab2.Height + 400
134     cmdNext.Top = C1TTab1Tab2.Height - 410
136     cmdNext.ZOrder 0
138     cmdBack.Top = C1TTab1Tab2.Height - 410
    
        '<EhFooter>
        Exit Sub

ComMapNum_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmMapProductsWiz.ComMapNum_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub InitForm()
Call Form_Load
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        
100     LoadExistingMapProducts

102     C1TTab1Tab2.CurrTab = 0

104     Call ComMapNum_Click
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmMapProductsWiz.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
