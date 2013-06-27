VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.UserControl strWho 
   ClientHeight    =   7620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10920
   ScaleHeight     =   7620
   ScaleWidth      =   10920
   Begin XpressEditorsLibCtl.dxTextEdit dxEdtLocationID 
      DataField       =   "LocationID"
      DataSource      =   "AdoOrganisation"
      Height          =   315
      Left            =   8595
      OleObjectBlob   =   "strWho.ctx":0000
      TabIndex        =   74
      Top             =   2430
      Visible         =   0   'False
      Width           =   1005
   End
   Begin XpressEditorsLibCtl.dxPickEdit dxPickMoveData 
      Height          =   315
      Left            =   2205
      OleObjectBlob   =   "strWho.ctx":0071
      TabIndex        =   39
      Top             =   5985
      Width           =   1995
   End
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   7620
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10920
      _cx             =   19262
      _cy             =   13441
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
      GridRows        =   1
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"strWho.ctx":0482
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab c1TabOrganisation 
         Height          =   7560
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   10860
         _cx             =   19156
         _cy             =   13335
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
         Caption         =   "General|Details|Summary"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   3
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
         Begin C1SizerLibCtl.C1Elastic elSummary 
            Height          =   7185
            Left            =   11805
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   10770
            _cx             =   18997
            _cy             =   12674
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
            Begin VB.TextBox txtTxtSummaryOffice 
               Appearance      =   0  'Flat
               BackColor       =   &H80000011&
               Enabled         =   0   'False
               Height          =   330
               Left            =   0
               TabIndex        =   6
               Text            =   "Amman, Jordan"
               Top             =   1215
               Width           =   2715
            End
            Begin VB.TextBox txtSummaryOrganisation 
               Appearance      =   0  'Flat
               BackColor       =   &H80000011&
               Enabled         =   0   'False
               Height          =   330
               Left            =   0
               TabIndex        =   5
               Text            =   "iMMAP"
               Top             =   495
               Width           =   2715
            End
            Begin VB.CommandButton cmdEdtOrganisation 
               Caption         =   "Edit"
               Height          =   285
               Left            =   1935
               TabIndex        =   4
               Top             =   135
               Width           =   780
            End
            Begin VB.CommandButton cmdEdtOffice 
               Caption         =   "Edit"
               Height          =   285
               Left            =   1935
               TabIndex        =   3
               Top             =   900
               Width           =   780
            End
            Begin VB.Label lblCurrentOrganisation 
               AutoSize        =   -1  'True
               Caption         =   "Current Organisation:"
               Height          =   195
               Left            =   0
               TabIndex        =   8
               Top             =   180
               Width           =   1485
            End
            Begin VB.Label lblCurrentOffice 
               AutoSize        =   -1  'True
               Caption         =   "Current Office:"
               Height          =   195
               Left            =   0
               TabIndex        =   7
               Top             =   945
               Width           =   1020
            End
         End
         Begin C1SizerLibCtl.C1Elastic elContact 
            Height          =   7185
            Left            =   11505
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   330
            Width           =   10770
            _cx             =   18997
            _cy             =   12674
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
            Begin VB.Frame FraStaff 
               Caption         =   "Staff:"
               Height          =   1320
               Left            =   45
               TabIndex        =   28
               Top             =   3240
               Width           =   3975
               Begin XpressEditorsLibCtl.dxTextEdit dxedtNatlStaff 
                  DataField       =   "natlStaff"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   2025
                  OleObjectBlob   =   "strWho.ctx":04BA
                  TabIndex        =   35
                  Top             =   675
                  Width           =   1815
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxedtIntlStaff 
                  DataField       =   "intlStaff"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   2025
                  OleObjectBlob   =   "strWho.ctx":0512
                  TabIndex        =   34
                  Top             =   225
                  Width           =   1815
               End
               Begin VB.Label lblTotalNationalStaff 
                  AutoSize        =   -1  'True
                  Caption         =   "Total National Staff:"
                  Height          =   195
                  Left            =   270
                  TabIndex        =   33
                  Top             =   855
                  Width           =   1410
               End
               Begin VB.Label lblTotalNational 
                  AutoSize        =   -1  'True
                  Caption         =   "Total National Staff:"
                  Height          =   195
                  Left            =   540
                  TabIndex        =   30
                  Top             =   1440
                  Width           =   1410
               End
               Begin VB.Label lblTotalInternational 
                  AutoSize        =   -1  'True
                  Caption         =   "Total international Staff:"
                  Height          =   195
                  Left            =   225
                  TabIndex        =   29
                  Top             =   315
                  Width           =   1635
               End
            End
            Begin VB.Frame FraOfficeUtilities 
               Caption         =   "Office Utilities:"
               Height          =   1095
               Left            =   45
               TabIndex        =   25
               Top             =   2025
               Width           =   3930
               Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpVehicleType 
                  Height          =   315
                  Left            =   1755
                  OleObjectBlob   =   "strWho.ctx":056A
                  TabIndex        =   32
                  Top             =   630
                  Width           =   2085
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxedttoNumVehicles 
                  DataField       =   "numOfVehicles"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   2115
                  OleObjectBlob   =   "strWho.ctx":0753
                  TabIndex        =   31
                  Top             =   225
                  Width           =   1680
               End
               Begin VB.Label lblTypesOfVehicles 
                  AutoSize        =   -1  'True
                  Caption         =   "Types of Vehicles:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   27
                  Top             =   675
                  Width           =   1305
               End
               Begin VB.Label lblTotalNumberVehicles 
                  AutoSize        =   -1  'True
                  Caption         =   "Total Number Of Vehicles:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   26
                  Top             =   270
                  Width           =   1860
               End
            End
            Begin VB.Frame FraOfficeContact 
               Caption         =   "Office Contact:"
               Height          =   1770
               Left            =   45
               TabIndex        =   21
               Top             =   135
               Width           =   4020
               Begin XpressEditorsLibCtl.dxTextEdit dxedtFamName 
                  DataField       =   "contactLname"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   1125
                  OleObjectBlob   =   "strWho.ctx":07AB
                  TabIndex        =   38
                  Top             =   1260
                  Width           =   2805
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxedtFirstName 
                  DataField       =   "contactFname"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   1125
                  OleObjectBlob   =   "strWho.ctx":0803
                  TabIndex        =   37
                  Top             =   720
                  Width           =   2805
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxedttitle 
                  DataField       =   "contactTitle"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   1125
                  OleObjectBlob   =   "strWho.ctx":085B
                  TabIndex        =   36
                  Top             =   270
                  Width           =   2805
               End
               Begin VB.Label lblContactTitle 
                  AutoSize        =   -1  'True
                  Caption         =   "Title:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   24
                  Top             =   360
                  Width           =   345
               End
               Begin VB.Label lblContactFamily 
                  AutoSize        =   -1  'True
                  Caption         =   "Family Name:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   23
                  Top             =   1350
                  Width           =   945
               End
               Begin VB.Label lblContactFirst 
                  AutoSize        =   -1  'True
                  Caption         =   "First Name:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   22
                  Top             =   765
                  Width           =   795
               End
            End
            Begin C1SizerLibCtl.C1Elastic elStaff 
               Height          =   5340
               Left            =   4320
               TabIndex        =   61
               TabStop         =   0   'False
               Top             =   135
               Width           =   6090
               _cx             =   10742
               _cy             =   9419
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
               Begin VB.Frame FraLocationVisibility 
                  Caption         =   "Location visibility level:"
                  Height          =   2685
                  Left            =   45
                  TabIndex        =   66
                  Top             =   855
                  Width           =   3570
                  Begin VB.CheckBox chkVisbilityLevel 
                     Caption         =   "Actual Location"
                     Height          =   375
                     Index           =   5
                     Left            =   270
                     TabIndex        =   72
                     Top             =   2160
                     Width           =   1680
                  End
                  Begin VB.CheckBox chkVisbilityLevel 
                     Caption         =   "Town"
                     Height          =   375
                     Index           =   4
                     Left            =   270
                     TabIndex        =   71
                     Top             =   1770
                     Width           =   1680
                  End
                  Begin VB.CheckBox chkVisbilityLevel 
                     Caption         =   "Sub District"
                     Height          =   375
                     Index           =   3
                     Left            =   270
                     TabIndex        =   70
                     Top             =   1380
                     Width           =   1680
                  End
                  Begin VB.CheckBox chkVisbilityLevel 
                     Caption         =   "District"
                     Height          =   375
                     Index           =   2
                     Left            =   270
                     TabIndex        =   69
                     Top             =   1005
                     Width           =   1680
                  End
                  Begin VB.CheckBox chkVisbilityLevel 
                     Caption         =   "Province"
                     Height          =   375
                     Index           =   1
                     Left            =   270
                     TabIndex        =   68
                     Top             =   615
                     Value           =   1  'Checked
                     Width           =   1680
                  End
                  Begin VB.CheckBox chkVisbilityLevel 
                     Caption         =   "National"
                     Height          =   375
                     Index           =   0
                     Left            =   270
                     TabIndex        =   67
                     Top             =   225
                     Value           =   1  'Checked
                     Width           =   1680
                  End
               End
               Begin VB.Frame FraPrivacy 
                  Caption         =   "Privacy:"
                  Height          =   705
                  Left            =   45
                  TabIndex        =   62
                  Top             =   0
                  Width           =   3570
                  Begin VB.OptionButton OptPrivacy 
                     Caption         =   "Public"
                     Height          =   330
                     Index           =   2
                     Left            =   2250
                     TabIndex        =   63
                     Top             =   225
                     Width           =   1050
                  End
                  Begin VB.OptionButton OptPrivacy 
                     Caption         =   "User group"
                     Height          =   330
                     Index           =   1
                     Left            =   1080
                     TabIndex        =   64
                     Top             =   225
                     Width           =   1230
                  End
                  Begin VB.OptionButton OptPrivacy 
                     Caption         =   "Private"
                     Height          =   330
                     Index           =   0
                     Left            =   135
                     TabIndex        =   65
                     Top             =   225
                     Value           =   -1  'True
                     Width           =   1050
                  End
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elOrganisation 
            Height          =   7185
            Left            =   45
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   330
            Width           =   10770
            _cx             =   18997
            _cy             =   12674
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
            Begin XpressEditorsLibCtl.dxTextEdit dxedtID 
               DataField       =   "id"
               DataSource      =   "AdoOrganisation"
               Height          =   315
               Left            =   3960
               OleObjectBlob   =   "strWho.ctx":08B3
               TabIndex        =   40
               Top             =   1755
               Visible         =   0   'False
               Width           =   1005
            End
            Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUporgtype 
               DataField       =   "organisationTypeId"
               DataSource      =   "AdoOrganisation"
               Height          =   315
               Left            =   90
               OleObjectBlob   =   "strWho.ctx":0924
               TabIndex        =   20
               Top             =   855
               Width           =   3795
            End
            Begin MSAdodcLib.Adodc AdoOrgType 
               Height          =   330
               Left            =   1080
               Top             =   4365
               Visible         =   0   'False
               Width           =   3300
               _ExtentX        =   5821
               _ExtentY        =   582
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
               Caption         =   "orgType"
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
            Begin XpressEditorsLibCtl.dxHyperLinkEdit dxHyperLinkWebsite 
               DataField       =   "website"
               DataSource      =   "AdoOrganisation"
               Height          =   315
               Left            =   1350
               OleObjectBlob   =   "strWho.ctx":0B0D
               TabIndex        =   19
               Top             =   1440
               Width           =   2490
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxedtOrgAcronym 
               DataField       =   "acronym"
               DataSource      =   "AdoOrganisation"
               Height          =   315
               Left            =   90
               OleObjectBlob   =   "strWho.ctx":0C5E
               TabIndex        =   18
               Top             =   1440
               Width           =   1095
            End
            Begin XpressEditorsLibCtl.dxTextEdit dxedtOrgname 
               DataField       =   "name"
               DataSource      =   "AdoOrganisation"
               Height          =   315
               Left            =   90
               OleObjectBlob   =   "strWho.ctx":0CB6
               TabIndex        =   17
               Top             =   315
               Width           =   3795
            End
            Begin MSComctlLib.ListView lstCluster 
               Height          =   2805
               Left            =   45
               TabIndex        =   11
               Top             =   2070
               Width           =   3795
               _ExtentX        =   6694
               _ExtentY        =   4948
               View            =   2
               LabelEdit       =   1
               Sorted          =   -1  'True
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               HideColumnHeaders=   -1  'True
               Checkboxes      =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   0
            End
            Begin MSAdodcLib.Adodc AdoOrganisation 
               Height          =   330
               Left            =   1035
               Top             =   5040
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
               Caption         =   "Org"
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
            Begin C1SizerLibCtl.C1Elastic elOffice 
               Height          =   7095
               Left            =   4320
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   -45
               Width           =   6135
               _cx             =   10821
               _cy             =   12515
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
               Begin VB.CommandButton cmdOfficeLocation 
                  Height          =   330
                  Left            =   3645
                  Picture         =   "strWho.ctx":0D0E
                  Style           =   1  'Graphical
                  TabIndex        =   51
                  Top             =   4725
                  Width           =   375
               End
               Begin MSAdodcLib.Adodc AdoOffType 
                  Height          =   375
                  Left            =   3420
                  Top             =   5085
                  Visible         =   0   'False
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   661
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
                  Caption         =   "Off Type"
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
               Begin MSAdodcLib.Adodc AdoOffStatus 
                  Height          =   375
                  Left            =   3465
                  Top             =   5535
                  Visible         =   0   'False
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   661
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
                  Caption         =   "Off Status"
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
               Begin XpressEditorsLibCtl.dxMemoExEdit dxMemoExAddress2 
                  DataField       =   "address2"
                  DataSource      =   "AdoOrganisation"
                  Height          =   360
                  Left            =   135
                  OleObjectBlob   =   "strWho.ctx":115A
                  TabIndex        =   42
                  Top             =   4320
                  Width           =   3930
               End
               Begin XpressEditorsLibCtl.dxMemoExEdit dxMemoExAddress1 
                  DataField       =   "address1"
                  DataSource      =   "AdoOrganisation"
                  Height          =   360
                  Left            =   135
                  OleObjectBlob   =   "strWho.ctx":12CC
                  TabIndex        =   43
                  Top             =   3645
                  Width           =   3930
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxEdtOffFax 
                  DataField       =   "fax"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   1845
                  OleObjectBlob   =   "strWho.ctx":143E
                  TabIndex        =   44
                  Top             =   3015
                  Width           =   2220
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxEdtOffEmail 
                  DataField       =   "email"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   135
                  OleObjectBlob   =   "strWho.ctx":1496
                  TabIndex        =   45
                  Top             =   3015
                  Width           =   1635
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxEdtOffPhone2 
                  DataField       =   "phone2"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   1845
                  OleObjectBlob   =   "strWho.ctx":14EE
                  TabIndex        =   46
                  Top             =   2385
                  Width           =   2220
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxEdtOffPhone1 
                  DataField       =   "phone1"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   135
                  OleObjectBlob   =   "strWho.ctx":1546
                  TabIndex        =   47
                  Top             =   2385
                  Width           =   1680
               End
               Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpOffStatus 
                  DataField       =   "officeStatus"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   135
                  OleObjectBlob   =   "strWho.ctx":159E
                  TabIndex        =   48
                  Top             =   1665
                  Width           =   3930
               End
               Begin XpressEditorsLibCtl.dxLookUpEdit dxLookUpOffType 
                  DataField       =   "officeType"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   135
                  OleObjectBlob   =   "strWho.ctx":1787
                  TabIndex        =   49
                  Top             =   1035
                  Width           =   3930
               End
               Begin XpressEditorsLibCtl.dxTextEdit dxedtOffPlaceName 
                  DataField       =   "offname"
                  DataSource      =   "AdoOrganisation"
                  Height          =   315
                  Left            =   135
                  OleObjectBlob   =   "strWho.ctx":1970
                  TabIndex        =   50
                  Top             =   360
                  Width           =   3885
               End
               Begin VB.Label lblAttachLocation 
                  Caption         =   "Attach Location..."
                  Height          =   285
                  Left            =   2250
                  TabIndex        =   73
                  Top             =   4770
                  Width           =   1365
               End
               Begin VB.Label lblOfficeStatus 
                  AutoSize        =   -1  'True
                  Caption         =   "Office Status:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   60
                  Top             =   1395
                  Width           =   960
               End
               Begin VB.Label lblOfficeType 
                  AutoSize        =   -1  'True
                  Caption         =   "Office Type:"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   135
                  TabIndex        =   59
                  Top             =   765
                  Width           =   870
               End
               Begin VB.Label lblOfficePlace 
                  AutoSize        =   -1  'True
                  Caption         =   "Office Place Name:"
                  ForeColor       =   &H00000000&
                  Height          =   195
                  Left            =   135
                  TabIndex        =   58
                  Top             =   135
                  Width           =   1380
               End
               Begin VB.Label lblPhone1 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone1:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   57
                  Top             =   2115
                  Width           =   600
               End
               Begin VB.Label lblPhone2 
                  AutoSize        =   -1  'True
                  Caption         =   "Phone2:"
                  Height          =   195
                  Left            =   1845
                  TabIndex        =   56
                  Top             =   2160
                  Width           =   600
               End
               Begin VB.Label lblEmail 
                  AutoSize        =   -1  'True
                  Caption         =   "Email:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   55
                  Top             =   2745
                  Width           =   420
               End
               Begin VB.Label lblFax 
                  AutoSize        =   -1  'True
                  Caption         =   "Fax"
                  Height          =   195
                  Left            =   1845
                  TabIndex        =   54
                  Top             =   2790
                  Width           =   255
               End
               Begin VB.Label lblAddress1 
                  AutoSize        =   -1  'True
                  Caption         =   "Address1:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   53
                  Top             =   3375
                  Width           =   705
               End
               Begin VB.Label lblAddress2 
                  AutoSize        =   -1  'True
                  Caption         =   "Address2:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   52
                  Top             =   4095
                  Width           =   705
               End
            End
            Begin XpressEditorsLibCtl.dxImageLists dxImageLists1 
               Left            =   5355
               OleObjectBlob   =   "strWho.ctx":19C8
               Top             =   360
            End
            Begin VB.Label lblOrganisationName 
               AutoSize        =   -1  'True
               Caption         =   "Organisation Name:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   135
               TabIndex        =   16
               Top             =   45
               Width           =   1395
            End
            Begin VB.Label lblOrganisationType 
               AutoSize        =   -1  'True
               Caption         =   "Organisation Type:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   90
               TabIndex        =   15
               Top             =   630
               Width           =   1335
            End
            Begin VB.Label lblAcronym 
               AutoSize        =   -1  'True
               Caption         =   "acronym:"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   90
               TabIndex        =   14
               Top             =   1215
               Width           =   645
            End
            Begin VB.Label lblWebsite 
               AutoSize        =   -1  'True
               Caption         =   "Website:"
               Height          =   195
               Left            =   1260
               TabIndex        =   13
               Top             =   1215
               Width           =   630
            End
            Begin VB.Label lblSectorCluster 
               AutoSize        =   -1  'True
               Caption         =   "Sector / Cluster Lead for:"
               Height          =   195
               Left            =   45
               TabIndex        =   12
               Top             =   1800
               Width           =   1785
            End
         End
      End
   End
End
Attribute VB_Name = "strWho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_cn As adodb.Connection

Private m_HasInitialized As Boolean
Public Event MoveFirstOK()
Public Event MovePreviousOK()
Public Event MoveNext()
Public Event MoveLast()
Public Event MoveAdd()
Public Event MoveDelete()
Public Event MoveUpdate()

Public Event Addlocation(sID As String)

Private M_RS As adodb.Recordset

Public Sub AttachLocationID(iID As Long)
    dxEdtLocationID.EditValue = iID
End Sub

Public Property Get EOF() As Boolean
    EOF = M_RS.EOF
End Property

Public Property Get BOF() As Boolean
    BOF = M_RS.BOF
End Property

Public Sub MoveFirst()
    dxPickMoveData_ButtonPressed 0, False
    RaiseEvent MoveFirstOK
    
End Sub

Public Sub MovePrevious()
    dxPickMoveData_ButtonPressed 1, False

    RaiseEvent MovePreviousOK
End Sub

Public Sub MoveNext()
    dxPickMoveData_ButtonPressed 2, False
    
    RaiseEvent MoveNext
End Sub

Public Sub MoveLast()
    dxPickMoveData_ButtonPressed 3, False

    
    RaiseEvent MoveLast
End Sub

Public Sub AddRecord()
    dxPickMoveData_ButtonPressed 4, False
    
    RaiseEvent MoveAdd
End Sub

Public Sub DeleteRecord()
    dxPickMoveData_ButtonPressed 5, False
    
    RaiseEvent MoveDelete
End Sub

Public Sub UpdateRecord()
    dxPickMoveData_ButtonPressed 1, False
    dxPickMoveData_ButtonPressed 2, False

    RaiseEvent MoveUpdate
End Sub

Private Sub cmdOfficeLocation_Click()
    RaiseEvent Addlocation(dxedtID.EditValue)
End Sub

Private Sub dxPickMoveData_ButtonPressed(ByVal ButtonIndex As Integer, _
                                         ByVal ButtonDown As Boolean)
'    On Error Resume Next


        With AdoOrganisation.Recordset

            Select Case ButtonIndex

                Case 2
                    .MoveNext

                    If .EOF Then .MoveLast

                Case 1
                    .MovePrevious

                    If .BOF Then .MoveFirst

                Case 0
                    .MoveFirst

                Case 3
                    .MoveLast

                Case 4
                    .AddNew
                    .UpDate

                Case 5
                    .Delete
                    .MoveNext

                    If .EOF Then .MovePrevious

                Case 6
'                    Dim ap
'                    RRequery = True
'                    ap = .AbsolutePosition
'                    .CancelUpdate
'                    .Requery
'                    RRequery = False
'                    .AbsolutePosition = ap
            End Select

        End With


End Sub

Private Sub InitDBS()

    With AdoOrgType
        .ConnectionString = m_Cnn.ConnectionString
        .RecordSource = "1organizationType"
        .Refresh
    End With

    With dxLookUporgtype

        Set .LookUpDataSource = AdoOrgType
        .LookUpKeyFieldName = "id"
        .LookUpDisplayFieldName = "name"
        .KeepInSync = True
        .ListFieldName = "name"

    End With
    
    With AdoOffType
        .ConnectionString = m_Cnn.ConnectionString
        .RecordSource = "1officeType"
        .Refresh
    End With

    With dxLookUpOffType

        Set .LookUpDataSource = AdoOffType
        .LookUpKeyFieldName = "id"
        .LookUpDisplayFieldName = "name"
        .KeepInSync = True
        .ListFieldName = "name"

    End With
    
    
    With AdoOffStatus
        .ConnectionString = m_Cnn.ConnectionString
        .RecordSource = "1officeStatus"
        .Refresh
    End With

    With dxLookUpOffStatus
        Set .LookUpDataSource = AdoOffStatus
        .LookUpKeyFieldName = "id"
        .LookUpDisplayFieldName = "name"
        .KeepInSync = True
        .ListFieldName = "name"
    End With
    
    With AdoOrganisation
        .ConnectionString = m_Cnn.ConnectionString '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
        .CommandType = adCmdTable
        .RecordSource = "organisations"
        .Refresh
    End With

End Sub

Public Sub Init(Optional CN As Connection, Optional bShowAll As Boolean = False, Optional bNavVisible As Boolean = False)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim RS As New adodb.Recordset
    
        Dim RSCurrItem As adodb.Recordset
        Dim lstItem As ListItem
        
        dxPickMoveData.Visible = bNavVisible
        
        
        If bShowAll Then
            cmdEdtOrganisation_Click
        Else
            c1TabOrganisation.TabVisible(0) = False
            c1TabOrganisation.TabVisible(1) = False
            c1TabOrganisation.TabVisible(2) = True
            c1TabOrganisation.CurrTab = 2
        End If

100     If Not CN Is Nothing Then Set m_cn = CN

        InitDBS
        
                '.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & "C:\OASIS\W3Import.mdb"
        
        RS.Open "SELECT name FROM 1sector", m_cn, adOpenDynamic, adLockReadOnly
        
        
        SafeMoveFirst RS
        
        lstCluster.ColumnHeaders.Add Text:="Cluster"
        
        Do While Not RS.EOF
            lstCluster.ListItems.Add Text:=RS.Fields.Item("name").Value
            RS.MoveNext
        Loop
    
        '102     RS.Open "SELECT name, id FROM 1organizationType", m_cn, adOpenForwardOnly, adLockReadOnly
        '
        '104     ComOrgType.Clear
        '106     ComVehicleType.Clear
        '108     ComOfficestatus.Clear
        '
        '110     RS.MoveFirst
        '
        '112     m_HasInitialized = True
        '
        '114     Do While Not RS.EOF
        '
        '116         With RS.Fields
        '118             ComOrgType.AddItem .Item("name").value
        '                'ComOrgType.itemData(ComOrgType.ListCount - 1) = CLng(.Item("id").value)
        '            End With
        '
        '120         RS.MoveNext
        '        Loop
        '
        '122     Set RS = New ADODB.Recordset
        '
        '124     RS.Open "SELECT name, id FROM 1officeStatus", m_cn, adOpenForwardOnly, adLockReadOnly
        '
        '126     Do While Not RS.EOF
        '
        '128         With RS.Fields
        '130             ComOfficestatus.AddItem .Item("name").value
        '                'ComOfficestatus.itemData(ComOfficestatus.ListCount - 1) = .Item("id").value
        '            End With
        '
        '132         RS.MoveNext
        '        Loop
        '
        '134     Set RS = New ADODB.Recordset
        '
        '136     RS.Open "SELECT name, id FROM 1vehicleType", m_cn, adOpenForwardOnly, adLockReadOnly
        '
        '138     Do While Not RS.EOF
        '
        '140         With RS.Fields
        '142             ComVehicleType.AddItem .Item("name").value
        '                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
        '            End With
        '
        '144         RS.MoveNext
        '        Loop
        '
        '146     Set RS = New ADODB.Recordset
        '
        '148     RS.Open "SELECT name, id FROM 1officeType", m_cn, adOpenForwardOnly, adLockReadOnly
        '
        '150     Do While Not RS.EOF
        '
        '152         With RS.Fields
        '154             ComOffType.AddItem .Item("name").value
        '                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
        '            End With
        '
        '156         RS.MoveNext
        '        Loop
        '
        '158     Set RS = New ADODB.Recordset
        '
        '160     lstCluster.ListItems.Clear
        '
        '162     RS.Open "SELECT name, id FROM 1sector", m_cn, adOpenForwardOnly, adLockReadOnly
        '
        '164     Do While Not RS.EOF
        '
        '166         With RS.Fields
        '168             lstCluster.ListItems.Add Text:=.Item("name").value
        '            End With
        '
        '170         RS.MoveNext
        '        Loop
        '
        '172     txtAcronym.Text = ""
        '174     txtAllIntStaff.Text = ""
        '176     txtAllNatStaff.Text = ""
        '178     txtContactFamily.Text = ""
        '180     txtContactFirst.Text = ""
        '182     txtContactTitle.Text = ""
        '184     txtOfficeNumOfVeh.Text = ""
        '186     txtOfficePlace.Text = ""
        '188     txtOfficeType.Text = ""
        '190     txtOrgName.Text = ""
        '192     txtOrgWebsite.Text = ""
        '194     txtTotIntStaff.Text = ""
        '196     txtTotNatStaff.Text = ""
        '198     txtTotNumOfVehicles.Text = ""
        '
        '        On Error Resume Next
        '200     ComOfficestatus.ListIndex = 0
        '202     ComOrgType.ListIndex = 0
        '204     ComStaffAlocOffice.ListIndex = 0
        '206     ComVehicleAllocOffice.ListIndex = 0
        '208     ComVehicleType.ListIndex = 0
        '210     ComOffType.ListIndex = 0
        '
        '        On Error GoTo Init_Err
        '
        '212     c1TabOrganisation.TabVisible(0) = True
        '214     g_RSAppSettings.MoveFirst
        '216     g_RSAppSettings.Find "SettingName = 'w3WHOOffice'"
        '218     c1TabOrganisation.TabVisible(1) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = 1, True, False)
        '
        '220     g_RSAppSettings.MoveFirst
        '222     g_RSAppSettings.Find "SettingName = 'w3WHOContact'"
        '224     c1TabOrganisation.TabVisible(2) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = 1, True, False)
        '
        '226     g_RSAppSettings.MoveFirst
        '228     g_RSAppSettings.Find "SettingName = 'w3WHOStaff'"
        '230     c1TabOrganisation.TabVisible(3) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = 1, True, False)
        '
        '232     g_RSAppSettings.MoveFirst
        '234     g_RSAppSettings.Find "SettingName = 'w3WHOTransport'"
        '236     c1TabOrganisation.TabVisible(4) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").value = 1, True, False)
        '
        '238     c1TabOrganisation.TabVisible(5) = False
        '
        '240     g_RSAppSettings.MoveFirst
        '242     g_RSAppSettings.Find "SettingName = 'w3OrgID'"
        '
        '244     If Not g_RSAppSettings.Fields.Item("SettingValue1").value = vbNull Then
        '246         If IsNumeric(g_RSAppSettings.Fields.Item("SettingValue1").value) Then
        '
        '248             Set RS = New ADODB.Recordset
        '
        '250             RS.Open "SELECT * FROM 1organisation WHERE id = " & CLng(g_RSAppSettings.Fields.Item("SettingValue1").value), m_cn, adOpenForwardOnly, adLockReadOnly
        '252             c1TabOrganisation.TabVisible(0) = False
        '254             c1TabOrganisation.TabVisible(1) = False
        '256             c1TabOrganisation.TabVisible(2) = False
        '258             c1TabOrganisation.TabVisible(3) = False
        '260             c1TabOrganisation.TabVisible(4) = False
        '262             c1TabOrganisation.TabVisible(5) = True
        '264             c1TabOrganisation.CurrTab = 5
        '266             txtSummaryOrganisation.Text = RS.Fields.Item("name").value
        '
        '268             Set RSCurrItem = New ADODB.Recordset
        '
        '270             RSCurrItem.Open "SELECT id, name FROM 1organizationType WHERE id = '" & RS.Fields.Item("organizationType").value & "'", m_cn
        '
        '                'ComOfficestatus
        '                'ComOffType
        '272             ItemInBox RSCurrItem.Fields.Item("name").value, ComOrgType
        '
        '274             With RS.Fields
        '276                 txtAcronym.Text = .Item("acronym")
        '278                 txtOfficeNumOfVeh.Text = .Item("numOfVehicles")
        '280                 txtOrgName.Text = .Item("name")
        '282                 txtOrgWebsite.Text = .Item("website")
        '284                 txtTotIntStaff.Text = .Item("intlStaff")
        '286                 txtTotNatStaff.Text = .Item("natlStaff")
        '288                 txtTotNumOfVehicles.Text = .Item("numOfVehicles")
        '                End With
        '
        '            End If
        '        End If
        '
        '290     g_RSAppSettings.MoveFirst
        '292     g_RSAppSettings.Find "SettingName = 'w3OfficeID'"
        '
        '294     If Not g_RSAppSettings.Fields.Item("SettingValue1").value = vbNull Then
        '296         If IsNumeric(g_RSAppSettings.Fields.Item("SettingValue1").value) Then
        '
        '298             Set RS = New ADODB.Recordset
        '
        '300             RS.Open "SELECT * FROM 1office WHERE id = " & CLng(g_RSAppSettings.Fields.Item("SettingValue1").value), m_cn, adOpenForwardOnly, adLockReadOnly
        '302             txtTxtSummaryOffice.Text = RS.Fields.Item("name").value
        '
        '304             Set RSCurrItem = New ADODB.Recordset
        '
        '306             RSCurrItem.Open "SELECT id, name FROM 1officeType WHERE id = '" & RS.Fields.Item("officeType").value & "'", m_cn
        '
        '                'ComOfficestatus
        '
        '308             ItemInBox RSCurrItem.Fields.Item("name").value, ComOffType
        '
        '310             Set RSCurrItem = New ADODB.Recordset
        '
        '                'TODO BELOW
        '312             'RSCurrItem.Open "SELECT id, name FROM 1officeStatus WHERE id = '" & RS.Fields.Item("officeStatus").value & "'", m_cn
        '
        '314             'ItemInBox RSCurrItem.Fields.Item("name").value, ComOfficestatus
        '
        '                On Error Resume Next
        '
        '316             With RS.Fields
        '318                 txtOfficeNumOfVeh.Text = .Item("numOfVehicles").value
        '320                 txtOfficePlace.Text = .Item("name").value
        '                End With
        '
        '            End If
        '        End If
        '
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox "OASISClient.strWho.Init LINE:" & Erl & " strWho component failure" & Err.Description
        '</EhFooter>
End Sub

Private Sub cmdAdd_Click()
'    txtAcronym.Text = ""
'    txtAllIntStaff.Text = "0"
'    txtAllNatStaff.Text = "0"
'    txtContactFamily.Text = ""
'    txtContactFirst.Text = ""
'    txtContactTitle.Text = ""
'    txtOfficeNumOfVeh.Text = "0"
'    txtOfficePlace.Text = ""
'    txtOfficeType.Text = ""
'    txtOrgName.Text = ""
'    txtOrgWebsite.Text = ""
'    txtTotIntStaff.Text = "0"
'    txtTotNatStaff.Text = "0"
'    txtTotNumOfVehicles.Text = "0"
'
'    On Error Resume Next
'    ComOfficestatus.ListIndex = 0
'    ComOrgType.ListIndex = 0
'    ComStaffAlocOffice.ListIndex = 0
'    ComVehicleAllocOffice.ListIndex = 0
'    ComVehicleType.ListIndex = 0
'
'    Dim i As Integer
'
'    For i = 1 To lstCluster.ListItems.Count
'        lstCluster.ListItems.Item(i).Checked = False
'    Next
'
'    'cmdAdd.Enabled = False
'
End Sub

Private Sub cmdEdtOffice_Click()
    c1TabOrganisation.TabVisible(0) = True
    c1TabOrganisation.TabVisible(1) = True
    c1TabOrganisation.TabVisible(2) = False
    c1TabOrganisation.CurrTab = 0
    'elAction.Visible = True
End Sub

Private Sub cmdEdtOrganisation_Click()
    c1TabOrganisation.TabVisible(0) = True
    c1TabOrganisation.TabVisible(1) = True
    c1TabOrganisation.TabVisible(2) = False
    c1TabOrganisation.CurrTab = 0
    'elAction.Visible = True
End Sub

Private Sub cmdOk_Click()
'    Dim sSQL As String
'    Dim sVal As String
'    Dim rs As ADODB.Recordset
'
'    'If cmdAdd.Enabled Then
'    'UPDATE
'    'Else
'    'INSERT
'
'    sVal = "INSERT INTO 1organisation (acronym, intlStaff, natlStaff, numOfVehicles, name, website, vehicleTypes, organizationType) VALUES ("
'
'    'TODO!!! clusterId
'
'    'INSERT INTO
'
'    sSQL = sSQL & "'" & txtAcronym.Text & "'"
'    sSQL = sSQL & ", '" & txtTotIntStaff.Text & "'"
'    sSQL = sSQL & ", '" & txtTotNatStaff.Text & "'"
'    sSQL = sSQL & ", '" & txtTotNumOfVehicles.Text & "'"
'    sSQL = sSQL & ", '" & txtOrgName.Text & "'"
'    sSQL = sSQL & ", '" & txtOrgWebsite.Text & "'"
'    'sSQL = sSQL & ", '" & txtTotIntStaff.Text & "'"
'    'sSQL = sSQL & ", '" & txtTotNatStaff.Text & "'"
'    'sSQL = sSQL & ", '" & txtTotNumOfVehicles.Text & "'"
'
'    Set rs = New ADODB.Recordset
'
'    rs.Open "SELECT * FROM 1vehicleType WHERE name = '" & ComVehicleType.List(ComVehicleType.ListIndex) & "'", m_cn
'
'    sSQL = sSQL & ", '" & rs.Fields.Item("id").value & "'"
'
'    Set rs = New ADODB.Recordset
'
'    rs.Open "SELECT * FROM 1organizationType WHERE name = '" & ComOrgType.List(ComOrgType.ListIndex) & "'", m_cn
'
'    sSQL = sSQL & ", '" & rs.Fields.Item("id").value & "'"
'    'sSQL = sSQL & ", '" & lstCluster.List(ComOrgType.ListIndex) & "'"
'
'    Set rs = m_cn.Execute(sVal & sSQL & ")")
'
'    '        sSQL = sSQL & "'" & txtAcronym.Text & "'"
'    '        sSQL = sSQL & ", '" & txtAllIntStaff.Text & "'"
'    '        sSQL = sSQL & ", '" & txtAllNatStaff.Text & "'"
'
'    '        sSQL = sSQL & ", '" & txtOfficeNumOfVeh.Text & "'"
'    '        sSQL = sSQL & ", '" & txtOfficePlace.Text & "'"
'    '        sSQL = sSQL & ", '" & txtOfficeType.Text & "'"
'    '        sSQL = sSQL & ", '" & txtOrgName.Text & "'"
'    '        sSQL = sSQL & ", '" & txtOrgWebsite.Text & "'"
'    '        sSQL = sSQL & ", '" & txtTotIntStaff.Text & "'"
'    '        sSQL = sSQL & ", '" & txtTotNatStaff.Text & "'"
'    '        sSQL = sSQL & ", '" & txtTotNumOfVehicles.Text & "'"
'
'    Set rs = New ADODB.Recordset
'
'    rs.Open "SELECT * FROM 1organisation WHERE name = '" & txtOrgName.Text & "'", m_cn
'
'    m_Cnn.Execute "UPDATE AppSettings SET SettingValue1 = '" & rs.Fields.Item("id").value & "' WHERE SettingName = 'w3OrgID'"
'
'    sVal = "INSERT INTO 1contact (firstName, lastName, title, orgId) VALUES ("
'
'    sSQL = "'" & txtContactFirst.Text & "'"
'    sSQL = sSQL & ", '" & txtContactFamily.Text & "'"
'    sSQL = sSQL & ", '" & txtContactTitle.Text & "'"
'    sSQL = sSQL & ", '" & rs.Fields.Item("id").value & "'"
'
'    Set rs = m_cn.Execute(sVal & sSQL & ")")
'
'    ' = ""
'    ' = ""
'
'    sVal = "INSERT INTO 1office (name, officeType, orgId) VALUES ("
'
'    sSQL = "'" & txtOfficePlace.Text & "'"
'
'    Set rs = New ADODB.Recordset
'    rs.Open "SELECT * FROM 1officeType WHERE name = '" & ComOffType.List(ComOffType.ListIndex) & "'", m_cn
'
'    sSQL = sSQL & ", '" & rs.Fields.Item("id").value & "'"
'
'    Set rs = New ADODB.Recordset
'
'    rs.Open "SELECT * FROM 1organisation WHERE name = '" & txtOrgName.Text & "'", m_cn
'
'    sSQL = sSQL & ", '" & rs.Fields.Item("id").value & "'"
'
'    m_cn.Execute sVal & sSQL & ")"
'
'    m_Cnn.Execute "UPDATE AppSettings SET SettingValue1 = '" & rs.Fields.Item("id").value & "' WHERE SettingName = 'w3OfficeID'"
'
'    '        On Error Resume Next
'    '        ComOfficestatus.ListIndex = 0
'    '
'    '        ComStaffAlocOffice.ListIndex = 0
'    '        ComVehicleAllocOffice.ListIndex = 0
'    '
'    '        Dim i As Integer
'    '
'    '        For i = 1 To lstCluster.ListItems.Count
'    '            lstCluster.ListItems.Item(i).Checked = False
'    '        Next
'
'    'End If
'
'    ' elAction.Visible = False
'    Init
End Sub

Public Function HasInitialized() As Boolean
    HasInitialized = m_HasInitialized
End Function

Private Sub UserControl_Resize()
On Error Resume Next
    dxPickMoveData.Move 100, UserControl.Height - ((dxPickMoveData.Height * 2) + 100)
'    dxPickMoveData.Move 100, UserControl.Height - (dxPickMoveData.Height + 100)
End Sub
