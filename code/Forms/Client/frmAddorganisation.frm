VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddorganisation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Who?"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   5445
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4560
      _cx             =   8043
      _cy             =   9604
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
      Begin C1SizerLibCtl.C1Elastic elAction 
         Height          =   645
         Left            =   225
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   4545
         Visible         =   0   'False
         Width           =   3930
         _cx             =   6932
         _cy             =   1138
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
         Begin VB.CommandButton cmdOk 
            Caption         =   "Ok"
            Height          =   375
            Left            =   2610
            TabIndex        =   59
            Top             =   0
            Width           =   870
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "Delete"
            Height          =   375
            Left            =   1755
            TabIndex        =   58
            Top             =   0
            Width           =   870
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   900
            TabIndex        =   57
            Top             =   0
            Width           =   870
         End
      End
      Begin C1SizerLibCtl.C1Tab c1TabOrganisation 
         Height          =   4155
         Left            =   180
         TabIndex        =   1
         Top             =   225
         Width           =   4065
         _cx             =   7170
         _cy             =   7329
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
         Caption         =   "Organisation|Office|Contact|Staff|Transport|Summary"
         Align           =   0
         CurrTab         =   5
         FirstTab        =   1
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
         Begin C1SizerLibCtl.C1Elastic elSummary 
            Height          =   3780
            Left            =   45
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   330
            Width           =   3975
            _cx             =   7011
            _cy             =   6668
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
            Begin VB.CommandButton cmdEdtOffice 
               Caption         =   "Edit"
               Height          =   285
               Left            =   1935
               TabIndex        =   54
               Top             =   900
               Width           =   780
            End
            Begin VB.CommandButton cmdEdtOrganisation 
               Caption         =   "Edit"
               Height          =   285
               Left            =   1935
               TabIndex        =   53
               Top             =   135
               Width           =   780
            End
            Begin VB.TextBox txtSummaryOrganisation 
               Appearance      =   0  'Flat
               BackColor       =   &H80000011&
               Enabled         =   0   'False
               Height          =   330
               Left            =   0
               TabIndex        =   52
               Text            =   "Organisation"
               Top             =   495
               Width           =   2715
            End
            Begin VB.TextBox txtTxtSummaryOffice 
               Appearance      =   0  'Flat
               BackColor       =   &H80000011&
               Enabled         =   0   'False
               Height          =   330
               Left            =   0
               TabIndex        =   51
               Text            =   "txtSummaryOffice"
               Top             =   1215
               Width           =   2715
            End
            Begin VB.Label lblCurrentOffice 
               AutoSize        =   -1  'True
               Caption         =   "Current Office:"
               Height          =   195
               Left            =   0
               TabIndex        =   56
               Top             =   945
               Width           =   1020
            End
            Begin VB.Label lblCurrentOrganisation 
               AutoSize        =   -1  'True
               Caption         =   "Current Organisation:"
               Height          =   195
               Left            =   0
               TabIndex        =   55
               Top             =   180
               Width           =   1485
            End
         End
         Begin C1SizerLibCtl.C1Elastic elvehicle 
            Height          =   3780
            Left            =   -4620
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   330
            Width           =   3975
            _cx             =   7011
            _cy             =   6668
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
            Begin VB.Frame FraAllocateTransportation 
               Caption         =   "Allocate transportation Vehicle:"
               Height          =   1590
               Left            =   90
               TabIndex        =   42
               Top             =   1890
               Width           =   3300
               Begin VB.ComboBox ComVehicleAllocOffice 
                  Height          =   315
                  Left            =   90
                  TabIndex        =   45
                  Text            =   "VehicleAllocOffice"
                  Top             =   495
                  Width           =   1770
               End
               Begin VB.TextBox txtOfficeNumOfVeh 
                  Height          =   285
                  Left            =   1395
                  TabIndex        =   44
                  Text            =   "OfficeNumOfVeh"
                  Top             =   990
                  Width           =   1050
               End
               Begin VB.CommandButton cmdApplyVehicles 
                  Caption         =   "Apply Vehicles"
                  Height          =   465
                  Left            =   2520
                  TabIndex        =   43
                  Top             =   900
                  Width           =   690
               End
               Begin VB.Label lblVehicleAllocation 
                  AutoSize        =   -1  'True
                  Caption         =   "Vehicle Allocation Office:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   47
                  Top             =   270
                  Width           =   1575
               End
               Begin VB.Label lblOfficeVehicles 
                  AutoSize        =   -1  'True
                  Caption         =   "Office Vehicles:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   46
                  Top             =   990
                  Width           =   1080
               End
            End
            Begin VB.ComboBox ComVehicleType 
               Height          =   315
               Left            =   270
               TabIndex        =   41
               Text            =   "VehicleType"
               Top             =   1080
               Width           =   2490
            End
            Begin VB.TextBox txtTotNumOfVehicles 
               Height          =   330
               Left            =   2385
               TabIndex        =   40
               Text            =   "TotNumOfVehicles "
               Top             =   180
               Width           =   510
            End
            Begin VB.Label lblTypesOfVehicles 
               AutoSize        =   -1  'True
               Caption         =   "Types of Vehicles:"
               Height          =   195
               Left            =   315
               TabIndex        =   49
               Top             =   720
               Width           =   1305
            End
            Begin VB.Label lblTotalNumberVehicles 
               AutoSize        =   -1  'True
               Caption         =   "Total Number Of Vehicles:"
               Height          =   195
               Left            =   270
               TabIndex        =   48
               Top             =   225
               Width           =   1860
            End
         End
         Begin C1SizerLibCtl.C1Elastic elStaff 
            Height          =   3780
            Left            =   -4920
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   330
            Width           =   3975
            _cx             =   7011
            _cy             =   6668
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
            Begin VB.Frame FraAllocateStaff 
               Caption         =   "Allocate Staff To Offices:"
               Height          =   1590
               Left            =   135
               TabIndex        =   30
               Top             =   1755
               Width           =   3300
               Begin VB.CommandButton cmdApply 
                  Caption         =   "Apply"
                  Height          =   285
                  Left            =   2700
                  TabIndex        =   34
                  Top             =   1080
                  Width           =   510
               End
               Begin VB.TextBox txtAllNatStaff 
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   33
                  Text            =   "AllNatStaff"
                  Top             =   1170
                  Width           =   1050
               End
               Begin VB.TextBox txtAllIntStaff 
                  Height          =   285
                  Left            =   1440
                  TabIndex        =   32
                  Text            =   "AllIntStaff"
                  Top             =   945
                  Width           =   1050
               End
               Begin VB.ComboBox ComStaffAlocOffice 
                  Height          =   315
                  Left            =   90
                  TabIndex        =   31
                  Text            =   "StaffAlocOffice"
                  Top             =   495
                  Width           =   1770
               End
               Begin VB.Label lblNumNatl 
                  AutoSize        =   -1  'True
                  Caption         =   "Num Natl Staff:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   39
                  Top             =   1170
                  Width           =   1080
               End
               Begin VB.Label lblNumIntl 
                  AutoSize        =   -1  'True
                  Caption         =   "Num Intl Staff:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   38
                  Top             =   945
                  Width           =   1005
               End
               Begin VB.Label lblStaffAllocation 
                  AutoSize        =   -1  'True
                  Caption         =   "Staff Allocation Office:"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   37
                  Top             =   270
                  Width           =   1575
               End
            End
            Begin VB.TextBox txtTotNatStaff 
               Height          =   330
               Left            =   1890
               TabIndex        =   29
               Text            =   "totNatStaff"
               Top             =   1035
               Width           =   600
            End
            Begin VB.TextBox txtTotIntStaff 
               Height          =   285
               Left            =   1890
               TabIndex        =   28
               Text            =   "totIntStaff"
               Top             =   585
               Width           =   645
            End
            Begin VB.Label lblTotalNational 
               AutoSize        =   -1  'True
               Caption         =   "Total National Staff:"
               Height          =   195
               Left            =   135
               TabIndex        =   36
               Top             =   1035
               Width           =   1410
            End
            Begin VB.Label lblTotalInternational 
               AutoSize        =   -1  'True
               Caption         =   "Total international Staff:"
               Height          =   195
               Left            =   90
               TabIndex        =   35
               Top             =   585
               Width           =   1680
            End
         End
         Begin C1SizerLibCtl.C1Elastic elContact 
            Height          =   3780
            Left            =   -5220
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   330
            Width           =   3975
            _cx             =   7011
            _cy             =   6668
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
            Begin VB.CommandButton cmdContactDetails 
               Caption         =   "Contact Details"
               Height          =   465
               Left            =   135
               TabIndex        =   27
               Top             =   3240
               Width           =   1185
            End
            Begin VB.TextBox txtContactTitle 
               Height          =   285
               Left            =   180
               TabIndex        =   26
               Text            =   "Contact title"
               Top             =   2610
               Width           =   3390
            End
            Begin VB.TextBox txtContactFamily 
               Height          =   285
               Left            =   180
               TabIndex        =   25
               Text            =   "Contact Family"
               Top             =   2025
               Width           =   3390
            End
            Begin VB.TextBox txtContactFirst 
               Height          =   285
               Left            =   180
               TabIndex        =   24
               Text            =   "Contact First Name"
               Top             =   1350
               Width           =   3390
            End
            Begin VB.Frame v 
               Caption         =   "Contact Privacy Status:"
               Height          =   690
               Left            =   135
               TabIndex        =   21
               Top             =   135
               Width           =   3390
               Begin VB.OptionButton OptPrivacy 
                  Caption         =   "Private"
                  Height          =   195
                  Index           =   1
                  Left            =   1755
                  TabIndex        =   23
                  Top             =   270
                  Width           =   915
               End
               Begin VB.OptionButton OptPrivacy 
                  Caption         =   "Public"
                  Height          =   195
                  Index           =   0
                  Left            =   180
                  TabIndex        =   22
                  Top             =   270
                  Value           =   -1  'True
                  Width           =   915
               End
            End
            Begin VB.Label lblContactTitle 
               AutoSize        =   -1  'True
               Caption         =   "Contact title:"
               Height          =   195
               Left            =   180
               TabIndex        =   65
               Top             =   2340
               Width           =   885
            End
            Begin VB.Label lblContactFamily 
               AutoSize        =   -1  'True
               Caption         =   "Contact family Name:"
               Height          =   195
               Left            =   180
               TabIndex        =   64
               Top             =   1755
               Width           =   1500
            End
            Begin VB.Label lblContactFirst 
               AutoSize        =   -1  'True
               Caption         =   "Contact First Name:"
               Height          =   195
               Left            =   180
               TabIndex        =   63
               Top             =   1035
               Width           =   1395
            End
         End
         Begin C1SizerLibCtl.C1Elastic elOffice 
            Height          =   3780
            Left            =   -5520
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   330
            Width           =   3975
            _cx             =   7011
            _cy             =   6668
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
            Begin VB.ComboBox ComOffType 
               Height          =   315
               Left            =   135
               Style           =   2  'Dropdown List
               TabIndex        =   67
               Top             =   990
               Width           =   3660
            End
            Begin VB.CommandButton cmdOfficeLocation 
               Caption         =   "Office Location"
               Height          =   465
               Left            =   135
               TabIndex        =   20
               Top             =   3195
               Width           =   1455
            End
            Begin VB.ComboBox ComOfficestatus 
               Height          =   315
               Left            =   135
               TabIndex        =   15
               Text            =   "Officestatus"
               Top             =   1665
               Width           =   3615
            End
            Begin VB.TextBox txtOfficeType 
               Height          =   315
               Left            =   135
               TabIndex        =   14
               Text            =   "Office Type"
               Top             =   990
               Width           =   3615
            End
            Begin VB.TextBox txtOfficePlace 
               Height          =   315
               Left            =   135
               TabIndex        =   13
               Text            =   "Office Place Name:"
               Top             =   405
               Width           =   3615
            End
            Begin VB.Label lblOfficeStatus 
               AutoSize        =   -1  'True
               Caption         =   "Office Status:"
               Height          =   195
               Left            =   135
               TabIndex        =   62
               Top             =   1395
               Width           =   960
            End
            Begin VB.Label lblOfficeType 
               AutoSize        =   -1  'True
               Caption         =   "Office Type:"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   135
               TabIndex        =   61
               Top             =   765
               Width           =   870
            End
            Begin VB.Label lblOfficePlace 
               AutoSize        =   -1  'True
               Caption         =   "Office Place Name:"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   135
               TabIndex        =   60
               Top             =   135
               Width           =   1380
            End
         End
         Begin C1SizerLibCtl.C1Elastic elOrganisation 
            Height          =   3780
            Left            =   -5820
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   3975
            _cx             =   7011
            _cy             =   6668
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
            Begin MSComctlLib.ListView lstCluster 
               Height          =   1500
               Left            =   90
               TabIndex        =   66
               Top             =   2070
               Width           =   3615
               _ExtentX        =   6376
               _ExtentY        =   2646
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
            Begin VB.TextBox txtOrgWebsite 
               Height          =   285
               Left            =   1305
               TabIndex        =   6
               Top             =   1440
               Width           =   2355
            End
            Begin VB.TextBox txtAcronym 
               Height          =   285
               Left            =   45
               TabIndex        =   5
               Top             =   1440
               Width           =   1095
            End
            Begin VB.ComboBox ComOrgType 
               Height          =   315
               Left            =   45
               TabIndex        =   4
               Text            =   "orgType"
               Top             =   855
               Width           =   3615
            End
            Begin VB.TextBox txtOrgName 
               Height          =   330
               Left            =   45
               TabIndex        =   3
               Top             =   270
               Width           =   3615
            End
            Begin VB.Label lblSectorCluster 
               AutoSize        =   -1  'True
               Caption         =   "Sector / Cluster Lead for:"
               Height          =   195
               Left            =   45
               TabIndex        =   11
               Top             =   1800
               Width           =   1785
            End
            Begin VB.Label lblWebsite 
               AutoSize        =   -1  'True
               Caption         =   "Website:"
               Height          =   195
               Left            =   1260
               TabIndex        =   10
               Top             =   1215
               Width           =   630
            End
            Begin VB.Label lblAcronym 
               AutoSize        =   -1  'True
               Caption         =   "acronym:"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   90
               TabIndex        =   9
               Top             =   1215
               Width           =   645
            End
            Begin VB.Label lblOrganisationType 
               AutoSize        =   -1  'True
               Caption         =   "Organisation Type:"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   90
               TabIndex        =   8
               Top             =   630
               Width           =   1335
            End
            Begin VB.Label lblOrganisationName 
               AutoSize        =   -1  'True
               Caption         =   "Organisation Name:"
               ForeColor       =   &H000000C0&
               Height          =   195
               Left            =   135
               TabIndex        =   7
               Top             =   45
               Width           =   1395
            End
         End
      End
   End
End
Attribute VB_Name = "frmAddorganisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cn As ADODB.Connection

Public Sub Init(Optional cn As Connection)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        Dim rs As New ADODB.Recordset
    
        Dim RSCurrItem As ADODB.Recordset
        Dim lstItem As ListItem

100     If Not cn Is Nothing Then Set m_cn = cn
    
102     rs.Open "SELECT name, id FROM 1organizationType", m_cn, adOpenForwardOnly, adLockReadOnly
    
104     ComOrgType.Clear
106     ComVehicleType.Clear
108     ComOfficestatus.Clear
    
110     SafeMoveFirst rs
    
112     Do While Not rs.EOF

114         With rs.Fields
116             ComOrgType.AddItem .Item("name").Value
                'ComOrgType.itemData(ComOrgType.ListCount - 1) = CLng(.Item("id").value)
            End With

118         rs.MoveNext
        Loop
    
120     Set rs = New ADODB.Recordset
    
122     rs.Open "SELECT name, id FROM 1officeStatus", m_cn, adOpenForwardOnly, adLockReadOnly
    
124     Do While Not rs.EOF

126         With rs.Fields
128             ComOfficestatus.AddItem .Item("name").Value
                'ComOfficestatus.itemData(ComOfficestatus.ListCount - 1) = .Item("id").value
            End With

130         rs.MoveNext
        Loop
    
132     Set rs = New ADODB.Recordset
    
134     rs.Open "SELECT name, id FROM 1vehicleType", m_cn, adOpenForwardOnly, adLockReadOnly
    
136     Do While Not rs.EOF

138         With rs.Fields
140             ComVehicleType.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

142         rs.MoveNext
        Loop
    
144     Set rs = New ADODB.Recordset
    
146     rs.Open "SELECT name, id FROM 1officeType", m_cn, adOpenForwardOnly, adLockReadOnly
    
148     Do While Not rs.EOF

150         With rs.Fields
152             ComOffType.AddItem .Item("name").Value
                'ComVehicleType.itemData(ComVehicleType.ListCount - 1) = .Item("id").value
            End With

154         rs.MoveNext
        Loop
    
156     Set rs = New ADODB.Recordset
    
158     lstCluster.ListItems.Clear
    
160     rs.Open "SELECT name, id FROM 1sector", m_cn, adOpenForwardOnly, adLockReadOnly
    
162     Do While Not rs.EOF

164         With rs.Fields
166             lstCluster.ListItems.Add Text:=.Item("name").Value
            End With

168         rs.MoveNext
        Loop
    
170     txtAcronym.Text = ""
172     txtAllIntStaff.Text = ""
174     txtAllNatStaff.Text = ""
176     txtContactFamily.Text = ""
178     txtContactFirst.Text = ""
180     txtContactTitle.Text = ""
182     txtOfficeNumOfVeh.Text = ""
184     txtOfficePlace.Text = ""
186     txtOfficeType.Text = ""
188     txtOrgName.Text = ""
190     txtOrgWebsite.Text = ""
192     txtTotIntStaff.Text = ""
194     txtTotNatStaff.Text = ""
196     txtTotNumOfVehicles.Text = ""
    
        On Error Resume Next
198     ComOfficestatus.ListIndex = 0
200     ComOrgType.ListIndex = 0
202     ComStaffAlocOffice.ListIndex = 0
204     ComVehicleAllocOffice.ListIndex = 0
206     ComVehicleType.ListIndex = 0
208     ComOffType.ListIndex = 0
    
210     c1TabOrganisation.TabVisible(0) = True
212     SafeMoveFirst g_RSAppSettings
214     g_RSAppSettings.Find "SettingName = 'w3WHOOffice'"
216     c1TabOrganisation.TabVisible(1) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = 1, True, False)
    
218     SafeMoveFirst g_RSAppSettings
220     g_RSAppSettings.Find "SettingName = 'w3WHOContact'"
222     c1TabOrganisation.TabVisible(2) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = 1, True, False)
    
224     SafeMoveFirst g_RSAppSettings
226     g_RSAppSettings.Find "SettingName = 'w3WHOStaff'"
228     c1TabOrganisation.TabVisible(3) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = 1, True, False)
    
230     SafeMoveFirst g_RSAppSettings
232     g_RSAppSettings.Find "SettingName = 'w3WHOTransport'"
234     c1TabOrganisation.TabVisible(4) = IIf(g_RSAppSettings.Fields.Item("SettingValue1").Value = 1, True, False)
    
236     c1TabOrganisation.TabVisible(5) = False
    
238     SafeMoveFirst g_RSAppSettings
240     g_RSAppSettings.Find "SettingName = 'w3OrgID'"
        
242     If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
244         If IsNumeric(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then
            
246             Set rs = New ADODB.Recordset
            
248             rs.Open "SELECT * FROM 1organisation WHERE id = " & CLng(g_RSAppSettings.Fields.Item("SettingValue1").Value), m_cn, adOpenForwardOnly, adLockReadOnly
250             c1TabOrganisation.TabVisible(0) = False
252             c1TabOrganisation.TabVisible(1) = False
254             c1TabOrganisation.TabVisible(2) = False
256             c1TabOrganisation.TabVisible(3) = False
258             c1TabOrganisation.TabVisible(4) = False
260             c1TabOrganisation.TabVisible(5) = True
262             c1TabOrganisation.CurrTab = 5
264             txtSummaryOrganisation.Text = rs.Fields.Item("name").Value
            
266             Set RSCurrItem = New ADODB.Recordset
            
268             RSCurrItem.Open "SELECT id, name FROM 1organisationType WHERE id = '" & rs.Fields.Item("organizationType").Value & "'", m_cn
            
                'ComOfficestatus
                'ComOffType
270             ItemInBox RSCurrItem.Fields.Item("name").Value, ComOrgType

272             With rs.Fields
274                 txtAcronym.Text = .Item("acronym")
276                 txtOfficeNumOfVeh.Text = .Item("numOfVehicles")
278                 txtOrgName.Text = .Item("name")
280                 txtOrgWebsite.Text = .Item("website")
282                 txtTotIntStaff.Text = .Item("intlStaff")
284                 txtTotNatStaff.Text = .Item("natlStaff")
286                 txtTotNumOfVehicles.Text = .Item("numOfVehicles")
                End With
            
            End If
        End If
    
288     SafeMoveFirst g_RSAppSettings
290     g_RSAppSettings.Find "SettingName = 'w3OfficeID'"
    
292     If Not g_RSAppSettings.Fields.Item("SettingValue1").Value = vbNull Then
294         If IsNumeric(g_RSAppSettings.Fields.Item("SettingValue1").Value) Then
            
296             Set rs = New ADODB.Recordset
            
298             rs.Open "SELECT * FROM 1office WHERE id = " & CLng(g_RSAppSettings.Fields.Item("SettingValue1").Value), m_cn, adOpenForwardOnly, adLockReadOnly
300             txtTxtSummaryOffice.Text = rs.Fields.Item("name").Value
            
302             Set RSCurrItem = New ADODB.Recordset
            
304             RSCurrItem.Open "SELECT id, name FROM 1officeType WHERE id = '" & rs.Fields.Item("officeType").Value & "'", m_cn
            
                'ComOfficestatus

306             ItemInBox RSCurrItem.Fields.Item("name").Value, ComOffType
            
308             Set RSCurrItem = New ADODB.Recordset
            
310             RSCurrItem.Open "SELECT id, name FROM 1officeStatus WHERE id = '" & rs.Fields.Item("officeStatus").Value & "'", m_cn

312             ItemInBox RSCurrItem.Fields.Item("name").Value, ComOfficestatus
            
                On Error Resume Next
314             With rs.Fields
316                 txtOfficeNumOfVeh.Text = .Item("numOfVehicles").Value
318                 txtOfficePlace.Text = .Item("name").Value
                End With
            
            End If
        End If
    
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddorganisation.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub cmdAdd_Click()
        '<EhHeader>
        On Error GoTo cmdAdd_Click_Err
        '</EhHeader>
100     txtAcronym.Text = ""
102     txtAllIntStaff.Text = "0"
104     txtAllNatStaff.Text = "0"
106     txtContactFamily.Text = ""
108     txtContactFirst.Text = ""
110     txtContactTitle.Text = ""
112     txtOfficeNumOfVeh.Text = "0"
114     txtOfficePlace.Text = ""
116     txtOfficeType.Text = ""
118     txtOrgName.Text = ""
120     txtOrgWebsite.Text = ""
122     txtTotIntStaff.Text = "0"
124     txtTotNatStaff.Text = "0"
126     txtTotNumOfVehicles.Text = "0"
        
        On Error Resume Next
128     ComOfficestatus.ListIndex = 0
130     ComOrgType.ListIndex = 0
132     ComStaffAlocOffice.ListIndex = 0
134     ComVehicleAllocOffice.ListIndex = 0
136     ComVehicleType.ListIndex = 0
    
        Dim i As Integer
    
138     For i = 1 To lstCluster.ListItems.Count
140         lstCluster.ListItems.Item(i).Checked = False
        Next
    
142     cmdAdd.Enabled = False
    
        '<EhFooter>
        Exit Sub

cmdAdd_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddorganisation.cmdAdd_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEdtOffice_Click()
        '<EhHeader>
        On Error GoTo cmdEdtOffice_Click_Err
        '</EhHeader>
100     c1TabOrganisation.TabVisible(0) = True
102     c1TabOrganisation.TabVisible(1) = True
104     c1TabOrganisation.TabVisible(2) = True
106     c1TabOrganisation.TabVisible(3) = True
108     c1TabOrganisation.TabVisible(4) = True
110     c1TabOrganisation.TabVisible(5) = False
112     c1TabOrganisation.CurrTab = 1
        'elAction.Visible = True
        '<EhFooter>
        Exit Sub

cmdEdtOffice_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddorganisation.cmdEdtOffice_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdEdtOrganisation_Click()
        '<EhHeader>
        On Error GoTo cmdEdtOrganisation_Click_Err
        '</EhHeader>
100     c1TabOrganisation.TabVisible(0) = True
102     c1TabOrganisation.TabVisible(1) = True
104     c1TabOrganisation.TabVisible(2) = True
106     c1TabOrganisation.TabVisible(3) = True
108     c1TabOrganisation.TabVisible(4) = True
110     c1TabOrganisation.TabVisible(5) = False
112     c1TabOrganisation.CurrTab = 0
        'elAction.Visible = True
        '<EhFooter>
        Exit Sub

cmdEdtOrganisation_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddorganisation.cmdEdtOrganisation_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdOK_Click()
        '<EhHeader>
        On Error GoTo cmdOK_Click_Err
        '</EhHeader>
        Dim sSQL As String
        Dim sVal As String
        Dim rs As ADODB.Recordset
    
100     If cmdAdd.Enabled Then
            'UPDATE
        Else
            'INSERT
        
102         sVal = "INSERT INTO 1organisation (acronym, intlStaff, natlStaff, numOfVehicles, name, website, vehicleTypes, organizationType) VALUES ("
    
            'TODO!!! clusterId
        
            'INSERT INTO
        
104         sSQL = sSQL & "'" & txtAcronym.Text & "'"
106         sSQL = sSQL & ", '" & txtTotIntStaff.Text & "'"
108         sSQL = sSQL & ", '" & txtTotNatStaff.Text & "'"
110         sSQL = sSQL & ", '" & txtTotNumOfVehicles.Text & "'"
112         sSQL = sSQL & ", '" & txtOrgName.Text & "'"
114         sSQL = sSQL & ", '" & txtOrgWebsite.Text & "'"
            'sSQL = sSQL & ", '" & txtTotIntStaff.Text & "'"
            'sSQL = sSQL & ", '" & txtTotNatStaff.Text & "'"
            'sSQL = sSQL & ", '" & txtTotNumOfVehicles.Text & "'"
        
116         Set rs = New ADODB.Recordset
        
118         rs.Open "SELECT * FROM 1vehicleType WHERE name = '" & ComVehicleType.List(ComVehicleType.ListIndex) & "'", m_cn
        
120         sSQL = sSQL & ", '" & rs.Fields.Item("id").Value & "'"
        
122         Set rs = New ADODB.Recordset
        
124         rs.Open "SELECT * FROM 1organizationType WHERE name = '" & ComOrgType.List(ComOrgType.ListIndex) & "'", m_cn
        
126         sSQL = sSQL & ", '" & rs.Fields.Item("id").Value & "'"
            'sSQL = sSQL & ", '" & lstCluster.List(ComOrgType.ListIndex) & "'"
       
128         Set rs = m_cn.Execute(sVal & sSQL & ")")
        
            '        sSQL = sSQL & "'" & txtAcronym.Text & "'"
            '        sSQL = sSQL & ", '" & txtAllIntStaff.Text & "'"
            '        sSQL = sSQL & ", '" & txtAllNatStaff.Text & "'"

            '        sSQL = sSQL & ", '" & txtOfficeNumOfVeh.Text & "'"
            '        sSQL = sSQL & ", '" & txtOfficePlace.Text & "'"
            '        sSQL = sSQL & ", '" & txtOfficeType.Text & "'"
            '        sSQL = sSQL & ", '" & txtOrgName.Text & "'"
            '        sSQL = sSQL & ", '" & txtOrgWebsite.Text & "'"
            '        sSQL = sSQL & ", '" & txtTotIntStaff.Text & "'"
            '        sSQL = sSQL & ", '" & txtTotNatStaff.Text & "'"
            '        sSQL = sSQL & ", '" & txtTotNumOfVehicles.Text & "'"
        
130         Set rs = New ADODB.Recordset
        
132         rs.Open "SELECT * FROM 1organisation WHERE name = '" & txtOrgName.Text & "'", m_cn
        
134         m_Cnn.Execute "UPDATE " & g_sAppSettingsTable & " SET SettingValue1 = '" & rs.Fields.Item("id").Value & "' WHERE SettingName = 'w3OrgID'"
        
136         sVal = "INSERT INTO 1contact (firstName, lastName, title, orgId) VALUES ("
            
138         sSQL = "'" & txtContactFirst.Text & "'"
140         sSQL = sSQL & ", '" & txtContactFamily.Text & "'"
142         sSQL = sSQL & ", '" & txtContactTitle.Text & "'"
144         sSQL = sSQL & ", '" & rs.Fields.Item("id").Value & "'"
                
146         Set rs = m_cn.Execute(sVal & sSQL & ")")
    
            ' = ""
            ' = ""
        
148         sVal = "INSERT INTO 1office (name, officeType, orgId) VALUES ("
            
150         sSQL = "'" & txtOfficePlace.Text & "'"
        
152         Set rs = New ADODB.Recordset
154         rs.Open "SELECT * FROM 1officeType WHERE name = '" & ComOffType.List(ComOffType.ListIndex) & "'", m_cn
        
156         sSQL = sSQL & ", '" & rs.Fields.Item("id").Value & "'"
                
158         Set rs = New ADODB.Recordset
        
160         rs.Open "SELECT * FROM 1organisation WHERE name = '" & txtOrgName.Text & "'", m_cn
        
        
        
162         sSQL = sSQL & ", '" & rs.Fields.Item("id").Value & "'"
        
164         m_cn.Execute sVal & sSQL & ")"
        
166         m_Cnn.Execute "UPDATE " & g_sAppSettingsTable & " SET SettingValue1 = '" & rs.Fields.Item("id").Value & "' WHERE SettingName = 'w3OfficeID'"
        
            '        On Error Resume Next
            '        ComOfficestatus.ListIndex = 0
            '
            '        ComStaffAlocOffice.ListIndex = 0
            '        ComVehicleAllocOffice.ListIndex = 0
            '
            '        Dim i As Integer
            '
            '        For i = 1 To lstCluster.ListItems.Count
            '            lstCluster.ListItems.Item(i).Checked = False
            '        Next
    
        End If


        'elAction.Visible = False
168     Init
        '<EhFooter>
        Exit Sub

cmdOK_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddorganisation.cmdOK_Click " & _
               "at line " & Erl
        'Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     If Not g_sLanguage = "" Then
102         If Not m_Cnn.State = adStateClosed Then
104             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddorganisation.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
