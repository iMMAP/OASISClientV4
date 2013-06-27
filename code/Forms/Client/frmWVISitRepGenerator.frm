VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmWVISitrepGenerator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "World Vision Situational Reports"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8745
   Icon            =   "frmWVISitRepGenerator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic C1Elastic 
      Height          =   4815
      Index           =   1
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8745
      _cx             =   15425
      _cy             =   8493
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
      GridRows        =   8
      GridCols        =   8
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmWVISitRepGenerator.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1TTab1Tab 
         Height          =   4815
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   8745
         _cx             =   15425
         _cy             =   8493
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
         Caption         =   "Existing|Create"
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
         TabHeight       =   1
         TabCaptionPos   =   4
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Begin C1SizerLibCtl.C1Elastic C1Elastic 
            Height          =   4725
            Index           =   2
            Left            =   -9300
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   45
            Width           =   8655
            _cx             =   15266
            _cy             =   8334
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic 
            Height          =   4725
            Index           =   0
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   8655
            _cx             =   15266
            _cy             =   8334
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
            BackColor       =   16777215
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Picture         =   "frmWVISitRepGenerator.frx":6927
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   0
            ChildSpacing    =   3
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
            PicturePos      =   6
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   12
            GridCols        =   8
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmWVISitRepGenerator.frx":78BF
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.ComboBox ComOfficeName 
               Height          =   315
               Left            =   5445
               TabIndex        =   20
               Text            =   "ComOfficeName"
               Top             =   1365
               Width           =   3210
            End
            Begin VB.ComboBox ComEnteredBy 
               Height          =   315
               Left            =   1095
               TabIndex        =   19
               Text            =   "ComEnteredBy"
               Top             =   1365
               Width           =   3210
            End
            Begin VB.TextBox txtFilter 
               Enabled         =   0   'False
               Height          =   330
               Left            =   1095
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "txtFilter"
               Top             =   990
               Width           =   7560
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic 
               Height          =   2235
               Index           =   3
               Left            =   0
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   2115
               Width           =   2130
               _cx             =   3757
               _cy             =   3942
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
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"frmWVISitRepGenerator.frx":79CA
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VB.ListBox lstTopics 
                  Height          =   2205
                  ItemData        =   "frmWVISitRepGenerator.frx":79FE
                  Left            =   0
                  List            =   "frmWVISitRepGenerator.frx":7A00
                  TabIndex        =   14
                  Top             =   0
                  Width           =   2130
               End
            End
            Begin VB.ComboBox ComSecRating 
               Height          =   315
               Left            =   4350
               Sorted          =   -1  'True
               TabIndex        =   6
               Text            =   "ComSecRating"
               Top             =   4005
               Visible         =   0   'False
               Width           =   4305
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "Cancel"
               Height          =   330
               Left            =   0
               TabIndex        =   5
               Top             =   4395
               Width           =   2130
            End
            Begin VB.CommandButton cmdOK 
               Caption         =   "Save"
               Height          =   330
               Left            =   5445
               TabIndex        =   4
               Top             =   4395
               Width           =   3210
            End
            Begin VB.CommandButton cmdPreview 
               Caption         =   "Preview"
               Height          =   330
               Left            =   2175
               TabIndex        =   3
               Top             =   4395
               Width           =   2130
            End
            Begin RichTextLib.RichTextBox RichTextBox 
               Height          =   1845
               Left            =   2175
               TabIndex        =   7
               Top             =   2115
               Visible         =   0   'False
               Width           =   6480
               _ExtentX        =   11430
               _ExtentY        =   3254
               _Version        =   393217
               Enabled         =   0   'False
               TextRTF         =   $"frmWVISitRepGenerator.frx":7A02
            End
            Begin C1SizerLibCtl.C1Elastic C1ESECURITYEVENT 
               Height          =   945
               Left            =   0
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   0
               Width           =   4305
               _cx             =   7594
               _cy             =   1667
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Times New Roman"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Enabled         =   -1  'True
               Appearance      =   0
               MousePointer    =   0
               Version         =   801
               BackColor       =   16777215
               ForeColor       =   33023
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   " SITUATIONAL REPORTS"
               Align           =   0
               AutoSizeChildren=   0
               BorderWidth     =   0
               ChildSpacing    =   4
               Splitter        =   0   'False
               FloodDirection  =   0
               FloodPercent    =   0
               CaptionPos      =   2
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
            Begin C1SizerLibCtl.C1Elastic lblMask 
               Height          =   2610
               Left            =   2175
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1740
               Width           =   6480
               _cx             =   11430
               _cy             =   4604
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
               BackColor       =   -2147483645
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "Please click on a Risk Category in the left list"
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
            End
            Begin VB.Label lblSecurityRating 
               BackColor       =   &H80000003&
               Caption         =   " Security Rating:"
               Height          =   345
               Left            =   2175
               TabIndex        =   9
               Top             =   4005
               Visible         =   0   'False
               Width           =   2130
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000003&
               Caption         =   " Filter:"
               Height          =   330
               Left            =   0
               TabIndex        =   18
               Top             =   990
               Width           =   1050
            End
            Begin VB.Label lblEnteredBy 
               BackColor       =   &H80000003&
               Caption         =   " Created by:"
               Height          =   330
               Left            =   0
               TabIndex        =   15
               Top             =   1365
               Width           =   1050
            End
            Begin VB.Label lblSitrepTopics 
               BackColor       =   &H80000003&
               Caption         =   " Risk Category:"
               Height          =   330
               Left            =   0
               TabIndex        =   11
               Top             =   1740
               Width           =   2130
            End
            Begin VB.Label lblNarrative 
               BackColor       =   &H80000003&
               Caption         =   " Narrative:"
               Height          =   330
               Left            =   2175
               TabIndex        =   10
               Top             =   1740
               Width           =   6480
            End
            Begin VB.Label lblDateTill 
               BackColor       =   &H80000003&
               Caption         =   " Office:"
               Height          =   330
               Left            =   4350
               TabIndex        =   8
               Top             =   1365
               Width           =   1050
            End
         End
      End
   End
End
Attribute VB_Name = "frmWVISitrepGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sNarrative_Social As String
Private sNarrative_Crime As String
Private sNarrative_Conflict As String
Private sNarrative_Terrorism As String
Private sNarrative_Kidnapping As String
Private sNarrative_HumSpace As String
Private sNarrative_Insfrast As String
Private sNarrative_Overall As String

Private sRating_Social As String
Private sRating_Crime As String
Private sRating_Conflict As String
Private sRating_Terrorism As String
Private sRating_Kidnapping As String
Private sRating_HumSpace As String
Private sRating_Insfrast As String
Private sRating_Overall As String

Private mCN As adodb.Connection

Private mPic As StdPicture

Private Sub PopulateCombo(ComOfficeName As ComboBox, _
                          sTableName As String, _
                          Optional sFieldName As String = "option")
        '<EhHeader>
        On Error GoTo PopulateCombo_Err
        '</EhHeader>

        Dim sSQL As String
        Dim oRS As New adodb.Recordset

100     ComOfficeName.Clear
102     sSQL = "SELECT DISTINCT [" & sFieldName & "] FROM [" & sTableName & "] ORDER BY [" & sTableName & "].[" & sFieldName & "]"
    
104     With oRS
    
106         .Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic
        
108         Do Until .EOF
        
110             If Not IsNull(.Fields(0).Value) Then
112                 ComOfficeName.AddItem .Fields(0).Value
                End If

114             .MoveNext
        
            Loop
     
116         .Close
    
        End With
    
118     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

PopulateCombo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.PopulateCombo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

'Public Sub Init(sUserName As String, dDateFrom As Date, dDateTill As Date, ppic As StdPicture)
Public Sub Init(sUserName As String, _
                sFilter As String, _
                ppic As StdPicture)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>

100     Set mCN = New adodb.Connection
102     mCN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\DynamicData\WorldVision.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
104     mCN.CursorLocation = g_sGlobalCursorLocation
106     mCN.Open
    
108     PopulateCombo ComEnteredBy, "dd_WVISec_mastertable", "DetailEnteredBy"
110     ComEnteredBy.Text = sUserName
112     ComEnteredBy.Enabled = True
114     ComEnteredBy.Locked = False
    
116     PopulateCombo ComOfficeName, "dd_WVISec_mastertable", "DetailResponsibleOffice"
118     ComOfficeName.Text = g_sRemoteTablePrefix
120     ComOfficeName.Enabled = True
122     ComOfficeName.Locked = False
    
        'txtDateFrom.Text = CStr(Format(dDateFrom, "Medium Date"))
        'txtDateTill.Text = CStr(Format(dDateTill, "Medium Date"))
124     txtFilter.Text = sFilter
    
126     Set mPic = ppic
128     FetchStrings
    
130     PopulateList lstTopics, "dd_WVISec_ddEventCategory"
132     lstTopics.AddItem "-- OVERALL --"
134     PopulateCombo ComSecRating, "dd_WVISec_ddZoneRatings", "Risk"
    
136     dxDateEditFrom = Format(Now() - 7, "Medium Date")
138     dxDateEditTill = Format(Now(), "Medium Date")
    
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.Init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub PopulateList(List1 As ListBox, _
                         sTableName As String, _
                         Optional sFieldName As String = "option")
        '<EhHeader>
        On Error GoTo PopulateList_Err
        '</EhHeader>

        Dim sSQL As String
        Dim oRS As New adodb.Recordset
    
100     List1.Clear
102     sSQL = "SELECT DISTINCT [" & sFieldName & "] FROM [" & sTableName & "] ORDER BY [" & sTableName & "].[" & sFieldName & "]"
    
104     With oRS
    
106         .Open sSQL, mCN, adOpenDynamic, adLockBatchOptimistic
        
108         Do Until .EOF

110             If Not IsNull(.Fields(0).Value) Then
112                 List1.AddItem .Fields(0).Value
                End If

114             .MoveNext
            Loop
     
116         .Close
    
        End With
    
118     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

PopulateList_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.PopulateList " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ShowWait(bShow As Boolean)
        '<EhHeader>
        On Error GoTo ShowWait_Err
        '</EhHeader>

100     C1TTab1Tab.Visible = Not bShow
102     C1TTab1Tab.Container.caption = "Please wait...."
104     C1TTab1Tab.Container.CaptionPos = cpCenterCenter
106     C1TTab1Tab.Container.FontSize = 26
108     C1TTab1Tab.Container.FontBold = True
110     C1TTab1Tab.Container.FontName = "Times New Roman"
    
        '<EhFooter>
        Exit Sub

ShowWait_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.ShowWait " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdPreview_Click()
        '<EhHeader>
        On Error GoTo cmdPreview_Click_Err
        '</EhHeader>

        Dim m_frmCannedReports As frmCannedReports
        Dim sFilter As String
100     Set m_frmCannedReports = New frmCannedReports
    
102     sFilter = txtFilter.Text & " OR ([DetailEventDate] = null)"

104     If Len(txtFilter.Text) < 1 Then sFilter = ""
106     m_frmCannedReports.InitWVISitRep g_sAppPath & "\data\templates\fixedtemplates\WVI.xml", ComEnteredBy.Text, sFilter, sNarrative_Social, sNarrative_Crime, sNarrative_Conflict, sNarrative_Terrorism, sNarrative_Kidnapping, sNarrative_HumSpace, sNarrative_Insfrast, sNarrative_Overall, sRating_Social, sRating_Crime, sRating_Conflict, sRating_Terrorism, sRating_Kidnapping, sRating_HumSpace, sRating_Insfrast, sRating_Overall, mPic

108     m_frmCannedReports.WindowState = 2
110     m_frmCannedReports.Show vbModal, Me
112     Set m_frmCannedReports = Nothing

        '<EhFooter>
        Exit Sub

cmdPreview_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.cmdPreview_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComSecRating_Click()
        '<EhHeader>
        On Error GoTo ComSecRating_Click_Err
        '</EhHeader>
100     SaveStrings
        '<EhFooter>
        Exit Sub

ComSecRating_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.ComSecRating_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComSecRating_LostFocus()
        '<EhHeader>
        On Error GoTo ComSecRating_LostFocus_Err
        '</EhHeader>
100     SaveStrings
        '<EhFooter>
        Exit Sub

ComSecRating_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.ComSecRating_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub lblSitrepTopics_Click()
        '<EhHeader>
        On Error GoTo lblSitrepTopics_Click_Err
        '</EhHeader>
100     FetchStrings
        '<EhFooter>
        Exit Sub

lblSitrepTopics_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.lblSitrepTopics_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub lstTopics_Click()
        '<EhHeader>
        On Error GoTo lstTopics_Click_Err
        '</EhHeader>
100     RichTextBox.Enabled = True
102     RichTextBox.Visible = True
104     lblSecurityRating.Visible = True
106     ComSecRating.Visible = True
108     lblMask.Visible = False
110     FetchStrings
        '<EhFooter>
        Exit Sub

lstTopics_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.lstTopics_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub RichTextBox_LostFocus()
        '<EhHeader>
        On Error GoTo RichTextBox_LostFocus_Err
        '</EhHeader>
100     SaveStrings
        '<EhFooter>
        Exit Sub

RichTextBox_LostFocus_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.RichTextBox_LostFocus " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SaveStrings()
        '<EhHeader>
        On Error GoTo SaveStrings_Err
        '</EhHeader>

100     Select Case lstTopics.Tag
    
            Case "Social & Political"
102             sNarrative_Social = RichTextBox.Text
104             sRating_Social = ComSecRating.Text

106         Case "Crime & Security"
108             sNarrative_Crime = RichTextBox.Text
110             sRating_Crime = ComSecRating.Text

112         Case "Conflict"
114             sNarrative_Conflict = RichTextBox.Text
116             sRating_Conflict = ComSecRating.Text

118         Case "Terrorism"
120             sNarrative_Terrorism = RichTextBox.Text
122             sRating_Terrorism = ComSecRating.Text

124         Case "Kidnapping"
126             sNarrative_Kidnapping = RichTextBox.Text
128             sRating_Kidnapping = ComSecRating.Text
  
130         Case "Humanitarian Space"
132             sNarrative_HumSpace = RichTextBox.Text
134             sRating_HumSpace = ComSecRating.Text

136         Case "Infrastructure"
138             sNarrative_Insfrast = RichTextBox.Text
140             sRating_Insfrast = ComSecRating.Text
            
142         Case "-- OVERALL --"
144             sNarrative_Overall = RichTextBox.Text
146             sRating_Overall = ComSecRating.Text
            
148         Case Else
150             sNarrative_Overall = ""
152             sRating_Overall = ""
            
        End Select
    
154     FetchStrings
    
        '<EhFooter>
        Exit Sub

SaveStrings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.SaveStrings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub FetchStrings()
        '<EhHeader>
        On Error GoTo FetchStrings_Err
        '</EhHeader>

100     Select Case lstTopics.Text
    
            Case "Social & Political"
102             RichTextBox = sNarrative_Social
104             ComSecRating.Text = sRating_Social

106         Case "Crime & Security"
108             RichTextBox = sNarrative_Crime
110             ComSecRating.Text = sRating_Crime

112         Case "Conflict"
114             RichTextBox = sNarrative_Conflict
116             ComSecRating.Text = sRating_Conflict

118         Case "Terrorism"
120             RichTextBox = sNarrative_Terrorism
122             ComSecRating.Text = sRating_Terrorism

124         Case "Kidnapping"
126             RichTextBox = sNarrative_Kidnapping
128             ComSecRating.Text = sRating_Kidnapping
  
130         Case "Humanitarian Space"
132             RichTextBox = sNarrative_HumSpace
134             ComSecRating.Text = sRating_HumSpace

136         Case "Infrastructure"
138             RichTextBox = sNarrative_Insfrast
140             ComSecRating.Text = sRating_Insfrast
            
142         Case "-- OVERALL --"
144             RichTextBox = sNarrative_Overall
146             ComSecRating.Text = sRating_Overall
        
148         Case Else
150             RichTextBox = ""
152             ComSecRating.Text = ""
            
        End Select
    
154     lstTopics.Tag = lstTopics.Text

        '<EhFooter>
        Exit Sub

FetchStrings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmWVISitrepGenerator.FetchStrings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

