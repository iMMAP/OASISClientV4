VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSendMess 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Alert Sender"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4110
   Icon            =   "frmSendMess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   4110
      _cx             =   7250
      _cy             =   10398
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
      Begin C1SizerLibCtl.C1Tab C1TTab 
         Height          =   5655
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   3915
         _cx             =   6906
         _cy             =   9975
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
         Caption         =   "Message|Recipients|Settings"
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
         Begin C1SizerLibCtl.C1Elastic elTab3 
            Height          =   5280
            Left            =   4860
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   330
            Width           =   3825
            _cx             =   6747
            _cy             =   9313
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
         End
         Begin C1SizerLibCtl.C1Elastic elTab2 
            Height          =   5280
            Left            =   4560
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   330
            Width           =   3825
            _cx             =   6747
            _cy             =   9313
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
            Begin VB.TextBox txtRecipient 
               Height          =   285
               Left            =   60
               TabIndex        =   10
               Text            =   "0093795976313"
               Top             =   120
               Width           =   2415
            End
            Begin VB.CommandButton cmdAddRecipient 
               Caption         =   "Add to list"
               Height          =   315
               Left            =   2580
               TabIndex        =   9
               Top             =   120
               Width           =   1095
            End
            Begin VB.ListBox lstRecipients 
               Height          =   1815
               Left            =   60
               TabIndex        =   8
               Top             =   480
               Width           =   2415
            End
         End
         Begin C1SizerLibCtl.C1Elastic elTab1 
            Height          =   5280
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   3825
            _cx             =   6747
            _cy             =   9313
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
            Begin VB.CommandButton cmdSendSMS 
               Caption         =   "Send!"
               Height          =   375
               Left            =   2340
               TabIndex        =   11
               Top             =   4740
               Width           =   1335
            End
            Begin VB.TextBox txtSMS 
               Height          =   1935
               Left            =   60
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   4
               Top             =   2700
               Width           =   3615
            End
            Begin MSComctlLib.ListView lvwAttributes 
               Height          =   2115
               Left            =   60
               TabIndex        =   3
               Top             =   540
               Width           =   3675
               _ExtentX        =   6482
               _ExtentY        =   3731
               View            =   3
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   1
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Field:"
                  Object.Width           =   2540
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Value"
                  Object.Width           =   2540
               EndProperty
            End
            Begin VB.Label lblCounter 
               Height          =   255
               Left            =   60
               TabIndex        =   5
               Top             =   4680
               Width           =   3615
            End
         End
      End
   End
End
Attribute VB_Name = "frmSendMess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private m_oShp As TatukGIS_XDK10.XGIS_Shape

Private Sub cmdAddRecipient_Click()
    lstRecipients.AddItem txtRecipient.Text
End Sub

Private Function FillGrid()
    Dim i As Integer
    
    For i = 0 To 10
        lvwAttributes.ListItems.Add , , "no:" & i
        lvwAttributes.ListItems.Item(lvwAttributes.ListItems.Count).SubItems(1) = Now
    Next
    Me.caption = ShowCredits
    
End Function

Public Function ShowCredits() As Long
    Dim SMS As Object
    Set SMS = CreateObject("ASPSMS.Booster")
    
    If SMS Is Nothing Then Exit Function
    
    
    SMS.UserKey = "MZ6T5UIZGWMQ"
    SMS.Password = "immap123"

    ShowCredits = CLng(SMS.Credits)

    If SMS.errorCode <> 1 Then
        MsgBox Str$(SMS.errorCode)
        MsgBox SMS.ErrorDescription
    End If

    Set SMS = Nothing
End Function


Private Sub cmdSendSMS_Click()
    Dim SMS As Object
    Dim i As Integer
    
    If lstRecipients.ListCount < 1 Then
        MsgBox "No recipeients!"
        Exit Sub
    End If
    
    Set SMS = CreateObject("ASPSMS.Booster")

    SMS.UserKey = "MZ6T5UIZGWMQ"
    SMS.Password = "immap123"
    SMS.Originator = "OASIS"
    SMS.MessageData = txtSMS.Text
    'SMS.FlashingSMS = True
    
    For i = 0 To lstRecipients.ListCount
        SMS.AddRecipient lstRecipients.List(i)
    Next
    
    'SMS.AddRecipient "0093795976313"
    'SMS.AddRecipient "0099591607404"
    SMS.SendTextSMS
    
    MsgBox SMS.errorCode
    
    SMS.DeleteAllRecipients
    Set SMS = Nothing
End Sub

Private Sub lvwAttributes_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim i As Integer
Dim sText As String

    For i = 1 To lvwAttributes.ListItems.Count
        If lvwAttributes.ListItems(i).Checked Then
            sText = sText & lvwAttributes.ListItems(i).SubItems(1) & vbLf
        End If
    Next

    txtSMS.Text = Mid$(sText, 1, Len(sText) - 1)
    
    End Sub

Private Sub txtSMS_Change()
    Dim sTemp As String
    Dim sOutput As String
    Dim i As Integer

    sTemp = txtSMS.Text

'    If Me.chkBlinkingSMS.Value = 1 Then    'Characters left is depending on using BlinkingSMS Feature or not
'        'BlinkingSMS used
'        '<BLINK> is replaced by <Space>, because it only uses 1 character
'        sTemp = Replace(sTemp, "<BLINK>", " ", , , vbTextCompare)
'        sTemp = Replace(sTemp, "</BLINK>", " ", , , vbTextCompare)
'
'        If 69 - Len(sTemp) >= 0 Then
'            lblCounter.caption = "Characters left: " & Trim(69 - Len(sTemp))
'        Else
'            lblCounter.caption = "Too much characters used"
'        End If
'
'    Else

        'BlinkingSMS not used
        If 160 - Len(sTemp) >= 0 Then
            lblCounter.caption = "Characters left: " & Trim(160 - Len(sTemp))
        Else
            lblCounter.caption = "Too many characters used"
        End If
    'End If

End Sub
