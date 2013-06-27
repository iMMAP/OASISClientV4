VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmOASISSynch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Data Pack Manager"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9435
   Icon            =   "frmOASISSynch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9435
   StartUpPosition =   1  'CenterOwner
   Begin OASISRemoteAdmin.Downloader Downloader1 
      Left            =   2025
      Top             =   2430
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin C1SizerLibCtl.C1Elastic c1Main 
      Height          =   5370
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9435
      _cx             =   16642
      _cy             =   9472
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
      _GridInfo       =   $"frmOASISSynch.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elNavs 
         Height          =   675
         Left            =   90
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4605
         Width           =   9255
         _cx             =   16325
         _cy             =   1191
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
         Begin VB.TextBox txtServerURL 
            Height          =   285
            Left            =   45
            TabIndex        =   16
            Text            =   "http://www.immap.org/"
            Top             =   270
            Width           =   4695
         End
         Begin VB.CommandButton cmdDo 
            Caption         =   "Dowload Data Packs"
            Enabled         =   0   'False
            Height          =   510
            Left            =   4905
            TabIndex        =   15
            Top             =   90
            Width           =   1410
         End
         Begin VB.CommandButton cmdSET 
            Caption         =   "Save Data Pack Definitions"
            Enabled         =   0   'False
            Height          =   510
            Left            =   6345
            TabIndex        =   14
            Top             =   90
            Width           =   1410
         End
         Begin VB.CommandButton cmdCommand1 
            Caption         =   "Get Data Pack Definitions"
            Height          =   510
            Left            =   7785
            TabIndex        =   13
            Top             =   90
            Width           =   1410
         End
         Begin VB.Label lblOASISServer 
            AutoSize        =   -1  'True
            Caption         =   "OASIS Server URL:"
            Height          =   195
            Left            =   90
            TabIndex        =   17
            Top             =   45
            Width           =   1410
         End
      End
      Begin C1SizerLibCtl.C1Tab C1TSynch 
         Height          =   4455
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   9255
         _cx             =   16325
         _cy             =   7858
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
         Appearance      =   1
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FrontTabColor   =   -2147483633
         BackTabColor    =   -2147483633
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "Settings|Files|Datasets|Profile|Data Packs"
         Align           =   0
         CurrTab         =   4
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
         Flags(0)        =   2
         Flags(1)        =   2
         Flags(2)        =   2
         Flags(3)        =   2
         Begin C1SizerLibCtl.C1Elastic elDataPacks 
            Height          =   4080
            Left            =   45
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   330
            Width           =   9165
            _cx             =   16166
            _cy             =   7197
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
            Begin DXDBGRIDLibCtl.dxDBGrid dxDataPacks 
               Height          =   4020
               Left            =   90
               OleObjectBlob   =   "frmOASISSynch.frx":6895
               TabIndex        =   11
               Top             =   45
               Width           =   8970
            End
         End
         Begin C1SizerLibCtl.C1Elastic elProfile 
            Height          =   4080
            Left            =   -9810
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   330
            Width           =   9165
            _cx             =   16166
            _cy             =   7197
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
            Begin DXDBGRIDLibCtl.dxDBGrid dxProfile 
               Height          =   3885
               Left            =   45
               OleObjectBlob   =   "frmOASISSynch.frx":753D
               TabIndex        =   10
               Top             =   90
               Width           =   4155
            End
         End
         Begin C1SizerLibCtl.C1Elastic elDatasets 
            Height          =   4080
            Left            =   -10110
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   330
            Width           =   9165
            _cx             =   16166
            _cy             =   7197
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
            Begin DXDBGRIDLibCtl.dxDBGrid dxDatasets 
               Height          =   3885
               Left            =   0
               OleObjectBlob   =   "frmOASISSynch.frx":81E5
               TabIndex        =   9
               Top             =   45
               Width           =   4290
            End
         End
         Begin C1SizerLibCtl.C1Elastic elFiles 
            Height          =   4080
            Left            =   -10410
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   330
            Width           =   9165
            _cx             =   16166
            _cy             =   7197
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
            Begin DXDBGRIDLibCtl.dxDBGrid dxFiles 
               Height          =   4020
               Left            =   90
               OleObjectBlob   =   "frmOASISSynch.frx":8E8D
               TabIndex        =   8
               Top             =   45
               Width           =   4155
            End
         End
         Begin C1SizerLibCtl.C1Elastic elSettings 
            Height          =   4080
            Left            =   -10710
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   9165
            _cx             =   16166
            _cy             =   7197
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
            Begin VB.CheckBox chkAutomaticDownload 
               Caption         =   "Automatic Download Updates"
               Height          =   375
               Left            =   135
               TabIndex        =   7
               Top             =   45
               Width           =   1950
            End
         End
      End
   End
End
Attribute VB_Name = "frmOASISSynch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RSDataPacks As ADODB.Recordset
Private WithEvents m_cUnzip As cUnzip
Attribute m_cUnzip.VB_VarHelpID = -1


Private Sub cmdCommand1_Click()
        '<EhHeader>
        On Error GoTo cmdCommand1_Click_Err
        '</EhHeader>
    Dim sWebsite As String

100     Set RSDataPacks = New ADODB.Recordset
    
102     sWebsite = txtServerURL.Text

104     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
106         sWebsite = sWebsite & "/"
        End If
    
108     RSDataPacks.open sWebsite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM DataPacks")
    
110     Set dxDataPacks.DataSource = RSDataPacks
112     dxDataPacks.Columns.RetrieveFields
114     dxDataPacks.Columns(0).Visible = False
116     dxDataPacks.Columns(1).Visible = False
        'dxDataPacks.Columns.Add gedCheckEdit
118     cmdSET.Enabled = True
120     cmdDo.Enabled = True
        '<EhFooter>
        Exit Sub

cmdCommand1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.cmdCommand1_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdDo_Click()
        '<EhHeader>
        On Error GoTo cmdDo_Click_Err
        '</EhHeader>
        Dim sWebsite As String
    
100     sWebsite = txtServerURL.Text

102     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
104         sWebsite = sWebsite & "/"
        End If
        
106     With dxDataPacks
        
    '108         m_frmDebug.DebugPrint .KeyField
        
108         .Dataset.First
        
110         Do While Not .Dataset.EOF

112             If .Dataset.FieldValues("Update") Then

114                 Downloader1.BeginDownload sWebsite & .Dataset.FieldValues("Path"), CreateAppPath & "\data\" & .Dataset.FieldValues("FolderConst") & "\" & .Dataset.FieldValues("Path")
                End If
            
116             .Dataset.next
            Loop
    
        End With

        '<EhFooter>
        Exit Sub

cmdDo_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.cmdDo_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function pOpen(ByVal sfile As String) As Boolean
        '<EhHeader>
        On Error GoTo pOpen_Err
        '</EhHeader>
    Dim i As Long
   
       ' Get the file directory:
100    Set m_cUnzip = New cUnzip
102    m_cUnzip.ZipFile = sfile
104    m_cUnzip.Directory
   
    '   'If m_cUnzip.FileCount > 0 Then
    '   '   m_cZipMRU.Add sFIle
    '   'End If
    '
    '   ' Display it in the ListView:
    '   For i = 1 To m_cUnzip.FileCount
    '      sFIle = m_cUnzip.Filename(i)
    '      If m_cUnzip.FileEncrypted(i) Then
    '         ' the way WinZip represents it.  I guess a nicer way would be
    '         ' to use overlay icons/state icons and/or colour changes in the LV
    '         'sFIle = sFIle & "+"
    '      End If
    '
    '      m_frmDebug.DebugPrint m_cUnzip.FileSize(i)
    '      m_frmDebug.DebugPrint Format$(m_cUnzip.FileDate(i), "short date") & " " & Format$(m_cUnzip.FileDate(i), "short time")
    '      m_frmDebug.DebugPrint m_cUnzip.FilePackedSize(i)
    '      m_frmDebug.DebugPrint m_cUnzip.FileDirectory(i)
    '   Next i
        '<EhFooter>
        Exit Function

pOpen_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.pOpen " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub ExtractZip()
        '<EhHeader>
        On Error GoTo ExtractZip_Err
        '</EhHeader>
    Dim bSel As Boolean
    Dim sFolder As String
    Dim iItem As Long

100       For iItem = 1 To m_cUnzip.FileCount
102          m_cUnzip.FileSelected(iItem) = True
104       Next iItem
   
       ' Get extract folder and do it:
   
     '  .Dataset.FieldValues ("Path") '& .Dataset.FieldValues("FolderConst")
   
106    sFolder = App.Path
   
    '   If (sFolder <> "") Then
108       m_cUnzip.OverwriteExisting = True
110       m_cUnzip.ExtractOnlyNewer = True
112       m_cUnzip.UseFolderNames = True
114       m_cUnzip.UnzipFolder = CreateAppPath & "\data\" & dxDataPacks.Dataset.FieldValues("FolderConst")
116       m_cUnzip.Unzip
    '   End If
   
        '<EhFooter>
        Exit Sub

ExtractZip_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.ExtractZip " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Downloader1_DownloadComplete(MaxBytes As Long, SaveFile As String)
        '<EhHeader>
        On Error GoTo Downloader1_DownloadComplete_Err
        '</EhHeader>
    
100     pOpen SaveFile
102     ExtractZip
     
104     frmLOG.txtLog.Text = frmLOG.txtLog.Text & vbCrLf & "OASIS Data Pack Manager: Completed " & SaveFile & ", Size = " & MaxBytes

        '<EhFooter>
        Exit Sub

Downloader1_DownloadComplete_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.Downloader1_DownloadComplete " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Downloader1_DownloadProgress(CurBytes As Long, _
                                         MaxBytes As Long, _
                                         SaveFile As String)
        '<EhHeader>
        On Error GoTo Downloader1_DownloadProgress_Err
        '</EhHeader>

        Dim i As Integer
        Dim RemBytes As Long

100     RemBytes = MaxBytes - CurBytes
                
102     If RemBytes < 2 ^ 20 Then
104         frmLOG.txtLog.Text = frmLOG.txtLog.Text & vbCrLf & "OASIS Data Pack Manager Progress: " & SaveFile & Format((MaxBytes - CurBytes) / 2 ^ 10, "#0.0 KB") & " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")"
        Else
106         frmLOG.txtLog.Text = frmLOG.txtLog.Text & vbCrLf & "OASIS Data Pack Manager Progress: " & SaveFile & Format((MaxBytes - CurBytes) / 2 ^ 20, "#0.00 MB") & " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")"
        End If

        '<EhFooter>
        Exit Sub

Downloader1_DownloadProgress_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.Downloader1_DownloadProgress " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Downloader1_DownloadError(SaveFile As String)
        '<EhHeader>
        On Error GoTo Downloader1_DownloadError_Err
        '</EhHeader>

100     frmLOG.txtLog.Text = frmLOG.txtLog.Text & vbCrLf & "Error downloading " & SaveFile

        '<EhFooter>
        Exit Sub

Downloader1_DownloadError_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.Downloader1_DownloadError " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSET_Click()
        '<EhHeader>
        On Error GoTo cmdSET_Click_Err
        '</EhHeader>

        Dim MsXmlHttp As New MSXML2.ServerXMLHTTP40
        Dim MsXmlDoc As New MSXML2.DOMDocument
        Dim sWebsite As String
        
        Dim sGUID() As String
        Dim sClientPath() As String
        Dim sServerPath() As String
        
100     sWebsite = txtServerURL.Text
    
102     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
104         sWebsite = sWebsite & "/"
        End If
          
106     RSDataPacks.Filter = adFilterPendingRecords
    
108     RSDataPacks.MoveFirst
        
110     With RSDataPacks

112         Do While Not .EOF

114             If IsNull(.fields.Item("GUID").Value) Then
116                 .fields.Item("GUID").Value = GUIDGen
118             ElseIf Len(.fields.Item("GUID").Value) < 5 Then
120                 .fields.Item("GUID").Value = GUIDGen
                End If

122             RSDataPacks.MoveNext
            Loop

        End With
    
124     MsXmlHttp.open "POST", sWebsite & "Oasis.asp", 0
126     RSDataPacks.Save MsXmlDoc, 1
        If frmDatabaseConnect.g_bProxyEnabled Then MsXmlHttp.setProxy 2, frmDatabaseConnect.g_sProxy
128     MsXmlHttp.send MsXmlDoc
    
130     If MsXmlHttp.responseText = "Data Updated" Then
132         MsgBox "Data Updates Succesfully on the server"
134         cmdCommand1_Click
        End If
    
136     Set MsXmlHttp = Nothing
138     Set MsXmlDoc = Nothing

        '<EhFooter>
        Exit Sub

cmdSET_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.cmdSET_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     txtServerURL.Text = WebSite 'GetSetting(App.EXEName, "Settings", "WebServerOASISSynch", "http://www.immap.org/")
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     SaveSetting App.EXEName, "Settings", "WebServerOASISSynch", txtServerURL.Text
        On Error Resume Next
102     RSDataPacks.Close
104     Set RSDataPacks = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
