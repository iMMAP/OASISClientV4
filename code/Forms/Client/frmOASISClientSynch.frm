VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmOASISClientSynch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Data Pack Manager"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9390
   Icon            =   "frmOASISClientSynch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin OASISClient.Downloader Downloader1 
      Left            =   765
      Top             =   3870
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin C1SizerLibCtl.C1Elastic c1Main 
      Height          =   5460
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9390
      _cx             =   16563
      _cy             =   9631
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
      _GridInfo       =   $"frmOASISClientSynch.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elNavs 
         Height          =   690
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4680
         Width           =   9210
         _cx             =   16245
         _cy             =   1217
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
         Begin VB.CommandButton cmdCommand1 
            Caption         =   "Get Data Pack Definitions"
            Height          =   510
            Left            =   6165
            TabIndex        =   5
            Top             =   180
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CommandButton cmdSET 
            Caption         =   "Save Data Pack Definitions"
            Enabled         =   0   'False
            Height          =   510
            Left            =   6345
            TabIndex        =   4
            Top             =   90
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CommandButton cmdDo 
            Caption         =   "Dowload..."
            Height          =   510
            Left            =   7785
            TabIndex        =   3
            Top             =   90
            Width           =   1410
         End
         Begin VB.TextBox txtServerURL 
            Height          =   285
            Left            =   45
            TabIndex        =   2
            Text            =   "http://www.immap.org/"
            Top             =   270
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label lblOASISServer 
            AutoSize        =   -1  'True
            Caption         =   "OASIS Server URL:"
            Height          =   195
            Left            =   90
            TabIndex        =   6
            Top             =   45
            Visible         =   0   'False
            Width           =   1410
         End
      End
      Begin C1SizerLibCtl.C1Tab C1TSynch 
         Height          =   4530
         Left            =   90
         TabIndex        =   7
         Top             =   90
         Width           =   9210
         _cx             =   16245
         _cy             =   7990
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
         Caption         =   "Settings|Available Files For Synch|Datasets|Profile|Available Data Packs"
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
            Height          =   4155
            Left            =   45
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   330
            Width           =   9120
            _cx             =   16087
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
               OleObjectBlob   =   "frmOASISClientSynch.frx":6895
               TabIndex        =   9
               Top             =   45
               Width           =   8970
            End
         End
         Begin C1SizerLibCtl.C1Elastic elProfile 
            Height          =   4155
            Left            =   -9765
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   330
            Width           =   9120
            _cx             =   16087
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
               OleObjectBlob   =   "frmOASISClientSynch.frx":753D
               TabIndex        =   11
               Top             =   90
               Width           =   4155
            End
         End
         Begin C1SizerLibCtl.C1Elastic elDatasets 
            Height          =   4155
            Left            =   -10065
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   330
            Width           =   9120
            _cx             =   16087
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
               OleObjectBlob   =   "frmOASISClientSynch.frx":81E5
               TabIndex        =   13
               Top             =   45
               Width           =   4290
            End
         End
         Begin C1SizerLibCtl.C1Elastic elFiles 
            Height          =   4155
            Left            =   -10365
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   330
            Width           =   9120
            _cx             =   16087
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
            Begin DXDBGRIDLibCtl.dxDBGrid dxSynchFolderURL 
               Height          =   2085
               Left            =   45
               OleObjectBlob   =   "frmOASISClientSynch.frx":8E8D
               TabIndex        =   17
               Top             =   1980
               Width           =   9015
            End
            Begin DXDBGRIDLibCtl.dxDBGrid dxDBSynch 
               Height          =   1815
               Left            =   90
               OleObjectBlob   =   "frmOASISClientSynch.frx":B46B
               TabIndex        =   18
               Top             =   45
               Width           =   8925
            End
         End
         Begin C1SizerLibCtl.C1Elastic elSettings 
            Height          =   4155
            Left            =   -10665
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   330
            Width           =   9120
            _cx             =   16087
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
               TabIndex        =   16
               Top             =   45
               Width           =   1950
            End
         End
      End
   End
End
Attribute VB_Name = "frmOASISClientSynch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private RSDataPacks As adodb.Recordset
Private WithEvents m_cUnzip As cUnzip
Attribute m_cUnzip.VB_VarHelpID = -1

Public Sub GetSynchFileFolders()
        '<EhHeader>
        On Error GoTo GetSynchFileFolders_Err
        '</EhHeader>
        Dim RSFileSynchURL As adodb.Recordset
        Dim sString As String

        Exit Sub

100     Set RSFileSynchURL = New adodb.Recordset

102     If Not Mid$(txtServerURL.Text, Len(txtServerURL.Text)) = "/" Then
104         txtServerURL.Text = txtServerURL.Text & "/"
        End If

        sString = txtServerURL.Text & "oasis.asp?getsf=" & CheckEncrypt(1) & "&us=" & CheckEncrypt(g_sUserName) & "&pwd=" & CheckEncrypt(g_sUserPass)
        Set RSFileSynchURL = OpenSilentHttpCommsRS(sString, True)
        
108     If RSFileSynchURL.State = adStateClosed Then Exit Sub
    
110     With dxDBSynch
112         Set .DataSource = RSFileSynchURL
114         .Columns.RetrieveFields
        End With

116     If Not RSFileSynchURL.Bof Then
118         SafeMoveFirst RSFileSynchURL
            
120         Do While Not RSFileSynchURL.EOF
122             RSFileSynchURL.Fields.Item("bUse").Value = False
124             RSFileSynchURL.MoveNext
            Loop
            
        End If

        '<EhFooter>
        Exit Sub

GetSynchFileFolders_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmOASISClientSynch.GetSynchFileFolders " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

'Private Sub GetSynchFiles()
'    Dim MsXmlHttp As Object 'New MSXML2.XMLHTTP
'    Dim MsXmlDoc As Object 'New MSXML2.DOMDocument
'
'    Set MsXmlHttp = CreateObject("MSXML2.XMLHTTP")
'    Set MsXmlDoc = CreateObject("MSXML2.DOMDocument")
'
'    MsXmlHttp.Open "GET", "http://www.immap.org/showfiles1.asp?folderConst=BCDE299F-6680-F04E-9533-E3806946960E", 0
'
'    MsXmlHttp.Send  'MsXmlDoc
'
'    If MsXmlHttp.responseText = "Data Updated" Then
'        MsgBox "Data Updates Succesfully on the server"
'
'    End If
'
'    Set MsXmlHttp = Nothing
'    Set MsXmlDoc = Nothing
'
'End Sub

Private Sub cmdCommand1_Click()
        '<EhHeader>
        On Error GoTo cmdCommand1_Click_Err
        '</EhHeader>
        Dim sWebsite As String
        Dim sString As String

100     Set RSDataPacks = New adodb.Recordset
    
102     sWebsite = txtServerURL.Text

104     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
106         sWebsite = sWebsite & "/"
        End If
        
        Stop
        ' CHECK THIS PETRI.....
        'this was:
        'sString = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & DataPacks")
        
        'corrected to:
        sString = sWebsite & "oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sUserName & "DataPacks")
        Set RSDataPacks = OpenSilentHttpCommsRS(sString, True)
        
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
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmOASISSynch.cmdCommand1_Click " & "at line " & Erl
        
        '</EhFooter>
End Sub

Public Sub ResetDownloadPossibilities()
    With dxDataPacks
        
       ' .Dataset.First
       ' Do While Not .Dataset.EOF
        '    .Dataset.FieldValues("Update") = False
        '    .Dataset.Next
       ' Loop
    
    End With
End Sub

Private Function StripFileName1(FilePath As String) As String

    Dim Path As Variant
    Path = Split(FilePath, "/")
    StripFileName1 = Path(UBound(Path))
End Function

Private Sub cmdDo_Click()
        '<EhHeader>
        On Error GoTo cmdDo_Click_Err
        '</EhHeader>
        Dim sWebsite As String
    
    
       ' If Not frmLog.Visible Then
       '     Me.Hide
       '     frmLog.Show  'vbModeless, Me
       ' End If
    
100     sWebsite = txtServerURL.Text

102     If Not Mid$(sWebsite, Len(sWebsite)) = "/" Then
104         sWebsite = sWebsite & "/"
        End If
        
106     With dxDataPacks
        
'108         m_frmDebug.debugprint .KeyField
        
110         .Dataset.First
        
112         Do While Not .Dataset.EOF

114             If .Dataset.FieldValues("Update") Then

116                 Downloader1.BeginDownload sWebsite & .Dataset.FieldValues("Path"), g_sAppPath & "\data\" & .Dataset.FieldValues("FolderConst") & "\" & StripFileName1(.Dataset.FieldValues("Path"))
                End If
            
118             .Dataset.Next
            Loop
    
        End With
                
        Me.Hide
                
        '<EhFooter>
        Exit Sub

cmdDo_Click_Err:
        Me.Hide
        frmLog.txtLog.Text = Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.cmdDo_Click " & _
               "at line " & Erl & vbCrLf & frmLog.txtLog.Text
        Resume Next
        '</EhFooter>
End Sub

Private Function pOpen(ByVal sFile As String) As Boolean
Dim i As Long
   
   ' Get the file directory:
   Set m_cUnzip = New cUnzip
   m_cUnzip.ZipFile = sFile
   m_cUnzip.Directory
   
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
   
106    sFolder = g_sAppPath
   
    '   If (sFolder <> "") Then
108       m_cUnzip.OverwriteExisting = True
110       m_cUnzip.ExtractOnlyNewer = True
112       m_cUnzip.UseFolderNames = True
114       m_cUnzip.UnzipFolder = g_sAppPath & "\data\" & dxDataPacks.Dataset.FieldValues("FolderConst")
          
116       m_cUnzip.Unzip
    '   End If
   
        '<EhFooter>
        Exit Sub

ExtractZip_Err:
        frmLog.txtLog.Text = Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.frmOASISSynch.ExtractZip " & _
               "at line " & Erl & vbCrLf & frmLog.txtLog.Text
        Resume Next
        '</EhFooter>
End Sub

Private Sub DoJunk()

End Sub

Private Sub Downloader1_DownloadComplete(MaxBytes As Long, _
                                         SaveFile As String)
        '<EhHeader>
        On Error GoTo Downloader1_DownloadComplete_Err
        '</EhHeader>
        Dim sName As String
        Dim sGUID As String
        Dim RSUpdater As adodb.Recordset
    
100     pOpen SaveFile
102     ExtractZip
     
104     frmLog.txtLog.Text = "OASIS Data Pack Manager: Completed " & SaveFile & ", Size = " & MaxBytes & vbCrLf & frmLog.txtLog.Text
        
        On Error Resume Next
        
        Kill SaveFile
        
        On Error GoTo Downloader1_DownloadComplete_Err
        
106     sName = StripFileName(SaveFile)
    
108     With dxDataPacks
                
110         .Dataset.First
        
112         Do While Not .Dataset.EOF

114             If StripFileName1(.Dataset.FieldValues("Path")) = sName Then
116                 sGUID = .Dataset.FieldValues("GUID")
118                 m_Cnn.Execute "INSERT INTO DataPacks ([GUID]) VALUES ('" & sGUID & "')"
                    Exit Do
                End If
            
120             .Dataset.Next
            Loop
    
        End With

        '

        '<EhFooter>
        Exit Sub

Downloader1_DownloadComplete_Err:
        frmLog.txtLog.Text = Err.Description & vbCrLf & _
               "in OASISClient.frmOASISClientSynch.Downloader1_DownloadComplete " & _
               "at line " & Erl & vbCrLf & frmLog.txtLog.Text
        Resume Next
        '</EhFooter>
End Sub

Private Sub Downloader1_DownloadProgress(CurBytes As Long, _
                                         MaxBytes As Long, _
                                         SaveFile As String)

    Dim i As Integer
    Dim RemBytes As Long

    RemBytes = MaxBytes - CurBytes
                
    If RemBytes < 2 ^ 20 Then
        frmLog.txtLog.Text = "OASIS Data Pack Manager: " & Format((MaxBytes - CurBytes) / 2 ^ 10, "#0.0 KB") & " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")" '& vbCrLf & frmLog.txtLog.Text
    Else
        frmLog.txtLog.Text = "OASIS Data Pack Manager: " & Format((MaxBytes - CurBytes) / 2 ^ 20, "#0.00 MB") & " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")" '& vbCrLf & frmLog.txtLog.Text
    End If

End Sub

Private Sub Downloader1_DownloadError(SaveFile As String)

    frmLog.txtLog.Text = "OASIS Data Pack Manager: " & "Error downloading " & SaveFile & vbCrLf & frmLog.txtLog.Text

End Sub

Private Sub cmdSET_Click()

'        Dim MsXmlHttp As New MSXML2.XMLHTTP
'        Dim MsXmlDoc As New MSXML2.DOMDocument
'        Dim sWebSite As String
'
'100     sWebSite = txtServerURL.Text
'
'        If Not Mid$(sWebSite, Len(sWebSite)) = "/" Then
'            sWebSite = sWebSite & "/"
'        End If
'
'
'
'102     RSDataPacks.Filter = adFilterPendingRecords
'
'104     MsXmlHttp.Open "POST", sWebSite & "setprofile.asp", 0
'106     RSDataPacks.Save MsXmlDoc, 1
'108     MsXmlHttp.Send MsXmlDoc
'
'
'110     If MsXmlHttp.responseText = "Data Updated" Then
'112         MsgBox "Data Updates Succesfully on the server"
'114         cmdCommand1_Click
'        End If
'
'116     Set MsXmlHttp = Nothing
'118     Set MsXmlDoc = Nothing
'        '<EhFooter>
                

End Sub

Private Sub dxDBSynch_OnChangeNode(ByVal OldNode As DXDBGRIDLibCtl.IdxGridNode, ByVal Node As DXDBGRIDLibCtl.IdxGridNode)
        '<EhHeader>
        On Error GoTo dxDBSynch_OnChangeNode_Err
        '</EhHeader>
100     m_frmDebug.DebugPrint Node.Strings(1) 'GUID
    
        Dim oRS As adodb.Recordset
        Dim sString As String

        sString = txtServerURL.Text & "oasis.asp?folderConst=" & CheckEncrypt(Node.Strings(1))
        Set oRS = OpenSilentHttpCommsRS(sString, True)

106     Set dxSynchFolderURL.DataSource = oRS
    
108     With dxDBSynch
    
110         .Columns.RetrieveFields
            
        End With

        '<EhFooter>
        Exit Sub

dxDBSynch_OnChangeNode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISClientSynch.dxDBSynch_OnChangeNode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
