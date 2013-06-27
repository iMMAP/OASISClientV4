VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OASISAttUploader 
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9660
   ScaleHeight     =   4140
   ScaleWidth      =   9660
   ToolboxBitmap   =   "ctrOASISAttUploader.ctx":0000
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   4140
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9660
      _cx             =   17039
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
      Align           =   5
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
      GridRows        =   3
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"ctrOASISAttUploader.ctx":0312
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elDecription 
         Height          =   1590
         Left            =   15
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2535
         Width           =   6795
         _cx             =   11986
         _cy             =   2805
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
         GridRows        =   5
         GridCols        =   5
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"ctrOASISAttUploader.ctx":036A
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.TextBox txtLog 
            Height          =   1215
            Left            =   3885
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   300
            Width           =   2910
         End
         Begin VB.TextBox txtDescription 
            Height          =   1215
            Left            =   0
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   300
            Width           =   3885
         End
         Begin VB.Label lblUploadLog 
            Caption         =   "Upload Log:"
            Height          =   240
            Left            =   3885
            TabIndex        =   13
            Top             =   60
            Width           =   1320
         End
         Begin VB.Label lblDescription 
            Caption         =   "Description:"
            Height          =   240
            Left            =   0
            TabIndex        =   10
            Top             =   60
            Width           =   3885
         End
      End
      Begin C1SizerLibCtl.C1Elastic elTools 
         Height          =   1590
         Left            =   6840
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2535
         Width           =   2805
         _cx             =   4948
         _cy             =   2805
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
         Begin VB.Frame FraUploadStatus 
            Caption         =   "Upload Progress:"
            Height          =   675
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   2475
            Begin MSComctlLib.ProgressBar ProgressBar1 
               Height          =   375
               Left            =   60
               TabIndex        =   12
               Top             =   240
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   661
               _Version        =   393216
               Appearance      =   1
            End
         End
         Begin VB.CommandButton cmdAddFile 
            Caption         =   "Add File"
            Height          =   315
            Left            =   1500
            TabIndex        =   7
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Remove All"
            Height          =   315
            Left            =   240
            TabIndex        =   6
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton AbortButton 
            Caption         =   "Abort"
            Height          =   315
            Left            =   1560
            TabIndex        =   5
            Top             =   1140
            Width           =   1215
         End
         Begin VB.CommandButton AsyncUpload 
            Caption         =   "Upload"
            Height          =   315
            Left            =   240
            TabIndex        =   4
            Top             =   1140
            Width           =   1215
         End
      End
      Begin C1SizerLibCtl.C1Elastic elFiles 
         Height          =   2490
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   15
         Width           =   9630
         _cx             =   16986
         _cy             =   4392
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
         _GridInfo       =   $"ctrOASISAttUploader.ctx":03ED
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin MSComctlLib.ListView lvFiles 
            Height          =   2460
            Left            =   15
            TabIndex        =   2
            Top             =   15
            Width           =   9600
            _ExtentX        =   16933
            _ExtentY        =   4339
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            Checkboxes      =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "File path"
               Object.Width           =   10584
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Title"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Category"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Size"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "OASISAttUploader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Winsock
Private intSock As Integer
Private strReceivedData As String
Private abortUpload As Long
Private WithEvents upload2 As ChilkatUpload
Attribute upload2.VB_VarHelpID = -1
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private colGUID As New Collection
Private m_LonIdleTimeOutMs As Long
Private m_LonChunkSize As Long
Private m_intPort As Integer
Private m_StrServerPath As String
Public Event UploadSuccess(bSuccess As Boolean)
Public Event PercentUploaded(Percentage As Integer)
Public Event LogUpdated()
Private m_StrRecordGUID As String
Private m_StrHostName As String
Private m_StrUploadLog As String
Private m_StrUserGroup As String
Private m_StrTableName As String
Private m_bUseBatchMode As Boolean
Private m_StrProxyDomain As String
Private m_StrProxyPassword As String
Private m_StrProxyUsername As String
Private m_intProxyPort As Integer
Private m_bUseProxy As Boolean
Private m_Strsubdomainname As String

Public Property Get subdomainname() As String
    subdomainname = m_Strsubdomainname
End Property

Public Property Let subdomainname(ByVal StrValue As String)
    m_Strsubdomainname = StrValue
End Property

Public Property Get UseProxy() As Boolean
    UseProxy = m_bUseProxy
End Property

Public Property Let UseProxy(ByVal bValue As Boolean)
    m_bUseProxy = bValue
End Property

Public Property Get ProxyPort() As Integer
    ProxyPort = m_intProxyPort
End Property

Public Property Let ProxyPort(ByVal intValue As Integer)
    m_intProxyPort = intValue
End Property

Public Property Get ProxyUsername() As String
    ProxyUsername = m_StrProxyUsername
End Property

Public Property Let ProxyUsername(ByVal StrValue As String)
    m_StrProxyUsername = StrValue
End Property

Public Property Get ProxyPassword() As String
    ProxyPassword = m_StrProxyPassword
End Property

Public Property Let ProxyPassword(ByVal StrValue As String)
    m_StrProxyPassword = StrValue
End Property

Public Property Get ProxyDomain() As String
    ProxyDomain = m_StrProxyDomain
End Property

Public Property Let ProxyDomain(ByVal StrValue As String)
    m_StrProxyDomain = StrValue
End Property

Public Property Get UseBatchMode() As Boolean
    UseBatchMode = m_bUseBatchMode
End Property

Public Property Let UseBatchMode(ByVal bValue As Boolean)
    m_bUseBatchMode = bValue
End Property

Public Property Get TableName() As String
    TableName = m_StrTableName
End Property

Public Property Let TableName(ByVal StrValue As String)
    m_StrTableName = StrValue
End Property

Public Property Get UserGroup() As String
    UserGroup = m_StrUserGroup
End Property

Public Property Let UserGroup(ByVal StrValue As String)
    m_StrUserGroup = StrValue
End Property

Public Property Get UploadLog() As String
    UploadLog = m_StrUploadLog
End Property

Public Property Let UploadLog(ByVal StrValue As String)
    m_StrUploadLog = StrValue
End Property

Public Property Get hostname() As String
    hostname = m_StrHostName
End Property

Public Property Let hostname(ByVal StrValue As String)
    m_StrHostName = StrValue
End Property

Public Property Get RecordGUID() As String
    RecordGUID = m_StrRecordGUID
End Property

Public Property Let RecordGUID(ByVal StrValue As String)
    m_StrRecordGUID = StrValue
End Property

Public Property Get ServerPath() As String
    ServerPath = m_StrServerPath
End Property

Public Property Let ServerPath(ByVal StrValue As String)
    m_StrServerPath = StrValue
End Property

Public Property Get Port() As Integer
    Port = m_intPort
End Property

Public Property Let Port(ByVal intValue As Integer)
    m_intPort = intValue
End Property

Public Property Get ChunkSize() As Long
    ChunkSize = m_LonChunkSize
End Property

Public Property Let ChunkSize(ByVal LonValue As Long)
    m_LonChunkSize = LonValue
End Property

Public Property Get IdleTimeOutMs() As Long
    IdleTimeOutMs = m_LonIdleTimeOutMs
End Property

Public Property Let IdleTimeOutMs(ByVal LonValue As Long)
    m_LonIdleTimeOutMs = LonValue
End Property

Private Sub ProxySettings(upload2 As ChilkatUpload)

    With upload2
        .ProxyDomain = m_StrProxyDomain
        .ProxyLogin = m_StrProxyUsername
        .ProxyPassword = m_StrProxyPassword
        .ProxyPort = m_intProxyPort
    End With

End Sub

Private Sub DoIT()
    Dim i As Integer
    
    abortUpload = 0
    Set upload2 = New ChilkatUpload
    upload2.IdleTimeOutMs = m_LonIdleTimeOutMs
    upload2.ChunkSize = m_LonChunkSize
    
    If m_bUseProxy Then ProxySettings upload2
    
    upload2.hostname = m_StrHostName
    upload2.Port = m_intPort
    upload2.Path = m_StrServerPath & "upload.php"
    
    For i = 1 To lvFiles.ListItems.Count

        If lvFiles.ListItems.Item(i).Checked Then
            upload2.AddFileReference "upload[]", lvFiles.ListItems.Item(i).Text
            upload2.AddParam "title[]", lvFiles.ListItems.Item(i).SubItems(1)
            upload2.AddParam "filecat[]", lvFiles.ListItems.Item(i).SubItems(2)
        End If

    Next
  
    'upload2.AddFileReference "upload[]", lvFiles.ListItems.Item(lvFiles.ListItems.Count).Text
    'upload2.AddParam "title[]", lvFiles.ListItems.Item(lvFiles.ListItems.Count).SubItems(1)
    'upload2.AddParam "filecat[]", lvFiles.ListItems.Item(lvFiles.ListItems.Count).SubItems(2)
    upload2.AddParam "extid", m_StrRecordGUID
    upload2.AddParam "brief", txtDescription.Text
    upload2.AddParam "usergroup", m_StrUserGroup
    upload2.AddParam "tablename", m_StrTableName
    upload2.AddParam "subdomain", m_Strsubdomainname
    upload2.AddParam "oasis_hidden", "335ffbcfac66fb164e9d5a54505cad8f"
    upload2.BeginUpload
    
    Do
        upload2.SleepMs 100
        ProgressBar1.value = upload2.PercentUploaded
        RaiseEvent PercentUploaded(CInt(upload2.PercentUploaded))
        DoEvents
        
        m_StrUploadLog = m_StrUploadLog & vbCrLf & " Upload Progress: " & upload2.PercentUploaded
        
        If (abortUpload = 1) Then
            upload2.abortUpload
        End If
        
    Loop Until upload2.UploadInProgress = 0
    
    ProgressBar1.value = 0
    
    If (upload2.UploadSuccess = 1) Then
        RaiseEvent UploadSuccess(True)
        If (upload2.ResponseStatus <> 200) Then
            upload2.SaveLastError "uploadError.xml"
            m_StrUploadLog = m_StrUploadLog & vbCrLf & "Failed to upload, HTTP response code = " + Str(upload2.ResponseStatus)
        Else
            m_StrUploadLog = m_StrUploadLog & vbCrLf & upload2.ResponseHeader
            Dim html As String
            html = upload2.VariantToString(upload2.responseBody, "iso-8859-1")
            m_StrUploadLog = m_StrUploadLog & vbCrLf & html

            If (InStr(html, "Upload completed") > 0) Then
                m_StrUploadLog = m_StrUploadLog & vbCrLf & "Upload complete!"
            Else
                ' An unexpected response was received.  Display the HTML source:
                m_StrUploadLog = m_StrUploadLog & vbCrLf & "Unexpected HTML response!"
            End If
            
        End If
        
    Else
        RaiseEvent UploadSuccess(False)
        m_StrUploadLog = m_StrUploadLog & vbCrLf & upload2.LastErrorText
        upload2.SaveLastError "uploadError.xml"
    End If
    
    If Len(html) > 0 Then
        'If Not WebBrowser.Document.Body Is Nothing Then WebBrowser.Document.Body.InnerHTML = html
    End If
    
    m_StrUploadLog = m_StrUploadLog & vbCrLf & upload2.ResponseHeader
    ClearAll

End Sub

Private Sub UploadWorker()
Dim html As String
Dim i As Integer

    m_StrUploadLog = m_StrUploadLog & vbCrLf & "Total Number of files:" & lvFiles.ListItems.Count
    
    For i = 1 To lvFiles.ListItems.Count
        If lvFiles.ListItems.Item(i).Checked Then
            m_StrUploadLog = m_StrUploadLog & vbCrLf & "Processing file #" & i
            html = Berserk(lvFiles.ListItems.Item(i).Text, lvFiles.ListItems.Item(i).SubItems(1), lvFiles.ListItems.Item(i).SubItems(2), m_StrRecordGUID)
            m_StrUploadLog = m_StrUploadLog & vbCrLf & "End Processing file #" & i
        End If
    Next

   ' If Len(html) > 0 Then
   '     If Not WebBrowser.Document.Body Is Nothing Then WebBrowser.Document.Body.InnerHTML = html
   '     m_StrUploadLog = m_StrUploadLog & vbCrLf & html
   ' End If
    
    ClearAll
    
End Sub

Private Sub ClearAll()
    'keith: petri, i think you have not checked in all code
    'URLEncode ""
    cmdClear_Click
'   ' txtGUID.Text = GetGUID
    txtDescription.Text = ""
    ProgressBar1.value = 0
'    lblProgress.caption = ""
End Sub

Private Function Berserk(sFile As String, _
                         sTitle As String, _
                         sfilecat As String, _
                         sGUID As String) As String
    Dim i As Integer
    
    abortUpload = 0
    
    Set upload2 = New ChilkatUpload
        
    upload2.IdleTimeOutMs = m_LonIdleTimeOutMs
    upload2.ChunkSize = m_LonChunkSize
    If m_bUseProxy Then ProxySettings upload2
    upload2.hostname = m_StrHostName
    upload2.Port = m_intPort
    upload2.Path = m_StrServerPath & "upload.php"

    upload2.AddFileReference "upload[]", sFile
    upload2.AddParam "title[]", sTitle
    upload2.AddParam "filecat[]", sfilecat
    'upload2.AddFileReference "upload[]", sFile
    'upload2.AddParam "title[]", sTitle
    'upload2.AddParam "filecat[]", sfilecat
    upload2.AddParam "usergroup", m_StrUserGroup
    upload2.AddParam "tablename", m_StrTableName
    upload2.AddParam "subdomain", m_Strsubdomainname
    upload2.AddParam "oasis_hidden", "335ffbcfac66fb164e9d5a54505cad8f"
    
    upload2.AddParam "extid", sGUID
    upload2.AddParam "brief", txtDescription.Text
    m_StrUploadLog = m_StrUploadLog & vbCrLf & "Total Upload Size:" & upload2.TotalUploadSize
    upload2.BeginUpload

    Do
        upload2.SleepMs 100

        DoEvents
        RaiseEvent PercentUploaded(CInt(upload2.PercentUploaded))
        m_StrUploadLog = m_StrUploadLog & vbCrLf & " Upload Progress: " & upload2.PercentUploaded

        If (abortUpload = 1) Then
            upload2.abortUpload
        End If
        
    Loop Until upload2.UploadInProgress = 0
        
    If (upload2.UploadSuccess = 1) Then
        RaiseEvent UploadSuccess(True)

        If (upload2.ResponseStatus <> 200) Then
        
            upload2.SaveLastError "uploadError.xml"
            
            m_StrUploadLog = m_StrUploadLog & vbCrLf & "Failed to upload, HTTP response code = " + Str(upload2.ResponseStatus)
            
        Else

            m_StrUploadLog = m_StrUploadLog & vbCrLf & upload2.ResponseHeader
            
            Dim html As String
            html = upload2.VariantToString(upload2.responseBody, "iso-8859-1")
    
        End If
        
    Else
        m_StrUploadLog = m_StrUploadLog & vbCrLf & upload2.LastErrorText
        RaiseEvent UploadSuccess(False)
        upload2.SaveLastError "uploadError.xml"
    End If
    
    m_StrUploadLog = m_StrUploadLog & vbCrLf & upload2.ResponseHeader

    Berserk = html

End Function

Private Sub AbortButton_Click()
    abortUpload = 1
End Sub

Private Sub AsyncUpload_Click()
    
    m_StrUploadLog = ""
    
    If Not m_bUseBatchMode Then
        UploadWorker
    Else
        DoIT
    End If

End Sub

Private Sub cmdAddFile_Click()
Dim sFile As String
Dim Title As String

    sFile = OpenDialog("All Files (*.*)|*.*|Text (*.txt)|*.txt|JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|CompuServe GIF (*.gif)|*.gif|AVI (*.avi)|*.avi|MPEG (*.mpg;*.mpeg)|*.mpg;*.mpeg|WMV (*.wmv)|*.wmv|QuickTime (*.mov)|*.mov|Flash (*.swf)|*.swf", "Select a file...", "*.*", App.Path, UserControl.Parent.hwnd)
    Title = GetFileFromPath(sFile)
    
    frmFileCategories.txtTitle.Text = Title
    frmFileCategories.Show vbModal, Me
    
    'txtParameter.Text = g_sFileTitle
    
    AddListItem lvFiles, sFile, Title, "Documents"

End Sub

Function AddListItem(lW As ListView, _
                            fname As String, _
                            sAlias As String, scategory As String)
        '<EhHeader>
        On Error GoTo AddListItem_Err
        '</EhHeader>
    
    
        Dim LVI As ListItem
    
100     If fname = "" Then Exit Function

102     Set LVI = lW.ListItems.Add

104     m_StrUploadLog = m_StrUploadLog & vbCrLf & "********Adding File To List*********"

106     LVI.Text = fname
108     m_StrUploadLog = m_StrUploadLog & vbCrLf & fname
        
110     LVI.SubItems(1) = sAlias
        
112     m_StrUploadLog = m_StrUploadLog & vbCrLf & sAlias
114     LVI.SubItems(2) = scategory
116     LVI.SubItems(3) = FileLen(fname) & " bytes"

118     m_StrUploadLog = m_StrUploadLog & vbCrLf & LVI.SubItems(2)
        
120     LVI.Checked = True
    
122     m_StrUploadLog = m_StrUploadLog & vbCrLf & "******End Adding File To List*******"
    
        '<EhFooter>
        Exit Function

AddListItem_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISAttUploader.AddListItem", _
                  "OASISAttUploader component failure"
        '</EhFooter>
End Function


Private Function GetFileFromPath(strPathWithFile As String) As String
On Error GoTo Error

    If Not right$(strPathWithFile, 1) = "\" Then
        strPathWithFile = Replace$(strPathWithFile, "/", "\")
        
        If InStrB(1, strPathWithFile, "\") Then
            GetFileFromPath = Mid$(strPathWithFile, InStrRev(strPathWithFile, "\") + 1)
        Else
            GetFileFromPath = strPathWithFile
        End If
    End If
    
    Exit Function
Error:
    GetFileFromPath = ""
End Function


Private Sub cmdClear_Click()
    lvFiles.ListItems.Clear
End Sub

Private Sub UserControl_Initialize()
    m_LonIdleTimeOutMs = 30000
    m_LonChunkSize = 2048
    m_intPort = 80
    m_bUseBatchMode = True
    m_bUseProxy = False
End Sub
