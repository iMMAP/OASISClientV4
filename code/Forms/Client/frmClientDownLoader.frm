VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDownLoader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8115
   Icon            =   "frmClientDownLoader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin OASISClient.Downloader Downloader1 
      Left            =   2655
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.ImageList SmallImages 
      Left            =   600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientDownLoader.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   6825
      TabIndex        =   2
      Top             =   2745
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5505
      TabIndex        =   1
      Top             =   2745
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstDownload 
      Height          =   2655
      Left            =   0
      TabIndex        =   3
      Top             =   45
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "SmallImages"
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Product"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   2716
      EndProperty
   End
   Begin VB.CommandButton cmdDownload 
      Caption         =   "Download"
      Height          =   375
      Left            =   4185
      TabIndex        =   0
      Top             =   2745
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmDownLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Event DownloadReady()

Private m_ColFiles As New Collection
Private m_ColFilesPath As New Collection

Private Sub cmdCancel_Click()

    cmdDownload.Enabled = True
    cmdCancel.Enabled = False
    Me.Downloader1.CancelAllDownload

End Sub

Public Sub ClearPrevs()
    Downloader1.ResetPreviousFiles
    Set m_ColFiles = New Collection
    Set m_ColFilesPath = New Collection
End Sub

Public Sub AddDownLoads()

End Sub

Private Sub cmdDownload_Click()

  Dim sURL As String
  Dim sFilename As String
  Dim sDescription As String

    cmdCancel.Enabled = True
    cmdDownload.Enabled = False
   
    lstDownload.ListItems.Clear

'    sURL = "http://download.microsoft.com/download/ie6sp1/finrel/6_sp1/W98NT42KMeXP/EN-US/ie6setup.exe"
'    sFileName = "ie6setup.exe"
'    sDescription = "Internet Explorer 6 Setup File"
'    Download_File sURL, g_sAppPath & "\" & sFileName, sDescription
'
'    sURL = "http://optusnet.dl.sourceforge.net/sourceforge/vnc-tight/tightvnc-1.2.9-setup.exe"
'    sFileName = "tightvnc-1.2.9-setup.exe"
'    sDescription = "TightVNC 1.2.9 Setup File"
'    Download_File sURL, g_sAppPath & "\" & sFileName, sDescription
'
'    sURL = "http://www.planet-source-code.com/vb/images/PscLogo1.jpg"
'    sFileName = "PscLogo1.jpg"
'    sDescription = "Planet Source Code image"
'    Download_File sURL, g_sAppPath & "\" & sFileName, sDescription
'
'    sURL = "http://www.winzip.com/index.htm"
'    sFileName = "index.html"
'    sDescription = "Winzip homepage"
'    Download_File sURL, g_sAppPath & "\" & sFileName, sDescription
'
'    sURL = "http://download.winzip.com/winzip90.exeX" '<--- do that on purpose to show error
'    sFileName = "winzip90.exe"
'    sDescription = "Winzip 9.0"
'    Download_File sURL, g_sAppPath & "\" & sFileName, sDescription

End Sub

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub Downloader1_DownloadAllComplete(FileNotDownload() As String)
        '<EhHeader>
        On Error GoTo Downloader1_DownloadAllComplete_Err
        '</EhHeader>

      Dim i As Integer
    Dim strDLError As String
    Dim colFiles As Collection

100     DebugPrint "Finished all download"
102     cmdDownload.Enabled = True
104     cmdCancel.Enabled = False

106     If UBound(FileNotDownload) > 0 Then
108         For i = 1 To UBound(FileNotDownload)
110             DebugPrint "File not downloaded: " & FileNotDownload(i)
112             strDLError = strDLError & "File not downloaded: " & FileNotDownload(i) & vbCrLf
114         Next i
        End If

116     If Not strDLError = "" Then
118         Set colFiles = Downloader1.DownloadedFilesAvailable
        
120         If Not colFiles Is Nothing Then
122             For i = 1 To colFiles.Count
                    On Error Resume Next
                    Kill m_ColFilesPath
124                 CopyFile m_ColFiles.Item(i), m_ColFilesPath.Item(i)
                Next
            End If
        Else
            MsgBox "The Update was Not Finalized due to the following files updated were not completed:" & vbCrLf & strDLError
        End If
    
    

        '<EhFooter>
        Exit Sub

Downloader1_DownloadAllComplete_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDownLoader.Downloader1_DownloadAllComplete " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()

    Me.caption = App.Title & " v" & App.major & "." & App.minor & " Build " & App.Revision
    DebugPrint ""

    lstDownload.ListItems.Clear
    
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If

End Sub

Private Sub Downloader1_DownloadComplete(MaxBytes As Long, SaveFile As String)

  Dim i As Integer

    DebugPrint "Completed " & SaveFile & ", Size = " & MaxBytes

    With lstDownload
        For i = 1 To .ListItems.Count
            If .ListItems(i).Key = SaveFile Then
                .ListItems(i).SubItems(1) = "Completed"
            End If
        Next i
    End With

End Sub

Private Sub Downloader1_DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)

  Dim i As Integer
  Dim RemBytes As Long

    With lstDownload
        For i = 1 To .ListItems.Count
            If .ListItems(i).Key = SaveFile Then
                RemBytes = MaxBytes - CurBytes
                If RemBytes < 2 ^ 20 Then
                    .ListItems(i).SubItems(1) = Format((MaxBytes - CurBytes) / 2 ^ 10, "#0.0 KB") & _
                               " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")"
                  Else
                    .ListItems(i).SubItems(1) = Format((MaxBytes - CurBytes) / 2 ^ 20, "#0.00 MB") & _
                               " (" & Format(CurBytes / MaxBytes, "#0.0%") & ")"
                End If
            End If
        Next i

    End With

End Sub

Private Sub Downloader1_DownloadError(SaveFile As String)

  Dim i As Integer

    DebugPrint "Error downloading " & SaveFile

    With lstDownload
        For i = 1 To .ListItems.Count
            If .ListItems(i).Key = SaveFile Then
                .ListItems(i).SubItems(1) = "Error"
            End If
        Next i

    End With

End Sub

Public Function Download_File(URL As String, SaveFile As String, Description As String, sCopypath As String)

    m_ColFiles.Add SaveFile
    m_ColFilesPath.Add sCopypath
    

    lstDownload.ListItems.Add , SaveFile, Description, , 1

    Me.Downloader1.BeginDownload URL, SaveFile

End Function

Private Function GetFileName(URL As String) As String

  Dim i As Integer

    For i = Len(URL) To 1 Step -1
        If Mid(URL, i, 1) = "/" Then
            GetFileName = Right(URL, Len(URL) - i)
            Exit For
        End If
    Next i

End Function
