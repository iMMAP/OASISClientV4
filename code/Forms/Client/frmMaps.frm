VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMaps 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
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
   Begin VB.PictureBox PicSrc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   855
      Left            =   2220
      ScaleHeight     =   795
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   4860
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picThumb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   1152
      Left            =   3300
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   0
      Top             =   4860
      Visible         =   0   'False
      Width           =   1536
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   1620
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView flb 
      Height          =   4125
      Left            =   90
      TabIndex        =   2
      Top             =   135
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   7276
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMaps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const MAX_PATH = 260

Private Type FILETIME
       dwLowDateTime As Long
       dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
       dwFileAttributes As Long
       ftCreationTime As FILETIME
       ftLastAccessTime As FILETIME
       ftLastWriteTime As FILETIME
       nFileSizeHigh As Long
       nFileSizeLow As Long
       dwReserved0 As Long
       dwReserved1 As Long
       cFileName As String * MAX_PATH
       cAlternate As String * 14
End Type

Dim EnablePreview As Boolean
Dim Filename As String
Dim INIPath As String
Dim lstFilesFocus As Boolean

Dim flbList As New Collection

Private Sub dlb_Change()
        '<EhHeader>
        On Error GoTo dlb_Change_Err
        '</EhHeader>
        Dim i As Long
        Dim FN As String
        Dim hHeight As Double, hWidth As Double
    
100     For i = flbList.Count To 1 Step -1
102         flbList.Remove (i)
        Next
    
104     flb.Icons = Nothing
106     ImgList.ListImages.Clear
    
108     flb.ListItems.Clear
110     flb.Refresh
    
112     If g_RSAppSettings.State = adStateClosed Then Exit Sub
    
114     SafeMoveFirst g_RSAppSettings
116     g_RSAppSettings.Find "SettingName = 'MapPreview'"
    
118     GetFiles g_sAppPath & g_RSAppSettings.Fields.Item("SettingValue1").Value
    
120     For i = flbList.Count To 1 Step -1
122         FN = LCase$(Right$(flbList.Item(i), 3))
124         If FN <> "jpg" And FN <> "bmp" And FN <> "cur" And FN <> "ico" Then
126             flbList.Remove (i)
            End If
        Next
    
128     For i = 1 To flbList.Count
130         PicSrc.Picture = LoadPicture(flbList(i))
        
132         hWidth = PicSrc.Width
134         hHeight = PicSrc.Height
        
136         If hHeight > 76.8 Then
138             hWidth = 76.8 * PicSrc.Width / PicSrc.Height
140             hHeight = 76.8
            End If
        
142         If hWidth > 102.4 Then
144             hHeight = 102.4 * PicSrc.Height / PicSrc.Width
146             hWidth = 102.4
            End If
        
148         picThumb.PaintPicture PicSrc, (picThumb.Width - hWidth) / 2, (picThumb.Height - hHeight) / 2, hWidth, hHeight
150         ImgList.ListImages.Add , , picThumb.Image
152         If flb.Icons Is Nothing Then flb.Icons = ImgList
154         flb.ListItems.Add , , GetFileName(flbList(i)), i
        
156         picThumb.Cls
        
158         caption = "GENERATING PREVIEWS  " & Format(Round(i / flbList.Count * 100, 2), "###.00") & "%"
        Next
    
160     flb.Arrange = lvwAutoTop
       ' lblInfo.caption = flb.ListItems.Count & " items in list"
        'caption = "Image Browser"
        '<EhFooter>
        Exit Sub

dlb_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMaps.dlb_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GetFiles(Path As String)
        '<EhHeader>
        On Error GoTo GetFiles_Err
        '</EhHeader>
       Dim WFD As WIN32_FIND_DATA
       Dim hFile As Long, fPath As String, fname As String
       Dim colFiles As Collection
       Dim varFile As Variant
   
100    fPath = AddBackslash(Path)
102    fname = fPath & "*.*"
104    Set colFiles = New Collection
   
106    hFile = FindFirstFile(fname, WFD)
108    If (hFile > 0) And ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
110        colFiles.Add fPath & StripNulls(WFD.cFileName)
       End If
   
112    While FindNextFile(hFile, WFD)
114        If ((WFD.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
116            colFiles.Add fPath & StripNulls(WFD.cFileName)
           End If
       Wend
   
118    FindClose hFile
   
120    For Each varFile In colFiles
122        flbList.Add varFile
       Next
124    Set colFiles = Nothing
        '<EhFooter>
        Exit Sub

GetFiles_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMaps.GetFiles " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function StripNulls(f As String) As String
   StripNulls = Left$(f, InStr(1, f, Chr$(0)) - 1)
End Function

Private Function AddBackslash(s As String) As String
   If Len(s) Then
      If Right$(s, 1) <> "\" Then
         AddBackslash = s & "\"
      Else
         AddBackslash = s
      End If
   Else
      AddBackslash = "\"
   End If
End Function

Private Function GetFileName(File As String) As String
    Dim i As Integer
    For i = Len(File) To 1 Step -1
        If Mid$(File, i, 1) = "\" Then
            i = i + 1
            Exit For
        End If
    Next
    
    GetFileName = Mid$(File, i)
End Function
Private Sub Form_Load()
    dlb_Change
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
End Sub
