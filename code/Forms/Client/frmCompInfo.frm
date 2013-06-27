VERSION 5.00
Begin VB.Form frmComputerInfo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8460
   Icon            =   "frmCompInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmMouse 
      Height          =   1320
      Left            =   4650
      TabIndex        =   43
      Top             =   3495
      Width           =   3735
      Begin VB.Label lblButtons 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1680
         TabIndex        =   49
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblMouseManuf 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1680
         TabIndex        =   48
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lblMouseDriver 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1680
         TabIndex        =   47
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Number of buttons:"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Mouse manufacturer:"
         Height          =   195
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   1500
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Mouse driver:"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame frmMonitor 
      Height          =   1680
      Left            =   120
      TabIndex        =   34
      Top             =   3135
      Width           =   4455
      Begin VB.Label lblDispManuf 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1680
         TabIndex        =   42
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label lblDispDriver 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1680
         TabIndex        =   41
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblColors 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1680
         TabIndex        =   40
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lblResolution 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1680
         TabIndex        =   39
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Driver manufacturer:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   1320
         Width           =   1440
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Display driver:"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Color depth:"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Screen resolution:"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.Frame frmTime 
      Height          =   975
      Left            =   4650
      TabIndex        =   21
      Top             =   0
      Width           =   3735
      Begin VB.Timer tmrTime 
         Interval        =   500
         Left            =   3480
         Top             =   1560
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   600
         TabIndex        =   25
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   600
         TabIndex        =   24
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   390
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Time:"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Frame frmDisks 
      Height          =   2535
      Left            =   4650
      TabIndex        =   1
      Top             =   960
      Width           =   3735
      Begin VB.ComboBox cmbDrives 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblFreeSpace 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   20
         Top             =   2160
         Width           =   45
      End
      Begin VB.Label lblTotalSpace 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   19
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Free diskspace:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total diskspace:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1170
      End
      Begin VB.Label lblSerial 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   1440
         Width           =   45
      End
      Begin VB.Label lblFSN 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   15
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serial Number:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "File System Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label lblDriveType 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Drive Type:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblDriveLabel 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1080
         TabIndex        =   10
         Top             =   240
         Width           =   45
      End
   End
   Begin VB.Frame frmGeneralInfo 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Memory used:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   2760
         Width           =   990
      End
      Begin VB.Label lblResources 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   32
         Top             =   2760
         Width           =   45
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Total RAM memory:"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   1395
      End
      Begin VB.Label lblTotRAM 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   30
         Top             =   2400
         Width           =   45
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   29
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblPlatform 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   28
         Top             =   600
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Windows Version:"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1275
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Windows Platform:"
         Height          =   195
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1320
      End
      Begin VB.Label lblTemp 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   9
         Top             =   2040
         Width           =   45
      End
      Begin VB.Label lblSystem 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label lblWindows 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   45
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Temporary path:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "System directory:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Windows directory:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1350
      End
      Begin VB.Label lblLogged 
         AutoSize        =   -1  'True
         Caption         =   "Logged in under ... on ..."
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmComputerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const REG_SZ = 1 ' Unicode nul terminated string
Const REG_BINARY = 3 ' Free form binary
Const HKEY_LOCAL_MACHINE = &H80000002
Const BITSPIXEL = 12
Const SM_CMOUSEBUTTONS = 43
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Declare Sub GlobalMemoryStatus _
                Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetDriveType _
                Lib "kernel32" _
                Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetVolumeInformation _
                Lib "kernel32" _
                Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
                                               ByVal lpVolumeNameBuffer As String, _
                                               ByVal nVolumeNameSize As Long, _
                                               lpVolumeSerialNumber As Long, _
                                               lpMaximumComponentLength As Long, _
                                               lpFileSystemFlags As Long, _
                                               ByVal lpFileSystemNameBuffer As String, _
                                               ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetLogicalDrives _
                Lib "kernel32" () As Long
Private Declare Function GetDiskFreeSpaceEx _
                Lib "kernel32" _
                Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
                                             lpFreeBytesAvailableToCaller As Currency, _
                                             lpTotalNumberOfBytes As Currency, _
                                             lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetWindowsDirectory _
                Lib "kernel32" _
                Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
                                              ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory _
                Lib "kernel32" _
                Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, _
                                             ByVal nSize As Long) As Long
Private Declare Function GetTempPath _
                Lib "kernel32" _
                Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                      ByVal lpBuffer As String) As Long
Private Declare Function GetUserName _
                Lib "advapi32.dll" _
                Alias "GetUserNameA" (ByVal lpBuffer As String, _
                                      nSize As Long) As Long
Private Declare Function GetComputerName _
                Lib "kernel32" _
                Alias "GetComputerNameA" (ByVal lpBuffer As String, _
                                          nSize As Long) As Long
Private Declare Function GetDiskFreeSpace _
                Lib "kernel32" _
                Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
                                           lpSectorsPerCluster As Long, _
                                           lpBytesPerSector As Long, _
                                           lpNumberOfFreeClusters As Long, _
                                           lpTtoalNumberOfClusters As Long) As Long
Private Declare Function LoadLibrary _
                Lib "kernel32" _
                Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary _
                Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress _
                Lib "kernel32" (ByVal hModule As Long, _
                                ByVal lpProcName As String) As Long
Private Declare Sub GetLocalTime _
                Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetVersionEx _
                Lib "kernel32" _
                Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RegCloseKey _
                Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyA" (ByVal hKey As Long, _
                                     ByVal lpSubKey As String, _
                                     phkResult As Long) As Long
Private Declare Function RegQueryValueEx _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          lpData As Any, _
                                          lpcbData As Long) As Long
Private Declare Function RegEnumKeyEx _
                Lib "advapi32.dll" _
                Alias "RegEnumKeyExA" (ByVal hKey As Long, _
                                       ByVal dwIndex As Long, _
                                       ByVal lpName As String, _
                                       lpcbName As Long, _
                                       ByVal lpReserved As Long, _
                                       ByVal lpClass As String, _
                                       lpcbClass As Long, _
                                       lpftLastWriteTime As FILETIME) As Long
Private Declare Function GetDeviceCaps _
                Lib "gdi32" (ByVal hdc As Long, _
                             ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics _
                Lib "user32" (ByVal nIndex As Long) As Long
Dim sTemp As String, nDisks As Long, Cnt As Long, DiskFreeExAvailable As Boolean

Private Sub cmbDrives_Click()
        '<EhHeader>
        On Error GoTo cmbDrives_Click_Err
        '</EhHeader>
        Dim sVolume As String, nSerial As Long, sFSN As String, TotalNumberOfClusters As Long
        Dim BytesFreeToCalller As Currency, TotalBytes As Currency, TotalFreeBytes As Currency
        Dim SectorsPerCluster As Long, BytesPerSector As Long, NumberOfFreeClusters As Long
100     sVolume = String(255, 0)
102     sFSN = String(255, 0)
        'Get information about the file system and volume
104     GetVolumeInformation cmbDrives.List(cmbDrives.ListIndex), sVolume, 255, nSerial, 0, 0, sFSN, 255

        'determine the type of the selected drive
106     Select Case GetDriveType(cmbDrives.List(cmbDrives.ListIndex))

            Case 2
108             lblDriveType.caption = "Removable"

110         Case 3
112             lblDriveType.caption = "Drive Fixed"

114         Case Is = 4
116             lblDriveType.caption = "Remote (network) drive"

118         Case Is = 5
120             lblDriveType.caption = "Cd-Rom"

122         Case Is = 6
124             lblDriveType.caption = "Ram disk"

126         Case Else
128             lblDriveType.caption = "Unrecognized"
        End Select

130     If nSerial = 0 Then
            'If the serial number is equal to 0, there's no information available
132         lblDriveLabel.caption = "NA"
134         lblFSN.caption = "NA"
136         lblSerial.caption = "NA"
138         lblTotalSpace.caption = "NA"
140         lblFreeSpace.caption = "NA"
        Else
142         lblDriveLabel.caption = sVolume
144         lblFSN = sFSN
146         lblSerial = Trim(str(nSerial))

            'if the GetDiskFreeSpaceEx-function is available, use it
148         If DiskFreeExAvailable Then
150             GetDiskFreeSpaceEx cmbDrives.List(cmbDrives.ListIndex), BytesFreeToCalller, TotalBytes, TotalFreeBytes
152             lblTotalSpace.caption = Format$(TotalBytes * 10000, "###,###,###,##0") + " bytes"
154             lblFreeSpace.caption = Format$(TotalFreeBytes * 10000, "###,###,###,##0") + " bytes"
            Else
156             GetDiskFreeSpace cmbDrives.List(cmbDrives.ListIndex), SectorsPerCluster, BytesPerSector, NumberOfFreeClusters, TotalNumberOfClusters
158             lblTotalSpace.caption = Format$(TotalNumberOfClusters * SectorsPerCluster * BytesPerSector, "###,###,###,##0") + " bytes"
160             lblFreeSpace.caption = Format$(NumberOfFreeClusters * SectorsPerCluster * BytesPerSector, "###,###,###,##0") + " bytes"
            End If
        End If

        'Vertical center
162     lblDriveLabel.Top = cmbDrives.Top + (cmbDrives.Height - lblDriveLabel.Height) / 2
        '<EhFooter>
        Exit Sub

cmbDrives_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.cmbDrives_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

        Dim hLib As Long
100     hLib = LoadLibrary("kernel32.dll")

        'Check if the GetDiskFreeSpaceExA-function is available
102     If GetProcAddress(hLib, "GetDiskFreeSpaceExA") <> 0 Then DiskFreeExAvailable = True
104     FreeLibrary hLib
        'Load all the information
106     ReadGeneralInfo
108     ReadDisks
110     ReadTime
112     ReadMonitor
114     ReadMouse
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub ReadGeneralInfo()
        '<EhHeader>
        On Error GoTo ReadGeneralInfo_Err
        '</EhHeader>
        Dim OSInfo As OSVERSIONINFO, MemStat As MEMORYSTATUS
100     sTemp = String(255, 0)
        'Get the current username
102     GetUserName sTemp, 255
104     lblLogged.caption = "Logged in under " + sTemp
106     sTemp = String(255, 0)
        'get the computername
108     GetComputerName sTemp, 255
110     lblLogged.caption = lblLogged.caption + " on " + sTemp
112     Me.caption = "Information about " + sTemp
114     sTemp = String(255, 0)
        'Get the windows-directory
116     GetWindowsDirectory sTemp, 255
118     lblWindows.caption = UCase$(sTemp)
120     sTemp = String(255, 0)
        'Get the system-directory
122     GetSystemDirectory sTemp, 255
124     lblSystem.caption = UCase$(sTemp)
126     sTemp = String(255, 0)
        'Get the temp-path
128     GetTempPath 255, sTemp
130     lblTemp.caption = UCase$(sTemp)
132     OSInfo.dwOSVersionInfoSize = Len(OSInfo)
        'Get information about the windows-version
134     GetVersionEx OSInfo

136     Select Case OSInfo.dwPlatformId

            Case 0
138             lblPlatform.caption = "Windows 32s on Windows 3.1"

140         Case 1
142             lblPlatform.caption = "Windows 95/98"

144         Case 2
146             lblPlatform.caption = "Windows NT"

148         Case Else
150             lblPlatform.caption = "Unknown"
        End Select

152     lblVersion.caption = Trim(str(OSInfo.dwMajorVersion)) + "." + Trim(str(OSInfo.dwMinorVersion))
154     MemStat.dwLength = Len(MemStat)
        'Get information about the memory
156     GlobalMemoryStatus MemStat

158     If MemStat.dwTotalPhys > 1024 ^ 2 Then
160         lblTotRAM.caption = Trim(str(Int(MemStat.dwTotalPhys / 1024 ^ 2))) + " mega bytes"
162     ElseIf MemStat.dwTotalPhys > 1024 Then
164         lblTotRAM.caption = Trim(str(Int(MemStat.dwTotalPhys / 1024))) + " kilo bytes"
        Else
166         lblTotRAM.caption = Trim(str(MemStat.dwTotalPhys)) + " bytes"
        End If

168     lblResources.caption = Trim(str(MemStat.dwMemoryLoad)) + "%"
        '<EhFooter>
        Exit Sub

ReadGeneralInfo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.ReadGeneralInfo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub ReadDisks()
        '<EhHeader>
        On Error GoTo ReadDisks_Err
        '</EhHeader>
        Dim bSelected As Boolean
        'Get all available drives
100     nDisks = GetLogicalDrives

102     For Cnt = 0 To 25

104         If (nDisks And 2 ^ Cnt) <> 0 Then
106             cmbDrives.AddItem Chr$(65 + Cnt) + ":\"

                'if this drive is a fixed drive, and there was no previous selected drive, select the current drive
108             If GetDriveType(Chr$(65 + Cnt) + ":\") = 3 And bSelected = False Then
110                 bSelected = True
112                 cmbDrives.ListIndex = cmbDrives.ListCount - 1
                End If
            End If

114     Next Cnt

        '<EhFooter>
        Exit Sub

ReadDisks_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.ReadDisks " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub ReadTime()
        '<EhHeader>
        On Error GoTo ReadTime_Err
        '</EhHeader>
        Dim MyTime As SYSTEMTIME, sDay As String
        'Get the local time
100     GetLocalTime MyTime
102     lblTime.caption = Format$(MyTime.wHour, "0#") + ":" + Format$(MyTime.wMinute, "0#") + ":" + Format$(MyTime.wSecond, "0#")
104     bAnimate = Not (bAnimate)

106     Select Case MyTime.wDayOfWeek

            Case 0
108             sDay = "Sunday"

110         Case 1
112             sDay = "Monday"

114         Case 2
116             sDay = "Tuesday"

118         Case 3
120             sDay = "Wednesday"

122         Case 4
124             sDay = "Thursday"

126         Case 5
128             sDay = "Friday"

130         Case 6
132             sDay = "Saturday"
        End Select

134     lblDate.caption = sDay + ", " + Format$(MyTime.wMonth, "0#") + "-" + Format$(MyTime.wDay, "0#") + "-" + Format$(MyTime.wYear, "0#")
        '<EhFooter>
        Exit Sub

ReadTime_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.ReadTime " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub ReadMonitor()
        '<EhHeader>
        On Error GoTo ReadMonitor_Err
        '</EhHeader>
        Dim sDriver As String, sManuf As String
        'retrieve the screen resolution
100     lblResolution = Trim(str(Screen.Width / Screen.TwipsPerPixelX)) + "x" + Trim(str(Screen.Height / Screen.TwipsPerPixelY))
        'retrieve the color depth
102     lblColors = Trim(str(GetDeviceCaps(Me.hdc, BITSPIXEL))) + " bit"
        'get the monitor's driver and manufacturer
104     GetDeviceDriver "Enum", "DISPLAY", sDriver, sManuf
106     lblDispDriver = sDriver
108     lblDispManuf = sManuf

        'The driver's name could be too long
110     If lblDispDriver.Width > frmMonitor.Width - lblDispDriver.Left - Label17.Left Then
112         lblDispDriver.Width = frmMonitor.Width - lblDispDriver.Left - Label17.Left
114         lblDispDriver.toolTipText = lblDispDriver.caption
        End If

        '<EhFooter>
        Exit Sub

ReadMonitor_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.ReadMonitor " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub ReadMouse()
        '<EhHeader>
        On Error GoTo ReadMouse_Err
        '</EhHeader>
        Dim sDriver As String, sManuf As String
        'Get the number of mousebuttons
100     lblButtons.caption = GetSystemMetrics(SM_CMOUSEBUTTONS)
        'get the mouse driver and manufacturer
102     GetDeviceDriver "Enum", "Mouse", sDriver, sManuf
104     lblMouseDriver.caption = sDriver
106     lblMouseManuf.caption = sManuf

        'The mouse name could be too long
108     If lblMouseDriver.Width > frmMouse.Width - lblMouseDriver.Left - Label20.Left Then
110         lblMouseDriver.Width = frmMouse.Width - lblMouseDriver.Left - Label20.Left
112         lblMouseDriver.toolTipText = lblMouseDriver.caption
        End If

        '<EhFooter>
        Exit Sub

ReadMouse_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.ReadMouse " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     tmrTime.Enabled = False
102     Set frmComputerInfo = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub tmrTime_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    ReadTime
End Sub

Private Function GetDeviceDriver(sPath As String, _
                                 sClass As String, _
                                 ByRef sDriver As String, _
                                 ByRef sManufacturer As String) As Boolean
        'to retrieve the driver and manufacturer, we need to do some digging in the registry
        '<EhHeader>
        On Error GoTo GetDeviceDriver_Err
        '</EhHeader>
        Dim FT As FILETIME, sName As String, Ret As Long, nCnt As Long
        'open the specified key
100     RegOpenKey HKEY_LOCAL_MACHINE, sPath, Ret

        'search for the specified class (can be 'mouse', 'display, 'system', 'hdc', 'fdc', ...)
102     If RegQueryStringValue(Ret, "Class") = sClass Then
            'if the class is found, get our data
104         sDriver = RegQueryStringValue(Ret, "DeviceDesc")
106         sManufacturer = RegQueryStringValue(Ret, "Mfg")
108         GetDeviceDriver = True
110         RegCloseKey Ret
            Exit Function
        End If

        Do
112         sName = String(255, 0)

            'anumerate all keys
114         If RegEnumKeyEx(Ret, nCnt, sName, 255, 0, vbNullString, 0, FT) <> 0 Then Exit Do
            'strip off the chr$(0)'s
116         sName = StripTerminator(sName)

            'search the new key
118         If GetDeviceDriver(sPath + "\" + sName, sClass, sDriver, sManufacturer) = True Then
120             GetDeviceDriver = True
                Exit Do
            End If

122         nCnt = nCnt + 1
        Loop

        'close our key
124     RegCloseKey Ret
        '<EhFooter>
        Exit Function

GetDeviceDriver_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.GetDeviceDriver " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function RegQueryStringValue(ByVal hKey As Long, _
                                     ByVal strValueName As String) As String
        '<EhHeader>
        On Error GoTo RegQueryStringValue_Err
        '</EhHeader>
        Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
100     lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)

102     If lResult = 0 Then
104         If lValueType = REG_SZ Then
106             strBuf = String(lDataBufSize, Chr$(0))
108             lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)

110             If lResult = 0 Then
112                 RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
                End If

114         ElseIf lValueType = REG_BINARY Then
                Dim strdata As Integer
116             lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strdata, lDataBufSize)

118             If lResult = 0 Then
120                 RegQueryStringValue = strdata
                End If
            End If
        End If

        '<EhFooter>
        Exit Function

RegQueryStringValue_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.RegQueryStringValue " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Function StripTerminator(sInput As String) As String
        '<EhHeader>
        On Error GoTo StripTerminator_Err
        '</EhHeader>
        Dim ZeroPos As Integer
100     ZeroPos = InStr(1, sInput, Chr$(0))

102     If ZeroPos > 0 Then
104         StripTerminator = Left$(sInput, ZeroPos - 1)
        End If

        '<EhFooter>
        Exit Function

StripTerminator_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmComputerInfo.StripTerminator " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
