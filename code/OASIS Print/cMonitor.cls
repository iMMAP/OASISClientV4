VERSION 1.0 CLASS
BEGIN
  MultiUse = 0   'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMonitor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type

Private Const CCHDEVICENAME = 32

Private Type MONITORINFOEXA
   cbSize As Long
   rcMonitor As RECT
   rcWork As RECT
   dwFlags As Long
   b(0 To CCHDEVICENAME - 1) As Byte
End Type

Private Type MONITORINFOEXW
   cbSize As Long
   rcMonitor As RECT
   rcWork As RECT
   dwFlags As Long
   b(0 To CCHDEVICENAME * 2 - 1) As Byte
End Type

Private Declare Function GetMonitorInfoA Lib "user32" ( _
      ByVal hMonitor As Long, _
      lpmi As MONITORINFOEXA _
   ) As Long
Private Declare Function GetMonitorInfoW Lib "user32" ( _
      ByVal hMonitor As Long, _
      lpmi As MONITORINFOEXW _
   ) As Long
Private Declare Function MonitorFromWindow Lib "user32" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function MonitorFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Const MONITOR_DEFAULTTONEAREST = 0

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Const MONITORINFOF_PRIMARY = &H1

Private Type OSVERSIONINFO
   dwVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion(0 To 127) As Byte
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2

Private m_hMonitor As Long
Private m_sName As String
Private m_rcMonitor As RECT
Private m_rcWork As RECT
Private m_bIsPrimary As Boolean

Private m_bIsNt As Boolean
Private m_bIsXp As Boolean

Public Sub ShowPrint(sPath As String, xmin As Double, xmax As Double, ymin As Double, ymax As Double)
    frmMainPrint.Show
    frmMainPrint.InitPrint sPath, xmin, xmax, ymin, ymax
End Sub

Public Property Get IsPrimary() As Boolean
   IsPrimary = m_bIsPrimary
End Property

Public Property Get hMonitor() As Long
   hMonitor = m_hMonitor
End Property
Public Property Get Name() As String
   Name = m_sName
End Property
Public Property Get left() As Long
   left = m_rcMonitor.left
End Property
Public Property Get top() As Long
   top = m_rcMonitor.top
End Property
Public Property Get Width() As Long
   Width = m_rcMonitor.right - m_rcMonitor.left
End Property
Public Property Get Height() As Long
   Height = m_rcMonitor.bottom - m_rcMonitor.top
End Property
Public Property Get WorkLeft() As Long
   WorkLeft = m_rcWork.left
End Property
Public Property Get WorkTop() As Long
   WorkTop = m_rcWork.top
End Property
Public Property Get WorkWidth() As Long
   WorkWidth = m_rcWork.right - m_rcMonitor.left
End Property
Public Property Get WorkHeight() As Long
   WorkHeight = m_rcWork.bottom - m_rcMonitor.top
End Property

Public Sub CreateFromPoint(ByVal x As Long, ByVal y As Long)
Dim hMon As Long
   On Error Resume Next ' For Windows 95 and NT
   hMon = MonitorFromPoint(x, y, MONITOR_DEFAULTTONEAREST)
   If Not (hMon = 0) Then
      fInit hMon
   End If
End Sub
Public Sub CreateFromWindow(ByVal hWnd As Long)
Dim hMon As Long
   On Error Resume Next ' For Windows 95 and NT
   hMon = MonitorFromWindow(hWnd, MONITOR_DEFAULTTONEAREST)
   If Not (hMon = 0) Then
      fInit hMon
   End If
End Sub

Friend Sub fInit( _
      ByVal hMonitor As Long _
   )
Dim iPos As Long
Dim sName As String
   On Error GoTo ErrorHandler ' For Windows 95 and NT
   m_hMonitor = hMonitor
   VerInitialise
   If (m_bIsNt) Then
      Dim tMIW As MONITORINFOEXW
      tMIW.cbSize = Len(tMIW)
      GetMonitorInfoW hMonitor, tMIW
      With tMIW
         LSet m_rcMonitor = .rcMonitor
         LSet m_rcWork = .rcWork
         m_bIsPrimary = ((.dwFlags And MONITORINFOF_PRIMARY) = MONITORINFOF_PRIMARY)
         sName = .b
         iPos = InStr(sName, vbNullChar)
      End With
   Else
      Dim tMIA As MONITORINFOEXA
      tMIA.cbSize = Len(tMIA)
      GetMonitorInfoA hMonitor, tMIA
      With tMIA
         LSet m_rcMonitor = .rcMonitor
         LSet m_rcWork = .rcWork
         m_bIsPrimary = ((.dwFlags And MONITORINFOF_PRIMARY) = MONITORINFOF_PRIMARY)
         sName = StrConv(.b, vbUnicode)
      End With
   End If
   iPos = InStr(sName, vbNullChar)
   If (iPos > 0) Then
      m_sName = left(sName, iPos - 1)
   Else
      m_sName = sName
   End If
   Exit Sub
   
ErrorHandler:
   m_hMonitor = 0
   Exit Sub
   
End Sub


Private Sub VerInitialise()
   
   Dim tOSV As OSVERSIONINFO
   tOSV.dwVersionInfoSize = Len(tOSV)
   GetVersionEx tOSV
   
   m_bIsNt = ((tOSV.dwPlatformId And VER_PLATFORM_WIN32_NT) = VER_PLATFORM_WIN32_NT)
   If (tOSV.dwMajorVersion > 5) Then
      'm_bHasGradientAndTransparency = True
      m_bIsXp = True
      'm_bIs2000OrAbove = True
   ElseIf (tOSV.dwMajorVersion = 5) Then
      'm_bHasGradientAndTransparency = True
      'm_bIs2000OrAbove = True
      If (tOSV.dwMinorVersion >= 1) Then
         m_bIsXp = True
      End If
   ElseIf (tOSV.dwMajorVersion = 4) Then ' NT4 or 9x/ME/SE
      'If (tOSV.dwMinorVersion >= 10) Then
      '   m_bHasGradientAndTransparency = True
      'End If
   Else ' Too old
   End If
   
End Sub



