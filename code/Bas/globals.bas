Attribute VB_Name = "globals"
Public g_sOGISFormat As String

Private Const Pi As Double = 3.1415926535979 'Pi constant, used in radians/degrees conversion
'Public Const Pi As Double = 3.14159265358979

Public GisUtils As New TatukGIS_XDK10.XGIS_Utils
Public ParamsUtils As New TatukGIS_XDK10.XGIS_ParamsUtils
Public minishp As TatukGIS_XDK10.XGIS_Shape               'minimap shape
Public minishpo As TatukGIS_XDK10.XGIS_Shape              'minimap shape outline
Public fminiMove As Boolean                'flag for move mini rectangle
Public lP1, lP2, lP3, lP4 As TatukGIS_XDK10.XGIS_Point    'large map extent points

Public g_PrevExt As TatukGIS_XDK10.XGIS_Extent

'GLOBAL APPLICATION SETTINGS
Public m_oColUserLayers As New Collection
Public g_RSAppSettings As New ADODB.Recordset
Public g_RSLocalAppSettings As New ADODB.Recordset
Public g_sAppSettingsTable As String
Public g_RSGISGridTableSettings As New ADODB.Recordset
Public g_RSThemeSettings As New ADODB.Recordset
Public g_RSThemeGroups As New ADODB.Recordset
Public g_sAppServerPath As String
Public g_RSMaps As ADODB.Recordset
Public g_sUserName As String
Public g_sUserPass As String
Public g_bUseSynch As Boolean
Public g_sRemoteTablePrefix As String
Public mRSUGSettings As New ADODB.Recordset
Public g_sLanguage As String
Public g_sAppPath As String
Public g_bFeedUpdate As Boolean
Public g_bMapTplUpdate As Boolean
Public g_bMapsUpdate As Boolean
Public g_bOnlineCheckedAtLogin As Boolean
Public g_bRefreshMapAfterSynch As Boolean
Public g_bMapLoadedCorrectly As Boolean
Public g_bDefaultMapChanged As Boolean
Public g_bDefaultMapChangedGUID As String

Public g_ClientDBPassword As String
Public bSQLServerInUse As Boolean
Public g_sGlobalConnectionString As String
Public g_sGlobalDialect As String
Public g_sGlobalCursorLocation As ADODB.CursorLocationEnum
Public g_bIncidentsV2 As Boolean
Public g_sIncidentsV2DDName As String
Public g_sIncidentsV2ConnectionString As String

Public g_sSQLServerDatabaseName As String
Public g_sManualSQLServerPath As String

Public g_ProxyChecked As Boolean
Public g_bProxyEnabled As Boolean
Public g_sProxyIP As String
Public g_sProxyPort As String
Public g_sProxyUser As String
Public g_sProxyPass As String
Public g_iProxyAuthType As Integer

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Server communications stuff
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public SilentHttpComms As clServerComms
'Public g_ServerConnTimeoutSeconds As Long
'Public g_ServerConnNoOfRetries As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public m_frmDebug As frmDebug

'GLOBAL Styles Used
Public g_RSStyles As New ADODB.Recordset

'GLOBAL Rendering Used
Public g_RSRender As New ADODB.Recordset

Public m_Cnn As New ADODB.Connection
Public g_CurrentUserID As Long
Public g_CurrentTool As OASIS_TOOLS
Public g_CurrentLocationType As OASISLocationType
Public g_CurrentFeatureType As OASISFeatureTypes

Public Const VK_CONTROL = 17
Public Const VK_DELETE = 46

Public Const MINIMAP_R_NAME = "minimap_rect"
Public Const MINIMAP_O_NAME = "minimap_rect_outline"

' API CALLS
Declare Function WindowFromPoint _
        Lib "user32.dll" (ByVal xPoint As Long, _
                          ByVal yPoint As Long) As Long

Declare Function CoCreateGuid _
        Lib "ole32.dll" (pguid As GUID) As Long
Declare Function StringFromGUID2 _
        Lib "ole32.dll" (rguid As Any, _
                         ByVal lpstrClsId As Long, _
                         ByVal cbMax As Long) As Long
Declare Function GetCursorPos _
        Lib "User32" (lpPoint As POINTAPI) As Long
Declare Function ShellExecute _
        Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hWnd As Long, _
                               ByVal lpOperation As String, _
                               ByVal lpFile As String, _
                               ByVal lpParameters As String, _
                               ByVal lpDirectory As String, _
                               ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

Public Declare Function SetParent _
               Lib "User32" (ByVal hWndChild As Long, _
                             ByVal hWndNewParent As Long) As Long

'GUID STRUCT
Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Public oCoordTransSettings As CoordTransSettings
Public oMapTipSetting As MapTipSetting
Public oLocatorSettings As LocatorSettings
Public oSelectionStyle As SelectionStyle
Public g_ZoomToSettings As ZoomToSettings
Public oUrlLayerSettings As UrlLayerSettings
Public oIncidentLayerSettings As IncidentLayerSettings
Public oMapSettings As MapSettings
Public oMapObjects As MapObjectsSettings
Public g_DatabaseSpecExtent As TatukGIS_XDK10.XGIS_Extent

'****************************************************************************
'
'   Set of functions to return UTC Time / Date using some API calls
'   Format of functions similar to standard Time / Date
'   funtions in VB (Time, Time$, Date, Date$, Now)
'
'   Date functions use TimeSerial, DateSerial and Named Formats in VB,
'   so the output of date functions will follow conventions for the users
'   country/regional settings.
'
'   Example: United States date is MM/DD/YYYY, Europe is DD/MM/YYYY.
'
'   Function examples are in United States format
'
'   Also there are two funtions that return the name of the
'   time zone the system is set for, one the short name (like "EST"),
'   the second for the long name (like "Eastern Standard Time")
'   Some of the time zone names don't realy work
'   well with short names.  But it works fine for most
'   U.S. and Canada time zones.  Change the code as you see fit.
'   Short and Long Time Zone Names change to "Daylight Time"
'   if a daylight time zone is selected.
'
'   List of new functions with there VB local time equivalent
'   UTC_time function   VB function     Format
'   UTCtime             Time$           24 Hour HH:MM:SS
'   UTCtime2            Time            Local format
'   UTCdate             Date            MM/DD/YYYY  (region dependent)
'   UTCnow              Now             MM/DD/YYYY HH:MM:SS AM/PM   (region dependent)
'   shortTZname         ----            XXX Ex: "EST", 3 to 5 letters
'   longTZname          ----            Long name Ex: "Eastern Standard Time"
'   ISOdate             Date            ISO 8601 format yyyy-mm-dd
'   ISOtime             Time            ISO 8601 format hh:mm:ssZ
'   ISOnow              Now             ISO 8601 format yyyy-mm-ddThh:mm:ssZ
'   UTCoffset           ----            Offset for local time in minutes
'
'   Time zone and UTC time functions assume the correct time zone is selected
'   in the Time/Date properties and clock set correct local time.
'***************************************************************************

Private Declare Sub GetSystemTime _
                Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function GetTimeZoneInformation _
                Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Type SYSTEMTIME
    wYear                                                             As Integer
    wMonth                                                            As Integer
    wDayOfWeek                                                        As Integer
    wDay                                                              As Integer
    wHour                                                             As Integer
    wMinute                                                           As Integer
    wSecond                                                           As Integer
    wMilliseconds                                                     As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias                                                              As Long
    StandardName(32)                                                  As Integer
    StandardDate                                                      As SYSTEMTIME
    StandardBias                                                      As Long
    DaylightName(32)                                                  As Integer
    DaylightDate                                                      As SYSTEMTIME
    DaylightBias                                                      As Long
End Type

Dim sysTime                                                               As SYSTEMTIME
Dim TZinfo                                                                As TIME_ZONE_INFORMATION

Public Const RELATE_INTERSECT = "T"

Public Declare Function Ellipse _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long
Public Declare Function SetROP2 _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal nDrawMode As Long) As Long
Public Const R2_BLACK = 1    ' 0
Public Const R2_COPYPEN = 13  ' P
Public Const R2_LAST = 16
Public Const R2_MASKNOTPEN = 3 ' DPna
Public Const R2_MASKPEN = 9   ' DPa
Public Const R2_MASKPENNOT = 5 ' PDna
Public Const R2_MERGENOTPEN = 12    ' DPno
Public Const R2_MERGEPEN = 15  ' DPo
Public Const R2_MERGEPENNOT = 14    ' PDno
Public Const R2_NOP = 11    ' D
Public Const R2_NOT = 6 ' Dn
Public Const R2_NOTCOPYPEN = 4 ' PN
Public Const R2_NOTMASKPEN = 8 ' DPan
Public Const R2_NOTMERGEPEN = 2 ' DPon
Public Const R2_WHITE = 16   ' 1
Public Const R2_XORPEN = 7   ' DPx

Public Declare Function LoadCursorFromFile _
               Lib "User32" _
               Alias "LoadCursorFromFileA" (ByVal lpFileName As String) As Long
Private Declare Function GetVersionExA _
                Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Private Declare Function InitCommonControls _
                Lib "comctl32.dll" () As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type GUID1
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const MOUSE_BUFFER = 300
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2

Public Declare Function GetWindowDC _
               Lib "User32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC _
               Lib "User32" (ByVal hWnd As Long, _
                             ByVal hdc As Long) As Long
Public Declare Function SetWindowPos _
               Lib "user32.dll" (ByVal hWnd As Long, _
                                 ByVal hWndInsertAfter As Long, _
                                 ByVal x As Long, _
                                 ByVal y As Long, _
                                 ByVal cx As Long, _
                                 ByVal cy As Long, _
                                 ByVal wFlags As Long) As Long
Public Declare Function StretchBlt _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal x As Long, _
                            ByVal y As Long, _
                            ByVal nWidth As Long, _
                            ByVal nHeight As Long, _
                            ByVal hSrcDC As Long, _
                            ByVal xSrc As Long, _
                            ByVal ySrc As Long, _
                            ByVal nSrcWidth As Long, _
                            ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
     
Public Const CTF_COINIT = &H8
Public Const CTF_INSIST = &H1
Public Const CTF_PROCESS_REF = &H4
Public Const CTF_THREAD_REF = &H2

Declare Function SHCreateThread _
        Lib "shlwapi.dll" (ByVal pfnThreadProc As Long, _
                           pData As Any, _
                           ByVal dwFlags As Long, _
                           ByVal pfnCallback As Long) As Long
     
Public Declare Function SetErrorMode _
               Lib "kernel32" (ByVal wMode As Long) As Long

Public Const SEM_NOGPFAULTERRORBOX = &H2&

Public g_bDemoLogin As Boolean
Public g_bHasEncrypt As Boolean
Public g_sKey As String
Public m_oAES As clsAES

Public g_lPinUID As Long

'This is used for the dynamic data form

Public ExcludeArray() As ExcludeType

Public DE9IM As New Dictionary

Public g_sPermanentLyrs As String
Public m_ColAnlFieldType As Collection

Public g_ColAlerts As Collection
Public arRangeColors() As Long
Public arRangeVals() As Long

Public m_PTUID As Long
Public g_udtSynchUpdateOptions As SynchUpdateOptions

Public g_bUpdateDynamicDataDefs As Boolean

Public Declare Function GetDesktopWindow _
               Lib "User32" () As Long

Public g_ssubdomain As String
Public m_bAdvancedDebug As Boolean

Public Sub CheckIfDebugEnhancedEnabled(mCN As ADODB.Connection)
        '<EhHeader>
        On Error GoTo CheckIfDebugEnhancedEnabled_Err
        '</EhHeader>
        
        Dim oRS As New ADODB.Recordset
100     oRS.Open "SELECT [SettingValue1] FROM [Appsettings] WHERE [SettingName] = 'ShowAdvancedDebug'", m_Cnn
102     m_bAdvancedDebug = False
    
104     If Not oRS.State = 0 Then
    
106         If Not oRS.EOF Then
        
108             If CStr(Trim(oRS.Fields(0).value)) = "1" Then m_bAdvancedDebug = True
        
            End If
            
            oRS.Close
    
        End If
        
        Set oRS = Nothing

        '<EhFooter>
        Exit Sub

CheckIfDebugEnhancedEnabled_Err:
        MsgBox "Globals.CheckIfDebugEnhancedEnabled_Err on line (" & Erl & ") " & Err.Description
        '</EhFooter>
End Sub

Public Sub DebugPrint(ByVal sText As Variant, _
                      Optional bAdvanced As Boolean = False)

If sText = left(sText, Len(">>>> Save")) = ">>>> Save" Then Stop

    If Not (bAdvanced And Not m_bAdvancedDebug) Then

        If m_frmDebug.Visible Then  ' Is Nothing Then
            m_frmDebug.DebugPrinter CStr(sText)
        Else
            Debug.Print "[" & Time & "] " & CStr(sText)
        End If
        
    End If

End Sub

Public Function HaversineDistance(ByVal Long1 As Double, _
                                  ByVal Long2 As Double, _
                                  ByVal Lat1 As Double, _
                                  ByVal Lat2 As Double) As Double
    
    Const R As Integer = 6371   'earth radius in km
        
    Dim DeltaLat As Double, DeltaLong As Double
    Dim a As Double, c As Double
    Dim Pi As Double
    
    On Error GoTo ErrorExit
    
    Pi = 4 * Atn(1)
    
    'convert Lat1, Lat2, Long1, Long2 from decimal degrees into radians
    Lat1 = Lat1 * Pi / 180
    Lat2 = Lat2 * Pi / 180
    Long1 = Long1 * Pi / 180
    Long2 = Long2 * Pi / 180
    
    'calculate change in Latitude and Longitude
    DeltaLat = Abs(Lat2 - Lat1)
    DeltaLong = Abs(Long2 - Long1)
    
    a = ((Sin(DeltaLat / 2)) ^ 2) + (Cos(Lat1) * Cos(Lat2) * ((Sin(DeltaLong / 2)) ^ 2))
    a = (Sin(DeltaLat / 2) * Sin(DeltaLat / 2)) + (Cos(Lat1) * Cos(Lat2) * Sin(DeltaLong / 2) * Sin(DeltaLong / 2))
    
    
    c = 2 * atan2(Sqr(1 - a), Sqr(a))                    'expressed as radians
    'c = 2 * Atn((Sqr(1 - a)) / (Sqr(a)))
    HaversineDistance = R * c
ErrorExit:

End Function

Public Function atan2(x As Double, _
                      y As Double) As Double

    If x > 0 Then
        atan2 = Atn(y / x)
    ElseIf x < 0 Then
        atan2 = Sgn(y) * (Pi - Atn(Abs(y / x)))
    ElseIf y = 0 Then
        atan2 = 0
    Else
        atan2 = Sgn(y) * Pi / 2
    End If

End Function

Public Function InstallAUpdate(hWnd As Long) As Boolean
    ' /PermissionManagerCheckInstalled
   
    'Shell
   If g_bOnlineCheckedAtLogin Then
    ShellExecute hWnd, vbNullString, g_sAppPath & "\OASIS_SynchNG_Client.exe", "CheckBackground", "C:\", 1
    ShellExecute hWnd, vbNullString, g_sAppPath & "\AUClient.exe", "CheckBackground", "C:\", 1
    
    End If
End Function

'
'
'Public Function GUIDGen() As String
'    Dim uGUID As GUID
'    Dim sGUID As String
'    Dim bGUID() As Byte
'    Dim lLen As Long
'    Dim retval As Long
'    lLen = 40
'    bGUID = String(lLen, 0)
'    CoCreateGuid uGUID
'    retval = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
'    sGUID = bGUID
'    If (Asc(Mid$(sGUID, retval, 1)) = 0) Then retval = retval - 1
'    GUIDGen = Left$(sGUID, retval)
'End Function

Public Function CreateManifest() As Boolean
    On Error Resume Next
    Dim EXEPath As String
    'Get The EXE Path
    EXEPath = App.Path & IIf(right(App.Path, 1) = "\", vbNullString, "\")
    EXEPath = EXEPath & App.EXEName & IIf(LCase(right(App.EXEName, 4)) = ".exe", ".manifest", ".exe.manifest")

    'Checks if the manifest has already been
    '     created
    If Dir(EXEPath, vbReadOnly Or vbSystem Or vbHidden) <> vbNullString Then GoTo ErrorHandler
    'Makes sure you are using windows xp

    If WinVersion = "Windows XP" Then
        Dim iFileNumber As Integer

        iFileNumber = FreeFile
        'Save the .manifest file
        Open EXEPath For Output As #iFileNumber
        Print #iFileNumber, FormatManifest
        CreateManifest = True
    Else
        Kill EXEPath
    End If

    'set the file to be hidden
    Close #iFileNumber
    SetAttr EXEPath, vbHidden Or vbSystem Or vbReadOnly Or vbArchive
ErrorHandler:
    Call InitCommonControls
End Function

Private Function WinVersion() As String
    Dim OSInfo As OSVERSIONINFO
    Dim retValue As Integer
    OSInfo.dwOSVersionInfoSize = 148
    OSInfo.szCSDVersion = Space$(128)
    retValue = GetVersionExA(OSInfo)

    With OSInfo

        Select Case .dwPlatformId

            Case 1

                If .dwMinorVersion = 0 Then
                    WinVersion = "Windows 95"
                ElseIf .dwMinorVersion = 10 Then
                    WinVersion = "Windows 98"
                End If

            Case 2

                If .dwMajorVersion = 3 Then
                    WinVersion = "Windows NT 3.51"
                ElseIf .dwMajorVersion = 4 Then
                    WinVersion = "Windows NT 4.0"
                ElseIf .dwMajorVersion >= 5 Then
                    WinVersion = "Windows XP"
                End If

            Case Else
                WinVersion = "Failed"
        End Select

    End With

End Function

Private Function FormatManifest() As String
    Dim Header As String
    Header = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>"
    Header = Header & vbCrLf & "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">"
    Header = Header & vbCrLf & "<assemblyIdentity"
    Header = Header & vbCrLf & "version=" & Chr(34) & "1.0.0.0" & Chr(34)
    Header = Header & vbCrLf & "processorArchitecture=" & Chr(34) & "X86" & Chr(34)
    Header = Header & vbCrLf & "name=" & Chr(34) & "Microsoft.VisualBasic6.IDE" & Chr(34)
    Header = Header & vbCrLf & "type=" & Chr(34) & "win32" & Chr(34)
    Header = Header & vbCrLf & "/>"
    Header = Header & vbCrLf & "<description>Microsoft Visual Basic 6 IDE</description>"
    Header = Header & vbCrLf & "<dependency>"
    Header = Header & vbCrLf & "<dependentAssembly>"
    Header = Header & vbCrLf & "<assemblyIdentity"
    Header = Header & vbCrLf & "type=" & Chr(34) & "win32" & Chr(34)
    Header = Header & vbCrLf & "name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34)
    Header = Header & vbCrLf & "version=" & Chr(34) & "6.0.0.0" & Chr(34)
    Header = Header & vbCrLf & "processorArchitecture=" & Chr(34) & "X86" & Chr(34)
    Header = Header & vbCrLf & "publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34)
    Header = Header & vbCrLf & "language=" & Chr(34) & "*" & Chr(34)
    Header = Header & vbCrLf & "/>"
    Header = Header & vbCrLf & "</dependentAssembly>"
    Header = Header & vbCrLf & "</dependency>"
    Header = Header & vbCrLf & "</assembly>"

    FormatManifest = Header
End Function

Public Function UTCtime()

    'Format: HH:MM:SS in 24 hour format (like time$ function)

    Call GetSystemTime(sysTime)
   
    UTCtime = Format(Str(sysTime.wHour), "00") & ":" & Format(Str(sysTime.wMinute), "00") & ":" & Format(Str(sysTime.wSecond), "00")
    
End Function

Public Function UTCtime2()

    'Format: HH:MM:SS in local format (like time function)

    Call GetSystemTime(sysTime)
    
    UTCtime2 = TimeSerial(sysTime.wHour, sysTime.wMinute, sysTime.wSecond)
    
End Function

Public Function UTCnow()

    'Format: Like "Now" function. ex: "1/2/2003 4:35:15 PM"
    'Functions returns format above (region dependent)

    Call GetSystemTime(sysTime)
    
    UTCnow = DateSerial(sysTime.wYear, sysTime.wMonth, sysTime.wDay) & " " & TimeSerial(sysTime.wHour, sysTime.wMinute, sysTime.wSecond)
       
    '****Can also show as MEDIUM DATE. Comment the line above and uncomment below
    'UTCnow = UCase(Format(DateSerial(sysTime.wYear, sysTime.wMonth, sysTime.wDay), "medium date")) & " " & TimeSerial(sysTime.wHour, sysTime.wMinute, sysTime.wSecond)
    
    '****Can also show as LONG DATE. Comment the lines above and uncomment below
    'UTCnow = Format(DateSerial(sysTime.wYear, sysTime.wMonth, sysTime.wDay), "long date") & " " & TimeSerial(sysTime.wHour, sysTime.wMinute, sysTime.wSecond)
    
End Function

Public Function UTCdate()

    'Format: MM/DD/YY ex: 1/1/2002 (like date function)
    'Functions returns format above (region dependent)

    Call GetSystemTime(sysTime)

    UTCdate = DateSerial(sysTime.wYear, sysTime.wMonth, sysTime.wDay)
    
End Function

Public Function ISOdate()

    'Format: YYYY-MM-DD ex: 2003-01-03
    'Functions returns format above (Fixed ISO Format)

    Call GetSystemTime(sysTime)

    ISOdate = Format(Str(sysTime.wYear), "0000") & "-" & Format(Str(sysTime.wMonth), "00") & "-" & Format(Str(sysTime.wDay), "00")
    
End Function

Public Function ISOtime()

    'Format: hh:mm:ssZ ex: 15:03:47Z ("Z"denotes ZULU or UTC time)
    'Function returns format above (Fixed ISO Format)

    Call GetSystemTime(sysTime)

    ISOtime = Format(Str(sysTime.wHour), "00") & ":" & Format(Str(sysTime.wMinute), "00") & ":" & Format(Str(sysTime.wSecond), "00") & "Z"
    
End Function

Public Function ISOnow()

    'Format: yyyy-mm-ddThh:mm:ssZ ex: 2003-01-03T15:03:47Z
    'Function returns format above (Fixed ISO Format)

    ISOnow = ISOdate & "T" & ISOtime
    
End Function

Public Function shortTZname()

    'Format: XYZ ex: "EST" for Eastern Standard Time
    'Some of the time zone names don't realy work
    'well with short names.  But it works fine for most
    'U.S. and Canada time zones.  Change the code as you see fit.

    Dim M                 As Integer
    Dim TZname            As String

    TZname = longTZname

    'Get first letter of each word of long time zone name to get short name
    'Fist letter, first word
    shortTZname = Mid(TZname, 1, 1)
    'Set pointer for second word
    M = InStr((M + 1), TZname, " ")
    'Loop thru the long time zone name and find first letter of remaining all words

    Do Until M = 0
        shortTZname = shortTZname & Mid(TZname, (M + 1), 1)
        M = InStr((M + 1), TZname, " ")
    Loop
    
    'Force uppercase for display. It should be already, but just incase
    shortTZname = UCase(shortTZname)
    
End Function

Public Function longTZname()

    'Format: Long Time Zone Name ex: "Eastern Standard Time"

    Dim TZResult            As Long
    Dim i                   As Long
    Dim tempname            As String
    
    TZResult = GetTimeZoneInformation(TZinfo)
    
    'Extract Time Zone Name from returned API call

    For i = 0 To 31

        If TZinfo.StandardName(i) = 0 Then Exit For
        longTZname = longTZname & Chr(TZinfo.StandardName(i))
    Next

    Select Case TZResult

        Case 0, 1 'Use standard time name
            
        Case 2 'Use daylight savings time name
            tempname = Mid(longTZname, 1, (InStr(1, longTZname, " ")))
            longTZname = tempname & "Daylight Time"
                
    End Select
    
    'Trim any spaces in longTZname. Shoud be free of spaces, but just incase
    longTZname = Trim(longTZname)
    
End Function

Public Function UTCoffset()

    'Get number of minutes your local Time Zone is offset from UTC

    Dim TZResult            As Long

    TZResult = GetTimeZoneInformation(TZinfo)

    Select Case TZResult

        Case 0, 1 'Use standard time UTC offset
            UTCoffset = TZinfo.Bias

        Case 2 'Use daylight savings time UTC offset
            UTCoffset = TZinfo.Bias - 60
                
    End Select

End Function

Public Function GetGuid() As String
    
    Dim udtGuid As GUID
    Dim sGUID As String
    Dim bytGuid() As Byte
    Dim lRet As Long
    Dim lLen As Long
    lLen = 40
    bytGuid = String(lLen, 0)
    CoCreateGuid udtGuid
    lRet = StringFromGUID2(udtGuid, VarPtr(bytGuid(0)), lLen)
    ' VarPtr is hidden VB function
    ' we need this function to format string (Basic String)
    sGUID = bytGuid

    If (Asc(Mid$(sGUID, lRet, 1)) = 0) Then
        lRet = lRet - 1
    End If

    GetGuid = left$(sGUID, lRet)
End Function

Public Function SafeMoveFirst(ByRef oRS As ADODB.Recordset) As Boolean
On Error GoTo FailMoveFirst
    If Not oRS.State = adStateClosed Then
    
        If Not oRS.EOF Or Not oRS.Bof Then
            oRS.MoveFirst
            SafeMoveFirst = True
        Else
            SafeMoveFirst = False
        End If

    Else
        SafeMoveFirst = False
    End If
    
    Exit Function
FailMoveFirst:
    SafeMoveFirst = False

End Function

Public Sub IncrementProfileSettingVersion(S1 As String, _
                                          S2 As String, _
                                          s3 As String)
    'Dummy sub for use in frmDynamicDataMenu
End Sub

Public Function DeleteRecordFromRSAndSave(rs1 As ADODB.Recordset, _
                                          S1 As String, _
                                          S2 As String) As Boolean
    'Dummy function for use in frmDynamicDataMenu
End Function


'Public Sub MapProjectDefUpdate(CN As Adodb.Connection, _
'                               RSRemote As Adodb.Recordset, _
'                               RS As Adodb.Recordset)
'
'    Dim lRetValDownload As VbMsgBoxResult
'    Dim lRetValSave As VbMsgBoxResult
'
'    'lRetValDownload = MsgBox("There is a new map project available for download. Do you want to download it?", vbYesNo, "New map project available")
'
'    'If lRetValDownload = vbYes Then
'
'    'lRetValSave = MsgBox("NOTE: This will override your existing map project. Do you want to save your existing map project first?", vbYesNoCancel, "New map project available")
'
'    '  If Not lRetValSave = vbCancel Then
'
'    '   If lRetValSave = vbYes Then
'
'    'Save it
'    '   LoadProjectFileFromDB g_sAppPath & "\data\user\maps\Map Backup on " & CStr(Format(Now(), "Medium Date")) & GUIDGen & ".ttkgp"
'    '   MsgBox "Your prior map project has been saved to: " & g_sAppPath & "\data\user\maps\Map Backup on " & CStr(Format(Now(), "Medium Date")) & GUIDGen & ".ttkgp"
'
'    '   End If
'
'    On Error Resume Next
'    Dim j As Integer
'    Dim sString As String
'
'    sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM [" & g_sRemoteTablePrefix & "ttkGISProjectDef]")
'    Set RSRemote = OpenSilentHttpCommsRS(sString, True)
'
'    If Not RSRemote Is Nothing Then
'
'        If Not RSRemote.State = adStateClosed Then
'
'            Set RS = New Adodb.Recordset
'            CN.Execute "delete from [ttkGISProjectDef]"
'
'            If Not RSRemote.EOF And Not RSRemote.BOF Then
'
'                If SafeMoveFirst(RSRemote) Then
'
'                    RS.Open "SELECT * FROM [ttkGISProjectDef]", CN, adOpenDynamic, adLockOptimistic
'
'                    If Not RSRemote.EOF Then
'
'                        RS.AddNew
'
'                        For j = 0 To RSRemote.Fields.Count - 1
'                            'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
'                            RS.Fields.Item(j).Value = RSRemote.Fields(RS.Fields.Item(j).Name).Value
'                        Next
'
'                    End If
'
'                    RS.UpdateBatch
'                End If
'
'                RSRemote.Close
'                RS.Close
'
'            End If
'
'        End If
'
'        SynchProfileSettingWithServer "SettingValue1", g_sRemoteTablePrefix, CN, "MapProjectDef"
'
'        ' End If
'
'        'End If
'
'    End If
'
'    Set RSRemote = Nothing
'    Set RS = Nothing
'
'End Sub

Public Sub UpdateSQLLayersInTTkGISLayerSQL(CN As ADODB.Connection)
                           
    On Error Resume Next
    
    Dim RSLocal As ADODB.Recordset
    Dim oRSttkGISLayerSQL As ADODB.Recordset
    Dim oCn As ADODB.Connection

    Set RSLocal = New ADODB.Recordset
    RSLocal.Open "SELECT * FROM [ttkGISLayerSQLInProject]", CN, adOpenDynamic, adLockOptimistic
                    
    Do While Not RSLocal.EOF
                            
        Set oCn = New ADODB.Connection
        oCn.Open IIf(bSQLServerInUse, g_sGlobalConnectionString, Replace(RSLocal.Fields("ADO").value, "CLIENTDBPATH", g_sAppPath))
        Set oRSttkGISLayerSQL = New ADODB.Recordset
                            
        If Not oCn.State = adStateClosed Then
                            
            oRSttkGISLayerSQL.Open "SELECT * from [ttkGISLayerSQL] WHERE [Name] = '" & RSLocal.Fields("LayerName").value & "'", oCn, adOpenDynamic, adLockBatchOptimistic
                            
            If Not oRSttkGISLayerSQL.State = adStateClosed Then
                            
                If oRSttkGISLayerSQL.EOF Then oRSttkGISLayerSQL.AddNew
                oRSttkGISLayerSQL.Fields("Name").value = RSLocal.Fields("LayerName").value
                oRSttkGISLayerSQL.Fields("XMIN").value = RSLocal.Fields("XMIN").value
                oRSttkGISLayerSQL.Fields("XMAX").value = RSLocal.Fields("XMAX").value
                oRSttkGISLayerSQL.Fields("YMIN").value = RSLocal.Fields("YMIN").value
                oRSttkGISLayerSQL.Fields("YMAX").value = RSLocal.Fields("YMAX").value
                oRSttkGISLayerSQL.Fields("SHAPETYPE").value = RSLocal.Fields("SHAPETYPE").value
                oRSttkGISLayerSQL.UpdateBatch adAffectCurrent
                oRSttkGISLayerSQL.Close
                oCn.Close
                            
            End If
                            
        End If
                            
        Set oRSttkGISLayerSQL = Nothing
        Set oCn = Nothing
                                
        RSLocal.MoveNext
                                
    Loop
                    
    RSLocal.UpdateBatch
    RSLocal.Close
    Set RSLocal = Nothing

End Sub

'Public Sub DynamDataDefsUpdate(CN As Adodb.Connection, _
'                               RSRemote As Adodb.Recordset, _
'                               RS As Adodb.Recordset)
'        '<EhHeader>
'        On Error GoTo DynamDataDefsUpdate_Err
'        '</EhHeader>
'        Dim j As Integer
'        Dim sString As String
'
'100     sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & g_sRemoteTablePrefix & "DynamicDataDefs")
'        Set RSRemote = OpenSilentHttpCommsRS(sString, True)
'
'102     If Not RSRemote.State = 0 Then
'104         Set RS = New Adodb.Recordset
'
'106         CN.Execute "delete from DynamicDataDefs"
'
'108         If Not RSRemote.EOF And Not RSRemote.BOF Then
'
'110             SafeMoveFirst RSRemote
'
'112             Err.Clear
'
'114             If Not Err.number > 0 Then
'116                 RS.Open "SELECT * FROM DynamicDataDefs", CN, adOpenDynamic, adLockOptimistic
'                    'MsgBox "remote: " & rsRemote.Fields.Count
'                    'MsgBox "local: " & RS.Fields.Count
'
'118                 Do While Not RSRemote.EOF
'
'120                     If Not Err.number > 0 Then
'122                         RS.AddNew
'
'124                         For j = 0 To RS.Fields.Count - 1
'
'
'                                If Not RSRemote.Fields.Item(RS.Fields.Item(j).Name) Is Nothing Then
'126                                 RS.Fields.Item(j).Value = RSRemote.Fields(RS.Fields.Item(j).Name).Value
'                                End If
'
'                            Next
'
'                        End If
'
'128                     RSRemote.MoveNext
'
'                    Loop
'
'130                 RS.UpdateBatch
'                End If
'
'132             RSRemote.Close
'134             RS.Close
'
'            End If
'
'            SynchProfileSettingWithServer "SettingValue7", g_sRemoteTablePrefix, CN
'
'136         Set RSRemote = Nothing
'138         Set RS = Nothing
'        End If
'
'        '<EhFooter>
'        Exit Sub
'
'DynamDataDefsUpdate_Err:
'        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLogin.DynamDataDefsUpdate " & "at line " & Erl
'        Resume Next
'        '</EhFooter>
'End Sub

Public Sub UGTableUpdate(sTableName As String, _
                         CN As ADODB.Connection, _
                         sSettingNum As String, _
                         sSettingName As String)
        '<EhHeader>
        On Error GoTo UGTableUpdate_Err
        '</EhHeader>
                         
        Dim rsRemote As ADODB.Recordset
        Dim RSLocal As ADODB.Recordset
        Dim oRS As ADODB.Recordset
        Dim iCountInUse As Integer
        Dim j As Integer
        Dim sString As String
    
100     DebugPrint "Updating table: " & sTableName
         
        'sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM [" & g_sRemoteTablePrefix & sTableName & "]")
        'Set rsRemote = OpenSilentHttpCommsRS(sString, True)
102     Set rsRemote = OpenServerRSCompressed(g_sAppServerPath & "/oasis4.asp", "id", "SELECT * FROM [" & g_sRemoteTablePrefix & sTableName & "]")
                
104     If Not rsRemote Is Nothing Then
    
106         If Not rsRemote.State = 0 Then
        
108             Set RSLocal = New ADODB.Recordset

110             If sTableName = "ttkGISProjectDef" Then
112                 CN.Execute "delete from [ttkGISProjectDef] WHERE bUGMap = true"

114                 Set oRS = CN.Execute("select count(sGUID) from [ttkGISProjectDef] WHERE InUse = true")
116                 iCountInUse = oRS.Fields(0).value
                Else

118                 CN.Execute "delete from [" & sTableName & "]"
                End If

120             If Not rsRemote.EOF And Not rsRemote.Bof Then

122                 SafeMoveFirst rsRemote
124                 Err.Clear

126                 If Not Err.number > 0 Then
128                     RSLocal.Open "SELECT * FROM [" & sTableName & "]", CN, adOpenDynamic, adLockBatchOptimistic
                    
130                     If RSLocal.State = adStateOpen Then
                    
132                         Do While Not rsRemote.EOF

134                             If Not Err.number > 0 Then
136                                 RSLocal.AddNew
                            
138                                 For j = 0 To RSLocal.Fields.Count - 1

140                                     If Not (bSQLServerInUse And UCase(RSLocal.Fields.Item(j).Name) = "ID") Then

142                                         If DoFieldExists(rsRemote, RSLocal.Fields.Item(j).Name) Then

144                                             If iCountInUse > 0 And sTableName = "ttkGISProjectDef" And RSLocal.Fields.Item(j).Name = "InUse" Then
146                                                 RSLocal.Fields.Item(j).value = False
                                                Else
148                                                 RSLocal.Fields.Item(j).value = rsRemote.Fields(RSLocal.Fields.Item(j).Name).value
                                                End If


                                            End If
                                        End If

                                    Next

150                                 RSLocal.UpdateBatch adAffectCurrent
                                End If
                                
152                             rsRemote.MoveNext
                                
                            Loop

154                         rsRemote.Close
156                         RSLocal.Close
                    
                        End If
                
                    End If
                    
                End If
            
158             If Len(sSettingNum) > 1 Then
160                 SynchProfileSettingWithServer sSettingNum, g_sRemoteTablePrefix, CN, sSettingName
                End If
            
            End If

162         Set rsRemote = Nothing
164         Set RSLocal = Nothing
       
        End If

        '<EhFooter>
        Exit Sub

UGTableUpdate_Err:
        DebugPrint "!!! Error found Globals.UGTableUpdate when updating table [" & sTableName & "] at line #" & Erl & " ..." & Err.Description
        Set rsRemote = Nothing
        Set RSLocal = Nothing
        Exit Sub
        Resume Next
        '</EhFooter>
End Sub

'Public Function DegToRad(dblAngle As Double) As Double
'    DegToRad = dblAngle * Pi / 180
'End Function

Public Function RadToDeg(dblAngle As Double) As Double
    RadToDeg = dblAngle * 180 / Pi
End Function
Public Function HandleNullString(sText As Variant)
    HandleNullString = IIf(IsNull(sText), "", sText)
End Function

