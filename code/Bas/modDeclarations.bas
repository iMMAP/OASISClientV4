Attribute VB_Name = "modDeclarations"
Option Explicit

Public Declare Function InternetGetConnectedState _
               Lib "wininet" (lpdwFlags As Long, _
                              ByVal dwReserved As Long) As Boolean

Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8
Public Const INTERNET_CONNECTION_CONFIGURED = &H40
Public Const INTERNET_CONNECTION_OFFLINE = &H20
Public Const INTERNET_RAS_INSTALLED = &H10

Public Const MAX_PATH                   As Long = 260
Public Const ERROR_SUCCESS              As Long = 0

'Treat entire URL param as one URL segment
Public Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
Public Const URL_ESCAPE_PERCENT         As Long = &H1000
Public Const URL_UNESCAPE_INPLACE       As Long = &H100000

'escape #'s in paths
Public Const URL_INTERNAL_PATH          As Long = &H800000
Public Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
Public Const URL_ESCAPE_SPACES_ONLY     As Long = &H4000000
Public Const URL_DONT_SIMPLIFY          As Long = &H8000000

Public m_oAES As clsAES

Declare Function CoCreateGuid _
        Lib "ole32.dll" (pguid As GUID) As Long
Declare Function StringFromGUID2 _
        Lib "ole32.dll" (rguid As Any, _
                         ByVal lpstrClsId As Long, _
                         ByVal cbMax As Long) As Long
Declare Function ShellExecute _
        Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hWnd As Long, _
                               ByVal lpOperation As String, _
                               ByVal lpFile As String, _
                               ByVal lpParameters As String, _
                               ByVal lpDirectory As String, _
                               ByVal nShowCmd As Long) As Long

'GUID STRUCT
Private Type GUID
    data1 As Long
    data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Public g_bHasEncrypt As Boolean
Public g_sKey As String

Public Const RT_enmPaddy = "00"
Public Const RT_IncidentSynch = "01"
Public Const RT_SQLLyrSynch = "02"
Public Const RT_GeoMarks = "03"
Public Const RT_InternetConnectionCheck = "04"
Public Const RT_IncidentNotifier = "05"
Public Const RT_IncidentDeleted = "06"
Public Const RT_IncidentEdited = "07"
Public Const RT_GeoMarksDeleted = "08"
Public Const RT_GeoMarksEdited = "09"

Public Function CheckEncrypt(sValue As String) As String
        '<EhHeader>
        On Error GoTo CheckEncrypt_Err
        '</EhHeader>
            
100     If g_bHasEncrypt Then
102         CheckEncrypt = m_oAES.AESEncyptString(sValue, g_sKey)
        Else
104         CheckEncrypt = sValue
        End If

        '<EhFooter>
        Exit Function

CheckEncrypt_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Synch.modDeclarations.CheckEncrypt " & "at line " & Erl
        Resume Next
        '</EhFooter>
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

    GetGuid = Left$(sGUID, lRet)
End Function

Public Function RFC3339DateTime() As String
    '  Get the current timedate and format it as RFC 3339

    Dim g_CurrentDateTime  As Date
    Dim iYear As Integer
    Dim sMonth As String
    Dim sDay As String
    Dim shour As String
    Dim sMinute As String
    Dim sSec As String

    iYear = Year(Date)
    sMonth = Month(Date)
    sDay = Day(Date)
    shour = Hour(Now())
    sMinute = Minute(Now())
    sSec = Second(Now())

    If (iYear < 70) Then
        iYear = iYear + 2000
    ElseIf (iYear < 1900) Then
        iYear = iYear + 1900
    End If

    'var g_Month = g_CurrentDateTime.getMonth() + 1;

    If (CInt(sMonth) <= 9) Then
        sMonth = "0" + sMonth
    End If

    If (CInt(sDay) <= 9) Then
        sDay = "0" + sDay
    End If

    If (CInt(shour) <= 9) Then
        shour = "0" + shour
    End If

    If (CInt(sMinute) <= 9) Then
        sMinute = "0" + sMinute
    End If

    If (CInt(sSec) <= 9) Then
        sSec = "0" + sSec
    End If

    RFC3339DateTime = iYear & "-" & sMonth & "-" & sDay & "T" & shour & ":" & sMinute & ":" & sSec & "Z"

End Function

Public Function RFC3339DateTimeEX(iYear As Integer, _
                                  sMonth As String, _
                                  sDay As String, _
                                  shour As String, _
                                  sMinute As String, _
                                  sSec As String) As String
    '  Get the current timedate and format it as RFC 3339

    Dim g_CurrentDateTime  As Date

    If (iYear < 70) Then
        iYear = iYear + 2000
    ElseIf (iYear < 1900) Then
        iYear = iYear + 1900
    End If

    'var g_Month = g_CurrentDateTime.getMonth() + 1;

    If (CInt(sMonth) <= 9) Then
        sMonth = "0" + sMonth
    End If

    If (CInt(sDay) <= 9) Then
        sDay = "0" + sDay
    End If

    If (CInt(shour) <= 9) Then
        shour = "0" + shour
    End If

    If (CInt(sMinute) <= 9) Then
        sMinute = "0" + sMinute
    End If

    If (CInt(sSec) <= 9) Then
        sSec = "0" + sSec
    End If

    RFC3339DateTimeEX = iYear & "-" & sMonth & "-" & sDay & "T" & shour & ":" & sMinute & ":" & sSec & "Z"

End Function

Public Sub SetNewSynchDBElement(oCN As ADODB.Connection, _
                                sNewGUID As String, _
                                sID As String, _
                                sTitle As String, _
                                sDescription As String, _
                                sBy As String, _
                                sRFC3339DateTime As String, _
                                sTableName As String, _
                                bIsGeoLayer As Boolean, _
                                Optional supdates As String = "'false'", _
                                Optional sDeleted As String = "'false'", _
                                Optional lSeq As Long = 1, Optional sSynchHistPrefix As String = "")
        '<EhHeader>
        On Error GoTo SetNewSynchDBElement_Err
        '</EhHeader>
        Dim sSQL As String

100     sSQL = "INSERT INTO [" & sSynchHistPrefix & "SynchHistory] (sID, sGUID, sTableName, swhen, sStatus, sequence"
102     sSQL = sSQL & ", sBy, sdelete, updates, noconflict)"
104     sSQL = sSQL & " VALUES ('" & sID & "','" & sNewGUID & "','" & sTableName
106     sSQL = sSQL & "','" & sRFC3339DateTime & "','pending'," & lSeq & ",'" & sDescription & " " & sBy & "'," & sDeleted & "," & supdates & ",'false')"
        
108     oCN.Execute sSQL

        '<EhFooter>
        Exit Sub

SetNewSynchDBElement_Err:
        MsgBox Err.Description & vbCrLf & "in OASIS_Synch.modDeclarations.SetNewSynchDBElement " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

