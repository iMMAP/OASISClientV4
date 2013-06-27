Attribute VB_Name = "Conversions"
Option Explicit

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

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(31) As Integer

    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long

End Type

Private Declare Function GetTimeZoneInformation _
    Lib "KERNEL32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Const TIME_ZONE_ID_INVALID& = &HFFFFFFFF
Private Const TIME_ZONE_ID_STANDARD& = 1
Private Const TIME_ZONE_ID_UNKNOWN& = 0

Private Const TIME_ZONE_ID_DAYLIGHT& = 2


Private Declare Function InternetTimeToSystemTime Lib "wininet.dll" _
       (ByVal lpszTime As String, _
        ByRef pst As SYSTEMTIME, _
        ByVal dwReserved As Long) _
        As Long

Public Function RGB2HTMLColor(R As Byte, G As Byte, _
   b As Byte) As String


'INPUT: Numeric (Base 10) Values for R, G, and B)

'RETURNS:
'A string that can be used as an HTML Color
'(i.e., "#" + the Hexadecimal equivalent)

Dim HexR, HexB, HexG As Variant
Dim sTemp As String
Dim HexA As Variant
On Error GoTo ErrorHandler

 'R
 HexR = Hex(R)
 If Len(HexR) < 2 Then HexR = "0" & HexR
 
 'Get Green Hex
 HexG = Hex(G)
If Len(HexG) < 2 Then HexG = "0" & HexG

HexB = Hex(b)
If Len(HexB) < 2 Then HexB = "0" & HexB

HexA = Hex(125)
If Len(HexA) < 2 Then HexA = "0" & HexA



    RGB2HTMLColor = "&" & HexA & HexR & HexG & HexB
ErrorHandler:
End Function

Public Function GetGmtTime(Optional StartingDate As Variant) As Date

    'Parameters: StartingDate (Optional).  The function will figure
    'out GMT time based on StartingDate
    'If StartingDate is not provided, the current time will be used
    
    Dim Difference As Long

    
    Difference = GetTimeDifference()
    
    If IsMissing(StartingDate) Then
        'use current time
        GetGmtTime = DateAdd("s", -Difference, Now)
    Else
        'use StartingDate

        GetGmtTime = DateAdd("s", -Difference, StartingDate)
    End If
End Function

Public Function GetTimeDifference() As Long

    'Returns  the time difference between
    'local & GMT time in seconds.
    'If the  result is negative, your time zone
    'lags behind GMT zone.
    'If the  result is positive, your time zone is ahead.
    
    Dim tz As TIME_ZONE_INFORMATION
    Dim retcode As Long

    Dim Difference As Long
    
    'retrieve the time zone information
    retcode = GetTimeZoneInformation(tz)
    
    'convert to seconds

    Difference = -tz.Bias * 60
    'cache the result

    GetTimeDifference = Difference
    
    'if we are in daylight  saving time, apply the bias.
    If retcode = TIME_ZONE_ID_DAYLIGHT& Then

        If tz.DaylightDate.wMonth <> 0 Then
            'if tz.DaylightDate.wMonth = 0 then the daylight
            'saving time change doesn't occur
            GetTimeDifference = Difference - tz.DaylightBias * 60
        End If

    End If
    
End Function

Public Function GetTimeHere(gmtTime As Date) As Date

    'Parameters:    gmtTime - Provides the time & date
    'from which to make calculations
    'Returns the time in your local time zone
    'which corresposponds to GMT time
    
    Dim Differerence As Long

    
    Differerence = GetTimeDifference()
    GetTimeHere = DateAdd("s", Differerence, gmtTime)
    
End Function

Public Function InternetTimeToVbLocalTime(ByVal DateString As String) As Date

    'Currently we process 2 formats
'    'Rfc822 and Iso8601
'
'    'Iso8601 is either 1997-07-16T19:20:30+01:00 (25 bytes) or 1997-07-16T19:20:30Z (20 bytes)
'    'Rfc822 is Tue, 23 Sep 2003 13:21:00 -07:00 (32 bytes) or Tue, 23 Sep 2003 13:21:00 GMT (29 bytes)
'
'    'The key difference is that Iso8661 time has a latin letter T in position 11
'
'
'
'    DateString = Trim$(DateString)
'
'    If Mid$(DateString, 11, 1) = "T" Then
'        InternetTimeToVbLocalTime = Iso8601TimeToLocalVbTime(DateString)
'    Else
'        InternetTimeToVbLocalTime = Rfc822TimeToLocalVbTime(DateString)
'    End If

    
End Function

Public Function TimeString(Seconds As Long, Optional Verbose _
As Boolean = False) As String

'if verbose = false, returns
'something like
'02:22.08
'if true, returns
'2 hours, 22 minutes, and 8 seconds

Dim lHrs As Long
Dim lMinutes As Long
Dim lSeconds As Long

lSeconds = Seconds

lHrs = Int(lSeconds / 3600)
lMinutes = (Int(lSeconds / 60)) - (lHrs * 60)
lSeconds = Int(lSeconds Mod 60)

Dim sAns As String


If lSeconds = 60 Then
    lMinutes = lMinutes + 1
    lSeconds = 0
End If

If lMinutes = 60 Then
    lMinutes = 0
    lHrs = lHrs + 1
End If

sAns = Format(CStr(lHrs), "#####0") & ":" & _
  Format(CStr(lMinutes), "00") & "." & _
  Format(CStr(lSeconds), "00")

If Verbose Then sAns = TimeStringtoEnglish(sAns)
TimeString = sAns

End Function

Private Function TimeStringtoEnglish(sTimeString As String) As String

    Dim sAns As String
    Dim sHour, sMin As String, sSec As String
    Dim iTemp As Integer, sTemp As String
    Dim iPos As Integer
    iPos = InStr(sTimeString, ":") - 1

    sHour = Left$(sTimeString, iPos)

    If CLng(sHour) <> 0 Then
        sAns = CLng(sHour) & " hour"

        If CLng(sHour) > 1 Then sAns = sAns & "s"
        sAns = sAns & ", "
    End If

    sMin = Mid$(sTimeString, iPos + 2, 2)

    iTemp = sMin

    If sMin = "00" Then
        sAns = IIf(Len(sAns), sAns & "0 minutes, and ", "")
    Else
        sTemp = IIf(iTemp = 1, " minute", " minutes")
        sTemp = IIf(Len(sAns), sTemp & ", and ", sTemp & " and ")
        sAns = sAns & Format$(iTemp, "##") & sTemp
    End If

    iTemp = Val(Right$(sTimeString, 2))
    sSec = Format$(iTemp, "#0")
    sAns = sAns & sSec & " second"

    If iTemp <> 1 Then sAns = sAns & "s"

    TimeStringtoEnglish = sAns

End Function


Public Function FixDBDate(DumbDate As Variant) As Variant

    Dim dAns As Variant
     'If date is invalid, return null

    If Not IsDate(DumbDate) Then
        dAns = Null

      'If date = default used by some datasources, return null
    ElseIf CStr(DumbDate) = "12:00:00 AM" Then
        dAns = Null
     'if date is valid, convert it to date type and return it.
    Else
        dAns = CDate(DumbDate)
    End If
    
    FixDBDate = dAns
    
End Function


Public Function BytesToMegabytes(Bytes As Double) As Double
   'This function gives an estimate to two decimal
   'places.  For a more precise answer, format to
   'more decimal places or just return dblAns
 
  Dim dblAns As Double
  dblAns = (Bytes / 1024) / 1024
  BytesToMegabytes = Format(dblAns, "###,###,##0.00")
  
End Function

Public Function TimeEarlierThan(TimeString As String, _
  CompareTo As String) As Boolean
'Takes two time strings as returns true if the
'first is earlier than the second

'Example Usage:
'Checks if it is before 6:00 PM
'If TimeEarlierThan(Format(Now, "Short Time"), "6:00 PM") Then
'    MsgBox "Good Afternoon (or morning)"
'Else
'    MsgBox "Good Evening"
'End If

If Not IsTime(TimeString) Or Not IsTime(CompareTo) Then _
    Exit Function

TimeEarlierThan = CDate(TimeString) < CDate(CompareTo)

End Function

Private Function IsTime(sTime As String) As Boolean

'http://www.freevbcode.com/ShowCode.Asp?ID=1321
'by Phil Fresle

    If Left(Trim(sTime), 1) Like "#" Then
        IsTime = IsDate(date & " " & sTime)
    End If
End Function

Public Function ChooseNotSelect(Choice As Integer) As String

'sAns = Choose(Choice, "1", "2", "3", "4")
'ChooseNotSelect = "" & sAns
'
''*********************************************
''INSTEAD OF THIS CODE
''Dim sAns As String
''Select Case Choice
''Case 1
''sAns = "1"
''Case 2
''sAns = "2"
''Case 3
''sAns = "3"
''Case 4
''sAns = "4"
''Case Else
''sAns = vbNullString
''
''End Select
''ChooseNotSelect = sAns

End Function

Public Sub ChooseNotDimStaticArray()
Dim iArray(1 To 10) As Integer
Dim i As Integer

For i = 1 To 10
    iArray(i) = Choose(i, 1000, 2000, 3000, 4000, 5000, 6000, 7000, 8000, _
        9000, 10000)
Next
'*********************************************************
'INSTEAD OF THIS CODE
'iArray(1) = 1000
'iArray(2) = 2000
'iArray(3) = 3000
'iArray(4) = 4000
'iArray(5) = 5000
'iArray(6) = 6000
'iArray(7) = 7000
'iArray(8) = 8000
'iArray(9) = 9000
'iArray(10) = 10000
End Sub

Public Function AppPath() As String
    
    Dim sAns As String
    sAns = g_sAppPath
    If Right(g_sAppPath, 1) <> "\" Then sAns = sAns & "\"
    AppPath = sAns

End Function
