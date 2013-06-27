Attribute VB_Name = "modSynch"
Option Explicit

Public Type sxHistory
         sID As String
         sGUID As String
         By As String
         sequence As Long
         when As String
         deleted As String
         updates As String
         noconflicts As String
End Type


Public Const XML_FOLDER As String = "\xml\"
Public Const XSL_FOLDER As String = "\xsl\"
Public Const SOURCES_XML_PATH As String = XML_FOLDER & "sources.xml"
Public Const RSS_SCHEMA As String = "http://my.netscape.com/rdf/simple/0.9/"
Public Const RDF_SCHEMA As String = "http://www.w3.org/1999/02/22-rdf-syntax-ns#"
Public Const SYNDICATION_SCHEMA As String = "http://purl.org/rss/modules/syndication/"

' Send message to a window, used to auto size listview columns
' based on contents
Private Declare Function SendMessage Lib "user32.dll" _
     Alias "SendMessageA" (ByVal hwnd As Long, _
     ByVal Msg As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long

Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE_USEHEADER = -2

Public gbCancelled As Boolean

Public Sub Main()
    
    frmSynchReader.Show
    'frmSynchReader.ClearAll
End Sub

Public Sub ResizeLVColumns(ByRef lvw As MSComctlLib.ListView)
    Dim lngCol As Long
    Dim lngHwnd As Long
    
    lngHwnd = lvw.hwnd
    For lngCol = 1 To lvw.ColumnHeaders.Count
        SendMessage lngHwnd, _
               LVM_SETCOLUMNWIDTH, _
               lngCol - 1, _
               LVSCW_AUTOSIZE_USEHEADER
    Next lngCol
End Sub

Public Function RFC3339DateTime() As String
    '  Get the current timedate and format it as RFC 3339

    Dim g_CurrentDateTime  As Date
    Dim iYear As Integer
    Dim sMonth As String
    Dim sDay As String
    Dim shour As String
    Dim sMinute As String
    Dim sSec As String

    iYear = Year(date)
    sMonth = Month(date)
    sDay = Day(date)
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

Public Function RFC3339DateTimeEX(iYear As Integer, sMonth As String, sDay As String, shour As String, sMinute As String, sSec As String) As String
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

