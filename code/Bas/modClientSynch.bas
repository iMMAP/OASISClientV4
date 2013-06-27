Attribute VB_Name = "modClientSynch"
Option Explicit

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
