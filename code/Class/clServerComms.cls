VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clServerComms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim MsXmlHttp As MSXML2.XMLHTTP ' ServerXMLHTTP60
Dim MsXmlDoc As MSXML2.DOMDocument

Private Declare Function GetTickCount _
                Lib "kernel32.dll" () As Long

Public Function OpenSilentHttpCommsRS(sSQL As String, _
                                      bUseMsXmlHttp As Boolean) As ADODB.Recordset
        
    Dim lMilliSecsTimed As Long
    lMilliSecsTimed = GetTickCount
    m_frmDebug.DebugPrint ".................................................................."
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Currently in the code this flag is set to TRUE each time this function is called
     
    If bUseMsXmlHttp Then

        Dim tFunctionTimer As Timer
        Dim bMaxAttemptsExceeded As Boolean
        Dim iAttemptCount As Integer
        Dim lMilliSeconds As Long
        Dim lMilliSecondsWait As Long
    
        On Error GoTo MsXmlHttpconnectionfails
        m_frmDebug.DebugPrint "OpenSilentHttpCommsRS: " & sSQL
    
        bMaxAttemptsExceeded = False
        iAttemptCount = 1
    
        MsXmlHttp.Open "GET", sSQL, False
        MsXmlHttp.setRequestHeader "Cache-Control", "no-cache"
        MsXmlHttp.setRequestHeader "pragma", "no-cache"
        lMilliSeconds = CLng(g_ServerConnTimeoutSeconds) * 1000 * CLng(g_ServerConnNoOfRetries)
        'MsXmlHttp.setTimeouts lMilliSeconds, lMilliSeconds, lMilliSeconds, lMilliSeconds
        MsXmlHttp.Send

        Do Until bMaxAttemptsExceeded Or MsXmlHttp.readyState = 4

            m_frmDebug.DebugPrint "***** Attempt " & iAttemptCount & " of " & g_ServerConnNoOfRetries & " *****"
            
            'do ticker count tomorrow
            lMilliSecondsWait = GetTickCount

            Do Until MsXmlHttp.readyState = 4 Or (GetTickCount - lMilliSecondsWait) > lMilliSeconds
                DoEvents
            Loop

            If Not MsXmlHttp.readyState = 4 Then iAttemptCount = iAttemptCount + 1

            If iAttemptCount = (g_ServerConnNoOfRetries + 1) Then
                bMaxAttemptsExceeded = True
                Set OpenSilentHttpCommsRS = Nothing
                MsXmlHttp.abort
                m_frmDebug.DebugPrint "FAILURE: Response from server not received."
            End If

        Loop
    
        If Not bMaxAttemptsExceeded Then
    
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Success: Response Received
            
            m_frmDebug.DebugPrint "succeeded on attempt: " & CStr(iAttemptCount)
        
            If Not MsXmlHttp.responseText = "-1" And Not MsXmlHttp.responseText = "" Then

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Response Desirable
                
                m_frmDebug.DebugPrint "Recordset returned"

                Set OpenSilentHttpCommsRS = RecordsetFromXMLString(MsXmlHttp.responseText)
            Else
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Response Undesirable
                
                m_frmDebug.DebugPrint "(" & MsXmlHttp.responseText & ") Recordset not received from server"

                Set OpenSilentHttpCommsRS = Nothing
            End If
    
        End If

        MsXmlHttp.abort

    Else
    
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Old-school comms

        On Error GoTo oRSOpenConnectionFails
        Dim oRS As New ADODB.Recordset
        m_frmDebug.DebugPrint "RS.OPEN: " & sSQL
        oRS.Open sSQL
        Set OpenSilentHttpCommsRS = oRS
        Set oRS = Nothing
        
    End If
    
    m_frmDebug.DebugPrint "         operation took " & Round(((GetTickCount - lMilliSecsTimed) / 1000), 2) & " seconds"
    m_frmDebug.DebugPrint ".................................................................."
    '    Stop
    Exit Function
    
MsXmlHttpconnectionfails:
    m_frmDebug.DebugPrint "comms failed"
    m_frmDebug.DebugPrint "Error message: " & Err.Description
    Set OpenSilentHttpCommsRS = Nothing
    MsXmlHttp.abort
    Exit Function
     
oRSOpenConnectionFails:
    m_frmDebug.DebugPrint "comms failed"
    m_frmDebug.DebugPrint "error message: " & Err.Description
    Set OpenSilentHttpCommsRS = Nothing
    Set oRS = Nothing
    Exit Function

End Function

Public Function OpenSilentHttpCommsResponse(sSQL As String, _
                                            bdummyboolean As Boolean) As String

    On Error GoTo connectionfails
    
    Dim bMaxAttemptsExceeded As Boolean
    Dim iAttemptCount As Integer
    Dim lMilliSeconds As Long
    Dim lMilliSecsTimed As Long
    Dim lMilliSecondsWait As Long
    
    lMilliSecsTimed = GetTickCount
    m_frmDebug.DebugPrint ".................................................................."
    m_frmDebug.DebugPrint "OpenSilentHttpCommsResponse: " & sSQL
    
    bMaxAttemptsExceeded = False
    iAttemptCount = 1
    
    MsXmlHttp.Open "GET", sSQL, True
    MsXmlHttp.setRequestHeader "Cache-Control", "no-cache"
    MsXmlHttp.setRequestHeader "pragma", "no-cache"
    lMilliSeconds = CLng(g_ServerConnTimeoutSeconds) * 1000 * CLng(g_ServerConnNoOfRetries)
    'MsXmlHttp.setTimeouts lMilliSeconds, lMilliSeconds, lMilliSeconds, lMilliSeconds
    MsXmlHttp.Send

    Do Until bMaxAttemptsExceeded Or MsXmlHttp.readyState = 4

        m_frmDebug.DebugPrint "***** Attempt " & iAttemptCount & " of " & g_ServerConnNoOfRetries & " *****"
        
        lMilliSecondsWait = GetTickCount

        Do Until MsXmlHttp.readyState = 4 Or (GetTickCount - lMilliSecondsWait) > lMilliSeconds
            DoEvents
            ' m_frmDebug.DebugPrint (GetTickCount - lMilliSecondsWait) & " ... " & lMilliSeconds
        Loop

        If Not MsXmlHttp.readyState = 4 Then iAttemptCount = iAttemptCount + 1

        If iAttemptCount = (g_ServerConnNoOfRetries + 1) Then
            bMaxAttemptsExceeded = True
            OpenSilentHttpCommsResponse = ""
            MsXmlHttp.abort
            m_frmDebug.DebugPrint "FAILURE: Response from server not received."
        End If

    Loop
    
    If Not bMaxAttemptsExceeded Then
    
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Success: Response Received
            
        m_frmDebug.DebugPrint "succeeded on attempt: " & CStr(iAttemptCount)
        m_frmDebug.DebugPrint "Response returned: " & MsXmlHttp.responseText
        OpenSilentHttpCommsResponse = MsXmlHttp.responseText
       
    End If

    MsXmlHttp.abort
    m_frmDebug.DebugPrint "         operation took " & Round(((GetTickCount - lMilliSecsTimed) / 1000), 2) & " seconds"
    m_frmDebug.DebugPrint ".................................................................."
    '    Stop
    Exit Function
    
connectionfails:
    m_frmDebug.DebugPrint "comms failed"
    m_frmDebug.DebugPrint "error message: " & Err.Description
    OpenSilentHttpCommsResponse = ""
    MsXmlHttp.abort
    Resume Next
End Function

Public Function SaveSilentHttpCommsRS(oRS As ADODB.Recordset, _
                                      sWebsite As String, _
                                      dummyboolean As Boolean) As Boolean

        On Error GoTo connectionfails
        
        Dim bMaxAttemptsExceeded As Boolean
        Dim iAttemptCount As Integer
        Dim lMilliSeconds As Long
        Dim lMilliSecondsWait As Long
  
100     iAttemptCount = 1

102     If Not oRS.EOF And Not oRS.BOF Then
    
104         m_frmDebug.DebugPrint "Saving RS with " & oRS.RecordCount & " record(s) to sWebsite: " & sWebsite
        
106         MsXmlHttp.Open "POST", sWebsite, True
108         oRS.Save MsXmlDoc, 1
110         MsXmlHttp.setRequestHeader "Cache-Control", "no-cache"
112         MsXmlHttp.setRequestHeader "pragma", "no-cache"

114         lMilliSeconds = CLng(g_ServerConnTimeoutSeconds) * 1000 * CLng(g_ServerConnNoOfRetries)
116         'MsXmlHttp.setTimeouts lMilliSeconds, lMilliSeconds, lMilliSeconds, lMilliSeconds
118         MsXmlHttp.Send MsXmlDoc

120         bMaxAttemptsExceeded = False
122         SaveSilentHttpCommsRS = True

124         Do Until bMaxAttemptsExceeded Or MsXmlHttp.readyState = 4

126             m_frmDebug.DebugPrint "***** Attempt " & iAttemptCount & " of " & g_ServerConnNoOfRetries & " *****"

                lMilliSecondsWait = GetTickCount

                Do Until MsXmlHttp.readyState = 4 Or (GetTickCount - lMilliSecondsWait) > lMilliSeconds
                    DoEvents
                Loop

128             'MsXmlHttp.waitForResponse g_ServerConnTimeoutSeconds
130             iAttemptCount = iAttemptCount + 1

132             If iAttemptCount = (g_ServerConnNoOfRetries + 1) Then
134                 bMaxAttemptsExceeded = True
136                 SaveSilentHttpCommsRS = False
138                 MsXmlHttp.abort
140                 m_frmDebug.DebugPrint "Save to server not confirmed...."
                End If

            Loop
            
142         If Not bMaxAttemptsExceeded Then m_frmDebug.DebugPrint "***** Success *****"
    
        Else
    
144         m_frmDebug.DebugPrint "SaveSilentHttpCommsRS: Recordset is either BOF or EOF - this may be a desired behaviour - assuming so..."
146         SaveSilentHttpCommsRS = True
        
        End If
    
        MsXmlHttp.abort
        Exit Function
    
connectionfails:
        m_frmDebug.DebugPrint "comms failed"
        m_frmDebug.DebugPrint "error message: " & Err.Description
        SaveSilentHttpCommsRS = False
        MsXmlHttp.abort

        '<EhFooter>
        Exit Function

SaveSilentHttpCommsRS_Err:
        Err.Raise vbObjectError + 100, "OASISClient.clServerComms.SaveSilentHttpCommsRS", "clServerComms component failure"
        '</EhFooter>
End Function

Private Sub Class_Initialize()
        '<EhHeader>
        On Error GoTo Class_Initialize_Err
        '</EhHeader>
100     Set MsXmlHttp = New MSXML2.XMLHTTP
102     Set MsXmlDoc = New MSXML2.DOMDocument

        '<EhFooter>
        Exit Sub

Class_Initialize_Err:
        Err.Raise vbObjectError + 100, "OASISClient.clServerComms.Class_Initialize", "clServerComms component failure"
        '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Set MsXmlHttp = Nothing
    Set MsXmlDoc = Nothing
End Sub
