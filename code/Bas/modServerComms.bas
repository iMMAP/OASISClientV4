Attribute VB_Name = "modServerComms"
Option Explicit
Private Const HTTPREQUEST_PROXYSETTING_PROXY = 2

Public Sub CheckProxySettings()
        '<EhHeader>
        On Error GoTo CheckProxySettings_Err
        '</EhHeader>
        Dim objShell, bKey
        Dim sKey As String
        Dim sProxyArray() As String
        Dim i As Integer
        
        Dim bAdvancedProxySettings As Boolean
        
100     g_bProxyEnabled = False
102     g_sProxyIP = ""
104     g_sProxyPort = ""
    
106     Set objShell = CreateObject("WScript.Shell")
    
108     If CStr(objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")) = "1" Then
    
110         bKey = objShell.RegRead("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer")
112         sKey = CStr(bKey)

114         If Len(sKey) > 1 And InStr(sKey, ";") > 0 Then

116             sProxyArray = Split(sKey, ";")
                
118             Do Until i = UBound(sProxyArray)
                
120                 If left(sProxyArray(i), 5) = "http=" Then
122                     sKey = Replace(sProxyArray(i), "http=", "")
                        Exit Do
                    End If

124                 i = i + 1
                Loop
                
            End If
            
126         sKey = Replace(sKey, "http=", "")
            
128         If Len(sKey) > 1 And InStr(sKey, ":") > 0 Then

130             sProxyArray = Split(sKey, ":")
132             g_bProxyEnabled = True
134             g_sProxyIP = sProxyArray(0)
136             g_sProxyPort = sProxyArray(1)

                SafeMoveFirst g_RSAppSettings
                g_RSAppSettings.Open "SELECT * FROM AppSettings", m_Cnn
                g_RSAppSettings.Find "SettingName = 'ProxyAuthentication'"

                If Not g_RSAppSettings.EOF Then
                    g_sProxyUser = g_RSAppSettings.Fields("SettingValue1").value
                    g_sProxyPass = g_RSAppSettings.Fields("SettingValue2").value
                    g_iProxyAuthType = g_RSAppSettings.Fields("SettingValue3").value
                Else
                    g_sProxyUser = ""
                    g_sProxyPass = ""
                End If
        
138             DebugPrint ">>>>>> Proxy detected at [" & sProxyArray(0) & ":" & sProxyArray(1) & "]"
    
            Else
140             DebugPrint ">>>>>> Proxy detected but not for http - ignoring"
            End If
    
        Else
142         DebugPrint ">>>>>> No proxy detected"
        End If
    
144     Set objShell = Nothing

        '<EhFooter>
        Exit Sub

CheckProxySettings_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in OASISClient.modServerComms.CheckProxySettings " & _
         "at line " & Erl
        'Resume Next
        '</EhFooter>
End Sub

Public Function OpenServerResponseCompressed(sWebsite As String, _
                                              sFunction As String, _
                                              sParameter As String) As String
        '<EhHeader>
        On Error GoTo OpenServerResponseCompressed_Err
        '</EhHeader>

        Dim oHttp As New WinHttpRequest
        Dim sRetval As String
        Dim sURLEncoded As String
        Dim bByte() As Byte
        Dim sSent As String
        Dim OASStringCompression As New OASISStringCompression.OASISCompression
        Dim sStringFromUTF As String
        
        sWebsite = Replace(sWebsite, "http://http://", "http://")
100     PrepareHttpComms oHttp, sWebsite, False
102     sURLEncoded = sFunction & "=" & sParameter
104     bByte = OASStringCompression.CompressStringToByteArray(sURLEncoded)
105     sSent = OASStringCompression.ConvertByteArrayToString(bByte)
106     oHttp.Send bByte
109     If Len(oHttp.responseBody) > 0 Then sStringFromUTF = OASStringCompression.ConvertByteArrayToString(oHttp.responseBody)
110     OpenServerResponseCompressed = OASStringCompression.DecompressStringToString(sStringFromUTF)

111     DebugPrint ">>>> OpenServerResponseCompressed before encryption: " & sURLEncoded, True
112     DebugPrint ">>>>> OpenServerResponseCompressed after encryption: " & sSent, True
114     DebugPrint ">>>> Data received encrypted: " & oHttp.responseBody, True
116     DebugPrint ">>>> Data received decrypted: " & OpenServerResponseCompressed, True
             
118     oHttp.abort
120     Set oHttp = Nothing
          
        '<EhFooter>
        Exit Function

OpenServerResponseCompressed_Err:
        DebugPrint "OpenServerResponseCompressed_Err: (" & Erl & ") " & Err.Description
        On Error Resume Next
        oHttp.abort
        Set oHttp = Nothing
        '</EhFooter>
End Function

Public Function OpenServerRSCompressed(sWebsite As String, _
                              sFunction As String, _
                              sParameter As String) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo OpenServerRSCompressed_Err
        '</EhHeader>

        Dim sRetval As String
        sRetval = OpenServerResponseCompressed(sWebsite, sFunction, sParameter)
           
128     If Not sRetval = "-1" And Not sRetval = "" Then
130         Set OpenServerRSCompressed = RecordsetFromXMLString(sRetval)
        Else
132         Set OpenServerRSCompressed = Nothing
        End If
        
        If OpenServerRSCompressed Is Nothing Then
            DebugPrint "!!!! OpenServerRSCompressed failed to create a recordset from retval: " & sRetval, True
        Else
            DebugPrint ">>>> OpenServerRSCompressed created a recordset from retval: " & sRetval, True
        End If
        '<EhFooter>
        Exit Function

OpenServerRSCompressed_Err:
        DebugPrint "!!     -- OpenServerRSCompressed_Err: (" & Erl & ") " & Err.Description
        'Resume Next
        '</EhFooter>
End Function

Private Sub PrepareHttpComms(oHttp As WinHttpRequest, sWebsite As String, bUpload As Boolean)
        '<EhHeader>
        On Error GoTo PrepareHttpComms_Err
        '</EhHeader>
    
        If Not g_ProxyChecked Then
            CheckProxySettings
            g_ProxyChecked = True
        End If
100     oHttp.abort
        Set oHttp = Nothing
        Set oHttp = New WinHttpRequest
102     If g_bProxyEnabled Then oHttp.setProxy HTTPREQUEST_PROXYSETTING_PROXY, g_sProxyIP & ":" & g_sProxyPort, "*.microsoft.com"
104     oHttp.setTimeouts -1, 360000, 360000, 360000
106     oHttp.Option(WinHttpRequestOption_EnableHttp1_1) = False
108     oHttp.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = 13056  'dec equivalent to hex 0x3300
110     oHttp.Open "POST", sWebsite, False
114     oHttp.setRequestHeader "Content-type", "application/x-www-form-urlencoded;charset=UTF-8"
116     oHttp.setRequestHeader "Expires", "0"
118     oHttp.setRequestHeader "Cache-Control", "no-cache"
120     oHttp.setRequestHeader "Pragma", "no-cache"
        'oHttp.SetCredentials sUser, sPwd, HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
    
        If Len(g_sProxyUser) > 0 Then
        'HTTPREQUEST_SETCREDENTIALS_FOR_SERVER = 0;
        'HTTPREQUEST_SETCREDENTIALS_FOR_PROXY = 1;
            oHttp.SetCredentials g_sProxyUser, g_sProxyPass, g_iProxyAuthType
        End If
        
        '<EhFooter>
        Exit Sub

PrepareHttpComms_Err:
        DebugPrint "!!     -- PrepareHttpComms_Err: (" & Erl & ") " & Err.Description
        Resume Next
        '</EhFooter>
End Sub

Public Function RecordsetFromXMLString(sXML As String) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo RecordsetFromXMLString_Err
        '</EhHeader>

        Dim oStream As ADODB.Stream
        Dim oRecordset As ADODB.Recordset
        
100     Set oStream = New ADODB.Stream
102     oStream.Open
104     oStream.WriteText sXML   'Give the XML string to the ADO Stream
106     oStream.Position = 0    'Set the stream position to the start

108     Set oRecordset = New ADODB.Recordset
110     oRecordset.Open oStream    'Open a recordset from the stream
112     oStream.Close

114     Set oStream = Nothing
116     Set RecordsetFromXMLString = oRecordset  'Return the recordset
118     Set oRecordset = Nothing

        '<EhFooter>
        Exit Function

RecordsetFromXMLString_Err:
       ' MsgBox Err.Description & vbCrLf & "in OASISClient.modServerComms.RecordsetFromXMLString " & "at line " & Erl
        Set RecordsetFromXMLString = Nothing
        '</EhFooter>
End Function

