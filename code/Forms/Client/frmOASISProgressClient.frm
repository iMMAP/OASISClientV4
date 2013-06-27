VERSION 5.00
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Begin VB.Form frmOASISProgress 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Server Communication in progress...."
   ClientHeight    =   1455
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmOASISProgressClient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin OASISClient.MsgScroller MsgScroller1 
      Height          =   315
      Left            =   60
      Top             =   1110
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   5220
      Top             =   750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "***** Attempt 1 of 3 *****"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   30
      TabIndex        =   1
      Top             =   780
      Width           =   5925
   End
   Begin CONTROLSLibCtl.dxProgressBar dxProgressBar1 
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5985
      _Version        =   65536
      _cx             =   10557
      _cy             =   1349
      ForeColor       =   0
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MinPos          =   0
      MaxPos          =   100
      Pos             =   50
      Step            =   10
      ShowText        =   0   'False
      Orientation     =   0
      StartColor      =   12648384
      EndColor        =   16384
      DrawBorderStyle =   1
      ShowTextStyle   =   0
      DrawBarStyle    =   2
      DrawBarBorderStyle=   2
   End
End
Attribute VB_Name = "frmOASISProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iAttemptCount As Integer
Dim bTerminateFlag As Boolean
Dim iSetTimeoutSeconds As Integer
Dim iSetNumberOfRetries  As Integer
Dim bSilent As Boolean
Dim bShowCommString As Boolean
Dim bResultReceived As Boolean

Dim RSReturned As ADODB.Recordset
Dim StringReturned As String
Dim BooleanReturned As Boolean
Dim bBreakout As Boolean

Dim lCountOfSeconds As Long

Public Sub SetShowAdvanced(bShow As Boolean)
    
    bShowCommString = bShow
    
    If bShow Then
        Me.Height = 1860
        MsgScroller1.Visible = True
    Else
        Me.Height = 1470
        MsgScroller1.Visible = False
    End If

End Sub

Public Function GetShowAdvanced() As Boolean
    GetShowAdvanced = bShowCommString
End Function

Public Sub MakeCommunicationSilent(bMakeSilent As Boolean)
        '<EhHeader>
        On Error GoTo MakeCommunicationSilent_Err
        '</EhHeader>

100     bSilent = bMakeSilent

        '<EhFooter>
        Exit Sub

MakeCommunicationSilent_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmOASISProgress.MakeCommunicationSilent " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub InitProgressBar()
        '<EhHeader>
        On Error GoTo InitProgressBar_Err

        '</EhHeader>
        If Not bSilent Then
100         lCountOfSeconds = 0
101         Timer1.Enabled = False
102         dxProgressBar1.Pos = 10
104         Timer1.Interval = 1000 'CLng(iSetTimeoutSeconds) * 90
106         Timer1.Enabled = True
        End If

        '<EhFooter>
        Exit Sub

InitProgressBar_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmOASISProgress.InitProgressBar " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub KillTimers()
        '<EhHeader>
        On Error GoTo KillTimers_Err
        '</EhHeader>
100     If Not bSilent Then Timer1.Enabled = False
        '<EhFooter>
        Exit Sub

KillTimers_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmOASISProgress.KillTimers " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetClientDBTimeoutSettings(sTimeout As String, _
                                      sRetries As String, _
                                      sClientDbFullPath As String)
        '<EhHeader>
        On Error GoTo SetClientDBTimeoutSettings_Err
        '</EhHeader>
    
        Dim oRs As ADODB.Recordset
        Dim ooConn As ADODB.Connection
        On Error Resume Next
    
100     Set oRs = New ADODB.Recordset
102     Set ooConn = New ADODB.Connection

104     With ooConn
106         .CursorLocation = adUseClient
108         .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sClientDbFullPath
110         .Open
        End With
        
112     With oRs
114         Set .ActiveConnection = ooConn
116         .CursorType = adOpenDynamic
118         .LockType = adLockBatchOptimistic
120         .Source = "SELECT * FROM AppSettings WHERE SettingName = 'ServerConnectionParameters'"
122         .CursorLocation = adUseClient
124         .Open
        End With
        
126     If oRs.State = adStateOpen Then
    
128         If Not oRs.EOF And Not oRs.BOF Then
    
130             oRs.Fields("SettingValue1").Value = sTimeout
132             oRs.Fields("SettingValue2").Value = sRetries

            Else
134             oRs.AddNew
136             oRs.Fields("SettingName").Value = "ServerConnectionParameters"
138             oRs.Fields("SettingValue1").Value = sTimeout
140             oRs.Fields("SettingValue2").Value = sRetries
            End If
        
        End If
    
142     oRs.UpdateBatch adAffectCurrent
144     oRs.Save
    
146     Set oRs = Nothing
148     Set ooConn = Nothing

150     InitialiseProgressForm sClientDbFullPath

        '<EhFooter>
        Exit Sub

SetClientDBTimeoutSettings_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmOASISProgress.SetClientDBTimeoutSettings " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub InitialiseProgressForm(sClientDbFullPath As String)
        '<EhHeader>
        On Error GoTo InitialiseProgressForm_Err
        '</EhHeader>

        Dim oRs As ADODB.Recordset
        Dim ooConn As ADODB.Connection

100     Set oRs = New ADODB.Recordset
102     Set ooConn = New ADODB.Connection

104     bSilent = False

106     With ooConn
108         .CursorLocation = adUseClient
110         .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sClientDbFullPath
112         .Open
        End With
        
114     With oRs
116         Set .ActiveConnection = ooConn
118         .CursorType = adOpenDynamic
120         .LockType = adLockBatchOptimistic
122         .Source = "SELECT * FROM AppSettings WHERE SettingName = 'ServerConnectionParameters'"
124         .CursorLocation = adUseClient
126         .Open
        End With
    
128     If oRs.State = adStateOpen Then
    
130         If Not oRs.EOF And Not oRs.BOF Then
    
132             If IsNumeric(oRs.Fields("SettingValue1").Value) Then
134                 iSetTimeoutSeconds = oRs.Fields("SettingValue1").Value
                Else
136                 iSetTimeoutSeconds = 20
                End If
            
138             If IsNumeric(oRs.Fields("SettingValue2").Value) Then
140                 iSetNumberOfRetries = oRs.Fields("SettingValue2").Value
                Else
142                 iSetNumberOfRetries = 3
                End If

            Else
        
144             iSetTimeoutSeconds = 20
146             iSetNumberOfRetries = 3
            End If
        
        Else
    
148         iSetTimeoutSeconds = 20
150         iSetNumberOfRetries = 3
        End If
    
152     Set oRs = Nothing
154     Set ooConn = Nothing

156     Debug.Print "Server communication timeout = " & iSetTimeoutSeconds
158     Debug.Print "Server communication retries = " & iSetNumberOfRetries

        '<EhFooter>
        Exit Sub

InitialiseProgressForm_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmOASISProgress.InitialiseProgressForm " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub



Private Sub Form_Load()
    MsgScroller1.Scroll = True
    MsgScroller1.ScrollSpeed = 1
    MsgScroller1.ScrollInterval = 10
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     KillTimers
        MsgScroller1.Clear
        Set RSReturned = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmOASISProgress.Form_Unload " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Timer1_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    lCountOfSeconds = lCountOfSeconds + 1

    Debug.Print lCountOfSeconds
    If lCountOfSeconds = Round((CLng(iSetTimeoutSeconds) * 0.8) / 10, 0) Then
    
        If (dxProgressBar1.Pos + 10) > 100 Then
        dxProgressBar1.Pos = 100
        Else
        dxProgressBar1.Pos = dxProgressBar1.Pos + 10
        End If
        Me.Refresh
        lCountOfSeconds = 0
    
    End If
    
End Sub

Public Function OpenHttpCommsRS(sSQL As String, _
                                bInit As Boolean) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo OpenHttpCommsRS_Err
        '</EhHeader>

        Dim MsXmlHttp As MSXML2.ServerXMLHTTP60
100     Set MsXmlHttp = New MSXML2.ServerXMLHTTP60
    
102     If bInit Then
    
104         If Not bSilent Then
106             Me.Show vbModeless
            End If
            
108         bResultReceived = False
110         iAttemptCount = 1
            Set RSReturned = Nothing
        End If

112     If Not bSilent Then

            InitProgressBar
114         Label1.caption = "***** Attempt " & iAttemptCount & " of " & iSetNumberOfRetries & " *****"
116         Debug.Print Label1.caption
        
            On Error Resume Next

120         If Not MsgScroller1.ListCount = 0 Then MsgScroller1.Clear
122         MsgScroller1.AddItem sSQL, "test", 0
124         Me.Refresh
        End If
        
        Debug.Print "sql: " & sSQL
        On Error GoTo connectionfails
126     MsXmlHttp.Open "GET", sSQL, True
128     MsXmlHttp.setRequestHeader "Cache-Control", "no-cache"
130     MsXmlHttp.setRequestHeader "pragma", "no-cache"
132     MsXmlHttp.setTimeouts CLng(iSetTimeoutSeconds) * 1000, CLng(iSetTimeoutSeconds) * 1000, CLng(iSetTimeoutSeconds) * 1000, CLng(iSetTimeoutSeconds) * 1000
134     MsXmlHttp.Send

136     If MsXmlHttp.waitForResponse(iSetTimeoutSeconds) Then
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Success: Response Received
            
            On Error GoTo OpenHttpCommsRS_Err
140         Debug.Print "succeeded on attempt: " & CStr(iAttemptCount)
        
142         If Not MsXmlHttp.responseText = "-1" And Not MsXmlHttp.responseText = "" Then

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Response Desirable
                
144             Debug.Print "Recordset returned"

146             If Not bResultReceived Then Set RSReturned = RecordsetFromXMLString(MsXmlHttp.responseText)
            Else
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Response Undesirable
                
148             Debug.Print "(" & MsXmlHttp.responseText & ") Recordset not received from server"

150             If Not bResultReceived Then Set RSReturned = Nothing
            End If

            bResultReceived = True
        Else
            
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ' Failure: Response not received
            
            On Error GoTo OpenHttpCommsRS_Err
152         iAttemptCount = iAttemptCount + 1
154         Debug.Print "failed attempt: " & CStr(iAttemptCount - 1)
       
156         If iAttemptCount = (iSetNumberOfRetries + 1) Then
    
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Exit function - maximum retries reached
                
158             If Not bResultReceived Then Set RSReturned = Nothing

            Else
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Retry function
                
160             Call OpenHttpCommsRS(sSQL, False)

            End If
    
        End If

162     MsXmlHttp.abort
164     Set MsXmlHttp = Nothing

166     If Not bSilent Then Me.Hide
        Set OpenHttpCommsRS = RSReturned
        KillTimers
        '<EhFooter>
        Exit Function
        
connectionfails:
        Debug.Print "Connection failed on attempt " & iAttemptCount & ". Err Message: " & Err.Description
        MsXmlHttp.abort
        Set MsXmlHttp = Nothing

        If Not bSilent Then Me.Hide
        Set OpenHttpCommsRS = RSReturned
        KillTimers
        Exit Function

OpenHttpCommsRS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmOASISProgress.OpenHttpCommsRS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function SaveHttpCommsRS(oRs As ADODB.Recordset, _
                                sWebsite As String, _
                                bInit As Boolean) As Boolean
        '<EhHeader>
        On Error GoTo SaveHttpCommsRS_Err
        '</EhHeader>

        Dim MsXmlHttp As MSXML2.ServerXMLHTTP60
        Dim MsXmlDoc As MSXML2.DOMDocument
        
        Dim lMilliSeconds As Long
    
100     Set MsXmlHttp = New MSXML2.ServerXMLHTTP60
102     Set MsXmlDoc = New MSXML2.DOMDocument
    
104     If bInit Then
    
106         If Not bSilent Then
108             Me.Show vbModeless ', frmContainer
            End If

114         iAttemptCount = 1
    
        End If

116     If Not oRs.EOF And Not oRs.BOF Then
    
122         Debug.Print "Saving RS with " & oRs.RecordCount & " record(s) to sWebsite: " & sWebsite
            On Error Resume Next
            
            If Not bSilent Then

124             If Not MsgScroller1.ListCount = 0 Then MsgScroller1.Clear
126             MsgScroller1.AddItem "Saving RS with " & oRs.RecordCount & " record(s) to sWebsite: " & sWebsite, "test", 0
128             Me.Refresh

            End If
        
            On Error GoTo connectionfails
130         MsXmlHttp.Open "POST", sWebsite, True
132         oRs.Save MsXmlDoc, 1
134         MsXmlHttp.setRequestHeader "Cache-Control", "no-cache"
136         MsXmlHttp.setRequestHeader "pragma", "no-cache"

            lMilliSeconds = CLng(iSetTimeoutSeconds) * 1000 * CLng(iSetNumberOfRetries)
138         MsXmlHttp.setTimeouts lMilliSeconds, lMilliSeconds, lMilliSeconds, lMilliSeconds
140         MsXmlHttp.Send MsXmlDoc

142         bBreakout = False
144         SaveHttpCommsRS = True

146         Do Until bBreakout Or MsXmlHttp.readyState = 4

                If Not bSilent Then
                    Label1.caption = "***** Attempt " & iAttemptCount & " of " & iSetNumberOfRetries & " *****"
                    Debug.Print Label1.caption
                End If

148             If Not bSilent Then InitProgressBar
150             MsXmlHttp.waitForResponse iSetTimeoutSeconds
152             iAttemptCount = iAttemptCount + 1

154             If iAttemptCount = (iSetNumberOfRetries + 1) Then
156                 bBreakout = True
158                 SaveHttpCommsRS = False
                    MsXmlHttp.abort
                    Debug.Print "Save to server not confirmed...."
                End If

            Loop
    
        Else
    
160         Debug.Print "SaveHttpCommsRS: Recordset is either BOF or EOF - this may be a desired behaviour"

162         SaveHttpCommsRS = True
        
        End If
    
        On Error Resume Next

164     If Not bSilent Then Me.Hide
166     Set MsXmlHttp = Nothing
168     Set MsXmlDoc = Nothing
        KillTimers

        '<EhFooter>
        Exit Function
        
connectionfails:

        On Error Resume Next
        Debug.Print "Connection failed on attempt " & iAttemptCount & ". Err Message: " & Err.Description
        SaveHttpCommsRS = False
        Resume Next

        If Not bSilent Then Me.Hide
        Set MsXmlHttp = Nothing
        Set MsXmlDoc = Nothing
        KillTimers
        Exit Function

SaveHttpCommsRS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmOASISProgress.SaveHttpCommsRS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function OpenHttpCommsResponse(sSQL As String, _
                                      bInit As Boolean) As String
        '<EhHeader>
        On Error GoTo OpenHttpCommsResponse_Err
        '</EhHeader>

        Dim MsXmlHttp As MSXML2.ServerXMLHTTP60
100     Set MsXmlHttp = New MSXML2.ServerXMLHTTP60

102     If bInit Then
    
104         If Not bSilent Then
106             Me.Show vbModeless ', frmContainer
108             bResultReceived = False
                StringReturned = "-1"
            End If

110         iAttemptCount = 1
    
        End If

112     If Not bSilent Then
            InitProgressBar
    
114         Label1.caption = "***** Attempt " & iAttemptCount & " of " & iSetNumberOfRetries & " *****"
116         Debug.Print Label1.caption
118         Debug.Print "sql: " & sSQL
            On Error Resume Next
       
120         If Not MsgScroller1.ListCount = 0 Then MsgScroller1.Clear
122         MsgScroller1.AddItem sSQL, "test", 0
124         Me.Refresh
        End If

        On Error GoTo connectionfails
126     MsXmlHttp.Open "GET", sSQL, True
128     MsXmlHttp.setRequestHeader "Cache-Control", "no-cache"
130     MsXmlHttp.setRequestHeader "pragma", "no-cache"
132     MsXmlHttp.setTimeouts CLng(iSetTimeoutSeconds) * 1000, CLng(iSetTimeoutSeconds) * 1000, CLng(iSetTimeoutSeconds) * 1000, CLng(iSetTimeoutSeconds) * 1000
134     MsXmlHttp.Send

        If MsXmlHttp.waitForResponse(iSetTimeoutSeconds) Then

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Success: Response Received
            
            On Error GoTo OpenHttpCommsResponse_Err
138         Debug.Print "succeeded on attempt: " & CStr(iAttemptCount) & " returned: " & MsXmlHttp.responseText

140         If Not bResultReceived Then StringReturned = MsXmlHttp.responseText
142         bResultReceived = True
        
        Else
        
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Failure: Response not received
            
            On Error GoTo OpenHttpCommsResponse_Err
144         iAttemptCount = iAttemptCount + 1
146         Debug.Print "failed attempt: " & CStr(iAttemptCount - 1)
        
148         If iAttemptCount = (iSetNumberOfRetries + 1) Then

                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Exit function - maximum retries reached
                    
150             If Not bResultReceived Then StringReturned = "-1"

            Else
            
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ' Retry function
                    
152             Call OpenHttpCommsResponse(sSQL, False)

            End If
        End If
        
154     If Not bSilent Then Me.Hide
156     MsXmlHttp.abort
158     Set MsXmlHttp = Nothing
        OpenHttpCommsResponse = StringReturned
        KillTimers
        Exit Function

connectionfails:
        Debug.Print "Connection failed on attempt " & iAttemptCount & ". Err Message: " & Err.Description
        
        If Not bSilent Then Me.Hide
        MsXmlHttp.abort
        Set MsXmlHttp = Nothing
        OpenHttpCommsResponse = StringReturned
        KillTimers
        Exit Function

OpenHttpCommsResponse_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmOASISProgress.OpenHttpCommsResponse " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function RecordsetFromXMLString(sXML As String) As Recordset
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
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmOASISProgress.RecordsetFromXMLString " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function






