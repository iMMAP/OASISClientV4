Attribute VB_Name = "Programs"
Option Explicit
Global Const gcReleaseVersion = True
Global gsApplicationName As String
Global gsPathAndDatabaseName As String
Global Const gcPlaceHolder = "<???>"
Global Const gcNoFilter = -10000
Global Const gcQuestionMarks = -5000
Global Const gcnNumberOfJobRemarkFields = 6

Global gdbMain As Database
Global grsSendlog As Dao.Recordset

Global gSMS As Booster

Type DeferredDeliveryTimeSettings
    bUseDeferredDeliveryTime As Boolean
    nTimeZone As Integer
    bSingleSMS As Boolean
    bPeriodicSMS As Boolean
    sDeliveryDate As String
    sStartingDate As String
    sNumberOfMessages As String
    sWaitingTime As String
    nWaitingPeriod As Integer
End Type

Type FontSettings
    bSpecificFontUsed As Boolean
    sFontName As String
    rFontSize As Double
    bFontBold As Boolean
    bFontItalic As Boolean
End Type

Type JobLogFilterSettings
    bJobRemarksFilterUsed(1 To gcnNumberOfJobRemarkFields) As Boolean
    sJobRemarks(1 To gcnNumberOfJobRemarkFields) As String
    bSMSTypeFilterUsed As Boolean
    sSMSType As String
    bSendingTypeFilterUsed As Boolean
    sSendingType As String
End Type

Type SendLogFilterSettings
    bRecipientNameFilterUsed As Boolean
    sRecipientName As String
    bRecipientFilterUsed As Boolean
    sRecipient As String
    bDeliveryStatusFilterUsed As Boolean
    nDeliveryStatus As Integer
    bReasonCodeFilterUsed As Boolean
    nReasonCode As Integer
    bJobIDFilterUsed As Boolean
    lJobId As Long
End Type

Type PhoneBookEntry
    sName As String
    sNumber As String
    sVariable1 As String
    sVariable2 As String
    sVariable3 As String
End Type

Type JobInfo
    lRecipients As Long
    lMessages As Long
    sSMSType As String
    tDeferredDeliveryTimeSettings As DeferredDeliveryTimeSettings
End Type

Type DBLogFieldsSettings
    bSendLog(1 To 16) As Boolean
    bJobLog(1 To 16) As Boolean
End Type

Enum ValidityPeriodMode
    UseSpecificSettingsAsLifeTime = 1
    UseSingleshotAsLifeTime = 2
    UseWaitingTimeAsLifeTime = 3
End Enum

Type ValidityPeriod
    bStandardPeriod As Boolean
    bUserDefinedSettingsOnlyForSpecificSMSTypes As Boolean
    lLifeTime As Long
    nLifeTimeUnit As Integer
End Type

Type ValidityPeriodSettings
    bSingleSMSUseUserDefinedLifeTime As Boolean
    bSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes As Boolean
    nSingleSMSValidityPeriodMode As ValidityPeriodMode
    lSingleSMSSpecificSettingLifeTime As Long
    nSingleSMSSpecificSettingLifeTimeUnit As Integer

    bPeriodicSMSUseUserDefinedLifeTime As Boolean
    bPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes As Boolean
    nPeriodicSMSValidityPeriodMode As ValidityPeriodMode
    lPeriodicSMSSpecificSettingLifeTime As Long
    nPeriodicSMSSpecificSettingLifeTimeUnit As Integer
End Type

Global gtValidityPeriodSettings As ValidityPeriodSettings

Global gsPhonebookVariableField(1 To 2, 1 To 3) As String

Global gtDBLogFieldsSettings As DBLogFieldsSettings

Global gtJobInfo As JobInfo

Global gtFontSettingsCurrent As FontSettings
Global gtFontSettingsDefault As FontSettings
Global gtMenuControl As MenuControl

Global gtJobLogFilterSettings As JobLogFilterSettings
Global gtSendLogFilterSettings As SendLogFilterSettings
Global gbSaveMessagesInSendLog As Boolean
Global glJobIDEditJobRemarks As Long

Global gtDeferredDeliveryTimeSettings As DeferredDeliveryTimeSettings
Global gnLanguage As Integer

Global gnSelectedRandomLogo As Integer
Global gsSQLSendjournal As String

Global gnPhoneBookSortOrder As Integer

Declare Function ShellExecute _
        Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hWnd As Long, _
                               ByVal lpOperation As String, _
                               ByVal lpFile As String, _
                               ByVal lpParameters As String, _
                               ByVal lpDirectory As String, _
                               ByVal nShowCmd As Long) As Long

Global Const gcErrorCodeOk = 1
Global Const gcErrorCodeConnectFailed = 2
Global Const gcErrorCodeAuthorizationFailed = 3
Global Const gcErrorCodeBinaryFileNotFound = 4
Global Const gcErrorCodeNotEnoughCreditsAvailable = 5
Global Const gcErrorCodeTimeOutError = 6
Global Const gcErrorCodeTransmissionErrorTryAgain = 7
Global Const gcErrorCodeInvalidUserKey = 8
Global Const gcErrorCodeInvalidPassword = 9
Global Const gcErrorCodeInvalidOriginator = 10
Global Const gcErrorCodeInvalidMessageData = 11
Global Const gcErrorCodeInvalidBinaryData = 12
Global Const gcErrorCodeInvalidBinaryFile = 13
Global Const gcErrorCodeInvalidMCC = 14
Global Const gcErrorCodeInvalidMNC = 15
Global Const gcErrorCodeInvalidxSer = 16
Global Const gcErrorCodeInvalidURLBufferedMessageNotification = 17
Global Const gcErrorCodeInvalidURLDeliveryNotification = 18
Global Const gcErrorCodeInvalidURLNonDeliveryNotification = 19
Global Const gcErrorCodeMissingRecipient = 20
Global Const gcErrorCodeMissingBinaryData = 21
Global Const gcErrorCodeInvalidDeferredDeliveryTime = 22
Global Const gcErrorCodeMissingTransactionReferenceNumber = 23
Global Const gcErrorCodeServiceTemporaryNotAvailable = 24
Global Const gcErrorCodeUserBarringActive = 25
Global Const gcErrorCodeNotAuthorizedForThisOperation = 26
Global Const gcErrorCodeMessageTooLong = 27
Global Const gcErrorCodeNoOriginatorRestrictions = 28
Global Const gcErrorCodeOriginatorAuthorizationPending = 29
Global Const gcErrorCodeOriginatorNotAuthorized = 30
Global Const gcErrorCodeOriginatorAlreadyAuthorized = 31

Declare Function SetWindowLong _
        Lib "user32" _
        Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                ByVal nIndex As Long, _
                                ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong _
        Lib "user32" _
        Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                ByVal nIndex As Long) As Long
Declare Function CallWindowProc _
        Lib "user32" _
        Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                 ByVal hWnd As Long, _
                                 ByVal msg As Long, _
                                 ByVal wParam As Long, _
                                 ByVal lParam As Long) As Long

Private Const WM_MOUSEWHEEL = &H20A ' window message for mouse wheel
Public Const GWL_WNDPROC = (-4)
Global gOldwndProcfrmSMSMain As Long
Global gOldwndProcfrmJobLog As Long
Global gOldwndProcfrmSendLog As Long
Global gOldwndProcfrmImport As Long


Public Sub FlexGridScrollDown(grdFlexGrid As MSFlexGrid)
        '<EhHeader>
        On Error GoTo FlexGridScrollDown_Err
        '</EhHeader>
        Dim lTopRow As Long
        Dim lNewTopRow As Long
        Dim lRows As Long
        Dim lRowHeight As Long

100     Select Case True

            Case grdFlexGrid.Rows > 2
102             lRowHeight = grdFlexGrid.RowPos(2) - grdFlexGrid.RowPos(1)
  
104         Case grdFlexGrid.Rows > 1
106             lRowHeight = grdFlexGrid.RowPos(1) - grdFlexGrid.RowPos(0)
  
        End Select

108     If (grdFlexGrid.Rows + 1) * lRowHeight <= grdFlexGrid.Height Then
            Exit Sub
        Else
            'Do nothing
        End If

110     lTopRow = grdFlexGrid.TopRow
112     lRows = grdFlexGrid.Rows

114     If lTopRow < lRows - 1 Then

116         Select Case True

                Case lRows > 5000
118                 lNewTopRow = lTopRow + 10
    
120             Case lRows > 1000
122                 lNewTopRow = lTopRow + 5
    
124             Case lRows > 500
126                 lNewTopRow = lTopRow + 4
    
128             Case lRows > 100
130                 lNewTopRow = lTopRow + 3
    
132             Case lRows > 50
134                 lNewTopRow = lTopRow + 2
    
136             Case Else
138                 lNewTopRow = lTopRow + 1
    
            End Select
  
140         If lNewTopRow >= lRows Then
142             lNewTopRow = lRows - 1
            End If
  
144         grdFlexGrid.TopRow = lNewTopRow
        End If

        '<EhFooter>
        Exit Sub

FlexGridScrollDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.FlexGridScrollDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub FlexGridScrollUp(grdFlexGrid As MSFlexGrid)
        ' scroll up..
        '<EhHeader>
        On Error GoTo FlexGridScrollUp_Err
        '</EhHeader>
        Dim lTopRow As Long
        Dim lNewTopRow As Long
        Dim lRows As Long

100     lTopRow = grdFlexGrid.TopRow
102     lRows = grdFlexGrid.Rows

104     If lTopRow > 1 Then

106         Select Case True

                Case lRows > 5000
108                 lNewTopRow = lTopRow - 10
    
110             Case lRows > 1000
112                 lNewTopRow = lTopRow - 5
    
114             Case lRows > 500
116                 lNewTopRow = lTopRow - 4
    
118             Case lRows > 100
120                 lNewTopRow = lTopRow - 3
    
122             Case lRows > 50
124                 lNewTopRow = lTopRow - 2
    
126             Case Else
128                 lNewTopRow = lTopRow - 1
    
            End Select
  
130         If lNewTopRow < 1 Then
132             lNewTopRow = 1
            End If
  
134         grdFlexGrid.TopRow = lNewTopRow
        End If

        '<EhFooter>
        Exit Sub

FlexGridScrollUp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.FlexGridScrollUp " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function MouseWheelSupportWndProcfrmSMSMain(ByVal hWnd As Long, _
                                                      ByVal wMsg As Long, _
                                                      ByVal wParam As Long, _
                                                      ByVal lParam As Long) As Long
        '<EhHeader>
        On Error GoTo MouseWheelSupportWndProcfrmSMSMain_Err
        '</EhHeader>
        Dim bMouseWheelUp As Boolean
        Dim lNewScrollValue As Long

100     If wMsg = WM_MOUSEWHEEL Then
102         If wParam > 0 Then bMouseWheelUp = True Else bMouseWheelUp = False
  
104         Select Case frmSMSMain.TabMain.Tab
  
                Case 1    'Phonebook Tab

106                 Select Case frmSMSMain.ActiveControl

                        Case frmSMSMain.grdRecipients

108                         If bMouseWheelUp Then
110                             FlexGridScrollUp frmSMSMain.grdRecipients
                            Else
112                             FlexGridScrollDown frmSMSMain.grdRecipients
                            End If
          
114                     Case Else

116                         If bMouseWheelUp Then
118                             FlexGridScrollUp frmSMSMain.grdPhonebook
                            Else
120                             FlexGridScrollDown frmSMSMain.grdPhonebook
                            End If

                    End Select

122             Case 10   'WAP Push Tab

124                 If frmSMSMain.scrlVertical.Visible = True Then
126                     lNewScrollValue = frmSMSMain.scrlVertical.Value
    
128                     Select Case bMouseWheelUp

                            Case True
130                             lNewScrollValue = lNewScrollValue - frmSMSMain.scrlVertical.SmallChange
      
132                         Case False
134                             lNewScrollValue = lNewScrollValue + frmSMSMain.scrlVertical.SmallChange
                
                        End Select

136                     If lNewScrollValue > frmSMSMain.scrlVertical.Max Then
138                         lNewScrollValue = frmSMSMain.scrlVertical.Max
                        End If
  
140                     If lNewScrollValue < frmSMSMain.scrlVertical.Min Then
142                         lNewScrollValue = frmSMSMain.scrlVertical.Min
                        End If

144                     frmSMSMain.scrlVertical.Value = lNewScrollValue
                    End If
  
146             Case Else
                    'Do nothing
  
            End Select

        End If
    
148     MouseWheelSupportWndProcfrmSMSMain = CallWindowProc(gOldwndProcfrmSMSMain, hWnd, wMsg, wParam, lParam)
        '<EhFooter>
        Exit Function

MouseWheelSupportWndProcfrmSMSMain_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.MouseWheelSupportWndProcfrmSMSMain " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function MouseWheelSupportWndProcfrmSendlog(ByVal hWnd As Long, _
                                                   ByVal wMsg As Long, _
                                                   ByVal wParam As Long, _
                                                   ByVal lParam As Long) As Long
        '<EhHeader>
        On Error GoTo MouseWheelSupportWndProcfrmSendlog_Err
        '</EhHeader>
        Dim bMouseWheelUp As Boolean

100     If wMsg = WM_MOUSEWHEEL Then
    
102         If wParam > 0 Then bMouseWheelUp = True Else bMouseWheelUp = False
    
104         Select Case bMouseWheelUp

                Case True
106               '  FlexGridScrollUp frmSendLog.grdSendLog
   
108             Case False
110              '   FlexGridScrollDown frmSendLog.grdSendLog
   
            End Select

        End If
    
112     MouseWheelSupportWndProcfrmSendlog = CallWindowProc(gOldwndProcfrmSendLog, hWnd, wMsg, wParam, lParam)
        '<EhFooter>
        Exit Function

MouseWheelSupportWndProcfrmSendlog_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.MouseWheelSupportWndProcfrmSendlog " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function MouseWheelSupportWndProcfrmJoblog(ByVal hWnd As Long, _
                                                  ByVal wMsg As Long, _
                                                  ByVal wParam As Long, _
                                                  ByVal lParam As Long) As Long
        '<EhHeader>
        On Error GoTo MouseWheelSupportWndProcfrmJoblog_Err
        '</EhHeader>
        Dim bMouseWheelUp As Boolean

100     If wMsg = WM_MOUSEWHEEL Then
    
102         If wParam > 0 Then bMouseWheelUp = True Else bMouseWheelUp = False
    
104         Select Case bMouseWheelUp

                Case True
106             '    FlexGridScrollUp frmJoblog.grdJobLog
    
108             Case False
110              '   FlexGridScrollDown frmJoblog.grdJobLog
    
            End Select
  
        End If
    
112     MouseWheelSupportWndProcfrmJoblog = CallWindowProc(gOldwndProcfrmJobLog, hWnd, wMsg, wParam, lParam)
        '<EhFooter>
        Exit Function

MouseWheelSupportWndProcfrmJoblog_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.MouseWheelSupportWndProcfrmJoblog " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function MouseWheelSupportWndProcfrmImport(ByVal hWnd As Long, _
                                                  ByVal wMsg As Long, _
                                                  ByVal wParam As Long, _
                                                  ByVal lParam As Long) As Long
        '<EhHeader>
        On Error GoTo MouseWheelSupportWndProcfrmImport_Err
        '</EhHeader>
        Dim bMouseWheelUp As Boolean

100     If wMsg = WM_MOUSEWHEEL Then
    
102         If wParam > 0 Then bMouseWheelUp = True Else bMouseWheelUp = False
    
104         Select Case bMouseWheelUp

                Case True
106             '    FlexGridScrollUp frmImport.grdImport
    
108             Case False
110              '   FlexGridScrollDown frmImport.grdImport
    
            End Select
  
        End If
    
112     MouseWheelSupportWndProcfrmImport = CallWindowProc(gOldwndProcfrmImport, hWnd, wMsg, wParam, lParam)
        '<EhFooter>
        Exit Function

MouseWheelSupportWndProcfrmImport_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.MouseWheelSupportWndProcfrmImport " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub CheckForRequiredDirectories(sPath As String)
        '<EhHeader>
        On Error GoTo CheckForRequiredDirectories_Err
        '</EhHeader>
        Dim sTempPath As String
        Dim sTempArray() As String
        Dim i As Integer

100     sTempArray = Split(sPath, "\")

102     For i = LBound(sTempArray) To UBound(sTempArray)
104         sTempPath = sTempPath & sTempArray(i) & "\"

106         If DirectoryExists(sTempPath) Then
                'Everything ok, do nothing
            Else
108             MkDir sTempPath
            End If

        Next

        '<EhFooter>
        Exit Sub

CheckForRequiredDirectories_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.CheckForRequiredDirectories " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function CreateWAPPushDestinationFileName(sFilename As String) As String
        '<EhHeader>
        On Error GoTo CreateWAPPushDestinationFileName_Err
        '</EhHeader>
        Dim sTempArray() As String
        Dim sFileNamePrefix As String
        Dim sFileServerName As String
        Dim sFileServerPath As String

        'Determine Filename according date
100     sFileNamePrefix = Trim(DatePart("yyyy", Now()))
102     sFileNamePrefix = sFileNamePrefix & Right("00" & Trim(DatePart("m", Now())), 2)
104     sFileNamePrefix = sFileNamePrefix & Right("00" & Trim(DatePart("d", Now())), 2)
106     sFileNamePrefix = sFileNamePrefix & Right("00" & Trim(DatePart("h", Now())), 2)
108     sFileNamePrefix = sFileNamePrefix & Right("00" & Trim(DatePart("n", Now())), 2)
110     sFileNamePrefix = sFileNamePrefix & Right("00" & Trim(DatePart("s", Now())), 2)

112     CreateWAPPushDestinationFileName = sFileNamePrefix & Trim(sFilename)
        '<EhFooter>
        Exit Function

CreateWAPPushDestinationFileName_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.CreateWAPPushDestinationFileName " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function DirectoryExists(sPath As String) As Boolean
        '<EhHeader>
        On Error GoTo DirectoryExists_Err
        '</EhHeader>
        Dim sDirResult As String

100     sDirResult = Dir$(sPath, vbDirectory)

102     If sDirResult <> "" Then
104         DirectoryExists = True
        Else
106         DirectoryExists = False
        End If

        '<EhFooter>
        Exit Function

DirectoryExists_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.DirectoryExists " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ExtractFilenameWithExtension(sFilename As String) As String
        '<EhHeader>
        On Error GoTo ExtractFilenameWithExtension_Err
        '</EhHeader>
        Dim sTempArray() As String
        Dim sTemp As String

        On Error GoTo ErrorTrap

100     sTempArray() = Split(sFilename, "\")

102     sTemp = sTempArray(UBound(sTempArray))

104     ExtractFilenameWithExtension = sTempArray(UBound(sTempArray))

ErrorTrap:
        Exit Function

        '<EhFooter>
        Exit Function

ExtractFilenameWithExtension_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ExtractFilenameWithExtension " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ExtractFilenameWithoutExtension(sFilename As String) As String
        '<EhHeader>
        On Error GoTo ExtractFilenameWithoutExtension_Err
        '</EhHeader>
        Dim sTempArray() As String
        Dim sTemp As String

        On Error GoTo ErrorTrap

100     sTempArray() = Split(sFilename, "\")

102     sTemp = sTempArray(UBound(sTempArray))
104     sTempArray = Split(sTemp, ".")

106     ExtractFilenameWithoutExtension = sTempArray(LBound(sTempArray))

ErrorTrap:
        Exit Function

        '<EhFooter>
        Exit Function

ExtractFilenameWithoutExtension_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ExtractFilenameWithoutExtension " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function FilesDropped(ByRef Data As TabDlg.DataObject) As String
        '<EhHeader>
        On Error GoTo FilesDropped_Err
        '</EhHeader>
        Dim v As TabDlg.DataObjectFiles, i As Long, s As String, f As String
        On Error GoTo Err_Init
        Dim sTemp As String
        Dim sTemp2 As String
        Dim nFile As Integer

100     Set v = Data.Files

102     If Not (v Is Nothing) Then

104         For i = 1 To v.Count
106             f = v.Item(i)

                'Resolve it first if it's a shortcut
108             If Len(f) > 3 Then
110                 If StrComp(Right$(f, 4), ".lnk", vbTextCompare) = 0 Then
                        'f = ResolveLink(f)
                    End If
                End If

112             If Dir(f, vbNormal Or vbArchive) = "" Then
114                 If Dir(f, vbDirectory) = "" Then
                        'Can't add it to the list
                    Else

                        'If it's a directory, add a "\" on the end of it.
116                     If Right$(f, 1) <> "\" Then
118                         f = f & "\"
                        End If

                        'Add the directory to the list
120                     s = s & f & vbCrLf
                    End If

                Else
                    'Add the file to the list
122                 s = s & f & vbCrLf
                End If

124         Next i

        End If

126     FilesDropped = s
        Exit Function
Err_Init:

128     If Err.number = 461 Then
130         MsgBox LoadLanguageSpecificString(gnLanguage, 741), vbExclamation, gsApplicationName
132         f = ""
134         Resume Next
        Else
136         MsgBox LoadLanguageSpecificString(gnLanguage, 741), vbExclamation, gsApplicationName
        End If

        '<EhFooter>
        Exit Function

FilesDropped_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.FilesDropped " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function CheckOriginatorBeforeSending() As Boolean
        '<EhHeader>
        On Error GoTo CheckOriginatorBeforeSending_Err
        '</EhHeader>
        Dim sMessage As String
        Dim nAnswer As Integer
100     Screen.MousePointer = vbHourglass

        Dim SMS As Booster
102     Set SMS = Nothing

104     Set SMS = New Booster
'106     SMS.Hosts = App.Path & "\smshosts.txt"
108     SMS.UserKey = frmSMSMain.txtUserkey.Text
110     SMS.Password = frmSMSMain.txtPassword.Text

112     SMS.Originator = frmSMSMain.txtOriginator
114     SMS.CheckOriginatorAuthorization
  
116     Screen.MousePointer = vbDefault

118     Select Case SMS.errorCode

            Case gcErrorCodeOk
120             CheckOriginatorBeforeSending = True
  
122         Case gcErrorCodeNoOriginatorRestrictions
124             CheckOriginatorBeforeSending = True
  
126         Case gcErrorCodeOriginatorAlreadyAuthorized
128             CheckOriginatorBeforeSending = True
  
130         Case gcErrorCodeOriginatorNotAuthorized, gcErrorCodeOriginatorAuthorizationPending
132             sMessage = LoadLanguageSpecificString(gnLanguage, 813) & " '" & frmSMSMain.txtOriginator & "' " & LoadLanguageSpecificString(gnLanguage, 814) & vbCrLf & LoadLanguageSpecificString(gnLanguage, 815)
134             nAnswer = MsgBox(sMessage, vbExclamation Or vbOKCancel Or vbDefaultButton2, gsApplicationName)

136             If nAnswer = vbCancel Then
138                 CheckOriginatorBeforeSending = False
                Else
140                 CheckOriginatorBeforeSending = True
                End If

142         Case Else
144             MsgBox LoadLanguageSpecificString(gnLanguage, 207) & vbCrLf & ErrorDescriptionFromASPSMSErrorCode(gnLanguage, CInt(SMS.errorCode)), vbExclamation, gsApplicationName
146             CheckOriginatorBeforeSending = False
  
        End Select

148     Set SMS = Nothing

        '<EhFooter>
        Exit Function

CheckOriginatorBeforeSending_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.CheckOriginatorBeforeSending " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub ShowServerFeedback(nStep As Integer, _
                       sOriginator As String, _
                       nErrorCode As Integer)
        '<EhHeader>
        On Error GoTo ShowServerFeedback_Err
        '</EhHeader>
        Dim sTemp As String

100     Select Case nErrorCode

            Case gcErrorCodeOk

102             Select Case nStep

                    Case 1
104                     sTemp = LoadLanguageSpecificString(gnLanguage, 783) & " '" & sOriginator & "' " & LoadLanguageSpecificString(gnLanguage, 784)

106                 Case 2
108                     sTemp = LoadLanguageSpecificString(gnLanguage, 803) & " '" & sOriginator & "' " & LoadLanguageSpecificString(gnLanguage, 804)

110                 Case 3
112                     sTemp = LoadLanguageSpecificString(gnLanguage, 801)
                End Select

114             MsgBox sTemp, vbInformation, gsApplicationName
  
116         Case gcErrorCodeNoOriginatorRestrictions

118             Select Case nStep

                    Case 1
120                     sTemp = LoadLanguageSpecificString(gnLanguage, 794) & " " & LoadLanguageSpecificString(gnLanguage, 795)

122                 Case 2
124                     sTemp = LoadLanguageSpecificString(gnLanguage, 794) & " " & LoadLanguageSpecificString(gnLanguage, 795)

126                 Case 3
128                     sTemp = LoadLanguageSpecificString(gnLanguage, 794) & " " & LoadLanguageSpecificString(gnLanguage, 795)
                End Select

130             MsgBox sTemp, vbInformation, gsApplicationName
    
132         Case gcErrorCodeOriginatorAuthorizationPending

134             Select Case nStep

                    Case 1
136                     sTemp = LoadLanguageSpecificString(gnLanguage, 790) & " '" & sOriginator & "' " & LoadLanguageSpecificString(gnLanguage, 791)

138                 Case 2

                        'Impossible
140                 Case 3
                        'Impossible
                End Select

142             MsgBox sTemp, vbInformation, gsApplicationName
  
144         Case gcErrorCodeOriginatorNotAuthorized

146             Select Case nStep

                    Case 1
148                     sTemp = LoadLanguageSpecificString(gnLanguage, 786) & " '" & sOriginator & "' " & LoadLanguageSpecificString(gnLanguage, 787)

150                 Case 2

                        '????????????
152                 Case 3
154                     sTemp = LoadLanguageSpecificString(gnLanguage, 802)
                End Select

156             MsgBox sTemp, vbExclamation, gsApplicationName
  
158         Case gcErrorCodeOriginatorAlreadyAuthorized

160             Select Case nStep

                    Case 1
162                     sTemp = LoadLanguageSpecificString(gnLanguage, 806)

164                 Case 2
166                     sTemp = LoadLanguageSpecificString(gnLanguage, 806)

168                 Case 3
170                     sTemp = LoadLanguageSpecificString(gnLanguage, 806)
                End Select

172             MsgBox sTemp, vbInformation, gsApplicationName
  
174         Case Else
176             MsgBox LoadLanguageSpecificString(gnLanguage, 207) & vbCrLf & ErrorDescriptionFromASPSMSErrorCode(gnLanguage, nErrorCode), vbExclamation, gsApplicationName

        End Select

        '<EhFooter>
        Exit Sub

ShowServerFeedback_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ShowServerFeedback " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function ActualValidityPeriod() As ValidityPeriod
        '<EhHeader>
        On Error GoTo ActualValidityPeriod_Err
        '</EhHeader>
        Dim lValidityPeriod As Long
        Dim bPeriodicJob As Boolean
        Dim sDecisionRule As String

100     bPeriodicJob = PeriodicJob()
        'Decisiontable
        '                                         R1  R2  R3  R4  R5  R6  R7  R8
        '
        'B1  bSingleSMSUseUserDefinedLifeTime     N   N   N   N   Y   Y   Y   Y
        'B2  bPeriodicSMSUseUserDefinedLifeTime   N   N   Y   Y   N   N   Y   Y
        'B3  bPeriodicJob                         N   Y   N   Y   N   Y   N   Y
        '
        'A1  Standard Lifetime                    Y   Y   Y   -   -   Y   -   -
        'A2  Userdefined Lifetime (Single)        -   -   -   -   Y   -   Y   -
        'A3  Userdefined Lifetime (Periodic)      -   -   -   Y   -   -   -   Y

102     If gtValidityPeriodSettings.bSingleSMSUseUserDefinedLifeTime = True Then
104         sDecisionRule = sDecisionRule & "Y"
        Else
106         sDecisionRule = sDecisionRule & "N"
        End If

108     If gtValidityPeriodSettings.bPeriodicSMSUseUserDefinedLifeTime = True Then
110         sDecisionRule = sDecisionRule & "Y"
        Else
112         sDecisionRule = sDecisionRule & "N"
        End If

114     If bPeriodicJob = True Then
116         sDecisionRule = sDecisionRule & "Y"
        Else
118         sDecisionRule = sDecisionRule & "N"
        End If

120     Select Case sDecisionRule

            Case "NNN"  'Standard Lifetime
122             ActualValidityPeriod.bStandardPeriod = True
      
124         Case "NNY"  'Standard Lifetime
126             ActualValidityPeriod.bStandardPeriod = True
      
128         Case "NYN"  'Standard Lifetime
130             ActualValidityPeriod.bStandardPeriod = True
      
132         Case "NYY"  'Userdefined Lifetime (Periodic)
134             ActualValidityPeriod = ValidityPeriodWithSpecificSettings(gtValidityPeriodSettings, gtDeferredDeliveryTimeSettings, True)
    
136         Case "YNN"  'Userdefined Lifetime (Single)
138             ActualValidityPeriod = ValidityPeriodWithSpecificSettings(gtValidityPeriodSettings, gtDeferredDeliveryTimeSettings, False)
  
140         Case "YNY"  'Standard Lifetime
142             ActualValidityPeriod.bStandardPeriod = True
      
144         Case "YYN"  'Userdefined Lifetime (Single)
146             ActualValidityPeriod = ValidityPeriodWithSpecificSettings(gtValidityPeriodSettings, gtDeferredDeliveryTimeSettings, False)
  
148         Case "YYY"  'Userdefined Lifetime (Periodic)
150             ActualValidityPeriod = ValidityPeriodWithSpecificSettings(gtValidityPeriodSettings, gtDeferredDeliveryTimeSettings, True)
  
152         Case Else
                'MsgBox "Case Else DecisionTable"
  
        End Select

        '<EhFooter>
        Exit Function

ActualValidityPeriod_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ActualValidityPeriod " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Sub CopySelectedGridAreaToClipboard(grdInput As MSFlexGrid)
        '<EhHeader>
        On Error GoTo CopySelectedGridAreaToClipboard_Err
        '</EhHeader>
        Dim sContent As String
        Dim sCell As String
        Dim lCurrentRow As Long
        Dim lCurrentCol As Long
        Dim lStartRow As Long
        Dim lEndRow As Long
        Dim lStartCol As Long
        Dim lEndCol As Long

100     If grdInput.RowSel = 0 Or grdInput.Row = 0 Then
            Exit Sub
        End If

        On Error GoTo ErrorTrap

102     Screen.MousePointer = vbHourglass

104     If grdInput.RowSel < grdInput.Row Then
106         lStartRow = grdInput.RowSel
108         lEndRow = grdInput.Row
        Else
110         lStartRow = grdInput.Row
112         lEndRow = grdInput.RowSel
        End If

114     If grdInput.ColSel < grdInput.col Then
116         lStartCol = grdInput.ColSel
118         lEndCol = grdInput.col
        Else
120         lStartCol = grdInput.col
122         lEndCol = grdInput.ColSel
        End If

124     For lCurrentCol = lStartCol To lEndCol
126         sCell = grdInput.TextMatrix(0, lCurrentCol)
128         sContent = sContent & sCell & vbTab
        Next

130     sContent = Left(sContent, Len(sContent) - 1) & vbCrLf

132     For lCurrentRow = lStartRow To lEndRow
134         For lCurrentCol = lStartCol To lEndCol
136             sCell = grdInput.TextMatrix(lCurrentRow, lCurrentCol)
138             sContent = sContent & sCell & vbTab
            Next

140         sContent = Left(sContent, Len(sContent) - 1) & vbCrLf
        Next

142     Clipboard.Clear
144     Clipboard.SetText sContent

146     Screen.MousePointer = vbDefault

ErrorTrap:
        Exit Sub
148     Screen.MousePointer = vbDefault
        '<EhFooter>
        Exit Sub

CopySelectedGridAreaToClipboard_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.CopySelectedGridAreaToClipboard " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub CurrentPositionWithinGrid(grdInput As MSFlexGrid, _
                              rx As Single, _
                              ry As Single, _
                              lCol As Long, _
                              lRow As Long)
        '<EhHeader>
        On Error GoTo CurrentPositionWithinGrid_Err
        '</EhHeader>
        Dim lTopRow As Long
        Dim lCurrentRow As Long
        Dim lCurrentCol As Long
        Dim lRowHeight As Long
        Dim lColWidth As Long

        On Error GoTo ErrorTrap

100     Select Case True

            Case grdInput.Rows > 2
102             lRowHeight = grdInput.RowPos(2) - grdInput.RowPos(1)
  
104         Case grdInput.Rows > 1
106             lRowHeight = grdInput.RowPos(1) - grdInput.RowPos(0)
  
        End Select

108     If lRowHeight > 0 Then
110         lCurrentRow = (Abs(grdInput.RowPos(1)) + ry - (lRowHeight / 2)) / lRowHeight

            'Because of first Row, which is fixed
112         If grdInput.RowPos(1) < 0 Then
114             lCurrentRow = lCurrentRow + 1
            Else
116             lCurrentRow = lCurrentRow - 1
            End If

118         If ry <= lRowHeight Then
120             lCurrentRow = 0
            End If

        Else
  
        End If

        Do

122         If lCurrentCol < grdInput.Cols Then
124             If grdInput.ColPos(lCurrentCol) < rx Then
126                 lCurrentCol = lCurrentCol + 1
                Else
                    Exit Do
                End If

            Else
                Exit Do
            End If

        Loop

128     lCurrentCol = lCurrentCol - 1

130     If lCurrentRow <= grdInput.Rows - 1 And lCurrentCol <= grdInput.Cols - 1 And lCurrentRow > -1 And lCurrentCol > -1 Then
132         grdInput.toolTipText = grdInput.TextMatrix(lCurrentRow, lCurrentCol)
        Else

        End If

134     lCol = lCurrentCol
136     lRow = lCurrentRow

        Exit Sub
ErrorTrap:
138     lCol = 0
140     lRow = 0
        Exit Sub
        '<EhFooter>
        Exit Sub

CurrentPositionWithinGrid_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.CurrentPositionWithinGrid " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function FileExists(sFile As String) As Integer
        '<EhHeader>
        On Error GoTo FileExists_Err
        '</EhHeader>
        On Error Resume Next

100     FileExists = (Dir$(sFile) <> "")

        '<EhFooter>
        Exit Function

FileExists_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.FileExists " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function CreatePersonalizedMessageData(sInput As String, lRecipient As Long, bVariableMessageData As Boolean) As String
        '<EhHeader>
        On Error GoTo CreatePersonalizedMessageData_Err
        '</EhHeader>
        Dim sMessage As String
        Dim bReplacementProcessed As Boolean

100     sMessage = sInput
102     sMessage = Replace(sMessage, LoadLanguageSpecificString(1, 234), frmSMSMain.grdRecipients.TextMatrix(lRecipient, 0))
104     sMessage = Replace(sMessage, LoadLanguageSpecificString(2, 234), frmSMSMain.grdRecipients.TextMatrix(lRecipient, 0))
106     sMessage = Replace(sMessage, LoadLanguageSpecificString(1, 235), frmSMSMain.grdRecipients.TextMatrix(lRecipient, 1))
108     sMessage = Replace(sMessage, LoadLanguageSpecificString(2, 235), frmSMSMain.grdRecipients.TextMatrix(lRecipient, 1))
  
110     sMessage = Replace(sMessage, "<" & gsPhonebookVariableField(1, 1) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 2))
112     sMessage = Replace(sMessage, "<" & gsPhonebookVariableField(2, 1) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 2))
  
114     sMessage = Replace(sMessage, "<" & gsPhonebookVariableField(1, 2) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 3))
116     sMessage = Replace(sMessage, "<" & gsPhonebookVariableField(2, 2) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 3))
  
118     sMessage = Replace(sMessage, "<" & gsPhonebookVariableField(1, 3) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 4))
120     sMessage = Replace(sMessage, "<" & gsPhonebookVariableField(2, 3) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 4))

122     sMessage = Replace(sMessage, "<" & LoadLanguageSpecificString(1, 213) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 2))
124     sMessage = Replace(sMessage, "<" & LoadLanguageSpecificString(2, 213) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 2))
126     sMessage = Replace(sMessage, "<" & LoadLanguageSpecificString(1, 214) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 3))
128     sMessage = Replace(sMessage, "<" & LoadLanguageSpecificString(2, 214) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 3))
130     sMessage = Replace(sMessage, "<" & LoadLanguageSpecificString(1, 215) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 3))
132     sMessage = Replace(sMessage, "<" & LoadLanguageSpecificString(2, 215) & ">", frmSMSMain.grdRecipients.TextMatrix(lRecipient, 3))

134     bVariableMessageData = PersonalizedSMSModeUsed(sInput)
136     CreatePersonalizedMessageData = sMessage

        '<EhFooter>
        Exit Function

CreatePersonalizedMessageData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.CreatePersonalizedMessageData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Function ActualValidityPeriodWithReturnValueInMinutes(bSpecificSMSType As Boolean) As Long
        '<EhHeader>
        On Error GoTo ActualValidityPeriodWithReturnValueInMinutes_Err
        '</EhHeader>
        Dim tValidityPeriod As ValidityPeriod

100     tValidityPeriod = ActualValidityPeriod()

102     Select Case True

            Case tValidityPeriod.bStandardPeriod = True
104             ActualValidityPeriodWithReturnValueInMinutes = 0
  
106         Case tValidityPeriod.bUserDefinedSettingsOnlyForSpecificSMSTypes = True And bSpecificSMSType = False
108             ActualValidityPeriodWithReturnValueInMinutes = 0
  
110         Case Else

112             Select Case tValidityPeriod.nLifeTimeUnit

                    Case 0  'Minutes
114                     ActualValidityPeriodWithReturnValueInMinutes = tValidityPeriod.lLifeTime
    
116                 Case 1  'Hours
118                     ActualValidityPeriodWithReturnValueInMinutes = (tValidityPeriod.lLifeTime) * 60
                End Select
  
        End Select

        '<EhFooter>
        Exit Function

ActualValidityPeriodWithReturnValueInMinutes_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ActualValidityPeriodWithReturnValueInMinutes " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function PeriodicJob() As Boolean
        '<EhHeader>
        On Error GoTo PeriodicJob_Err
        '</EhHeader>

100     If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
102         If gtDeferredDeliveryTimeSettings.bSingleSMS = True Then
104             PeriodicJob = False
            End If

106         If gtDeferredDeliveryTimeSettings.bPeriodicSMS = True Then
108             PeriodicJob = True
            End If

        Else
110         PeriodicJob = False
        End If

        '<EhFooter>
        Exit Function

PeriodicJob_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.PeriodicJob " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ProcessSendAction(cRecipientData As Collection, bVariableMessageData As Boolean, bVariableDeferredDeliveryTimeData As Boolean, lJobId As Long) As Integer
        '<EhHeader>
        On Error GoTo ProcessSendAction_Err
        '</EhHeader>

        Dim lRecipient As Long
        Dim sMessage As String
        Dim cRecipientTemp As Recipient
        Dim sVersionSpecific As String
        Dim lCounter As Long
        Dim lTemp As Long

        Dim sMessageData As String
        Dim sDeferredDeliveryTime As String

        Dim sP1 As String
        Dim sP2 As String
        Dim sP3 As String
        Dim sP4 As String
        Dim sP5 As String
        Dim sP6 As String
        Dim sNotifyURL As String

100     Screen.MousePointer = vbHourglass

102     Set gSMS = Nothing

104     Set gSMS = New Booster
106     gSMS.Hosts = App.Path & "\smshosts.txt"
108     gSMS.UserKey = frmSMSMain.txtUserkey.Text
110     gSMS.Password = frmSMSMain.txtPassword.Text
112     gSMS.Originator = frmSMSMain.txtOriginator.Text

114     For Each cRecipientTemp In cRecipientData

116         Select Case True
  
                Case bVariableMessageData = False And bVariableDeferredDeliveryTimeData = False
118                 gSMS.AddRecipient cRecipientTemp.sRecipient, cRecipientTemp.sTransactionReferenceNumber
    
120                 If cRecipientTemp.sDeferredDeliveryTime <> "" And gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
122                     gSMS.DeferredDeliveryTime = ConvertToASPSMSTimeFormat(cRecipientTemp.sDeferredDeliveryTime)
124                     gSMS.TimeZone = TimeZoneFromListBoxSetting(gtDeferredDeliveryTimeSettings.nTimeZone)
                    End If
    
126             Case bVariableMessageData = True And bVariableDeferredDeliveryTimeData = True
128                 gSMS.AddRecipientEx cRecipientTemp.sRecipient, cRecipientTemp.sTransactionReferenceNumber, cRecipientTemp.sMessageData, ConvertToASPSMSTimeFormat(cRecipientTemp.sDeferredDeliveryTime)
    
130             Case bVariableMessageData = True And bVariableDeferredDeliveryTimeData = False
132                 gSMS.AddRecipientEx cRecipientTemp.sRecipient, cRecipientTemp.sTransactionReferenceNumber, cRecipientTemp.sMessageData

134                 If cRecipientTemp.sDeferredDeliveryTime <> "" Then
136                     gSMS.DeferredDeliveryTime = ConvertToASPSMSTimeFormat(cRecipientTemp.sDeferredDeliveryTime)
138                     gSMS.TimeZone = TimeZoneFromListBoxSetting(gtDeferredDeliveryTimeSettings.nTimeZone)
                    End If
   
140             Case bVariableMessageData = False And bVariableDeferredDeliveryTimeData = True
142                 gSMS.AddRecipientEx cRecipientTemp.sRecipient, cRecipientTemp.sTransactionReferenceNumber, , ConvertToASPSMSTimeFormat(cRecipientTemp.sDeferredDeliveryTime)
                       
144             Case Else
                    'Do nothing
    
            End Select
    
146         Select Case True
  
                Case frmSMSMain.optSMSType(0).Value = True 'Text SMS

148                 If bVariableMessageData = True Then
150                     sMessage = cRecipientTemp.sMessageData
                    Else
152                     sMessage = frmSMSMain.txtSMS.Text
                    End If

154                 sP1 = "2"
  
156             Case frmSMSMain.optSMSType(1).Value = True 'Operatorlogo
158                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
160                 sP1 = "3"
  
162             Case frmSMSMain.optSMSType(2).Value = True 'Grouplogo
164                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
166                 sP1 = "12"
  
168             Case frmSMSMain.optSMSType(3).Value = True 'Ringtone
170                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
172                 sP1 = "5"
  
174             Case frmSMSMain.optSMSType(4).Value = True 'Picturemessage
176                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">: " & frmSMSMain.txtPictureMessageText.Text
178                 sP1 = "4"
  
180             Case frmSMSMain.optSMSType(5).Value = True 'VCard
182                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
184                 sP1 = "10"
  
186             Case frmSMSMain.optSMSType(6).Value = True 'Unicode Text SMS
188                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
190                 sP1 = "17"
      
192             Case frmSMSMain.optSMSType(7).Value = True 'WAP Push SMS
194                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
196                 sP1 = "18"
      
198             Case frmSMSMain.optSMSType(8).Value = True 'Binarydata
200                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
202                 sP1 = "6"
      
204             Case frmSMSMain.optSMSType(9).Value = True
206                 VersionSpecificAction 20, , , sVersionSpecific
208                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
210                 sP1 = sVersionSpecific
    
212             Case frmSMSMain.optSMSType(10).Value = True
214                 VersionSpecificAction 21, , , sVersionSpecific
216                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
218                 sP1 = sVersionSpecific
    
220             Case frmSMSMain.optSMSType(11).Value = True
222                 VersionSpecificAction 22, , , sVersionSpecific
224                 sMessage = "<" & SelectedSMSType(gnLanguage) & ">"
226                 sP1 = sVersionSpecific
    
            End Select
  
228         If gbSaveMessagesInSendLog = True Then
230             SaveOutgoingMessage cRecipientTemp.sRecipient, cRecipientTemp.sRecipientName, cRecipientTemp.sTransactionReferenceNumber, sMessage, cRecipientTemp.sDeferredDeliveryTime, lJobId   'Puts Message into Sendlog
            End If

        Next

        'Specify OTA Deliverynotification URL
232     If frmSMSMain.chkUseOTADeliveryNotifications.Value = 1 Then
            'P2: Userkey
234         sP2 = frmSMSMain.txtUserkey.Text

            'P3: Password
236         sP3 = frmSMSMain.txtPassword.Text

            'P4: Phonenumber, which is used to receive notification
238         sP4 = URLEncode(frmSMSMain.txtRecipientDeliveryNotification.Text)

            'P5: Language
240         sP5 = Trim(Str$(gnLanguage))

242         sNotifyURL = "http://www.aspsms.com/notifications/notify.asp?P1=" & sP1 & "&P2=" & sP2 & "&P3=" & sP3 & "&SCTS=<SCTS>&DSCTS=<DSCTS>&RSN=<RSN>&DST=<DST>&RCPNT=<RCPNT>&P4=" & sP4 & "&P5=" & sP5 & "&TRN=<TRN>"

244         If frmSMSMain.chkDeliveryNotificationDelivered.Value = 1 Then
246             gSMS.URLDeliveryNotification = sNotifyURL
            Else
                'Do nothing
            End If

248         If frmSMSMain.chkDeliveryNotificationBuffered.Value = 1 Then
250             gSMS.URLBufferedMessageNotification = sNotifyURL
            Else
                'Do nothing
            End If

252         If frmSMSMain.chkDeliveryNotificationNotDelivered.Value = 1 Then
254             gSMS.URLNonDeliveryNotification = sNotifyURL
            Else
                'Do nothing
            End If

        Else   'Update Log automatically
            'Do nothing
        End If

256     VersionSpecificAction 60, , lTemp
258     gSMS.AffiliateID = lTemp

260     Select Case True

            Case frmSMSMain.optSMSType(0).Value = True 'Text SMS

262             If bVariableMessageData = True Then
                    'Do nothing, already set
                Else
264                 gSMS.MessageData = frmSMSMain.txtSMS.Text
                End If

266             If frmSMSMain.chkFlashingSMS.Value = 1 Then gSMS.FlashingSMS = True Else gSMS.FlashingSMS = False
268             If frmSMSMain.chkBlinkingSMS.Value = 1 Then gSMS.BlinkingSMS = True Else gSMS.BlinkingSMS = False
270             If frmSMSMain.chkReplaceMessage.Value = 1 Then
272                 gSMS.ReplaceMessage = 5
                Else
                    'Do nothing
                End If

274             gSMS.LifeTime = ActualValidityPeriodWithReturnValueInMinutes(False)
276             gSMS.SendTextSMS
    
278         Case frmSMSMain.optSMSType(1).Value = True 'Operatorlogo
280             gSMS.MCC = Val(frmSMSMain.txtMCC.Text)
282             gSMS.MNC = Val(frmSMSMain.txtMNC.Text)
284             gSMS.LifeTime = ActualValidityPeriodWithReturnValueInMinutes(False)
  
286             If frmSMSMain.chkRandomLogo.Value = 0 Then
288                 gSMS.BinaryFileLocation = frmSMSMain.txtPathLogo.Text
290                 gSMS.SendLogo
                Else
292                 gSMS.BinaryFileLocation = App.Path & "\" & Right("0000" & Trim(Str$(gnSelectedRandomLogo)), 4) & ".bmp"
294                 gSMS.SendLogo
                End If

296         Case frmSMSMain.optSMSType(2).Value = True 'Grouplogo
298             gSMS.LifeTime = ActualValidityPeriodWithReturnValueInMinutes(False)

300             If frmSMSMain.chkRandomLogo.Value = 0 Then
302                 gSMS.BinaryFileLocation = frmSMSMain.txtPathLogo.Text
304                 gSMS.SendGroupLogo
                Else
306                 gSMS.BinaryFileLocation = App.Path & "\" & Right("0000" & Trim(Str$(gnSelectedRandomLogo)), 4) & ".bmp"
308                 gSMS.SendGroupLogo
                End If

310         Case frmSMSMain.optSMSType(3).Value = True 'Ringtone
312             gSMS.LifeTime = ActualValidityPeriodWithReturnValueInMinutes(False)
314             gSMS.BinaryFileLocation = frmSMSMain.txtPathRingtone.Text
316             gSMS.SendRingtone
  
318         Case frmSMSMain.optSMSType(4).Value = True 'Picturemessage
320             gSMS.LifeTime = ActualValidityPeriodWithReturnValueInMinutes(False)
322             gSMS.MessageData = frmSMSMain.txtPictureMessageText.Text
324             gSMS.BinaryFileLocation = frmSMSMain.txtPathPictureMessage.Text
326             gSMS.SendPictureMessage
  
328         Case frmSMSMain.optSMSType(5).Value = True 'VCard
330             gSMS.LifeTime = ActualValidityPeriodWithReturnValueInMinutes(False)
332             gSMS.VCard.Name = frmSMSMain.txtVCardName.Text
334             gSMS.VCard.PhoneNumber = frmSMSMain.txtVCardPhoneNumber.Text

336             If frmSMSMain.chkVCardBlinkingSMS.Value = 1 Then
338                 gSMS.BlinkingSMS = True
                Else
340                 gSMS.BlinkingSMS = False
                End If

342             gSMS.SendVCard
    
344         Case frmSMSMain.optSMSType(6).Value = True 'Unicode / UCS 2 Message
346             gSMS.LifeTime = ActualValidityPeriodWithReturnValueInMinutes(False)
348             sMessage = frmSMSMain.txtUnicode.Text
350             gSMS.MessageData = ConvertUnicodeScreenDataToUCS2MessageData(sMessage)
  
352             If frmSMSMain.chkFlashingSMSUnicode.Value Then
354                 gSMS.XSer = "020118"
                Else
356                 gSMS.XSer = "020108"
                End If
  
358             If frmSMSMain.chkReplaceMessageUnicode = 1 Then
360                 gSMS.ReplaceMessage = 5
                Else
                    'Do nothing
                End If

362             gSMS.SendBinaryData
  
364         Case frmSMSMain.optSMSType(7).Value = True 'WAP Push SMS
366             gSMS.LifeTime = ActualValidityPeriodWithReturnValueInMinutes(False)
368             gSMS.WAPPushSettings.Description = frmSMSMain.txtWAPPushSMSDescription
370             gSMS.BinaryFileLocation = frmSMSMain.txtWAPPushSMSPicturePath
372             gSMS.SendWAPPushSMS
    
374         Case frmSMSMain.optSMSType(8).Value = True 'Binarydata
376             gSMS.LifeTime = ActualValidityPeriodWithReturnValueInMinutes(False)
378             sMessage = frmSMSMain.txtMessageData.Text  'Remove Carriage Returns
380             sMessage = UCase(Replace(sMessage, vbCrLf, ""))
382             gSMS.MessageData = sMessage
384             gSMS.XSer = UCase(frmSMSMain.txtXSer.Text)
386             gSMS.SendBinaryData
  
388         Case frmSMSMain.optSMSType(9).Value = True
390             VersionSpecificAction 10
   
392         Case frmSMSMain.optSMSType(10).Value = True
394             VersionSpecificAction 11
     
396         Case frmSMSMain.optSMSType(11).Value = True
398             VersionSpecificAction 12
    
        End Select

400     Screen.MousePointer = vbDefault

402     If gSMS.errorCode <> 1 Then
404         MsgBox LoadLanguageSpecificString(gnLanguage, 207) & vbCrLf & ErrorDescriptionFromASPSMSErrorCode(gnLanguage, CInt(gSMS.errorCode)), vbExclamation, gsApplicationName
406         ProcessSendAction = False
        Else
408         ProcessSendAction = True
        End If

410     gSMS.DeleteAllRecipients
  
412     Set gSMS = Nothing
        '<EhFooter>
        Exit Function

ProcessSendAction_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ProcessSendAction " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Sub CenterForm(InputForm As Form)
        '<EhHeader>
        On Error GoTo CenterForm_Err
        '</EhHeader>

        'Centers Window independent of Screenresolution
100     If InputForm.WindowState <> 2 Then
102         InputForm.Move (Screen.Width - InputForm.Width) / 2, (Screen.Height - InputForm.Height) / 2
        End If

        '<EhFooter>
        Exit Sub

CenterForm_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.CenterForm " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function CheckForDuplicatesWithinCurrentRecipients() As Boolean
        '<EhHeader>
        On Error GoTo CheckForDuplicatesWithinCurrentRecipients_Err
        '</EhHeader>
        Dim sOrig() As PhoneBookEntry
        Dim sDest() As PhoneBookEntry
        Dim i As Long
        Dim sBackup As String
        Dim lIndex As Long
        Dim lDuplicates As Long
        Dim sTemp As String
        Dim sMessage As String
        Dim sInfoDuplicates As String
        Dim nShowDuplicates As Integer
        Dim nAnswer As Integer
        Const cMaxDuplicatesToShow = 20

100     If frmSMSMain.grdRecipients.Rows <= 1 Then Exit Function

102     ReDim tOrigPhoneBookEntry(1 To frmSMSMain.grdRecipients.Rows - 1) As PhoneBookEntry
104     ReDim tDestPhoneBookEntry(1 To frmSMSMain.grdRecipients.Rows - 1) As PhoneBookEntry
106     ReDim tFoundDuplicatesPhoneBookEntry(1 To cMaxDuplicatesToShow) As PhoneBookEntry

108     Screen.MousePointer = vbHourglass

        'Fill Array
110     For i = LBound(tOrigPhoneBookEntry) To UBound(tOrigPhoneBookEntry)
112         tOrigPhoneBookEntry(i).sName = frmSMSMain.grdRecipients.TextMatrix(i, 0)
114         tOrigPhoneBookEntry(i).sNumber = frmSMSMain.grdRecipients.TextMatrix(i, 1)
        Next

        'Sort Array by PhoneNumber
116     QuickSortPhoneBookEntryArrayByPhoneNumber tOrigPhoneBookEntry, LBound(tOrigPhoneBookEntry), UBound(tOrigPhoneBookEntry)

        'Remove Duplicates
118     lIndex = LBound(tOrigPhoneBookEntry)

120     For i = LBound(tOrigPhoneBookEntry) To UBound(tOrigPhoneBookEntry)

122         If tOrigPhoneBookEntry(i).sNumber <> sBackup Then
124             tDestPhoneBookEntry(lIndex) = tOrigPhoneBookEntry(i)
126             lIndex = lIndex + 1
            Else
128             lDuplicates = lDuplicates + 1

130             If lDuplicates <= cMaxDuplicatesToShow Then
132                 tFoundDuplicatesPhoneBookEntry(lDuplicates) = tOrigPhoneBookEntry(i)
                End If
            End If

134         sBackup = tOrigPhoneBookEntry(i).sNumber
        Next

136     Screen.MousePointer = vbDefault

138     If lDuplicates > 0 Then
140         If lDuplicates > cMaxDuplicatesToShow Then
142             nShowDuplicates = cMaxDuplicatesToShow
            Else
144             nShowDuplicates = lDuplicates
            End If
  
146         For i = 1 To nShowDuplicates

148             If tFoundDuplicatesPhoneBookEntry(i).sName = "" Then
150                 sInfoDuplicates = sInfoDuplicates & tFoundDuplicatesPhoneBookEntry(i).sNumber & vbCrLf
                Else
152                 sInfoDuplicates = sInfoDuplicates & tFoundDuplicatesPhoneBookEntry(i).sName & " / " & tFoundDuplicatesPhoneBookEntry(i).sNumber & vbCrLf
                End If

            Next

154         If lDuplicates > nShowDuplicates Then
156             sInfoDuplicates = sInfoDuplicates & LoadLanguageSpecificString(gnLanguage, 459) & vbCrLf  'And more...
            End If
  
158         If lDuplicates = 1 Then
160             sMessage = LoadLanguageSpecificString(gnLanguage, 453) & vbCrLf & vbCrLf & sInfoDuplicates & vbCrLf & LoadLanguageSpecificString(gnLanguage, 454) 'One Duplicate has been found. Remove?
            Else
162             sMessage = LoadLanguageSpecificString(gnLanguage, 455) & " " & Trim(Str$(lDuplicates)) & " " & LoadLanguageSpecificString(gnLanguage, 456) & vbCrLf & vbCrLf & sInfoDuplicates & vbCrLf & LoadLanguageSpecificString(gnLanguage, 457) 'Duplicates have been found. Remove?
            End If

164         nAnswer = MsgBox(sMessage, vbQuestion Or vbOKCancel, gsApplicationName)

166         If nAnswer = vbOK Then
                'Proceed
            Else
                Exit Function
            End If

        Else
168         CheckForDuplicatesWithinCurrentRecipients = True
            Exit Function
        End If
  
        '<EhFooter>
        Exit Function

CheckForDuplicatesWithinCurrentRecipients_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.CheckForDuplicatesWithinCurrentRecipients " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ConvertUnicodeScreenDataToUCS2MessageData(sInput As String) As String
        '<EhHeader>
        On Error GoTo ConvertUnicodeScreenDataToUCS2MessageData_Err
        '</EhHeader>
        Dim i As Integer
        Dim sTemp As String
        Dim lTemp As Long
        Dim sResult As String

100     sTemp = sInput

102     For i = 1 To Len(sTemp)
104         lTemp = AscW(Mid(sTemp, i, 1))
106         sResult = sResult & IntegerToHex(lTemp)
        Next

        'Cut down to 140 bytes, 70 chars
108     sResult = Left(sResult, 280)

110     ConvertUnicodeScreenDataToUCS2MessageData = sResult

        '<EhFooter>
        Exit Function

ConvertUnicodeScreenDataToUCS2MessageData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ConvertUnicodeScreenDataToUCS2MessageData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ExportSendLogIntoTextfile(nLanguage As Integer, nOutputFile As Integer, lJobId As Long, bWriteTitle As Boolean) As Boolean
        '<EhHeader>
        On Error GoTo ExportSendLogIntoTextfile_Err
        '</EhHeader>
        Dim sLine As String
        Dim i As Integer

        Dim rsMain As Recordset
        Dim sField As String
        Dim sEntry As String

        On Error GoTo ErrorTrap

100     Set rsMain = gdbMain.OpenRecordset("select * from Jobs where lID = " & Str$(lJobId))

102     If bWriteTitle Then
            'Write Header
104         sLine = LoadLanguageSpecificString(nLanguage, 551) & ": " & Now() & vbCrLf
106         Put #nOutputFile, , sLine
108         sLine = String(Len(sLine) - Len(vbCrLf), "-") & vbCrLf
110         Put #nOutputFile, , sLine
112         Put #nOutputFile, , vbCrLf
        Else
114         sLine = String(300, "-") & vbCrLf
116         Put #nOutputFile, , sLine
118         Put #nOutputFile, , vbCrLf
        End If

120     sLine = "Job ID: " & vbTab & " " & Trim(Str$(lJobId)) & vbCrLf
122     Put #nOutputFile, , sLine

124     For i = 1 To gcnNumberOfJobRemarkFields
126         sLine = GetJobRemarksFieldFromDatabase(nLanguage, i) & ": " & vbTab & rsMain("JobRemarksField" & Right("00" & Trim(Str$(i)), 2)) & vbCrLf
128         Put #nOutputFile, , sLine
        Next
  
        'Remarks
130     sLine = LoadLanguageSpecificString(nLanguage, 581) & ": " & vbTab & rsMain("JobRemarksMemo") & "" & vbCrLf & vbCrLf
132     Put #nOutputFile, , sLine
    
        'Date of Job
134     sLine = LoadLanguageSpecificString(nLanguage, 568) & ": " & vbTab & rsMain("FJobDate") & "" & vbCrLf
136     Put #nOutputFile, , sLine
    
        'Number of Recipients
138     sLine = LoadLanguageSpecificString(nLanguage, 569) & ": " & vbTab & rsMain("FNumberOfRecipients") & "" & vbCrLf
140     Put #nOutputFile, , sLine
    
        'Number of SMS
142     sLine = LoadLanguageSpecificString(nLanguage, 576) & ": " & vbTab & rsMain("FNumberOfMessages") & "" & vbCrLf
144     Put #nOutputFile, , sLine
    
        'SMS Type
146     sLine = LoadLanguageSpecificString(nLanguage, 570) & ": " & vbTab & rsMain("FSMSType") & "" & vbCrLf
148     Put #nOutputFile, , sLine
    
        'Sending Type
150     sLine = LoadLanguageSpecificString(nLanguage, 571) & ": " & vbTab & rsMain("FSendingType") & "" & vbCrLf
152     Put #nOutputFile, , sLine

        'Starting Date
154     If rsMain("FWaitingTime") <> "" Then
156         sLine = LoadLanguageSpecificString(nLanguage, 573) & ": " & vbTab & rsMain("FStartingDate") & "" & vbCrLf  'StartingDate
158         Put #nOutputFile, , sLine
        Else

160         If rsMain("FStartingDate") <> "" Then
162             sLine = LoadLanguageSpecificString(nLanguage, 604) & ": " & vbTab & rsMain("FStartingDate") & "" & vbCrLf  'DeferredDeliveryTime
164             Put #nOutputFile, , sLine
            End If
        End If
  
        'Ending Date
166     If rsMain("FWaitingTime") <> "" Then
168         sLine = LoadLanguageSpecificString(nLanguage, 574) & ": " & vbTab & rsMain("FEndingDate") & "" & vbCrLf
170         Put #nOutputFile, , sLine
        End If
  
        'Waiting Time between SMS
172     If rsMain("FWaitingTime") <> "" Then
174         sLine = LoadLanguageSpecificString(nLanguage, 575) & ": " & vbTab & rsMain("FWaitingTime") & "" & vbCrLf
176         Put #nOutputFile, , sLine
        End If
    
178     sLine = vbCrLf
180     Put #nOutputFile, , sLine
  
182     rsMain.Close
 
        'Write Header
184     sLine = "Details" & vbCrLf
186     Put #nOutputFile, , sLine
188     sLine = String(Len(sLine) - Len(vbCrLf), "-") & vbCrLf
190     Put #nOutputFile, , sLine

192     sLine = vbCrLf
194     Put #nOutputFile, , sLine

196     sLine = LoadLanguageSpecificString(nLanguage, 600) & vbTab
198     sLine = sLine & LoadLanguageSpecificString(nLanguage, 601) & vbTab
200     sLine = sLine & LoadLanguageSpecificString(nLanguage, 602) & vbTab
202     sLine = sLine & LoadLanguageSpecificString(nLanguage, 603) & vbTab
204     sLine = sLine & LoadLanguageSpecificString(nLanguage, 604) & vbTab
206     sLine = sLine & LoadLanguageSpecificString(nLanguage, 605) & vbTab
208     sLine = sLine & LoadLanguageSpecificString(nLanguage, 606) & vbTab
210     sLine = sLine & LoadLanguageSpecificString(nLanguage, 611) & vbTab
212     sLine = sLine & LoadLanguageSpecificString(nLanguage, 612) & vbTab
214     sLine = sLine & LoadLanguageSpecificString(nLanguage, 607) & vbTab
216     sLine = sLine & LoadLanguageSpecificString(nLanguage, 613) & vbTab
218     sLine = sLine & LoadLanguageSpecificString(nLanguage, 608) & vbTab
220     sLine = sLine & LoadLanguageSpecificString(nLanguage, 609) & vbTab
222     sLine = sLine & LoadLanguageSpecificString(nLanguage, 610) & vbCrLf

224     Put #nOutputFile, , sLine

226     Set rsMain = gdbMain.OpenRecordset("select * from SendJournal where lJobID = " & Str$(lJobId) & " order by dSubmissionDate, lID")

228     Do While Not rsMain.EOF
230         sLine = rsMain("sName") & vbTab
232         sLine = sLine & rsMain("sRecipient") & vbTab
234         sLine = sLine & rsMain("sTransactionReferenceNumber") & vbTab
236         sLine = sLine & rsMain("sMessage") & vbTab
238         sLine = sLine & rsMain("dDeferredDeliveryTime") & vbTab
240         sLine = sLine & rsMain("dSubmissionDate") & vbTab
242         sLine = sLine & rsMain("dNotificationDate") & vbTab
244         sLine = sLine & rsMain("sDeliveryStatus") & vbTab
246         sLine = sLine & rsMain("sDeliveryStatusText") & vbTab
248         sLine = sLine & rsMain("sReasonCode") & vbTab
250         sLine = sLine & rsMain("sReasonText") & vbTab
252         sLine = sLine & rsMain("dNotificationDateBuffered") & vbTab
254         sLine = sLine & rsMain("sReasoncodeBuffered") & vbTab
256         sLine = sLine & rsMain("sReasonTextBuffered")
  
258         sLine = Replace(sLine, vbCrLf, " ") & vbCrLf
260         Put #nOutputFile, , sLine
262         rsMain.MoveNext
        Loop

264     rsMain.Close
266     ExportSendLogIntoTextfile = True
        Exit Function

ErrorTrap:
268     ExportSendLogIntoTextfile = False
        Exit Function
        '<EhFooter>
        Exit Function

ExportSendLogIntoTextfile_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ExportSendLogIntoTextfile " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function GetPhonebookVariableFieldFromDatabase(nLanguage As Integer, nIndex As Integer) As String
        '<EhHeader>
        On Error GoTo GetPhonebookVariableFieldFromDatabase_Err
        '</EhHeader>
        Dim sPhonebookVariableField As String

100     sPhonebookVariableField = GetSettingFromDatabase("PhonebookVariableField0" & Trim(Str$(nIndex)))

102     If sPhonebookVariableField = "DEFAULT" Then
104         GetPhonebookVariableFieldFromDatabase = LoadLanguageSpecificString(nLanguage, 212 + nIndex)
        Else
106         GetPhonebookVariableFieldFromDatabase = sPhonebookVariableField
        End If

        '<EhFooter>
        Exit Function

GetPhonebookVariableFieldFromDatabase_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.GetPhonebookVariableFieldFromDatabase " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function GetJobRemarksFieldFromDatabase(nLanguage As Integer, nIndex As Integer) As String
        '<EhHeader>
        On Error GoTo GetJobRemarksFieldFromDatabase_Err
        '</EhHeader>
        Dim sJobRemarksField As String

100     sJobRemarksField = GetSettingFromDatabase("JobRemarksField0" & Trim(Str$(nIndex)))

102     If sJobRemarksField = "DEFAULT" Then
104         VersionSpecificAction 47 + nIndex, nLanguage, , sJobRemarksField
        End If

106     GetJobRemarksFieldFromDatabase = sJobRemarksField

        '<EhFooter>
        Exit Function

GetJobRemarksFieldFromDatabase_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.GetJobRemarksFieldFromDatabase " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub InitJobLogColWidthSettings()
        '<EhHeader>
        On Error GoTo InitJobLogColWidthSettings_Err
        '</EhHeader>
        Dim i As Integer
        Dim nUsedDBFields As Integer
        Dim rColWidth As Double

100     For i = 1 To 16

102         If gtDBLogFieldsSettings.bJobLog(i) = True Then
104             nUsedDBFields = nUsedDBFields + 1
            End If

        Next

106     For i = 0 To 15
108         rColWidth = Screen.Width / nUsedDBFields * 0.97
110         PutSettingIntoDataBase "ColWidthJobs" & Right("00" & Trim(Str$(i)), 2), rColWidth
        Next

        '<EhFooter>
        Exit Sub

InitJobLogColWidthSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.InitJobLogColWidthSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub InitSendLogColWidthSettings()
        '<EhHeader>
        On Error GoTo InitSendLogColWidthSettings_Err
        '</EhHeader>
        Dim i As Integer
        Dim nUsedDBFields As Integer
        Dim rColWidth As Double

100     For i = 1 To 16

102         If gtDBLogFieldsSettings.bSendLog(i) = True Then
104             nUsedDBFields = nUsedDBFields + 1
            End If

        Next

106     For i = 0 To 15
108         rColWidth = Screen.Width / nUsedDBFields * 0.97
110         PutSettingIntoDataBase "ColWidthSendLog" & Right("00" & Trim(Str$(i)), 2), rColWidth
        Next

        '<EhFooter>
        Exit Sub

InitSendLogColWidthSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.InitSendLogColWidthSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function IntegerToHex(ByVal lInput As Long) As String
        '<EhHeader>
        On Error GoTo IntegerToHex_Err
        '</EhHeader>
        Dim lTemp As Long
        Dim sTemp As String
        Dim nNibbleMask1 As Long
        Dim nNibbleMask2 As Long
        Dim nNibbleMask3 As Long
        Dim nNibbleMask4 As Long

        'Convert to unsigned integer
100     lTemp = lInput And &HFFFF

102     nNibbleMask1 = &HF
104     nNibbleMask2 = &HF0
106     nNibbleMask3 = &HF00
108     nNibbleMask4 = &HF000

110     sTemp = NibbleDezToHex((lTemp And nNibbleMask1))
112     sTemp = NibbleDezToHex((lTemp And nNibbleMask2) / (&H10)) & sTemp
114     sTemp = NibbleDezToHex((lTemp And nNibbleMask3) / (&H100)) & sTemp
116     sTemp = NibbleDezToHex((lTemp And nNibbleMask4) / (&H1000)) & sTemp

118     IntegerToHex = sTemp
        '<EhFooter>
        Exit Function

IntegerToHex_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.IntegerToHex " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function JobInfoCurrentTime() As Date
        '<EhHeader>
        On Error GoTo JobInfoCurrentTime_Err
        '</EhHeader>
100     JobInfoCurrentTime = Now()
        '<EhFooter>
        Exit Function

JobInfoCurrentTime_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobInfoCurrentTime " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function JobInfoEndingDate() As Variant
        '<EhHeader>
        On Error GoTo JobInfoEndingDate_Err
        '</EhHeader>
        Dim lRecipientMultiplier As Long
        Dim dStart As Date
        Dim dEnd As Date

100     lRecipientMultiplier = 1

102     If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
104         If gtDeferredDeliveryTimeSettings.bSingleSMS = True Then
106             JobInfoEndingDate = ""
            End If

108         If gtDeferredDeliveryTimeSettings.bPeriodicSMS = True Then
                'Sending Type
110             lRecipientMultiplier = Val(gtDeferredDeliveryTimeSettings.sNumberOfMessages) - 1
        
112             dStart = gtDeferredDeliveryTimeSettings.sStartingDate
          
114             Select Case gtDeferredDeliveryTimeSettings.nWaitingPeriod

                    Case 0
116                     dEnd = DateAdd("s", lRecipientMultiplier * gtDeferredDeliveryTimeSettings.sWaitingTime, dStart)
            
118                 Case 1
120                     dEnd = DateAdd("n", lRecipientMultiplier * gtDeferredDeliveryTimeSettings.sWaitingTime, dStart)
      
122                 Case 2
124                     dEnd = DateAdd("h", lRecipientMultiplier * gtDeferredDeliveryTimeSettings.sWaitingTime, dStart)
      
126                 Case 3
128                     dEnd = DateAdd("d", lRecipientMultiplier * gtDeferredDeliveryTimeSettings.sWaitingTime, dStart)
      
130                 Case 4
132                     dEnd = DateAdd("ww", lRecipientMultiplier * gtDeferredDeliveryTimeSettings.sWaitingTime, dStart)
      
134                 Case 5
136                     dEnd = DateAdd("m", lRecipientMultiplier * gtDeferredDeliveryTimeSettings.sWaitingTime, dStart)
    
138                 Case Else
      
                End Select
    
140             JobInfoEndingDate = dEnd
            Else
142             JobInfoEndingDate = ""
            End If

        Else
144         JobInfoEndingDate = ""
        End If

        Exit Function

ErrorTrap:
        Exit Function
        '<EhFooter>
        Exit Function

JobInfoEndingDate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobInfoEndingDate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Function JobInfoNumberOfMessages() As Long
        '<EhHeader>
        On Error GoTo JobInfoNumberOfMessages_Err
        '</EhHeader>
        Dim lCountRecipients As Long
        Dim lRecipientMultiplier As Long
        Dim lCountMessages As Long

        'Number of Recipients
100     lCountRecipients = frmSMSMain.grdRecipients.Rows - 1

102     lRecipientMultiplier = 1

104     If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
106         If gtDeferredDeliveryTimeSettings.bPeriodicSMS = True Then
108             lRecipientMultiplier = Val(gtDeferredDeliveryTimeSettings.sNumberOfMessages)
            End If
        End If

110     lCountMessages = lCountRecipients * lRecipientMultiplier
112     JobInfoNumberOfMessages = lCountMessages
        '<EhFooter>
        Exit Function

JobInfoNumberOfMessages_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobInfoNumberOfMessages " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Function JobInfoNumberOfRecipients() As Long
        '<EhHeader>
        On Error GoTo JobInfoNumberOfRecipients_Err
        '</EhHeader>
100     JobInfoNumberOfRecipients = frmSMSMain.grdRecipients.Rows - 1
        '<EhFooter>
        Exit Function

JobInfoNumberOfRecipients_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobInfoNumberOfRecipients " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function JobInfoSendingType(nLanguage As Integer) As String
        '<EhHeader>
        On Error GoTo JobInfoSendingType_Err
        '</EhHeader>

100     If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
102         If gtDeferredDeliveryTimeSettings.bSingleSMS = True Then
                'Sending Type
104             JobInfoSendingType = LoadLanguageSpecificString(nLanguage, 577)
            End If

106         If gtDeferredDeliveryTimeSettings.bPeriodicSMS = True Then
108             JobInfoSendingType = LoadLanguageSpecificString(nLanguage, 578)
            End If

        Else
            'Sending Type
110         JobInfoSendingType = LoadLanguageSpecificString(nLanguage, 579)
        End If

        '<EhFooter>
        Exit Function

JobInfoSendingType_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobInfoSendingType " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Function JobInfoSMSType(nLanguage As Integer)
        'SMS Type
        '<EhHeader>
        On Error GoTo JobInfoSMSType_Err
        '</EhHeader>
100     JobInfoSMSType = SelectedSMSType(nLanguage)
        '<EhFooter>
        Exit Function

JobInfoSMSType_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobInfoSMSType " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function JobInfoStartingDate() As Variant
        '<EhHeader>
        On Error GoTo JobInfoStartingDate_Err
        '</EhHeader>

100     If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
102         If gtDeferredDeliveryTimeSettings.bSingleSMS = True Then
104             JobInfoStartingDate = CVDate(gtDeferredDeliveryTimeSettings.sDeliveryDate)
            End If

106         If gtDeferredDeliveryTimeSettings.bPeriodicSMS = True Then
108             JobInfoStartingDate = CVDate(gtDeferredDeliveryTimeSettings.sStartingDate)
            End If

        Else
        End If

        '<EhFooter>
        Exit Function

JobInfoStartingDate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobInfoStartingDate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Function JobInfoWaitingTimeBetweenMessages(nLanguage As Integer) As String
        '<EhHeader>
        On Error GoTo JobInfoWaitingTimeBetweenMessages_Err
        '</EhHeader>
        Dim sWaitingPeriodInfo As String

100     If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
102         If gtDeferredDeliveryTimeSettings.bSingleSMS = True Then
            End If

104         If gtDeferredDeliveryTimeSettings.bPeriodicSMS = True Then
    
106             Select Case gtDeferredDeliveryTimeSettings.nWaitingPeriod

                    Case 0
108                     sWaitingPeriodInfo = LoadLanguageSpecificString(nLanguage, 187)
                  
110                 Case 1
112                     sWaitingPeriodInfo = LoadLanguageSpecificString(nLanguage, 188)
            
114                 Case 2
116                     sWaitingPeriodInfo = LoadLanguageSpecificString(nLanguage, 189)
            
118                 Case 3
120                     sWaitingPeriodInfo = LoadLanguageSpecificString(nLanguage, 190)
            
122                 Case 4
124                     sWaitingPeriodInfo = LoadLanguageSpecificString(nLanguage, 191)
            
126                 Case 5
128                     sWaitingPeriodInfo = LoadLanguageSpecificString(nLanguage, 192)
          
130                 Case Else
      
                End Select
    
132             JobInfoWaitingTimeBetweenMessages = gtDeferredDeliveryTimeSettings.sWaitingTime & " " & sWaitingPeriodInfo

            End If

        Else
            'Do nothing
        End If

        '<EhFooter>
        Exit Function

JobInfoWaitingTimeBetweenMessages_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobInfoWaitingTimeBetweenMessages " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function JobInfoWaitingTimeInfo(nLanguage As Integer) As String
        '<EhHeader>
        On Error GoTo JobInfoWaitingTimeInfo_Err
        '</EhHeader>
        Dim sWaitingTimeInfo As String

100     Select Case gtDeferredDeliveryTimeSettings.nWaitingPeriod
  
            Case 0
102             sWaitingTimeInfo = LoadLanguageSpecificString(nLanguage, 187)
  
104         Case 1
106             sWaitingTimeInfo = LoadLanguageSpecificString(nLanguage, 188)
  
108         Case 2
110             sWaitingTimeInfo = LoadLanguageSpecificString(nLanguage, 189)
  
112         Case 3
114             sWaitingTimeInfo = LoadLanguageSpecificString(nLanguage, 190)
  
116         Case 4
118             sWaitingTimeInfo = LoadLanguageSpecificString(nLanguage, 191)
  
120         Case 5
122             sWaitingTimeInfo = LoadLanguageSpecificString(nLanguage, 192)
  
124         Case Else
        End Select

126     JobInfoWaitingTimeInfo = sWaitingTimeInfo
        '<EhFooter>
        Exit Function

JobInfoWaitingTimeInfo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobInfoWaitingTimeInfo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function JobLogSQLStatement(nLanguage As Integer) As String
        '<EhHeader>
        On Error GoTo JobLogSQLStatement_Err
        '</EhHeader>
        Dim sSQL As String
        Dim sWhere As String
        Dim sField As String
        Dim i As Integer
       
100     ReDim sJobRemarksField(1 To gcnNumberOfJobRemarkFields) As String
       
102     For i = 1 To gcnNumberOfJobRemarkFields
104         sJobRemarksField(i) = GetJobRemarksFieldFromDatabase(nLanguage, i)
        Next

106     sSQL = sSQL & "SELECT lID as ID, "

108     If gtDBLogFieldsSettings.bJobLog(2) = True Then
110         sSQL = sSQL & "JobRemarksField01 as [" & sJobRemarksField(1) & "], "
        End If

112     If gtDBLogFieldsSettings.bJobLog(3) = True Then
114         sSQL = sSQL & "JobRemarksField02 as [" & sJobRemarksField(2) & "], "
        End If

116     If gtDBLogFieldsSettings.bJobLog(4) = True Then
118         sSQL = sSQL & "JobRemarksField03 as [" & sJobRemarksField(3) & "], "
        End If

120     If gtDBLogFieldsSettings.bJobLog(5) = True Then
122         sSQL = sSQL & "JobRemarksField04 as [" & sJobRemarksField(4) & "], "
        End If

124     If gtDBLogFieldsSettings.bJobLog(6) = True Then
126         sSQL = sSQL & "JobRemarksField05 as [" & sJobRemarksField(5) & "], "
        End If

128     If gtDBLogFieldsSettings.bJobLog(7) = True Then
130         sSQL = sSQL & "JobRemarksField06 as [" & sJobRemarksField(6) & "], "
        End If

132     If gtDBLogFieldsSettings.bJobLog(8) = True Then
134         sSQL = sSQL & "JobRemarksMemo as [" & LoadLanguageSpecificString(nLanguage, 581) & "], "
        End If

136     If gtDBLogFieldsSettings.bJobLog(9) = True Then
138         sSQL = sSQL & "FSendingType as [" & LoadLanguageSpecificString(nLanguage, 552) & "], "
        End If

140     If gtDBLogFieldsSettings.bJobLog(10) = True Then
142         sSQL = sSQL & "FSMSType as [" & LoadLanguageSpecificString(nLanguage, 553) & "], "
        End If

144     If gtDBLogFieldsSettings.bJobLog(11) = True Then
146         sSQL = sSQL & "FJobDate as [" & LoadLanguageSpecificString(nLanguage, 554) & "], "
        End If

148     If gtDBLogFieldsSettings.bJobLog(12) = True Then
150         sSQL = sSQL & "FStartingDate as [" & LoadLanguageSpecificString(nLanguage, 555) & "], "
        End If

152     If gtDBLogFieldsSettings.bJobLog(13) = True Then
154         sSQL = sSQL & "FEndingDate as [" & LoadLanguageSpecificString(nLanguage, 556) & "], "
        End If

156     If gtDBLogFieldsSettings.bJobLog(14) = True Then
158         sSQL = sSQL & "FWaitingTime as [" & LoadLanguageSpecificString(nLanguage, 557) & "], "
        End If

160     If gtDBLogFieldsSettings.bJobLog(15) = True Then
162         sSQL = sSQL & "FNumberOfMessages as [" & LoadLanguageSpecificString(nLanguage, 558) & "], "
        End If

164     If gtDBLogFieldsSettings.bJobLog(16) = True Then
166         sSQL = sSQL & "FNumberOfRecipients as [" & LoadLanguageSpecificString(nLanguage, 559) & "], "
        End If

168     sSQL = Left(sSQL, Len(sSQL) - 2)

170     sSQL = sSQL & " FROM Jobs "

172     For i = 1 To gcnNumberOfJobRemarkFields

174         If gtJobLogFilterSettings.bJobRemarksFilterUsed(i) = True Then
176             sField = "JobRemarksField" & Right("0" & Trim(Str$(i)), 2)
178             sWhere = sWhere & "(" & sField & " = '" & gtJobLogFilterSettings.sJobRemarks(i) & "') and "
            Else
                'Do nothing
            End If

        Next

180     If gtJobLogFilterSettings.bSMSTypeFilterUsed = True Then
182         sWhere = sWhere & "(FSMSType = '" & gtJobLogFilterSettings.sSMSType & "') and "
        End If

184     If gtJobLogFilterSettings.bSendingTypeFilterUsed = True Then
186         sWhere = sWhere & "(FSendingType = '" & gtJobLogFilterSettings.sSendingType & "') and "
        End If

188     If sWhere <> "" Then 'Cut off last 'and'
190         sWhere = "WHERE " & Left(sWhere, Len(sWhere) - 4)
        End If

192     JobLogSQLStatement = sSQL & sWhere & " ORDER BY lID"
        '<EhFooter>
        Exit Function

JobLogSQLStatement_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.JobLogSQLStatement " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function NibbleDezToHex(nInput As Integer) As String
        '<EhHeader>
        On Error GoTo NibbleDezToHex_Err
        '</EhHeader>

100     Select Case nInput

            Case 0 To 9
102             NibbleDezToHex = Trim(Str$(nInput))
  
104         Case 10
106             NibbleDezToHex = "A"
  
108         Case 11
110             NibbleDezToHex = "B"
  
112         Case 12
114             NibbleDezToHex = "C"
  
116         Case 13
118             NibbleDezToHex = "D"
  
120         Case 14
122             NibbleDezToHex = "E"
  
124         Case 15
126             NibbleDezToHex = "F"
  
128         Case Else
                '"Interner Fehler NibbleDezToHex"
  
        End Select
  
        '<EhFooter>
        Exit Function

NibbleDezToHex_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.NibbleDezToHex " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function PersonalizedSMSModeUsed(sInput As String) As Boolean
        '<EhHeader>
        On Error GoTo PersonalizedSMSModeUsed_Err
        '</EhHeader>
        Dim bPlaceholdersUsed As Boolean

100     Select Case True

            Case InStr(sInput, LoadLanguageSpecificString(1, 234)) >= 1
102             bPlaceholdersUsed = True
  
104         Case InStr(sInput, LoadLanguageSpecificString(2, 234)) >= 1
106             bPlaceholdersUsed = True
  
108         Case InStr(sInput, LoadLanguageSpecificString(1, 235)) >= 1
110             bPlaceholdersUsed = True
  
112         Case InStr(sInput, LoadLanguageSpecificString(2, 235)) >= 1
114             bPlaceholdersUsed = True
  
116         Case InStr(sInput, "<" & LoadLanguageSpecificString(1, 213) & ">") >= 1
118             bPlaceholdersUsed = True
  
120         Case InStr(sInput, "<" & LoadLanguageSpecificString(2, 213) & ">") >= 1
122             bPlaceholdersUsed = True
  
124         Case InStr(sInput, "<" & LoadLanguageSpecificString(1, 214) & ">") >= 1
126             bPlaceholdersUsed = True
  
128         Case InStr(sInput, "<" & LoadLanguageSpecificString(2, 214) & ">") >= 1
130             bPlaceholdersUsed = True
  
132         Case InStr(sInput, "<" & LoadLanguageSpecificString(1, 215) & ">") >= 1
134             bPlaceholdersUsed = True
  
136         Case InStr(sInput, "<" & LoadLanguageSpecificString(2, 215) & ">") >= 1
138             bPlaceholdersUsed = True
  
140         Case InStr(sInput, "<" & LoadLanguageSpecificString(1, 216) & ">") >= 1
142             bPlaceholdersUsed = True
  
144         Case InStr(sInput, "<" & LoadLanguageSpecificString(2, 216) & ">") >= 1
146             bPlaceholdersUsed = True
  
148         Case InStr(sInput, "<" & gsPhonebookVariableField(1, 1) & ">") >= 1
150             bPlaceholdersUsed = True
  
152         Case InStr(sInput, "<" & gsPhonebookVariableField(2, 1) & ">") >= 1
154             bPlaceholdersUsed = True
  
156         Case InStr(sInput, "<" & gsPhonebookVariableField(1, 2) & ">") >= 1
158             bPlaceholdersUsed = True
  
160         Case InStr(sInput, "<" & gsPhonebookVariableField(2, 2) & ">") >= 1
162             bPlaceholdersUsed = True
  
164         Case InStr(sInput, "<" & gsPhonebookVariableField(1, 3) & ">") >= 1
166             bPlaceholdersUsed = True
  
168         Case InStr(sInput, "<" & gsPhonebookVariableField(2, 3) & ">") >= 1
170             bPlaceholdersUsed = True

        End Select

172     If bPlaceholdersUsed = True Then
174         PersonalizedSMSModeUsed = True
        Else
176         PersonalizedSMSModeUsed = False
        End If

        '<EhFooter>
        Exit Function

PersonalizedSMSModeUsed_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.PersonalizedSMSModeUsed " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function PhonebookSQLStatement() As String
        '<EhHeader>
        On Error GoTo PhonebookSQLStatement_Err
        '</EhHeader>
        Dim sSQL As String
        Dim sTemp As String

100     sTemp = frmSMSMain.txtFilter.Text
102     sTemp = Replace(sTemp, "'", "")

104     If sTemp <> "" Then
106         sSQL = "SELECT ID, "
108         sSQL = sSQL & "sName as [" & LoadLanguageSpecificString(gnLanguage, 196) & "], "
110         sSQL = sSQL & "sNumber as [" & LoadLanguageSpecificString(gnLanguage, 197) & "], "
112         sSQL = sSQL & "Variable1 as [" & gsPhonebookVariableField(gnLanguage, 1) & "], "
114         sSQL = sSQL & "Variable2 as [" & gsPhonebookVariableField(gnLanguage, 2) & "], "
116         sSQL = sSQL & "Variable3 as [" & gsPhonebookVariableField(gnLanguage, 3) & "] "
118         sSQL = sSQL & "FROM Phonebook "
120         sSQL = sSQL & "WHERE sName like '*" & sTemp & "*' or "
122         sSQL = sSQL & "sNumber like '*" & sTemp & "*' or "
124         sSQL = sSQL & "Variable1 like '*" & sTemp & "*' or "
126         sSQL = sSQL & "Variable2 like '*" & sTemp & "*' or "
128         sSQL = sSQL & "Variable3 like '*" & sTemp & "*' "
130         sSQL = sSQL & SQLOrderPartFromPhonebookOrder()
        Else
132         sSQL = "SELECT ID, "
134         sSQL = sSQL & "sName as [" & LoadLanguageSpecificString(gnLanguage, 196) & "], "
136         sSQL = sSQL & "sNumber as [" & LoadLanguageSpecificString(gnLanguage, 197) & "], "
138         sSQL = sSQL & "Variable1 as [" & gsPhonebookVariableField(gnLanguage, 1) & "], "
140         sSQL = sSQL & "Variable2 as [" & gsPhonebookVariableField(gnLanguage, 2) & "], "
142         sSQL = sSQL & "Variable3 as [" & gsPhonebookVariableField(gnLanguage, 3) & "] "
144         sSQL = sSQL & "FROM Phonebook " & SQLOrderPartFromPhonebookOrder()
        End If

146     PhonebookSQLStatement = sSQL
        '<EhFooter>
        Exit Function

PhonebookSQLStatement_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.PhonebookSQLStatement " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub PutJobRemarksFieldIntoDatabase(nLanguage As Integer, _
                                   nIndex As Integer, _
                                   sJobRemarksField As String)
        '<EhHeader>
        On Error GoTo PutJobRemarksFieldIntoDatabase_Err
        '</EhHeader>
        Dim sTemp As String

100     VersionSpecificAction 47 + nIndex, nLanguage, , sTemp

102     If sJobRemarksField = sTemp Then
104         PutSettingIntoDataBase "JobRemarksField" & Right("00" & Trim(Str$(nIndex)), 2), "DEFAULT"
        Else
106         PutSettingIntoDataBase "JobRemarksField" & Right("00" & Trim(Str$(nIndex)), 2), sJobRemarksField
        End If

        '<EhFooter>
        Exit Sub

PutJobRemarksFieldIntoDatabase_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.PutJobRemarksFieldIntoDatabase " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub QuickSortPhoneBookEntryArrayByName(sInputArray() As PhoneBookEntry, _
                                       l As Long, _
                                       r As Long)
        '<EhHeader>
        On Error GoTo QuickSortPhoneBookEntryArrayByName_Err
        '</EhHeader>
        Dim i As Long
        Dim j As Long
        Dim x As PhoneBookEntry
        Dim tPhoneBookEntryTemp As PhoneBookEntry

100     i = l
102     j = r
104     x = sInputArray((l + r) / 2)
106     While (i <= j)
108         While (sInputArray(i).sName < x.sName And i < r)
110             i = i + 1
            Wend

112         While (x.sName < sInputArray(j).sName And j > l)
114             j = j - 1
            Wend

116         If (i <= j) Then
118             tPhoneBookEntryTemp = sInputArray(i)           'Swap
120             sInputArray(i) = sInputArray(j)                'Swap
122             sInputArray(j) = tPhoneBookEntryTemp           'Swap
124             i = i + 1
126             j = j - 1
            End If

        Wend

128     If (l < j) Then Call QuickSortPhoneBookEntryArrayByName(sInputArray(), l, j)
130     If (i < r) Then Call QuickSortPhoneBookEntryArrayByName(sInputArray(), i, r)

        '<EhFooter>
        Exit Sub

QuickSortPhoneBookEntryArrayByName_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.QuickSortPhoneBookEntryArrayByName " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub QuickSortPhoneBookEntryArrayByPhoneNumber(sInputArray() As PhoneBookEntry, _
                                              l As Long, _
                                              r As Long)
        '<EhHeader>
        On Error GoTo QuickSortPhoneBookEntryArrayByPhoneNumber_Err
        '</EhHeader>
        Dim i As Long
        Dim j As Long
        Dim x As PhoneBookEntry
        Dim tPhoneBookEntryTemp As PhoneBookEntry

100     i = l
102     j = r
104     x = sInputArray((l + r) / 2)
106     While (i <= j)
108         While (sInputArray(i).sNumber < x.sNumber And i < r)
110             i = i + 1
            Wend

112         While (x.sNumber < sInputArray(j).sNumber And j > l)
114             j = j - 1
            Wend

116         If (i <= j) Then
118             tPhoneBookEntryTemp = sInputArray(i)           'Swap
120             sInputArray(i) = sInputArray(j)                'Swap
122             sInputArray(j) = tPhoneBookEntryTemp           'Swap
124             i = i + 1
126             j = j - 1
            End If

        Wend

128     If (l < j) Then Call QuickSortPhoneBookEntryArrayByPhoneNumber(sInputArray(), l, j)
130     If (i < r) Then Call QuickSortPhoneBookEntryArrayByPhoneNumber(sInputArray(), i, r)

        '<EhFooter>
        Exit Sub

QuickSortPhoneBookEntryArrayByPhoneNumber_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.QuickSortPhoneBookEntryArrayByPhoneNumber " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function ErrorDescriptionFromASPSMSErrorCode(nLanguage As Integer, nErrorCode As Integer) As String
        '<EhHeader>
        On Error GoTo ErrorDescriptionFromASPSMSErrorCode_Err
        '</EhHeader>

100     Select Case nErrorCode

            Case gcErrorCodeConnectFailed
102             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 420)
  
104         Case gcErrorCodeAuthorizationFailed
106             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 421)
  
108         Case gcErrorCodeBinaryFileNotFound
110             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 422)
  
112         Case gcErrorCodeNotEnoughCreditsAvailable
114             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 423)
  
116         Case gcErrorCodeTimeOutError
118             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 424)
  
120         Case gcErrorCodeTransmissionErrorTryAgain
122             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 425)
  
124         Case gcErrorCodeInvalidUserKey
126             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 426)
  
128         Case gcErrorCodeInvalidPassword
130             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 427)
  
132         Case gcErrorCodeInvalidOriginator
134             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 428)
  
136         Case gcErrorCodeInvalidMessageData
138             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 429)
  
140         Case gcErrorCodeInvalidBinaryData
142             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 430)
  
144         Case gcErrorCodeInvalidBinaryFile
146             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 431)
  
148         Case gcErrorCodeInvalidMCC
150             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 432)
  
152         Case gcErrorCodeInvalidMNC
154             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 433)
  
156         Case gcErrorCodeInvalidxSer
158             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 434)
  
160         Case gcErrorCodeInvalidURLBufferedMessageNotification
162             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 435)
  
164         Case gcErrorCodeInvalidURLDeliveryNotification
166             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 436)
  
168         Case gcErrorCodeInvalidURLNonDeliveryNotification
170             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 437)
  
172         Case gcErrorCodeMissingRecipient
174             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 438)
  
176         Case gcErrorCodeMissingBinaryData
178             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 439)
  
180         Case gcErrorCodeInvalidDeferredDeliveryTime
182             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 440)
  
184         Case gcErrorCodeMissingTransactionReferenceNumber
186             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 441)
  
188         Case gcErrorCodeServiceTemporaryNotAvailable
190             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 442)
  
192         Case gcErrorCodeUserBarringActive
194             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 443)
  
196         Case gcErrorCodeNotAuthorizedForThisOperation
198             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 444)
  
200         Case gcErrorCodeMessageTooLong
202             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 445)
  
204         Case gcErrorCodeNoOriginatorRestrictions
206             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 446)
  
208         Case gcErrorCodeOriginatorAuthorizationPending
210             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 447)
  
212         Case gcErrorCodeOriginatorNotAuthorized
214             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 448)
  
216         Case gcErrorCodeOriginatorAlreadyAuthorized
218             ErrorDescriptionFromASPSMSErrorCode = LoadLanguageSpecificString(nLanguage, 449)
    
220         Case Else
  
        End Select

        '<EhFooter>
        Exit Function

ErrorDescriptionFromASPSMSErrorCode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ErrorDescriptionFromASPSMSErrorCode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function FieldExists(rsInput As Recordset, sField As String) As Boolean
        '<EhHeader>
        On Error GoTo FieldExists_Err
        '</EhHeader>
        Dim vTemp As Variant
        On Error GoTo ErrorTrap

100     vTemp = rsInput(sField)

102     FieldExists = True

        Exit Function
ErrorTrap:

104     FieldExists = False
        Exit Function
        '<EhFooter>
        Exit Function

FieldExists_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.FieldExists " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub RefreshImportCounter()
        '<EhHeader>
        On Error GoTo RefreshImportCounter_Err
        '</EhHeader>
100    ' frmImport.fraImport.caption = LoadLanguageSpecificString(gnLanguage, 641) & " / " & Trim(Str$(frmImport.grdImport.Rows)) & " " & LoadLanguageSpecificString(gnLanguage, 642)
        '<EhFooter>
        Exit Sub

RefreshImportCounter_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.RefreshImportCounter " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub RestoreJobInfoSettings()
        '<EhHeader>
        On Error GoTo RestoreJobInfoSettings_Err
        '</EhHeader>
        Dim i As Integer

100     For i = 1 To gcnNumberOfJobRemarkFields
102         UpdateComboBoxRemarksfrmSMSMain i
        Next

104     For i = 1 To gcnNumberOfJobRemarkFields
106         frmSMSMain.cboJobRemarks(i).Text = GetSettingFromDatabase("JobRemarksDefault" & Right("00" & Trim(Str$(i)), 2))
        Next

108     RestoreOTANotificationSettings

        '<EhFooter>
        Exit Sub

RestoreJobInfoSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.RestoreJobInfoSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub RestoreLogSettings()
        '<EhHeader>
        On Error GoTo RestoreLogSettings_Err
        '</EhHeader>
        Dim sRelevantSetting As String
        Dim i As Integer

100     For i = 1 To 16
102         sRelevantSetting = "SendLogFieldEnabled" & Right("00" & Trim(Str$(i)), 2)
104         gtDBLogFieldsSettings.bSendLog(i) = GetSettingFromDatabase(sRelevantSetting)
  
106         sRelevantSetting = "JobLogFieldEnabled" & Right("00" & Trim(Str$(i)), 2)
108         gtDBLogFieldsSettings.bJobLog(i) = GetSettingFromDatabase(sRelevantSetting)
        Next

        '<EhFooter>
        Exit Sub

RestoreLogSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.RestoreLogSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub RestoreOTANotificationSettings()
        '<EhHeader>
        On Error GoTo RestoreOTANotificationSettings_Err
        '</EhHeader>

100     If GetSettingFromDatabase("OTANotificationSettingsUseService") = True Then
102         frmSMSMain.chkUseOTADeliveryNotifications.Value = 1
        Else
104         frmSMSMain.chkUseOTADeliveryNotifications.Value = 0
        End If

106     frmSMSMain.txtRecipientDeliveryNotification.Text = GetSettingFromDatabase("OTANotificationSettingsMobilenumber")

108     If GetSettingFromDatabase("OTANotificationSettingsEventBuffered") = True Then
110         frmSMSMain.chkDeliveryNotificationBuffered.Value = 1
        Else
112         frmSMSMain.chkDeliveryNotificationBuffered.Value = 0
        End If

114     If GetSettingFromDatabase("OTANotificationSettingsEventDelivered") = True Then
116         frmSMSMain.chkDeliveryNotificationDelivered.Value = 1
        Else
118         frmSMSMain.chkDeliveryNotificationDelivered.Value = 0
        End If

120     If GetSettingFromDatabase("OTANotificationSettingsEventNotDelivered") = True Then
122         frmSMSMain.chkDeliveryNotificationNotDelivered.Value = 1
        Else
124         frmSMSMain.chkDeliveryNotificationNotDelivered.Value = 0
        End If

        '<EhFooter>
        Exit Sub

RestoreOTANotificationSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.RestoreOTANotificationSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function SaveCurrentJobInfo(nLanguage As Integer) As Long
        '<EhHeader>
        On Error GoTo SaveCurrentJobInfo_Err
        '</EhHeader>
        Dim rsMain As Recordset
        Dim sField As String
        Dim sEntry As String
        Dim i As Integer

        On Error GoTo ErrorTrap

100     Set rsMain = gdbMain.OpenRecordset("select * from Jobs order by lID")

102     rsMain.AddNew

104     For i = 1 To gcnNumberOfJobRemarkFields
106         rsMain("JobRemarksField" & Right("00" & Trim(Str$(i)), 2)) = Left(frmSMSMain.cboJobRemarks(i).Text, 80)
        Next

108     rsMain("JobRemarksMemo") = frmSMSMain.txtJobRemarksMemo.Text
110     rsMain("FSendingType") = JobInfoSendingType(nLanguage)
112     rsMain("FSMSType") = JobInfoSMSType(nLanguage)
114     rsMain("FJobDate") = JobInfoCurrentTime()

116     If JobInfoStartingDate() <> "" Then
118         rsMain("FStartingDate") = JobInfoStartingDate()
        End If
  
120     If JobInfoEndingDate() <> "" Then
122         rsMain("FEndingDate") = JobInfoEndingDate()
        End If

124     rsMain("FWaitingTime") = JobInfoWaitingTimeBetweenMessages(nLanguage)
126     rsMain("FNumberOfMessages") = JobInfoNumberOfMessages()
128     rsMain("FNumberOfRecipients") = JobInfoNumberOfRecipients()

130     rsMain.UpDate
132     rsMain.MoveLast
134     SaveCurrentJobInfo = rsMain("lID")

136     rsMain.Close

        Exit Function
ErrorTrap:
        Exit Function

        '<EhFooter>
        Exit Function

SaveCurrentJobInfo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.SaveCurrentJobInfo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Sub SetListIndexWithText(cboInput As ComboBox, _
                         sText As String)
        '<EhHeader>
        On Error GoTo SetListIndexWithText_Err
        '</EhHeader>
        Dim nListIndex As Integer
        Dim bFound As Boolean

100     If sText <> "" Then
102         cboInput.Text = sText
        Else
104         nListIndex = 0

106         Do While Not bFound And (nListIndex + 1) <= cboInput.ListCount

108             If cboInput.List(nListIndex) = sText Then
110                 cboInput.ListIndex = nListIndex
112                 bFound = True
                End If

114             nListIndex = nListIndex + 1
            Loop

        End If

        '<EhFooter>
        Exit Sub

SetListIndexWithText_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.SetListIndexWithText " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function SQLOrderPartFromPhonebookOrder() As String
        '<EhHeader>
        On Error GoTo SQLOrderPartFromPhonebookOrder_Err
        '</EhHeader>

100     Select Case gnPhoneBookSortOrder

            Case 1
102             SQLOrderPartFromPhonebookOrder = " order by sName asc"
  
104         Case 2
106             SQLOrderPartFromPhonebookOrder = " order by sName desc"
   
108         Case 3
110             SQLOrderPartFromPhonebookOrder = " order by sNumber asc"
  
112         Case 4
114             SQLOrderPartFromPhonebookOrder = " order by sNumber desc"
  
116         Case Else
118             gnPhoneBookSortOrder = 1
120             SQLOrderPartFromPhonebookOrder = SQLOrderPartFromPhonebookOrder()
  
        End Select
  
        '<EhFooter>
        Exit Function

SQLOrderPartFromPhonebookOrder_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.SQLOrderPartFromPhonebookOrder " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function SQLValidCharsEncode(sInput As String) As String
        '<EhHeader>
        On Error GoTo SQLValidCharsEncode_Err
        '</EhHeader>
        Const cValidChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789äöüÄÖÜ/()"

        Dim i As Integer
        Dim sResult As String
        Dim sTemp As String

100     If Len(sInput) > 0 Then

102         For i = 1 To Len(sInput)
104             sTemp = Mid(sInput, i, 1)

106             If InStr(1, cValidChars, sTemp, vbBinaryCompare) = 0 Then
108                 sTemp = " "
                End If

110             sResult = sResult & sTemp
            Next

112         SQLValidCharsEncode = sResult
        End If

        '<EhFooter>
        Exit Function

SQLValidCharsEncode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.SQLValidCharsEncode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function TrimValidityPeriodToCorrectValues(tValidityPeriod As ValidityPeriod) As ValidityPeriod
        '<EhHeader>
        On Error GoTo TrimValidityPeriodToCorrectValues_Err
        '</EhHeader>
        Dim tValidityPeriodTemp As ValidityPeriod

100     tValidityPeriodTemp = tValidityPeriod
102     tValidityPeriodTemp.nLifeTimeUnit = tValidityPeriod.nLifeTimeUnit
104     tValidityPeriodTemp.bStandardPeriod = tValidityPeriod.bStandardPeriod
106     tValidityPeriodTemp.bUserDefinedSettingsOnlyForSpecificSMSTypes = tValidityPeriod.bUserDefinedSettingsOnlyForSpecificSMSTypes

108     Select Case tValidityPeriod.nLifeTimeUnit

            Case 0  'Minutes

110             If tValidityPeriod.lLifeTime < 3 Then
112                 tValidityPeriodTemp.lLifeTime = 3
                End If

114             If tValidityPeriod.lLifeTime > 1440 Then
116                 tValidityPeriodTemp.nLifeTimeUnit = 1
118                 tValidityPeriodTemp.lLifeTime = 24
                End If
    
120         Case 1  'Hours

122             If tValidityPeriod.lLifeTime < 1 Then
124                 tValidityPeriodTemp.nLifeTimeUnit = 0
126                 tValidityPeriodTemp.lLifeTime = 3
                End If

128             If tValidityPeriod.lLifeTime > 24 Then
130                 tValidityPeriodTemp.lLifeTime = 24
                End If

        End Select

132     TrimValidityPeriodToCorrectValues = tValidityPeriodTemp
        '<EhFooter>
        Exit Function

TrimValidityPeriodToCorrectValues_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.TrimValidityPeriodToCorrectValues " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub UpdateComboBoxRemarksfrmSMSMain(nIndex As Integer)
        '<EhHeader>
        On Error GoTo UpdateComboBoxRemarksfrmSMSMain_Err
        '</EhHeader>
        Dim rsMain As Recordset
        Dim sField As String
        Dim sEntry As String
        Dim sSelectedField As String

        On Error GoTo ErrorTrap

100     sField = "JobRemarksField" & Right("0" & Trim(Str$(nIndex)), 2)

102     Set rsMain = gdbMain.OpenRecordset("select " & sField & " from Jobs group by " & sField & " order by " & sField)

104     sSelectedField = frmSMSMain.cboJobRemarks(nIndex).Text

106     If rsMain.BOF And rsMain.EOF Then
            'Do nothing
        Else
108         frmSMSMain.cboJobRemarks(nIndex).Clear

110         Do While Not rsMain.EOF
112             sEntry = rsMain(sField) & ""
114             frmSMSMain.cboJobRemarks(nIndex).AddItem sEntry
116             rsMain.MoveNext
            Loop

        End If

118     frmSMSMain.cboJobRemarks(nIndex).Text = sSelectedField

120     rsMain.Close

        Exit Sub
ErrorTrap:
        Exit Sub
        '<EhFooter>
        Exit Sub

UpdateComboBoxRemarksfrmSMSMain_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.UpdateComboBoxRemarksfrmSMSMain " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub UpdateComboBoxfrmFilterJoblog(nIndex As Integer, _
                                  cboInput As ComboBox, _
                                  sFieldInput As String)
        '<EhHeader>
        On Error GoTo UpdateComboBoxfrmFilterJoblog_Err
        '</EhHeader>
        Dim rsMain As Recordset
        Dim sField As String
        Dim sEntry As String
        Dim sSelectedField As String

        On Error GoTo ErrorTrap

100     If nIndex = -1 Then
102         sField = sFieldInput
        Else
104         sField = sFieldInput & Right("0" & Trim(Str$(nIndex)), 2)
        End If

106     Set rsMain = gdbMain.OpenRecordset("select " & sField & " from Jobs group by " & sField & " order by " & sField)

108     If rsMain.BOF And rsMain.EOF Then
            'Do nothing
        Else
110         cboInput.Clear
112         cboInput.AddItem LoadLanguageSpecificString(gnLanguage, 321)
114         cboInput.ItemData(cboInput.NewIndex) = gcNoFilter

116         Do While Not rsMain.EOF
118             sEntry = rsMain(sField) & ""
120             cboInput.AddItem sEntry
122             rsMain.MoveNext
            Loop

        End If

124     rsMain.Close

        Exit Sub
ErrorTrap:
        Exit Sub
        '<EhFooter>
        Exit Sub

UpdateComboBoxfrmFilterJoblog_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.UpdateComboBoxfrmFilterJoblog " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub UpdateJobInfoScreen(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo UpdateJobInfoScreen_Err
        '</EhHeader>
        Dim sWaitingPeriodInfo As String
        Dim lCountRecipients As Long
        Dim lRecipientMultiplier As Long
        Dim lCountMessages As Long
        Dim dStart As Date
        Dim dEnd As Date

        'CurrentTime
100     frmSMSMain.lblAutoGeneratedRemarks(1).caption = LoadLanguageSpecificString(nLanguage, 568) & ": " & JobInfoCurrentTime

        'Number of Recipients
102     lCountRecipients = frmSMSMain.grdRecipients.Rows - 1
104     frmSMSMain.lblAutoGeneratedRemarks(2).caption = LoadLanguageSpecificString(nLanguage, 569) & ": " & Trim(Str$(JobInfoNumberOfRecipients))

        'SMS Type
106     frmSMSMain.lblAutoGeneratedRemarks(4).caption = LoadLanguageSpecificString(nLanguage, 570) & ": " & JobInfoSMSType(nLanguage)

        'Sending Type
108     frmSMSMain.lblAutoGeneratedRemarks(5).caption = ""

        'Starting Date or Delivery Date
110     frmSMSMain.lblAutoGeneratedRemarks(6).caption = ""

        'Ending Date

        'Waiting Time between Messages

112     lRecipientMultiplier = 1

114     If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = True Then
116         If gtDeferredDeliveryTimeSettings.bSingleSMS = True Then
                'Sending Type
118             frmSMSMain.lblAutoGeneratedRemarks(5).caption = LoadLanguageSpecificString(nLanguage, 571) & ": " & JobInfoSendingType(nLanguage)
                'Starting Date or Delivery Date
120             frmSMSMain.lblAutoGeneratedRemarks(6).caption = LoadLanguageSpecificString(nLanguage, 572) & ": " & gtDeferredDeliveryTimeSettings.sDeliveryDate
122             frmSMSMain.lblAutoGeneratedRemarks(6).Visible = True
124             frmSMSMain.lblAutoGeneratedRemarks(7).Visible = False
126             frmSMSMain.lblAutoGeneratedRemarks(8).Visible = False
            End If

128         If gtDeferredDeliveryTimeSettings.bPeriodicSMS = True Then
                'Sending Type
130             lRecipientMultiplier = Val(gtDeferredDeliveryTimeSettings.sNumberOfMessages)
132             frmSMSMain.lblAutoGeneratedRemarks(5).caption = LoadLanguageSpecificString(nLanguage, 571) & ": " & JobInfoSendingType(nLanguage)
134             frmSMSMain.lblAutoGeneratedRemarks(6).caption = LoadLanguageSpecificString(nLanguage, 573) & ": " & JobInfoStartingDate()
    
136             dStart = gtDeferredDeliveryTimeSettings.sStartingDate
          
138             dEnd = JobInfoEndingDate()
    
140             frmSMSMain.lblAutoGeneratedRemarks(7).caption = LoadLanguageSpecificString(nLanguage, 574) & ": " & dEnd
142             frmSMSMain.lblAutoGeneratedRemarks(8).caption = LoadLanguageSpecificString(nLanguage, 575) & ": " & gtDeferredDeliveryTimeSettings.sWaitingTime & " " & JobInfoWaitingTimeInfo(nLanguage)
144             frmSMSMain.lblAutoGeneratedRemarks(6).Visible = True
146             frmSMSMain.lblAutoGeneratedRemarks(7).Visible = True
148             frmSMSMain.lblAutoGeneratedRemarks(8).Visible = True
            End If

        Else
            'Sending Type
150         frmSMSMain.lblAutoGeneratedRemarks(5).caption = LoadLanguageSpecificString(nLanguage, 571) & ": " & JobInfoSendingType(nLanguage)
152         frmSMSMain.lblAutoGeneratedRemarks(6).Visible = False
154         frmSMSMain.lblAutoGeneratedRemarks(7).Visible = False
156         frmSMSMain.lblAutoGeneratedRemarks(8).Visible = False
        End If

        'Number of Messages
158     frmSMSMain.lblAutoGeneratedRemarks(3).caption = LoadLanguageSpecificString(nLanguage, 576) & ": " & Trim(Str$(JobInfoNumberOfMessages))

        Exit Sub
        '<EhFooter>
        Exit Sub

UpdateJobInfoScreen_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.UpdateJobInfoScreen " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Function URLEncode(sInput As String) As String
        '<EhHeader>
        On Error GoTo URLEncode_Err
        '</EhHeader>
        Const cValidChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

        Dim i As Integer
        Dim sResult As String
        Dim sTemp As String

100     If Len(sInput) > 0 Then

102         For i = 1 To Len(sInput)
104             sTemp = Mid(sInput, i, 1)

106             If InStr(1, cValidChars, sTemp, vbBinaryCompare) = 0 Then
108                 sTemp = "%" & Right("00" & Hex(Asc(sTemp)), 2)
                End If

110             sResult = sResult & sTemp
            Next

112         URLEncode = sResult
        End If

        '<EhFooter>
        Exit Function

URLEncode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.URLEncode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub LoadSavedRandomLogosOnScreen()
        '<EhHeader>
        On Error GoTo LoadSavedRandomLogosOnScreen_Err
        '</EhHeader>
        Dim i As Integer
        Dim sFilename As String

        On Error Resume Next

100     For i = 0 To 20
102         sFilename = App.Path & "\" & Right("0000" & Trim(Str$(i)), 4) & ".bmp"
104         frmSMSMain.imgRandomLogo(i).Picture = LoadPicture(sFilename)
        Next

106     frmSMSMain.imgRandomLogo_Click 0
108     gnSelectedRandomLogo = 0
        '<EhFooter>
        Exit Sub

LoadSavedRandomLogosOnScreen_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.LoadSavedRandomLogosOnScreen " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Function NumberOfRemainingCharsInLastConcatenationPart(nLenght As Integer, bBlinkingSMS As Boolean, bBlinkingTagSet As Boolean) As Integer
        '<EhHeader>
        On Error GoTo NumberOfRemainingCharsInLastConcatenationPart_Err
        '</EhHeader>
        Dim nTemp As Integer
        Dim nLengthWork As Integer
        Dim nLenghtFirstSMS As Integer
        Dim nLenghtConcatenationPart As Integer
        Dim nLenghtSMS As Integer
        Dim nUsedCharsWithinThisPart As Integer

        'Explanation of this function
        '1 SMS may consist out of up to 160 chars
        'If several SMS are concatenated together, the lenght of each concatenationpart has to be reduced to 153 chars, because
        'the concatenation infoheader also takes space
        'Similar explanation for blinking SMS

100     If bBlinkingSMS = True Then
102         If bBlinkingTagSet = True Then
104             nLengthWork = nLenght
            Else
106             nLengthWork = nLenght + 1 'Because in this case BlinkingTag will be automatically set and needs one char
            End If

        Else
108         nLengthWork = nLenght
        End If

110     If bBlinkingSMS = False Then
112         nLenghtFirstSMS = 160
114         nLenghtConcatenationPart = 153
        Else
116         nLenghtFirstSMS = 70
118         nLenghtConcatenationPart = 66
        End If

120     If nLengthWork > nLenghtFirstSMS Then
122         nUsedCharsWithinThisPart = nLengthWork Mod nLenghtConcatenationPart
124         nLenghtSMS = nLenghtConcatenationPart
        Else
126         nUsedCharsWithinThisPart = nLengthWork Mod nLenghtFirstSMS
128         nLenghtSMS = nLenghtFirstSMS
        End If

130     Select Case True

            Case nUsedCharsWithinThisPart <> 0
132             NumberOfRemainingCharsInLastConcatenationPart = nLenghtSMS - nUsedCharsWithinThisPart
  
134         Case nLengthWork = 0
136             NumberOfRemainingCharsInLastConcatenationPart = nLenghtSMS
  
138         Case Else
140             NumberOfRemainingCharsInLastConcatenationPart = 0
  
        End Select

        '<EhFooter>
        Exit Function

NumberOfRemainingCharsInLastConcatenationPart_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.NumberOfRemainingCharsInLastConcatenationPart " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Function NumberOfConcatenationParts(nLenght As Integer, bBlinkingSMS As Boolean) As Integer
        '<EhHeader>
        On Error GoTo NumberOfConcatenationParts_Err
        '</EhHeader>
        Dim nTemp As Integer
        Dim nLenghtOneSMS As Integer
        Dim nLenghtConcatenationPart As Integer

100     If bBlinkingSMS = False Then
102         nLenghtOneSMS = 160
104         nLenghtConcatenationPart = 153
        Else
106         nLenghtOneSMS = 70
108         nLenghtConcatenationPart = 66
        End If
  
110     If nLenght > nLenghtOneSMS Then
112         If nLenght Mod nLenghtConcatenationPart = 0 Then
114             NumberOfConcatenationParts = nLenght / nLenghtConcatenationPart
            Else
116             NumberOfConcatenationParts = Int((nLenght / nLenghtConcatenationPart) + 1)
            End If

        Else
118         NumberOfConcatenationParts = 1
        End If

        '<EhFooter>
        Exit Function

NumberOfConcatenationParts_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.NumberOfConcatenationParts " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub AdjustFontControls(frmInput As Form)
        '<EhHeader>
        On Error GoTo AdjustFontControls_Err
        '</EhHeader>
        Dim tFontSettingsTemp As FontSettings
        Dim Control As Object
        Dim i As Integer
        Dim bExpectedFontFound As Boolean
        Dim sFont As String

100     If gtFontSettingsCurrent.bSpecificFontUsed = False Then
102         tFontSettingsTemp = gtFontSettingsDefault
        Else
104         tFontSettingsTemp = gtFontSettingsCurrent
        End If

106     frmInput.Font.Name = tFontSettingsTemp.sFontName
108     frmInput.Font.Size = tFontSettingsTemp.rFontSize
110     frmInput.Font.Bold = tFontSettingsTemp.bFontBold
112     frmInput.Font.Italic = tFontSettingsTemp.bFontItalic

114     For Each Control In frmInput.Controls
  
116         Select Case True
  
                Case TypeOf Control Is CommandButton
118                 Control.Font.Name = tFontSettingsTemp.sFontName
120                 Control.Font.Size = tFontSettingsTemp.rFontSize
122                 Control.Font.Bold = tFontSettingsTemp.bFontBold
124                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
      
126             Case TypeOf Control Is Label
128                 Control.Font.Name = tFontSettingsTemp.sFontName
130                 Control.Font.Size = tFontSettingsTemp.rFontSize
132                 Control.Font.Bold = tFontSettingsTemp.bFontBold
134                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
    
136             Case TypeOf Control Is Frame
138                 Control.Font.Name = tFontSettingsTemp.sFontName
140                 Control.Font.Size = tFontSettingsTemp.rFontSize
142                 Control.Font.Bold = tFontSettingsTemp.bFontBold
144                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
    
146             Case TypeOf Control Is TextBox
148                 Control.Font.Name = tFontSettingsTemp.sFontName
150                 Control.Font.Size = tFontSettingsTemp.rFontSize
152                 Control.Font.Bold = tFontSettingsTemp.bFontBold
154                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
    
156             Case TypeOf Control Is CheckBox
158                 Control.Font.Name = tFontSettingsTemp.sFontName
160                 Control.Font.Size = tFontSettingsTemp.rFontSize
162                 Control.Font.Bold = tFontSettingsTemp.bFontBold
164                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
    
166             Case TypeOf Control Is OptionButton
168                 Control.Font.Name = tFontSettingsTemp.sFontName
170                 Control.Font.Size = tFontSettingsTemp.rFontSize
172                 Control.Font.Bold = tFontSettingsTemp.bFontBold
174                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
    
176             Case TypeOf Control Is ComboBox
178                 Control.Font.Name = tFontSettingsTemp.sFontName
180                 Control.Font.Size = tFontSettingsTemp.rFontSize
182                 Control.Font.Bold = tFontSettingsTemp.bFontBold
184                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
    
186             Case TypeOf Control Is ListBox
188                 Control.Font.Name = tFontSettingsTemp.sFontName
190                 Control.Font.Size = tFontSettingsTemp.rFontSize
192                 Control.Font.Bold = tFontSettingsTemp.bFontBold
194                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
    
196             Case TypeOf Control Is MSFlexGrid
198                 Control.Font = tFontSettingsTemp.sFontName
200                 Control.Font.Size = tFontSettingsTemp.rFontSize
202                 Control.Font.Bold = tFontSettingsTemp.bFontBold
204                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
    
206             Case TypeOf Control Is SSTab
208                 Control.Font = tFontSettingsTemp.sFontName
210                 Control.Font.Size = tFontSettingsTemp.rFontSize
212                 Control.Font.Bold = tFontSettingsTemp.bFontBold
214                 Control.Font.Italic = tFontSettingsTemp.bFontItalic
    
216             Case Else
                    'Do nothing
    
            End Select
  
            'Special cases, these settings do override everything else
218         If Control Is frmSMSMain.cmdSend Then
220             Control.Font.Bold = True
            End If
  
222         If Control Is frmSMSMain.txtUnicode Then
224             Control.Font.Size = Int(tFontSettingsTemp.rFontSize * 1.5)
            End If
  
226     Next Control

        '<EhFooter>
        Exit Sub

AdjustFontControls_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.AdjustFontControls " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Function ControlIsLoaded(ctlInput As Control) As Boolean
        '<EhHeader>
        On Error GoTo ControlIsLoaded_Err
        '</EhHeader>
        Dim sTemp As String
        On Error GoTo ControlDoesntExist

100     sTemp = ctlInput.Tag

102     ControlIsLoaded = True

        Exit Function

ControlDoesntExist:
104     ControlIsLoaded = False
        '<EhFooter>
        Exit Function

ControlIsLoaded_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ControlIsLoaded " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function GetNameFromPhoneBook(sPhoneNumber As String) As String
        '<EhHeader>
        On Error GoTo GetNameFromPhoneBook_Err
        '</EhHeader>
        On Error Resume Next

100     frmSMSMain.dtaPhonebook.Recordset.FindFirst LoadLanguageSpecificString(gnLanguage, 197) & " = '" & sPhoneNumber & "'"

102     If frmSMSMain.dtaPhonebook.Recordset.NoMatch Then
104         GetNameFromPhoneBook = ""
        Else
106         GetNameFromPhoneBook = frmSMSMain.dtaPhonebook.Recordset(LoadLanguageSpecificString(gnLanguage, 196)) & ""
        End If

108     frmSMSMain.dtaPhonebook.Recordset.MoveFirst
        '<EhFooter>
        Exit Function

GetNameFromPhoneBook_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.GetNameFromPhoneBook " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function GetSettingFromDatabase(sSettingType As String) As Variant
        '<EhHeader>
        On Error GoTo GetSettingFromDatabase_Err
        '</EhHeader>
        Dim sVersionSpecific1 As String
        Dim sVersionSpecific2 As String
        Dim sVersionSpecific3 As String

100     VersionSpecificAction 17, , , sVersionSpecific1
102     VersionSpecificAction 18, , , sVersionSpecific2
104     VersionSpecificAction 19, , , sVersionSpecific3

106     frmSMSMain.dtaSettings.Recordset.FindFirst "SettingType = '" & UCase(sSettingType) & "'"

108     If frmSMSMain.dtaSettings.Recordset.NoMatch Then   'Settingrecord not found in database, applying default value

110         Select Case UCase(sSettingType)

                Case "USERKEY", "PASSWORD", "ORIGINATOR", "FONTNAME", "OTANOTIFICATIONSETTINGSMOBILENUMBER"
112                 GetSettingFromDatabase = ""
  
114             Case "JOBREMARKSFIELD01", "JOBREMARKSFIELD02", "JOBREMARKSFIELD03", "JOBREMARKSFIELD04", "JOBREMARKSFIELD05", "JOBREMARKSFIELD06", "JOBREMARKSMEMO"
116                 GetSettingFromDatabase = "DEFAULT"
  
118             Case "PHONEBOOKVARIABLEFIELD01", "PHONEBOOKVARIABLEFIELD02", "PHONEBOOKVARIABLEFIELD03"
120                 GetSettingFromDatabase = "DEFAULT"
  
122             Case "JOBREMARKSDEFAULT01", "JOBREMARKSDEFAULT02", "JOBREMARKSDEFAULT03", "JOBREMARKSDEFAULT04", "JOBREMARKSDEFAULT05", "JOBREMARKSDEFAULT06"
124                 GetSettingFromDatabase = ""
  
126             Case "FONTBOLD", "FONTITALIC", "FONTSTRIKETHRU", "FONTUNDERLINE", "SPECIFICFONTUSED", "TEXTSMSENABLED", "OPERATORLOGOENABLED", "GROUPLOGOENABLED", "RINGTONEENABLED", "PICTUREMESSAGEENABLED", "VCARDENABLED", "UNICODEENABLED", "WAPPUSHSMSENABLED", "BINARYDATAENABLED", sVersionSpecific1, sVersionSpecific2, sVersionSpecific3, "SAVEMESSAGESINSENDJOURNAL", "OTANOTIFICATIONSETTINGSUSESERVICE", "OTANOTIFICATIONSETTINGSEVENTBUFFERED", "OTANOTIFICATIONSETTINGSEVENTDELIVERED", "OTANOTIFICATIONSETTINGSEVENTNOTDELIVERED", "OTANOTIFICATIONSETTINGSREPLACE"
128                 GetSettingFromDatabase = False
  
130             Case "FONTSIZE", "LANGUAGE"
132                 GetSettingFromDatabase = 0
  
134             Case "COLWIDTHJOBS00", "COLWIDTHJOBS01", "COLWIDTHJOBS02", "COLWIDTHJOBS03", "COLWIDTHJOBS04", "COLWIDTHJOBS05", "COLWIDTHJOBS06", "COLWIDTHJOBS07", "COLWIDTHJOBS08", "COLWIDTHJOBS09", "COLWIDTHJOBS10", "COLWIDTHJOBS11", "COLWIDTHJOBS12", "COLWIDTHJOBS13", "COLWIDTHJOBS14", "COLWIDTHJOBS15"
136                 GetSettingFromDatabase = -1
  
138             Case "COLWIDTHSENDLOG00", "COLWIDTHSENDLOG01", "COLWIDTHSENDLOG02", "COLWIDTHSENDLOG03", "COLWIDTHSENDLOG04", "COLWIDTHSENDLOG05", "COLWIDTHSENDLOG06", "COLWIDTHSENDLOG07", "COLWIDTHSENDLOG08", "COLWIDTHSENDLOG09", "COLWIDTHSENDLOG10", "COLWIDTHSENDLOG11", "COLWIDTHSENDLOG12", "COLWIDTHSENDLOG13", "COLWIDTHSENDLOG14", "COLWIDTHSENDLOG15"
140                 GetSettingFromDatabase = -1
  
142             Case "SENDLOGFIELDENABLED01", "SENDLOGFIELDENABLED02", "SENDLOGFIELDENABLED03", "SENDLOGFIELDENABLED04", "SENDLOGFIELDENABLED05", "SENDLOGFIELDENABLED06", "SENDLOGFIELDENABLED07", "SENDLOGFIELDENABLED08", "SENDLOGFIELDENABLED09", "SENDLOGFIELDENABLED10", "SENDLOGFIELDENABLED11", "SENDLOGFIELDENABLED12", "SENDLOGFIELDENABLED13", "SENDLOGFIELDENABLED14", "SENDLOGFIELDENABLED15", "SENDLOGFIELDENABLED16"
144                 GetSettingFromDatabase = True
  
146             Case "JOBLOGFIELDENABLED01", "JOBLOGFIELDENABLED02", "JOBLOGFIELDENABLED03", "JOBLOGFIELDENABLED04", "JOBLOGFIELDENABLED05", "JOBLOGFIELDENABLED06", "JOBLOGFIELDENABLED07", "JOBLOGFIELDENABLED08", "JOBLOGFIELDENABLED09", "JOBLOGFIELDENABLED10", "JOBLOGFIELDENABLED11", "JOBLOGFIELDENABLED12", "JOBLOGFIELDENABLED13", "JOBLOGFIELDENABLED14", "JOBLOGFIELDENABLED15", "JOBLOGFIELDENABLED16"
148                 GetSettingFromDatabase = True
    
150             Case "SINGLESMSUSEUSERDEFINEDLIFETIME", "PERIODICSMSUSERDEFINEDLIFETIME"
152                 GetSettingFromDatabase = False
  
154             Case "SINGLESMSUSEUDSETTINGSONLYFORSPECIFICSMSTYPES", "PERIODICSMSUSEUDSETTINGSONLYFORSPECIFICSMSTYPES"
156                 GetSettingFromDatabase = True
    
158             Case "SINGLESMSVALIDITYPERIODMODE"
160                 GetSettingFromDatabase = ValidityPeriodMode.UseSpecificSettingsAsLifeTime

162             Case "PERIODICSMSVALIDITYPERIODMODE"
164                 GetSettingFromDatabase = ValidityPeriodMode.UseWaitingTimeAsLifeTime
  
166             Case "SINGLESMSSPECIFICSETTINGLIFETIME", "PERIODICSMSSPECIFICSETTINGLIFETIME"
168                 GetSettingFromDatabase = 3
  
170             Case "SINGLESMSSPECIFICSETTINGLIFETIMEUNIT", "PERIODICSMSSPECIFICSETTINGLIFETIMEUNIT"
172                 GetSettingFromDatabase = 0

174             Case Else
176                 GetSettingFromDatabase = 0
  
            End Select

        Else  'Everything ok, Settingrecord found in database
  
178         Select Case UCase(sSettingType)

                Case "USERKEY", "PASSWORD", "ORIGINATOR", "FONTNAME", "OTANOTIFICATIONSETTINGSMOBILENUMBER"
180                 GetSettingFromDatabase = frmSMSMain.dtaSettings.Recordset("SettingValue") & ""
  
182             Case "JOBREMARKSFIELD01", "JOBREMARKSFIELD02", "JOBREMARKSFIELD03", "JOBREMARKSFIELD04", "JOBREMARKSFIELD05", "JOBREMARKSFIELD06", "JOBREMARKSMEMO"
184                 GetSettingFromDatabase = frmSMSMain.dtaSettings.Recordset("SettingValue") & ""
  
186             Case "PHONEBOOKVARIABLEFIELD01", "PHONEBOOKVARIABLEFIELD02", "PHONEBOOKVARIABLEFIELD03"
188                 GetSettingFromDatabase = frmSMSMain.dtaSettings.Recordset("SettingValue") & ""
  
190             Case "JOBREMARKSDEFAULT01", "JOBREMARKSDEFAULT02", "JOBREMARKSDEFAULT03", "JOBREMARKSDEFAULT04", "JOBREMARKSDEFAULT05", "JOBREMARKSDEFAULT06"
192                 GetSettingFromDatabase = frmSMSMain.dtaSettings.Recordset("SettingValue") & ""
  
194             Case "FONTBOLD", "FONTITALIC", "FONTSTRIKETHRU", "FONTUNDERLINE", "SPECIFICFONTUSED", "TEXTSMSENABLED", "OPERATORLOGOENABLED", "GROUPLOGOENABLED", "RINGTONEENABLED", "PICTUREMESSAGEENABLED", "VCARDENABLED", "UNICODEENABLED", "WAPPUSHSMSENABLED", "BINARYDATAENABLED", sVersionSpecific1, sVersionSpecific2, sVersionSpecific3, "SAVEMESSAGESINSENDJOURNAL", "OTANOTIFICATIONSETTINGSUSESERVICE", "OTANOTIFICATIONSETTINGSEVENTBUFFERED", "OTANOTIFICATIONSETTINGSEVENTDELIVERED", "OTANOTIFICATIONSETTINGSEVENTNOTDELIVERED", "OTANOTIFICATIONSETTINGSREPLACE"

196                 If frmSMSMain.dtaSettings.Recordset("SettingValue") & "" = "TRUE" Then
198                     GetSettingFromDatabase = True
                    Else
200                     GetSettingFromDatabase = False
                    End If
  
202             Case "FONTSIZE", "LANGUAGE"
204                 GetSettingFromDatabase = Val(frmSMSMain.dtaSettings.Recordset("SettingValue") & "")
  
206             Case "COLWIDTHJOBS00", "COLWIDTHJOBS01", "COLWIDTHJOBS02", "COLWIDTHJOBS03", "COLWIDTHJOBS04", "COLWIDTHJOBS05", "COLWIDTHJOBS06", "COLWIDTHJOBS07", "COLWIDTHJOBS08", "COLWIDTHJOBS09", "COLWIDTHJOBS10", "COLWIDTHJOBS11", "COLWIDTHJOBS12", "COLWIDTHJOBS13", "COLWIDTHJOBS14", "COLWIDTHJOBS15"
208                 GetSettingFromDatabase = Val(frmSMSMain.dtaSettings.Recordset("SettingValue") & "")
  
210             Case "COLWIDTHSENDLOG00", "COLWIDTHSENDLOG01", "COLWIDTHSENDLOG02", "COLWIDTHSENDLOG03", "COLWIDTHSENDLOG04", "COLWIDTHSENDLOG05", "COLWIDTHSENDLOG06", "COLWIDTHSENDLOG07", "COLWIDTHSENDLOG08", "COLWIDTHSENDLOG09", "COLWIDTHSENDLOG10", "COLWIDTHSENDLOG11", "COLWIDTHSENDLOG12", "COLWIDTHSENDLOG13", "COLWIDTHSENDLOG14", "COLWIDTHSENDLOG15"
212                 GetSettingFromDatabase = Val(frmSMSMain.dtaSettings.Recordset("SettingValue") & "")
  
214             Case "SENDLOGFIELDENABLED01", "SENDLOGFIELDENABLED02", "SENDLOGFIELDENABLED03", "SENDLOGFIELDENABLED04", "SENDLOGFIELDENABLED05", "SENDLOGFIELDENABLED06", "SENDLOGFIELDENABLED07", "SENDLOGFIELDENABLED08", "SENDLOGFIELDENABLED09", "SENDLOGFIELDENABLED10", "SENDLOGFIELDENABLED11", "SENDLOGFIELDENABLED12", "SENDLOGFIELDENABLED13", "SENDLOGFIELDENABLED14", "SENDLOGFIELDENABLED15", "SENDLOGFIELDENABLED16"

216                 If frmSMSMain.dtaSettings.Recordset("SettingValue") & "" = "TRUE" Then
218                     GetSettingFromDatabase = True
                    Else
220                     GetSettingFromDatabase = False
                    End If
  
222             Case "JOBLOGFIELDENABLED01", "JOBLOGFIELDENABLED02", "JOBLOGFIELDENABLED03", "JOBLOGFIELDENABLED04", "JOBLOGFIELDENABLED05", "JOBLOGFIELDENABLED06", "JOBLOGFIELDENABLED07", "JOBLOGFIELDENABLED08", "JOBLOGFIELDENABLED09", "JOBLOGFIELDENABLED10", "JOBLOGFIELDENABLED11", "JOBLOGFIELDENABLED12", "JOBLOGFIELDENABLED13", "JOBLOGFIELDENABLED14", "JOBLOGFIELDENABLED15", "JOBLOGFIELDENABLED16"

224                 If frmSMSMain.dtaSettings.Recordset("SettingValue") & "" = "TRUE" Then
226                     GetSettingFromDatabase = True
                    Else
228                     GetSettingFromDatabase = False
                    End If
  
230             Case "SINGLESMSUSEUSERDEFINEDLIFETIME", "SINGLESMSUSEUDSETTINGSONLYFORSPECIFICSMSTYPES", "PERIODICSMSUSEUSERDEFINEDLIFETIME", "PERIODICSMSUSEUDSETTINGSONLYFORSPECIFICSMSTYPES"

232                 If frmSMSMain.dtaSettings.Recordset("SettingValue") & "" = "TRUE" Then
234                     GetSettingFromDatabase = True
                    Else
236                     GetSettingFromDatabase = False
                    End If
  
238             Case "SINGLESMSVALIDITYPERIODMODE", "SINGLESMSSPECIFICSETTINGLIFETIME", "SINGLESMSSPECIFICSETTINGLIFETIMEUNIT", "PERIODICSMSVALIDITYPERIODMODE", "PERIODICSMSSPECIFICSETTINGLIFETIME", "PERIODICSMSSPECIFICSETTINGLIFETIMEUNIT"
240                 GetSettingFromDatabase = Val(frmSMSMain.dtaSettings.Recordset("SettingValue") & "")

242             Case Else
244                 GetSettingFromDatabase = 0
  
            End Select
  
        End If

        '<EhFooter>
        Exit Function

GetSettingFromDatabase_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.GetSettingFromDatabase " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub InitApp()
        '<EhHeader>
        On Error GoTo InitApp_Err
        '</EhHeader>
        Dim sSQL As String
        Dim nDefaultTimeZone As Integer

100     VersionSpecificAction 2
102     VersionSpecificAction 3

104     VersionSpecificAction 1, 1
106     VersionSpecificAction 42
108     VersionSpecificAction 59

        'Init Deferred Delivery Time
110     gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = False
112     VersionSpecificAction 36, nDefaultTimeZone
114     gtDeferredDeliveryTimeSettings.nTimeZone = nDefaultTimeZone
116     gtDeferredDeliveryTimeSettings.bSingleSMS = True
118     gtDeferredDeliveryTimeSettings.bPeriodicSMS = False
120     gtDeferredDeliveryTimeSettings.sDeliveryDate = Now()
122     gtDeferredDeliveryTimeSettings.sStartingDate = Now()
124     gtDeferredDeliveryTimeSettings.sNumberOfMessages = "10"
126     gtDeferredDeliveryTimeSettings.sWaitingTime = "1"
128     gtDeferredDeliveryTimeSettings.nWaitingPeriod = 3

        'Init Filtersettings Sendjournal
130     gtSendLogFilterSettings.bDeliveryStatusFilterUsed = False
132     gtSendLogFilterSettings.bReasonCodeFilterUsed = False
134     gtSendLogFilterSettings.bRecipientNameFilterUsed = False
136     gtSendLogFilterSettings.bRecipientFilterUsed = False
138     gtSendLogFilterSettings.bJobIDFilterUsed = False
140     gtSendLogFilterSettings.lJobId = 0

        'Simply take any desired control to save default value
142     gtFontSettingsDefault.sFontName = frmSMSMain.txtUserkey.Font
144     gtFontSettingsDefault.rFontSize = frmSMSMain.txtUserkey.FontSize
146     gtFontSettingsDefault.bFontBold = frmSMSMain.txtUserkey.FontBold
148     gtFontSettingsDefault.bFontItalic = frmSMSMain.txtUserkey.FontItalic

        'DatabaseSettings General Settings
150     frmSMSMain.dtaSettings.DatabaseName = gsPathAndDatabaseName
152     frmSMSMain.dtaSettings.RecordSource = "select * from Settings"
154     frmSMSMain.dtaSettings.Refresh

156     RestoreGeneralSettings

158     RestoreLogSettings

160     Call frmSMSMain.txtSMS_Change
162     Call frmSMSMain.txtUnicode_Change

        'Sendlog
164     Set gdbMain = OpenDatabase(gsPathAndDatabaseName)
166     Set grsSendlog = gdbMain.OpenRecordset("select * from SendJournal")


        'DatabaseSettings Phonebook
168     gnPhoneBookSortOrder = 1
170     frmSMSMain.dtaPhonebook.DatabaseName = gsPathAndDatabaseName
172     frmSMSMain.dtaPhonebook.RecordSource = PhonebookSQLStatement()
174     frmSMSMain.dtaPhonebook.Refresh

176     frmSMSMain.grdPhonebook.Refresh

        'DatabaseSettings Available Countrys
178     frmSMSMain.dtaCountrys.DatabaseName = gsPathAndDatabaseName
180     frmSMSMain.dtaCountrys.RecordSource = "select Country from Networks group by Country order by Country"
182     frmSMSMain.dtaCountrys.Refresh

184     frmSMSMain.dtaCountrys.Recordset.MoveFirst

186     Do While Not frmSMSMain.dtaCountrys.Recordset.EOF
188         frmSMSMain.lstCountrys.AddItem frmSMSMain.dtaCountrys.Recordset("Country")
190         frmSMSMain.dtaCountrys.Recordset.MoveNext
        Loop

        'DatabaseSettings Available Networks
192     frmSMSMain.dtaNetworks.DatabaseName = gsPathAndDatabaseName
194     frmSMSMain.dtaNetworks.RecordSource = "select * from Networks order by Country, MCC, MNC"
196     frmSMSMain.dtaNetworks.Refresh

        'DatabaseSettings Templates
198     frmSMSMain.dtaTemplates.DatabaseName = gsPathAndDatabaseName
200     frmSMSMain.dtaTemplates.RecordSource = "select * from Templates order by ID"
202     frmSMSMain.dtaTemplates.Refresh
204     frmSMSMain.grdTemplates.ColWidth(0) = 0
206     frmSMSMain.grdTemplates.ColWidth(1) = frmSMSMain.grdTemplates.Width / 1.1

        'Phonebook
208     frmSMSMain.grdPhonebook.ColWidth(0) = 0
210     frmSMSMain.grdPhonebook.ColWidth(1) = frmSMSMain.grdPhonebook.Width / 4
212     frmSMSMain.grdPhonebook.ColWidth(2) = frmSMSMain.grdPhonebook.Width / 4
214     frmSMSMain.grdPhonebook.ColWidth(3) = frmSMSMain.grdPhonebook.Width / 7
216     frmSMSMain.grdPhonebook.ColWidth(4) = frmSMSMain.grdPhonebook.Width / 7
218     frmSMSMain.grdPhonebook.ColWidth(5) = frmSMSMain.grdPhonebook.Width / 7

220     frmSMSMain.grdRecipients.ColWidth(0) = frmSMSMain.grdRecipients.Width / 2.3
222     frmSMSMain.grdRecipients.ColWidth(1) = frmSMSMain.grdRecipients.Width / 2.3
224     frmSMSMain.grdRecipients.ColWidth(2) = 0
226     frmSMSMain.grdRecipients.ColWidth(3) = 0
228     frmSMSMain.grdRecipients.ColWidth(4) = 0
230     frmSMSMain.grdRecipients.ColAlignment(0) = flexAlignLeftCenter
232     frmSMSMain.grdRecipients.ColAlignment(1) = flexAlignLeftCenter

234     frmSMSMain.dtaTemplates.DatabaseName = gsPathAndDatabaseName
236     frmSMSMain.dtaTemplates.RecordSource = "select * from Templates order by ID"
238     frmSMSMain.dtaTemplates.Refresh

240     Call frmSMSMain.grdPhonebook_Click
242     LoadSavedRandomLogosOnScreen

244     AdjustFontControls frmSMSMain

246     frmSMSMain.AdjustLanguageSettings gnLanguage
248     RestoreJobInfoSettings
        '<EhFooter>
        Exit Sub

InitApp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.InitApp " & _
               "at line " & Erl
        
        Resume Next
        '</EhFooter>
End Sub
Function LoadLanguageSpecificString(nLanguage As Integer, nString As Integer) As String
        '<EhHeader>
        On Error GoTo LoadLanguageSpecificString_Err
        '</EhHeader>
        On Error GoTo ErrorTrap
        Dim nLookup As Integer

100     nLookup = nLanguage * 1000 + nString
102     LoadLanguageSpecificString = LoadResString(nLookup)

        Exit Function
ErrorTrap:
        'No Ressource found
104     LoadLanguageSpecificString = ""
        Exit Function

        '<EhFooter>
        Exit Function

LoadLanguageSpecificString_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.LoadLanguageSpecificString " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ParseReceivedLogoControlDataOld(sInput As String, bSuccess As Boolean) As Long
        '<EhHeader>
        On Error GoTo ParseReceivedLogoControlDataOld_Err
        '</EhHeader>
        Dim nPositionContentLenghtBegin As Integer
        Dim nPositionContentLenghtEnd As Integer
        Dim sContentLenghtToken As String
        Dim nContentLenght As Integer
        Dim sContent As String
        Dim sTemp As String
        Dim sTempArray() As String
        Dim sRelevantItems() As String

100     sTemp = UCase(sInput)

102     If InStr(1, sTemp, "OBJECT NOT FOUND") >= 1 Then
104         bSuccess = False
            Exit Function
        End If

106     nPositionContentLenghtBegin = InStr(1, sTemp, UCase("Content-Length"))
108     nPositionContentLenghtEnd = InStr(nPositionContentLenghtBegin, sTemp, vbCrLf & vbCrLf)

110     sContentLenghtToken = Mid(sInput, nPositionContentLenghtBegin, (nPositionContentLenghtEnd - nPositionContentLenghtBegin))
112     sTempArray() = Split(sContentLenghtToken, ":")
114     nContentLenght = Val(sTempArray(1))

116     sContent = Mid(sInput, nPositionContentLenghtEnd + Len(vbCrLf & vbCrLf), nContentLenght)

118     If Left(sContent, 5) = "Logos" Then
120         sRelevantItems() = Split(sContent, ":")
122         ParseReceivedLogoControlDataOld = Val(sRelevantItems(1))
124         bSuccess = True
        Else
126         bSuccess = False
        End If

        Exit Function
ErrorTrap:
        Exit Function
        '<EhFooter>
        Exit Function

ParseReceivedLogoControlDataOld_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ParseReceivedLogoControlDataOld " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ParseReceivedLogoControlData(sInput As String, bSuccess As Boolean) As Long
        '<EhHeader>
        On Error GoTo ParseReceivedLogoControlData_Err
        '</EhHeader>
        Dim nPositionContentLenghtBegin As Integer
        Dim nPositionContentLenghtEnd As Integer
        Dim sContentLenghtToken As String
        Dim nContentLenght As Integer
        Dim sContent As String
        Dim sTemp As String
        Dim sTempArray() As String
        Dim sRelevantItems() As String

100     sTemp = UCase(sInput)

102     If InStr(1, sTemp, "OBJECT NOT FOUND") >= 1 Then
104         bSuccess = False
            Exit Function
        End If

106     If UCase(Left(sTemp, 5)) = UCase("Logos") Then
108         sRelevantItems() = Split(sTemp, ":")
110         ParseReceivedLogoControlData = Val(sRelevantItems(1))
112         bSuccess = True
        Else
114         bSuccess = False
        End If

        Exit Function
ErrorTrap:
        Exit Function
        '<EhFooter>
        Exit Function

ParseReceivedLogoControlData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ParseReceivedLogoControlData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ParseReceivedPictureData(sInput As String, bSuccess As Boolean) As String
        '<EhHeader>
        On Error GoTo ParseReceivedPictureData_Err
        '</EhHeader>
        Dim nPositionContentLenghtBegin As Integer
        Dim nPositionContentLenghtEnd As Integer
        Dim sContentLenghtToken As String
        Dim nContentLenght As Integer
        Dim sContent As String
        Dim sTemp As String
        Dim sTempArray() As String

100     sTemp = UCase(sInput)

102     If InStr(1, sTemp, "OBJECT NOT FOUND") >= 1 Then
104         bSuccess = False
            Exit Function
        End If

106     nPositionContentLenghtBegin = InStr(1, sTemp, UCase("Content-Length"))
108     nPositionContentLenghtEnd = InStr(nPositionContentLenghtBegin, sTemp, vbCrLf & vbCrLf)

110     sContentLenghtToken = Mid(sInput, nPositionContentLenghtBegin, (nPositionContentLenghtEnd - nPositionContentLenghtBegin))
112     sTempArray() = Split(sContentLenghtToken, ":")
114     nContentLenght = Val(sTempArray(1))

116     sContent = Mid(sInput, nPositionContentLenghtEnd + Len(vbCrLf & vbCrLf), nContentLenght)

118     If Left(sContent, 2) = "BM" Then
120         ParseReceivedPictureData = sContent
122         bSuccess = True
        Else
124         bSuccess = False
        End If

        Exit Function
ErrorTrap:
        Exit Function
        '<EhFooter>
        Exit Function

ParseReceivedPictureData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ParseReceivedPictureData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub ProcessRandomLogoDisplayUpdate(nIndex As Integer, _
                                   sPicture As String)
        '<EhHeader>
        On Error GoTo ProcessRandomLogoDisplayUpdate_Err
        '</EhHeader>
        Dim sFilename As String
        Dim nFile As Integer

100     sFilename = App.Path & "\" & Right("0000" & Trim(Str$(nIndex)), 4) & ".bmp"

        On Error GoTo ErrorTrap

102     nFile = FreeFile

104     Open sFilename For Binary As nFile
106     Put nFile, , sPicture

108     Close nFile

        Exit Sub

ErrorTrap:
        Exit Sub

        '<EhFooter>
        Exit Sub

ProcessRandomLogoDisplayUpdate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ProcessRandomLogoDisplayUpdate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub PutSettingIntoDataBase(sSettingType As String, _
                           vSettingValue As Variant)
        '<EhHeader>
        On Error GoTo PutSettingIntoDataBase_Err
        '</EhHeader>
        Dim sVersionSpecific1 As String
        Dim sVersionSpecific2 As String
        Dim sVersionSpecific3 As String

100     VersionSpecificAction 17, , , sVersionSpecific1
102     VersionSpecificAction 18, , , sVersionSpecific2
104     VersionSpecificAction 19, , , sVersionSpecific3

106     frmSMSMain.dtaSettings.Recordset.FindFirst "SettingType = '" & UCase(sSettingType) & "'"

108     If frmSMSMain.dtaSettings.Recordset.NoMatch Then
110         frmSMSMain.dtaSettings.Recordset.AddNew
        Else
112         frmSMSMain.dtaSettings.Recordset.Edit
        End If

114     frmSMSMain.dtaSettings.Recordset("SettingType") = sSettingType

116     Select Case UCase(sSettingType)

            Case "USERKEY", "PASSWORD", "ORIGINATOR", "FONTNAME", "OTANOTIFICATIONSETTINGSMOBILENUMBER"
118             frmSMSMain.dtaSettings.Recordset("SettingValue") = vSettingValue & ""
  
120         Case "JOBREMARKSFIELD01", "JOBREMARKSFIELD02", "JOBREMARKSFIELD03", "JOBREMARKSFIELD04", "JOBREMARKSFIELD05", "JOBREMARKSFIELD06", "JOBREMARKSMEMO"
122             frmSMSMain.dtaSettings.Recordset("SettingValue") = vSettingValue & ""
  
124         Case "PHONEBOOKVARIABLEFIELD01", "PHONEBOOKVARIABLEFIELD02", "PHONEBOOKVARIABLEFIELD03"
126             frmSMSMain.dtaSettings.Recordset("SettingValue") = vSettingValue & ""
  
128         Case "JOBREMARKSDEFAULT01", "JOBREMARKSDEFAULT02", "JOBREMARKSDEFAULT03", "JOBREMARKSDEFAULT04", "JOBREMARKSDEFAULT05", "JOBREMARKSDEFAULT06"
130             frmSMSMain.dtaSettings.Recordset("SettingValue") = vSettingValue & ""
   
132         Case "FONTBOLD", "FONTITALIC", "FONTSTRIKETHRU", "FONTUNDERLINE", "SPECIFICFONTUSED", "TEXTSMSENABLED", "OPERATORLOGOENABLED", "GROUPLOGOENABLED", "RINGTONEENABLED", "PICTUREMESSAGEENABLED", "VCARDENABLED", "WAPPUSHSMSENABLED", "BINARYDATAENABLED", "UNICODEENABLED", sVersionSpecific1, sVersionSpecific2, sVersionSpecific3, "SAVEMESSAGESINSENDJOURNAL", "OTANOTIFICATIONSETTINGSUSESERVICE", "OTANOTIFICATIONSETTINGSEVENTBUFFERED", "OTANOTIFICATIONSETTINGSEVENTDELIVERED", "OTANOTIFICATIONSETTINGSEVENTNOTDELIVERED", "OTANOTIFICATIONSETTINGSREPLACE"

134             If vSettingValue = True Then
136                 frmSMSMain.dtaSettings.Recordset("SettingValue") = "TRUE"
                Else
138                 frmSMSMain.dtaSettings.Recordset("SettingValue") = "FALSE"
                End If
  
140         Case "FONTSIZE", "LANGUAGE"
142             frmSMSMain.dtaSettings.Recordset("SettingValue") = Str$(vSettingValue)
  
144         Case "COLWIDTHJOBS00", "COLWIDTHJOBS01", "COLWIDTHJOBS02", "COLWIDTHJOBS03", "COLWIDTHJOBS04", "COLWIDTHJOBS05", "COLWIDTHJOBS06", "COLWIDTHJOBS07", "COLWIDTHJOBS08", "COLWIDTHJOBS09", "COLWIDTHJOBS10", "COLWIDTHJOBS11", "COLWIDTHJOBS12", "COLWIDTHJOBS13", "COLWIDTHJOBS14", "COLWIDTHJOBS15", "COLWIDTHJOBS15"
146             frmSMSMain.dtaSettings.Recordset("SettingValue") = Str$(vSettingValue)
  
148         Case "COLWIDTHSENDLOG00", "COLWIDTHSENDLOG01", "COLWIDTHSENDLOG02", "COLWIDTHSENDLOG03", "COLWIDTHSENDLOG04", "COLWIDTHSENDLOG05", "COLWIDTHSENDLOG06", "COLWIDTHSENDLOG07", "COLWIDTHSENDLOG08", "COLWIDTHSENDLOG09", "COLWIDTHSENDLOG10", "COLWIDTHSENDLOG11", "COLWIDTHSENDLOG12", "COLWIDTHSENDLOG13", "COLWIDTHSENDLOG14", "COLWIDTHSENDLOG15"
150             frmSMSMain.dtaSettings.Recordset("SettingValue") = Str$(vSettingValue)
  
152         Case "SENDLOGFIELDENABLED01", "SENDLOGFIELDENABLED02", "SENDLOGFIELDENABLED03", "SENDLOGFIELDENABLED04", "SENDLOGFIELDENABLED05", "SENDLOGFIELDENABLED06", "SENDLOGFIELDENABLED07", "SENDLOGFIELDENABLED08", "SENDLOGFIELDENABLED09", "SENDLOGFIELDENABLED10", "SENDLOGFIELDENABLED11", "SENDLOGFIELDENABLED12", "SENDLOGFIELDENABLED13", "SENDLOGFIELDENABLED14", "SENDLOGFIELDENABLED15", "SENDLOGFIELDENABLED16"

154             If vSettingValue = True Then
156                 frmSMSMain.dtaSettings.Recordset("SettingValue") = "TRUE"
                Else
158                 frmSMSMain.dtaSettings.Recordset("SettingValue") = "FALSE"
                End If
  
160         Case "JOBLOGFIELDENABLED01", "JOBLOGFIELDENABLED02", "JOBLOGFIELDENABLED03", "JOBLOGFIELDENABLED04", "JOBLOGFIELDENABLED05", "JOBLOGFIELDENABLED06", "JOBLOGFIELDENABLED07", "JOBLOGFIELDENABLED08", "JOBLOGFIELDENABLED09", "JOBLOGFIELDENABLED10", "JOBLOGFIELDENABLED11", "JOBLOGFIELDENABLED12", "JOBLOGFIELDENABLED13", "JOBLOGFIELDENABLED14", "JOBLOGFIELDENABLED15", "JOBLOGFIELDENABLED16"

162             If vSettingValue = True Then
164                 frmSMSMain.dtaSettings.Recordset("SettingValue") = "TRUE"
                Else
166                 frmSMSMain.dtaSettings.Recordset("SettingValue") = "FALSE"
                End If
  
168         Case "SINGLESMSUSEUSERDEFINEDLIFETIME", "SINGLESMSUSEUDSETTINGSONLYFORSPECIFICSMSTYPES", "PERIODICSMSUSEUSERDEFINEDLIFETIME", "PERIODICSMSUSEUDSETTINGSONLYFORSPECIFICSMSTYPES"

170             If vSettingValue = True Then
172                 frmSMSMain.dtaSettings.Recordset("SettingValue") = "TRUE"
                Else
174                 frmSMSMain.dtaSettings.Recordset("SettingValue") = "FALSE"
                End If
  
176         Case "SINGLESMSVALIDITYPERIODMODE", "SINGLESMSSPECIFICSETTINGLIFETIME", "SINGLESMSSPECIFICSETTINGLIFETIMEUNIT", "PERIODICSMSVALIDITYPERIODMODE", "PERIODICSMSSPECIFICSETTINGLIFETIME", "PERIODICSMSSPECIFICSETTINGLIFETIMEUNIT"
178             frmSMSMain.dtaSettings.Recordset("SettingValue") = Str$(vSettingValue)
  
180         Case Else
                'Do nothing
    
        End Select

182     frmSMSMain.dtaSettings.Recordset.UpDate

        '<EhFooter>
        Exit Sub

PutSettingIntoDataBase_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.PutSettingIntoDataBase " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub ReadInRecipientsFile(sInputFile As String)
        '<EhHeader>
        On Error GoTo ReadInRecipientsFile_Err
        '</EhHeader>
        Dim nInputFile As Integer
        Dim sLine As String
        Dim nCounter As Integer

100     nInputFile = FreeFile

102     Open sInputFile For Input As nInputFile

104     If FileLen(sInputFile) = 0 Then
            Exit Sub
        End If

106     Do While Not EOF(nInputFile)
108         Line Input #nInputFile, sLine
110         frmSMSMain.grdRecipients.AddItem LoadLanguageSpecificString(gnLanguage, 198) & Chr$(9) & sLine
112         nCounter = nCounter + 1
        Loop

114     Close nInputFile

116     MsgBox Str$(nCounter) & " " & LoadLanguageSpecificString(gnLanguage, 199), vbInformation, gsApplicationName
        '<EhFooter>
        Exit Sub

ReadInRecipientsFile_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ReadInRecipientsFile " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub RefreshPreviewGrid(sInputFile As String)
'        '<EhHeader>
'        On Error GoTo RefreshPreviewGrid_Err
'        '</EhHeader>
'        Dim nInputFile As Integer
'        Dim sLine As String
'        Dim lCounter As Long
'        Dim sSeparator As String
'        Dim nFormat As Integer
'        Dim sLineParts() As String
'        Dim nMaxLineParts As Integer
'        Dim nMaxVisibleCols As Integer
'        Dim i As Integer
'
'        On Error GoTo ErrorTrap
'
'100     Screen.MousePointer = vbHourglass
'
'102     nInputFile = FreeFile
'
'104     Open sInputFile For Input As nInputFile
'
'106     If FileLen(sInputFile) = 0 Then
'108         Close nInputFile
'            Exit Sub
'        End If
'
'        'The used Separator
''110     Select Case True
'
''            Case UCase(frmImport.cboSeparator.Text) = "<TAB>"
''112             sSeparator = vbTab
''
''114         Case UCase(frmImport.cboSeparator.Text) = "<SPACE>"
''116             sSeparator = " "
''
''118         Case Else
''120             sSeparator = Left(frmImport.cboSeparator.Text, 1)
''
''        End Select
'
'122     frmImport.grdImport.Cols = 1
'124     frmImport.grdImport.Rows = 1
'
'126     Do While Not EOF(nInputFile)
'128         Line Input #nInputFile, sLine
'
'130         sLineParts = Split(sLine, sSeparator)
'
'132         If UBound(sLineParts) <> -1 Then
'134             lCounter = lCounter + 1
'            End If
'
'136         If UBound(sLineParts) + 1 > nMaxLineParts Then
'138             nMaxLineParts = UBound(sLineParts) + 1
'            End If
'
'        Loop
'
'140     Close
'
'142     If nMaxLineParts <= 5 Then
'144         nMaxVisibleCols = nMaxLineParts
'        Else
'146         nMaxVisibleCols = nMaxLineParts
'        End If
'
'148     For i = 1 To frmImport.cboDatabaseField.UBound
'150         Unload frmImport.cboDatabaseField(i)
'        Next
'
'152     frmImport.grdImport.Rows = 0
'154     frmImport.grdImport.Cols = 0
'
'156     frmImport.cboDatabaseField(0).Width = frmImport.grdImport.Width / nMaxVisibleCols * 0.9
'158     frmImport.cboDatabaseField(0).Clear
'160     frmImport.cboDatabaseField(0).AddItem LoadLanguageSpecificString(gnLanguage, 20)
'162     frmImport.cboDatabaseField(0).AddItem LoadLanguageSpecificString(gnLanguage, 196)
'164     frmImport.cboDatabaseField(0).AddItem LoadLanguageSpecificString(gnLanguage, 197)
'166     frmImport.cboDatabaseField(0).AddItem gsPhonebookVariableField(gnLanguage, 1)
'168     frmImport.cboDatabaseField(0).AddItem gsPhonebookVariableField(gnLanguage, 2)
'170     frmImport.cboDatabaseField(0).AddItem gsPhonebookVariableField(gnLanguage, 3)
'172     frmImport.cboDatabaseField(0).Visible = True
'174     frmImport.cboDatabaseField(0).ListIndex = 0
'
'176     For i = 1 To nMaxLineParts - 1
'178         Load frmImport.cboDatabaseField(i)
'180         frmImport.cboDatabaseField(i).Top = frmImport.cboDatabaseField(0).Top
'182         frmImport.cboDatabaseField(i).Left = frmImport.cboDatabaseField(0).Left + (frmImport.grdImport.Width / nMaxVisibleCols * 0.95) * (i)
'184         frmImport.cboDatabaseField(i).Width = frmImport.grdImport.Width / nMaxVisibleCols * 0.9
'186         frmImport.cboDatabaseField(i).Visible = True
'188         frmImport.cboDatabaseField(i).AddItem LoadLanguageSpecificString(gnLanguage, 20)
'190         frmImport.cboDatabaseField(i).AddItem LoadLanguageSpecificString(gnLanguage, 196)
'192         frmImport.cboDatabaseField(i).AddItem LoadLanguageSpecificString(gnLanguage, 197)
'194         frmImport.cboDatabaseField(i).AddItem gsPhonebookVariableField(gnLanguage, 1)
'196         frmImport.cboDatabaseField(i).AddItem gsPhonebookVariableField(gnLanguage, 2)
'198         frmImport.cboDatabaseField(i).AddItem gsPhonebookVariableField(gnLanguage, 3)
'200         frmImport.cboDatabaseField(i).ListIndex = 0
'        Next
'
'202     DoEvents
'
'204     For i = 1 To frmImport.grdImport.Cols - 1
'206         frmImport.cboDatabaseField(i).Left = frmImport.grdImport.ColWidth(i) * (i) + frmImport.cboDatabaseField(0).Left + (i * 6)
'        Next
'
'208     frmImport.grdImport.Cols = nMaxLineParts
'
'210     For i = 0 To nMaxLineParts - 1
'212         frmImport.grdImport.ColWidth(i) = frmImport.grdImport.Width / nMaxVisibleCols * 0.95
'214         frmImport.grdImport.ColAlignment(i) = flexAlignLeftCenter
'        Next
'
'216     frmImport.grdImport.Rows = lCounter
'218     lCounter = 0
'
'220     Open sInputFile For Input As nInputFile
'
'222     Do While Not EOF(nInputFile)
'224         Line Input #nInputFile, sLine
'
'            'Remove Separators used twice
'
'226         sLineParts = Split(sLine, sSeparator)
'
'228         Select Case UBound(sLineParts)
'
'                Case -1
'                    'Do nothing
'
'230             Case Else
'
'232                 For i = 0 To UBound(sLineParts)
'234                     frmImport.grdImport.TextMatrix(lCounter, i) = sLineParts(i)
'                    Next
'
'236                 lCounter = lCounter + 1
'
'            End Select
'
'        Loop
'
'238     frmImport.grdImport.col = 0
'240     frmImport.grdImport.Row = 0
'242     frmImport.grdImport.ColSel = frmImport.grdImport.Cols - 1
'244     frmImport.grdImport.RowSel = frmImport.grdImport.Rows - 1
'
'246     frmImport.grdImport.CellForeColor = &H808080
'
'248     frmImport.grdImport.ColSel = 0
'250     frmImport.grdImport.RowSel = 0
'
'252     Screen.MousePointer = vbDefault
'
'254     RefreshImportCounter
'
'256     frmImport.cmdRemoveFromImportList.Enabled = True
'258     frmImport.cmdSelectAll.Enabled = True
'260     frmImport.cmdSaveRecords(1).Enabled = True
'262     frmImport.cmdSaveRecords(2).Enabled = True
'264     frmImport.cmdRefreshWindow.Enabled = True
'
'        Exit Sub
'ErrorTrap:
'
'266     Select Case Err.number
'
'            Case 30006
'268             MsgBox LoadLanguageSpecificString(gnLanguage, 26), vbCritical, gsApplicationName
'
'270         Case Else
'                'Do nothing
'
'        End Select
'
'272     Close
274     Screen.MousePointer = vbDefault
        '<EhFooter>
        Exit Sub

RefreshPreviewGrid_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.RefreshPreviewGrid " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub RefreshDataImportWindow(sInputFile As String)
'        '<EhHeader>
'        On Error GoTo RefreshDataImportWindow_Err
'        '</EhHeader>
'        Dim nInputFile As Integer
'        Dim sLine As String
'        Dim lCounter As Long
'        Dim sSeparator As String
'        Dim nFormat As Integer
'        Dim sLineParts() As String
'
'100     Screen.MousePointer = vbHourglass
'
'102     nInputFile = FreeFile
'
'104     Open sInputFile For Input As nInputFile
'
'106     If FileLen(sInputFile) = 0 Then
'            Exit Sub
'        End If
'
'        'The used Separator
'108     Select Case True
'
'            Case UCase(frmImport.cboSeparator.Text) = "<TAB>"
'110             sSeparator = vbTab
'
'112         Case UCase(frmImport.cboSeparator.Text) = "<SPACE>"
'114             sSeparator = " "
'
'116         Case Else
'118             sSeparator = Left(frmImport.cboSeparator.Text, 1)
'
'        End Select
'
'120     Do While Not EOF(nInputFile)
'
'122         Line Input #nInputFile, sLine
'
'            'Remove Separators used twice
'124         Do While InStr(1, sLine, sSeparator & sSeparator) > 0
'126             sLine = Replace(sLine, sSeparator & sSeparator, sSeparator)
'            Loop
'
'128         Select Case nFormat
'
'                Case 0 'Name <Separator> PhoneNumber
'130                 sLineParts = Split(sLine, sSeparator)
'
'132                 Select Case UBound(sLineParts)
'
'                        Case 1
'
'134                     Case 0
'
'136                     Case Else
'                            'Do nothing
'
'                    End Select
'
'138             Case 1 'PhoneNumber <Separator> Name
'140                 sLineParts = Split(sLine, sSeparator)
'
'142                 Select Case UBound(sLineParts)
'
'                        Case 1
'
'144                     Case 0
'
'146                     Case Else
'                            'Do nothing
'
'                    End Select
'
'148             Case 2 'Only Name
'150                 sLine = Replace(sLine, vbTab, " ")
'
'152             Case 3 'Only Phonenumber
'154                 sLine = Replace(sLine, vbTab, " ")
'
'156             Case Else
'                    'Unexecpted, do nothing, probably an application bug
'
'            End Select
'
'158         lCounter = lCounter + 1
'
'        Loop
'
'160     Close nInputFile
'162     Screen.MousePointer = vbDefault
'        '<EhFooter>
'        Exit Sub
'
'RefreshDataImportWindow_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.Programs.RefreshDataImportWindow " & _
'               "at line " & Erl
'        Resume Next
        '</EhFooter>
End Sub

Sub RestoreGeneralSettings()
        '<EhHeader>
        On Error GoTo RestoreGeneralSettings_Err
        '</EhHeader>
        Dim i As Integer
        Dim k As Integer
        Dim nIndexLastUsedControl As Integer
        Dim nIndexFirstEnabledControl As Integer
        Dim bEnabledControlFound As Boolean
        Dim sTemp As String
        Dim sJobRemarksField As String
100     ReDim boptSMSTypeCheck(0 To 11) As Boolean

102     frmSMSMain.txtUserkey.Text = "MZ6T5UIZGWMQ" 'GetSettingFromDatabase("Userkey")
        'frmSMSMain.txtPassword.Text = GetSettingFromDatabase("Password")
104     frmSMSMain.txtOriginator.Text = GetSettingFromDatabase("Originator")

106     gnLanguage = GetSettingFromDatabase("Language")

108     Select Case gnLanguage

            Case 1
110             frmSMSMain.mnuLanguage(1).Checked = True
112             frmSMSMain.mnuLanguage(2).Checked = False

114         Case 2
116             frmSMSMain.mnuLanguage(1).Checked = False
118             frmSMSMain.mnuLanguage(2).Checked = True
    
120         Case Else
    
        End Select

122     gtFontSettingsCurrent.bSpecificFontUsed = GetSettingFromDatabase("SpecificFontUsed")
124     gtFontSettingsCurrent.sFontName = GetSettingFromDatabase("FontName")
126     gtFontSettingsCurrent.rFontSize = GetSettingFromDatabase("FontSize")
128     gtFontSettingsCurrent.bFontBold = GetSettingFromDatabase("FontBold")
130     gtFontSettingsCurrent.bFontItalic = GetSettingFromDatabase("FontItalic")

132     For i = frmSMSMain.optSMSType.LBound To frmSMSMain.optSMSType.UBound

134         If frmSMSMain.optSMSType(i).Value = True Then
136             nIndexLastUsedControl = i
                Exit For
            End If

        Next

138     gtMenuControl.bTextSMSEnabled = GetSettingFromDatabase("TextSMSEnabled")
140     gtMenuControl.bOperatorLogoEnabled = GetSettingFromDatabase("OperatorLogoEnabled")
142     gtMenuControl.bGroupLogoEnabled = GetSettingFromDatabase("GroupLogoEnabled")
144     gtMenuControl.bRingtoneEnabled = GetSettingFromDatabase("RingtoneEnabled")
146     gtMenuControl.bPictureMessageEnabled = GetSettingFromDatabase("PictureMessageEnabled")
148     gtMenuControl.bVCardEnabled = GetSettingFromDatabase("VCardEnabled")
150     gtMenuControl.bUnicodeEnabled = GetSettingFromDatabase("UnicodeEnabled")
152     gtMenuControl.bWAPPushSMSEnabled = GetSettingFromDatabase("WAPPushSMSEnabled")
154     gtMenuControl.bBinaryDataEnabled = GetSettingFromDatabase("BinaryDataEnabled")

156     VersionSpecificAction 13

158     frmSMSMain.optSMSType(0).Enabled = gtMenuControl.bTextSMSEnabled
160     frmSMSMain.optSMSType(1).Enabled = gtMenuControl.bOperatorLogoEnabled
162     frmSMSMain.optSMSType(2).Enabled = gtMenuControl.bGroupLogoEnabled
164     frmSMSMain.optSMSType(3).Enabled = gtMenuControl.bRingtoneEnabled
166     frmSMSMain.optSMSType(4).Enabled = gtMenuControl.bPictureMessageEnabled
168     frmSMSMain.optSMSType(5).Enabled = gtMenuControl.bVCardEnabled
170     frmSMSMain.optSMSType(6).Enabled = gtMenuControl.bUnicodeEnabled
172     frmSMSMain.optSMSType(7).Enabled = gtMenuControl.bWAPPushSMSEnabled
174     frmSMSMain.optSMSType(8).Enabled = gtMenuControl.bBinaryDataEnabled

176     VersionSpecificAction 4

178     boptSMSTypeCheck(0) = gtMenuControl.bTextSMSEnabled
180     boptSMSTypeCheck(1) = gtMenuControl.bOperatorLogoEnabled
182     boptSMSTypeCheck(2) = gtMenuControl.bGroupLogoEnabled
184     boptSMSTypeCheck(3) = gtMenuControl.bRingtoneEnabled
186     boptSMSTypeCheck(4) = gtMenuControl.bPictureMessageEnabled
188     boptSMSTypeCheck(5) = gtMenuControl.bVCardEnabled
190     boptSMSTypeCheck(6) = gtMenuControl.bUnicodeEnabled
192     boptSMSTypeCheck(7) = gtMenuControl.bWAPPushSMSEnabled
194     boptSMSTypeCheck(8) = gtMenuControl.bBinaryDataEnabled

196     VersionSpecificAction 14, , , , boptSMSTypeCheck(9)
198     VersionSpecificAction 15, , , , boptSMSTypeCheck(10)
200     VersionSpecificAction 16, , , , boptSMSTypeCheck(11)

202     If boptSMSTypeCheck(nIndexLastUsedControl) = True Then
            'Do nothing, everything ok
204         frmSMSMain.optSMSType(nIndexLastUsedControl).Value = True
        Else

            'Search for available control
206         For i = LBound(boptSMSTypeCheck) To UBound(boptSMSTypeCheck)

208             If boptSMSTypeCheck(i) = True Then
210                 nIndexFirstEnabledControl = i
212                 bEnabledControlFound = True
                    Exit For
                End If

            Next

214         If bEnabledControlFound Then  'At least one available
216             frmSMSMain.optSMSType(nIndexFirstEnabledControl).Value = True
218             frmSMSMain.cmdSend.Enabled = True
            Else

                'No control found
220             For i = frmSMSMain.optSMSType.LBound To frmSMSMain.optSMSType.UBound
222                 frmSMSMain.optSMSType(i).Value = False
                Next

224             frmSMSMain.cmdSend.caption = "Send"
226             frmSMSMain.cmdSend.Enabled = False
            End If
        End If

228     gbSaveMessagesInSendLog = GetSettingFromDatabase("SaveMessagesInSendJournal")

230     For i = 1 To 2  'Language
232         For k = LBound(gsPhonebookVariableField, 2) To UBound(gsPhonebookVariableField, 2) 'Number of Variable Fields
234             gsPhonebookVariableField(i, k) = GetPhonebookVariableFieldFromDatabase(i, k)
            Next
        Next

236     gtValidityPeriodSettings.bSingleSMSUseUserDefinedLifeTime = GetSettingFromDatabase("SingleSMSUseUserDefinedLifetime")
238     gtValidityPeriodSettings.bSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes = GetSettingFromDatabase("SingleSMSUseUDSettingsOnlyForSpecificSMSTypes")
240     gtValidityPeriodSettings.nSingleSMSValidityPeriodMode = GetSettingFromDatabase("SingleSMSValidityPeriodMode")
242     gtValidityPeriodSettings.lSingleSMSSpecificSettingLifeTime = GetSettingFromDatabase("SingleSMSSpecificSettingLifeTime")
244     gtValidityPeriodSettings.nSingleSMSSpecificSettingLifeTimeUnit = GetSettingFromDatabase("SingleSMSSpecificSettingLifeTimeUnit")

246     gtValidityPeriodSettings.bPeriodicSMSUseUserDefinedLifeTime = GetSettingFromDatabase("PeriodicSMSUseUserDefinedLifeTime")
248     gtValidityPeriodSettings.bPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes = GetSettingFromDatabase("PeriodicSMSUseUDSettingsOnlyForSpecificSMSTypes")
250     gtValidityPeriodSettings.nPeriodicSMSValidityPeriodMode = GetSettingFromDatabase("PeriodicSMSValidityPeriodMode")
252     gtValidityPeriodSettings.lPeriodicSMSSpecificSettingLifeTime = GetSettingFromDatabase("PeriodicSMSSpecificSettingLifeTime")
254     gtValidityPeriodSettings.nPeriodicSMSSpecificSettingLifeTimeUnit = GetSettingFromDatabase("PeriodicSMSSpecificSettingLifeTimeUnit")
        '<EhFooter>
        Exit Sub

RestoreGeneralSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.RestoreGeneralSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Sub SaveOutgoingMessage(sRecipient As String, _
                        sRecipientNameFromGrid As String, _
                        sTransactionReferenceNumber As String, _
                        sMessage As String, _
                        sDeferredDeliveryTime As String, _
                        lJobId As Long)
        '<EhHeader>
        On Error GoTo SaveOutgoingMessage_Err
        '</EhHeader>
100     grsSendlog.AddNew

102     If sRecipientNameFromGrid = "" Then
104         grsSendlog("sName") = GetNameFromPhoneBook(sRecipient)
        Else
106         grsSendlog("sName") = sRecipientNameFromGrid
        End If

108     grsSendlog("sRecipient") = sRecipient & ""
110     grsSendlog("sTransactionReferenceNumber") = sTransactionReferenceNumber & ""
112     grsSendlog("sMessage") = sMessage & ""

114     If sDeferredDeliveryTime <> "" Then
116         grsSendlog("dDeferredDeliveryTime") = CVDate(sDeferredDeliveryTime) & ""
        End If

118     grsSendlog("sDeliveryStatus") = "???"
120     grsSendlog("sDeliveryStatusText") = "???"
122     grsSendlog("sReasonCode") = "???"
124     grsSendlog("sReasonText") = "???"
126     grsSendlog("sReasonCodeBuffered") = "???"
128     grsSendlog("sReasonTextBuffered") = "???"
130     grsSendlog("lJobID") = lJobId
132     grsSendlog.UpDate
        '<EhFooter>
        Exit Sub

SaveOutgoingMessage_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.SaveOutgoingMessage " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function ConvertStringToDate(sDDMMYYYYHHmmss As String) As Date
        '<EhHeader>
        On Error GoTo ConvertStringToDate_Err
        '</EhHeader>
        Dim sWork As String
        Dim dWork As Date

        Dim nYear As Integer
        Dim nMonth As Integer
        Dim nDay As Integer
        Dim nHour As Integer
        Dim nMinute As Integer
        Dim nSecond As Integer

100     sWork = sDDMMYYYYHHmmss

        'Check Length
102     If Len(sWork) > 14 Then
104         sWork = Left(sWork, 14)
        End If

106     If Len(sWork) < 14 Then
108         ConvertStringToDate = False
            Exit Function
        End If

110     nDay = Val(Mid(sWork, 1, 2))
112     nMonth = Val(Mid(sWork, 3, 2))
114     nYear = Val(Mid(sWork, 5, 4))
116     nHour = Val(Mid(sWork, 9, 2))
118     nMinute = Val(Mid(sWork, 11, 2))
120     nSecond = Val(Mid(sWork, 13, 2))

122     dWork = DateSerial(nYear, nMonth, nDay)
124     dWork = DateAdd("h", nHour, dWork)
126     dWork = DateAdd("n", nMinute, dWork)
128     dWork = DateAdd("s", nSecond, dWork)

130     ConvertStringToDate = dWork
        '<EhFooter>
        Exit Function

ConvertStringToDate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ConvertStringToDate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function DeliveryStatusFromDeliveryStatusCode(sDeliveryStatus As String) As String
        '<EhHeader>
        On Error GoTo DeliveryStatusFromDeliveryStatusCode_Err
        '</EhHeader>

100     Select Case sDeliveryStatus

            Case "-1"
102             DeliveryStatusFromDeliveryStatusCode = LoadLanguageSpecificString(gnLanguage, 203)

104         Case "0"
106             DeliveryStatusFromDeliveryStatusCode = LoadLanguageSpecificString(gnLanguage, 204)

108         Case "1"
110             DeliveryStatusFromDeliveryStatusCode = LoadLanguageSpecificString(gnLanguage, 205)

112         Case "2"
114             DeliveryStatusFromDeliveryStatusCode = LoadLanguageSpecificString(gnLanguage, 206)

116         Case Else
                'Unexecpted, do nothing, probably an application bug

        End Select

        '<EhFooter>
        Exit Function

DeliveryStatusFromDeliveryStatusCode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.DeliveryStatusFromDeliveryStatusCode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ReasonRemarkFromReasonCodeTill(sReasonCode As String, sDST As String) As String
        '<EhHeader>
        On Error GoTo ReasonRemarkFromReasonCodeTill_Err
        '</EhHeader>

100     Select Case sReasonCode

            Case ""
102             ReasonRemarkFromReasonCodeTill = ""

104         Case "000"
106             ReasonRemarkFromReasonCodeTill = "Es wurde versucht, eine Mitteilung an eine Nummer zu schicken, die nicht oder nicht mehr existiert. Bitte überprüfen Sie, ob Sie die Nummer richtig eingegeben haben."

108         Case "001"
110             ReasonRemarkFromReasonCodeTill = "-"

112         Case "002"
114             ReasonRemarkFromReasonCodeTill = "-"

116         Case "003"
118             ReasonRemarkFromReasonCodeTill = "-"

120         Case "004"
122             ReasonRemarkFromReasonCodeTill = "-"

124         Case "005"
126             ReasonRemarkFromReasonCodeTill = "-"

128         Case "006"
130             ReasonRemarkFromReasonCodeTill = "-"

132         Case "007"
134             ReasonRemarkFromReasonCodeTill = "-"

136         Case "008"
138             ReasonRemarkFromReasonCodeTill = "-"

140         Case "009"
142             ReasonRemarkFromReasonCodeTill = "-"

144         Case "010"
146             ReasonRemarkFromReasonCodeTill = "-"

148         Case "100"
150             ReasonRemarkFromReasonCodeTill = "-"

152         Case "101"
154             ReasonRemarkFromReasonCodeTill = "-"

156         Case "102"
158             ReasonRemarkFromReasonCodeTill = "-"

160         Case "103"
162             ReasonRemarkFromReasonCodeTill = "Es wurde versucht, eine Mitteilung an eine Nummer zu schicken, die nicht oder nicht mehr existiert. Bitte überprüfen Sie, ob Sie die Nummer richtig eingegeben haben."

164         Case "104"
166             ReasonRemarkFromReasonCodeTill = "-"

168         Case "105"
170             ReasonRemarkFromReasonCodeTill = "-"

172         Case "106"
174             ReasonRemarkFromReasonCodeTill = "-"

176         Case "107"
178             ReasonRemarkFromReasonCodeTill = "Es wurde versucht, eine Mitteilung an ein Mobiltelefon zu schicken, das momentan nicht eingeschaltet ist oder keinen Empfang hat."

180         Case "108"

182             Select Case sDST

                    Case "1"
184                     ReasonRemarkFromReasonCodeTill = "Es wurde versucht, eine Mitteilung zu schicken, die SMS konnte jedoch nicht im Mobiltelefon gespeichert werden." & " " & "Dieser Reasoncode wurde schon häufig beobachtet, wenn der Handyspeicher voll war und kein Platz für weitere Mitteilungen vorhanden war. Falls der Empfänger eine oder mehrere SMS löscht, kann die Mitteilung wahrscheinlich trotzdem noch ausgeliefert werden."
    
186                 Case "2"
188                     ReasonRemarkFromReasonCodeTill = "Die Gültigkeitsdauer der Mitteilung wurde überschritten."
    
                End Select

190         Case "109"
192             ReasonRemarkFromReasonCodeTill = "-"

194         Case "110"
196             ReasonRemarkFromReasonCodeTill = "Dieser Reasoncode wurde schon häufig im Zusammenhang mit Nokia 91xx Mobiltelefonen und der Verwendung von alphanumerischen Absendern beobacht, es sind jedoch auch andere Fehlerursachen denkbar."

198         Case "111"
200             ReasonRemarkFromReasonCodeTill = "-"

202         Case "112"
204             ReasonRemarkFromReasonCodeTill = "-"

206         Case "113"
208             ReasonRemarkFromReasonCodeTill = "-"

210         Case "114"
212             ReasonRemarkFromReasonCodeTill = "-"

214         Case "115"
216             ReasonRemarkFromReasonCodeTill = "-"

218         Case "116"
220             ReasonRemarkFromReasonCodeTill = "-"

222         Case "117"
224             ReasonRemarkFromReasonCodeTill = "-"

226         Case "118"
228             ReasonRemarkFromReasonCodeTill = "Es wurde versucht, eine Mitteilung zu schicken, es ist jedoch ein nicht näher spezifizierbarer Fehler im Netzwerk aufgetreten. Dieser Fehler wurde schon bei nicht erreichbaren oder überlasteten Netzwerken und kriegerischen Ereignissen beobachtet."

230         Case "119"
232             ReasonRemarkFromReasonCodeTill = "Es wurde versucht, eine Mitteilung zu schicken, ist es jedoch ein Fehler im Netzwerk (Public Land Mobile Network) aufgetreten."

234         Case "120"
236             ReasonRemarkFromReasonCodeTill = "Es wurde versucht, eine Mitteilung zu schicken, es ist jedoch ein schwerer Fehler im Netzwerk (Home Location Register) aufgetreten." & " " & "Dieser Reasoncode wurde schon häufig beobachtet, wenn versucht wurde, eine SMS in ein Netzwerk zu schicken, das nicht erreichbar ist. In der Regel wird dann häufig nach Ablauf der Gültigkeitsdauer ein Übertragungsfehler (Code 108) angezeigt."

238         Case "121"
240             ReasonRemarkFromReasonCodeTill = "Es wurde versucht, eine Mitteilung zu schicken, es ist jedoch ein Fehler im Netzwerk (Visiting Location Register) aufgetreten." & " " & "Dieser Reasoncode wurde schon im Zusammenhang mit Roamingproblemen beobachtet, z.B. wenn sich der Empfänger im Ausland befand und sein Mobiltelefon in einem Netzwerk eingebucht war, dass keinen SMS Roamingvertrag mit seinem Heimnetzwerk abgeschlossen hatte."

242         Case "122"
244             ReasonRemarkFromReasonCodeTill = "-"

246         Case "123"
248             ReasonRemarkFromReasonCodeTill = "Es wurde versucht, eine Mitteilung zu schicken, es ist jedoch ein Fehler im Netzwerk (Controlling MSC) aufgetreten." & " " & "Dieser Reasoncode wurde schon im Zusammenhang mit Roamingproblemen beobachtet, z.B. wenn sich der Empfänger im Ausland befand und sein Mobiltelefon in einem Netzwerk eingebucht war, dass keinen SMS Roamingvertrag mit seinem Heimnetzwerk abgeschlossen hatte."

250         Case "124"
252             ReasonRemarkFromReasonCodeTill = "-"

254         Case "125"
256             ReasonRemarkFromReasonCodeTill = "-"

258         Case "126"
260             ReasonRemarkFromReasonCodeTill = "-"

262         Case "127"
264             ReasonRemarkFromReasonCodeTill = "-"

266         Case "200"
268             ReasonRemarkFromReasonCodeTill = "-"

270         Case "201"
272             ReasonRemarkFromReasonCodeTill = "-"

274         Case "202"
276             ReasonRemarkFromReasonCodeTill = "-"

278         Case "203"
280             ReasonRemarkFromReasonCodeTill = "-"

282         Case "204"
284             ReasonRemarkFromReasonCodeTill = "-"

286         Case "205"
288             ReasonRemarkFromReasonCodeTill = "-"

290         Case "206"
292             ReasonRemarkFromReasonCodeTill = "-"

294         Case "207"
296             ReasonRemarkFromReasonCodeTill = "-"

298         Case "208"
300             ReasonRemarkFromReasonCodeTill = "-"

302         Case Else
304             ReasonRemarkFromReasonCodeTill = "Unbekannter Fehlercode"

        End Select

        '<EhFooter>
        Exit Function

ReasonRemarkFromReasonCodeTill_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ReasonRemarkFromReasonCodeTill " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ReasonRemarkFromReasonCode(sReasonCode As String, sDST As String) As String
        '<EhHeader>
        On Error GoTo ReasonRemarkFromReasonCode_Err
        '</EhHeader>

100     Select Case sReasonCode

            Case ""
102             ReasonRemarkFromReasonCode = ""

104         Case "000"
106             ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 351)

108         Case "001"
110             ReasonRemarkFromReasonCode = "-"

112         Case "002"
114             ReasonRemarkFromReasonCode = "-"

116         Case "003"
118             ReasonRemarkFromReasonCode = "-"

120         Case "004"
122             ReasonRemarkFromReasonCode = "-"

124         Case "005"
126             ReasonRemarkFromReasonCode = "-"

128         Case "006"
130             ReasonRemarkFromReasonCode = "-"

132         Case "007"
134             ReasonRemarkFromReasonCode = "-"

136         Case "008"
138             ReasonRemarkFromReasonCode = "-"

140         Case "009"
142             ReasonRemarkFromReasonCode = "-"

144         Case "010"
146             ReasonRemarkFromReasonCode = "-"

148         Case "100"
150             ReasonRemarkFromReasonCode = "-"

152         Case "101"
154             ReasonRemarkFromReasonCode = "-"

156         Case "102"
158             ReasonRemarkFromReasonCode = "-"

160         Case "103"
162             ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 352)

164         Case "104"
166             ReasonRemarkFromReasonCode = "-"

168         Case "105"
170             ReasonRemarkFromReasonCode = "-"

172         Case "106"
174             ReasonRemarkFromReasonCode = "-"

176         Case "107"
178             ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 353)

180         Case "108"

182             Select Case sDST

                    Case "1"
184                     ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 354) & " " & LoadLanguageSpecificString(gnLanguage, 355)
    
186                 Case "2"
188                     ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 356)
    
                End Select

190         Case "109"
192             ReasonRemarkFromReasonCode = "-"

194         Case "110"
196             ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 358)

198         Case "111"
200             ReasonRemarkFromReasonCode = "-"

202         Case "112"
204             ReasonRemarkFromReasonCode = "-"

206         Case "113"
208             ReasonRemarkFromReasonCode = "-"

210         Case "114"
212             ReasonRemarkFromReasonCode = "-"

214         Case "115"
216             ReasonRemarkFromReasonCode = "-"

218         Case "116"
220             ReasonRemarkFromReasonCode = "-"

222         Case "117"
224             ReasonRemarkFromReasonCode = "-"

226         Case "118"
228             ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 359)

230         Case "119"
232             ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 360)

234         Case "120"
236             ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 361) & " " & LoadLanguageSpecificString(gnLanguage, 362)

238         Case "121"
240             ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 363) & " " & LoadLanguageSpecificString(gnLanguage, 364)

242         Case "122"
244             ReasonRemarkFromReasonCode = "-"

246         Case "123"
248             ReasonRemarkFromReasonCode = "-"

250         Case "124"
252             ReasonRemarkFromReasonCode = "-"

254         Case "125"
256             ReasonRemarkFromReasonCode = "-"

258         Case "126"
260             ReasonRemarkFromReasonCode = "-"

262         Case "127"
264             ReasonRemarkFromReasonCode = "-"

266         Case "200"
268             ReasonRemarkFromReasonCode = "-"

270         Case "201"
272             ReasonRemarkFromReasonCode = "-"

274         Case "202"
276             ReasonRemarkFromReasonCode = "-"

278         Case "203"
280             ReasonRemarkFromReasonCode = "-"

282         Case "204"
284             ReasonRemarkFromReasonCode = "-"

286         Case "205"
288             ReasonRemarkFromReasonCode = "-"

290         Case "206"
292             ReasonRemarkFromReasonCode = "-"

294         Case "207"
296             ReasonRemarkFromReasonCode = "-"

298         Case "208"
300             ReasonRemarkFromReasonCode = "-"

302         Case Else
304             ReasonRemarkFromReasonCode = LoadLanguageSpecificString(gnLanguage, 310)

        End Select

        '<EhFooter>
        Exit Function

ReasonRemarkFromReasonCode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ReasonRemarkFromReasonCode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ReasonFromReasonCode(sReasonCode As String) As String
        '<EhHeader>
        On Error GoTo ReasonFromReasonCode_Err
        '</EhHeader>

100     Select Case sReasonCode

            Case ""
102             ReasonFromReasonCode = ""

104         Case "000"
106             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 261)

108         Case "001"
110             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 262)

112         Case "002"
114             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 263)

116         Case "003"
118             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 264)

120         Case "004"
122             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 265)

124         Case "005"
126             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 266)

128         Case "006"
130             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 267)

132         Case "007"
134             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 268)

136         Case "008"
138             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 269)

140         Case "009"
142             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 270)

144         Case "010"
146             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 271)

148         Case "100"
150             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 272)

152         Case "101"
154             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 273)

156         Case "102"
158             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 274)

160         Case "103"
162             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 275)

164         Case "104"
166             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 276)

168         Case "105"
170             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 277)

172         Case "106"
174             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 278)

176         Case "107"
178             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 279)

180         Case "108"
182             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 281)

184         Case "109"
186             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 282)

188         Case "110"
190             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 283)

192         Case "111"
194             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 284)

196         Case "112"
198             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 285)

200         Case "113"
202             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 286)

204         Case "114"
206             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 287)

208         Case "115"
210             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 288)

212         Case "116"
214             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 289)

216         Case "117"
218             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 290)

220         Case "118"
222             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 291)

224         Case "119"
226             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 292)

228         Case "120"
230             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 293)

232         Case "121"
234             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 294)

236         Case "122"
238             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 295)

240         Case "123"
242             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 296)

244         Case "124"
246             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 297)

248         Case "125"
250             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 298)

252         Case "126"
254             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 299)

256         Case "127"
258             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 300)

260         Case "200"
262             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 301)

264         Case "201"
266             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 302)

268         Case "202"
270             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 303)

272         Case "203"
274             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 304)

276         Case "204"
278             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 305)

280         Case "205"
282             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 306)

284         Case "206"
286             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 307)

288         Case "207"
290             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 308)

292         Case "208"
294             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 309)

296         Case Else
298             ReasonFromReasonCode = LoadLanguageSpecificString(gnLanguage, 310)

        End Select

        '<EhFooter>
        Exit Function

ReasonFromReasonCode_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ReasonFromReasonCode " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ConvertToASPSMSTimeFormat(sInput As String) As String
        '<EhHeader>
        On Error GoTo ConvertToASPSMSTimeFormat_Err
        '</EhHeader>
        Dim dTemp As Date
        Dim sDay As String
        Dim sMonth As String
        Dim sYear As String
        Dim shour As String
        Dim sMinute As String
        Dim sSecond As String

        'The dateformat has to be: ddmmyyyyhhmmss
100     If IsDate(sInput) Then
102         dTemp = CVDate(sInput)
104         sDay = Right("00" & DatePart("d", dTemp), 2)
106         sMonth = Right("00" & DatePart("m", dTemp), 2)
108         sYear = DatePart("yyyy", dTemp)
110         shour = Right("00" & DatePart("h", dTemp), 2)
112         sMinute = Right("00" & DatePart("n", dTemp), 2)
114         sSecond = Right("00" & DatePart("s", dTemp), 2)
116         ConvertToASPSMSTimeFormat = sDay & sMonth & sYear & shour & sMinute & sSecond
        Else
118         ConvertToASPSMSTimeFormat = ""
        End If

        '<EhFooter>
        Exit Function

ConvertToASPSMSTimeFormat_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ConvertToASPSMSTimeFormat " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function SendLogSQLStatement(nLanguage As Integer, sSQLWhere As String) As String
        '<EhHeader>
        On Error GoTo SendLogSQLStatement_Err
        '</EhHeader>
        Dim sSQL As String
        Dim sWhere As String
        Dim sTemp As String
100     sSQL = "SELECT lID as ID, "
102     sSQL = sSQL & "sName as [" & LoadLanguageSpecificString(gnLanguage, 600) & "], "
104     sSQL = sSQL & "sRecipient as [" & LoadLanguageSpecificString(gnLanguage, 601) & "], "
106     sSQL = sSQL & "sTransactionReferenceNumber as [" & LoadLanguageSpecificString(gnLanguage, 602) & "], "
108     sSQL = sSQL & "sMessage as [" & LoadLanguageSpecificString(gnLanguage, 603) & "], "
110     sSQL = sSQL & "dDeferredDeliveryTime as [" & LoadLanguageSpecificString(gnLanguage, 604) & "], "
112     sSQL = sSQL & "dSubmissionDate as [" & LoadLanguageSpecificString(gnLanguage, 605) & "], "
114     sSQL = sSQL & "dNotificationDate as [" & LoadLanguageSpecificString(gnLanguage, 606) & "], "
116     sSQL = sSQL & "sReasoncode as [" & LoadLanguageSpecificString(gnLanguage, 607) & "], "
118     sSQL = sSQL & "dNotificationDateBuffered as [" & LoadLanguageSpecificString(gnLanguage, 608) & "], "
120     sSQL = sSQL & "sReasonCodeBuffered as [" & LoadLanguageSpecificString(gnLanguage, 609) & "], "
122     sSQL = sSQL & "sReasonTextBuffered as [" & LoadLanguageSpecificString(gnLanguage, 610) & "], "
124     sSQL = sSQL & "sDeliveryStatus as [" & LoadLanguageSpecificString(gnLanguage, 611) & "], "
126     sSQL = sSQL & "sDeliveryStatusText as [" & LoadLanguageSpecificString(gnLanguage, 612) & "], "
128     sSQL = sSQL & "sReasonText as [" & LoadLanguageSpecificString(gnLanguage, 613) & "], "
130     sSQL = sSQL & "lJobID "
132     sSQL = sSQL & "FROM SendJournal "
       
134     sSQL = "SELECT lID as ID, "

136     If gtDBLogFieldsSettings.bSendLog(2) = True Then
138         sSQL = sSQL & "lJobID as [" & LoadLanguageSpecificString(nLanguage, 615) & "], "
        End If

140     If gtDBLogFieldsSettings.bSendLog(3) = True Then
142         sSQL = sSQL & "sName as [" & LoadLanguageSpecificString(gnLanguage, 600) & "], "
        End If

144     If gtDBLogFieldsSettings.bSendLog(4) = True Then
146         sSQL = sSQL & "sRecipient as [" & LoadLanguageSpecificString(gnLanguage, 601) & "], "
        End If

148     If gtDBLogFieldsSettings.bSendLog(5) = True Then
150         sSQL = sSQL & "sTransactionReferenceNumber as [" & LoadLanguageSpecificString(gnLanguage, 602) & "], "
        End If

152     If gtDBLogFieldsSettings.bSendLog(6) = True Then
154         sSQL = sSQL & "sMessage as [" & LoadLanguageSpecificString(gnLanguage, 603) & "], "
        End If

156     If gtDBLogFieldsSettings.bSendLog(7) = True Then
158         sSQL = sSQL & "dDeferredDeliveryTime as [" & LoadLanguageSpecificString(gnLanguage, 604) & "], "
        End If

160     If gtDBLogFieldsSettings.bSendLog(8) = True Then
162         sSQL = sSQL & "dSubmissionDate as [" & LoadLanguageSpecificString(gnLanguage, 605) & "], "
        End If

164     If gtDBLogFieldsSettings.bSendLog(9) = True Then
166         sSQL = sSQL & "dNotificationDate as [" & LoadLanguageSpecificString(gnLanguage, 606) & "], "
        End If

168     If gtDBLogFieldsSettings.bSendLog(10) = True Then
170         sSQL = sSQL & "sReasoncode as [" & LoadLanguageSpecificString(gnLanguage, 607) & "], "
        End If

172     If gtDBLogFieldsSettings.bSendLog(11) = True Then
174         sSQL = sSQL & "dNotificationDateBuffered as [" & LoadLanguageSpecificString(gnLanguage, 608) & "], "
        End If

176     If gtDBLogFieldsSettings.bSendLog(12) = True Then
178         sSQL = sSQL & "sReasonCodeBuffered as [" & LoadLanguageSpecificString(gnLanguage, 609) & "], "
        End If

180     If gtDBLogFieldsSettings.bSendLog(13) = True Then
182         sSQL = sSQL & "sReasonTextBuffered as [" & LoadLanguageSpecificString(gnLanguage, 610) & "], "
        End If

184     If gtDBLogFieldsSettings.bSendLog(14) = True Then
186         sSQL = sSQL & "sDeliveryStatus as [" & LoadLanguageSpecificString(gnLanguage, 611) & "], "
        End If

188     If gtDBLogFieldsSettings.bSendLog(15) = True Then
190         sSQL = sSQL & "sDeliveryStatusText as [" & LoadLanguageSpecificString(gnLanguage, 612) & "], "
        End If

192     If gtDBLogFieldsSettings.bSendLog(16) = True Then
194         sSQL = sSQL & "sReasonText as [" & LoadLanguageSpecificString(gnLanguage, 613) & "], "
        End If

196     sSQL = Left(sSQL, Len(sSQL) - 2)

198     sSQL = sSQL & " FROM SendJournal "
       
200     If gtSendLogFilterSettings.bRecipientNameFilterUsed Then
202         sWhere = sWhere & "(sName = '" & gtSendLogFilterSettings.sRecipientName & "') and "
        End If

204     If gtSendLogFilterSettings.bRecipientFilterUsed Then
206         sWhere = sWhere & "(sRecipient = '" & gtSendLogFilterSettings.sRecipient & "') and "
        End If

208     If gtSendLogFilterSettings.bDeliveryStatusFilterUsed Then
210         If gtSendLogFilterSettings.nDeliveryStatus = gcQuestionMarks Then
212             sWhere = sWhere & "(sDeliveryStatus = '???') and "
            Else
214             sWhere = sWhere & "(sDeliveryStatus = '" & gtSendLogFilterSettings.nDeliveryStatus & "') and "
            End If
        End If

216     If gtSendLogFilterSettings.bReasonCodeFilterUsed Then
218         If gtSendLogFilterSettings.nReasonCode = gcQuestionMarks Then
220             sWhere = sWhere & "(sReasoncode = '???' or sReasonCodeBuffered = '???') and "
            Else

222             If gtSendLogFilterSettings.nReasonCode = 0 Then 'Special case 0, Unknown subscriber
224                 sWhere = sWhere & "((sReasoncode = '000' or sReasonCodeBuffered = '000') and "
226                 sWhere = sWhere & "(sDeliveryStatus = '1' or sDeliveryStatus = '2')) and "
                Else
228                 sWhere = sWhere & "(sReasoncode = '" & gtSendLogFilterSettings.nReasonCode & "' or sReasonCodeBuffered = '" & gtSendLogFilterSettings.nReasonCode & "') and "
                End If
            End If
        End If

230     If gtSendLogFilterSettings.bJobIDFilterUsed Then
232         sWhere = sWhere & "(lJobID = " & gtSendLogFilterSettings.lJobId & ") and "
        End If

234     If sWhere <> "" Then 'Cut off last 'and'
236         sWhere = "WHERE " & Left(sWhere, Len(sWhere) - 4)
        End If

238     sSQLWhere = sWhere
240     SendLogSQLStatement = sSQL & sWhere & " ORDER BY lID"

        '<EhFooter>
        Exit Function

SendLogSQLStatement_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.SendLogSQLStatement " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Sub SetListIndexWithItemData(cboInput As ComboBox, _
                             nItemdata As Integer)
        '<EhHeader>
        On Error GoTo SetListIndexWithItemData_Err
        '</EhHeader>
        Dim nListIndex As Integer
        Dim bFound As Boolean

100     nListIndex = 0

102     Do While Not bFound And (nListIndex + 1) <= cboInput.ListCount

104         If cboInput.ItemData(nListIndex) = nItemdata Then
106             cboInput.ListIndex = nListIndex
108             bFound = True
            End If

110         nListIndex = nListIndex + 1
        Loop

        '<EhFooter>
        Exit Sub

SetListIndexWithItemData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.SetListIndexWithItemData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Sub ShowCredits(bEnableMessagebox As Boolean, Optional bShowLabel As Boolean)
        '<EhHeader>
        On Error GoTo ShowCredits_Err
        '</EhHeader>
        Dim SMS As Booster
        Dim rCredits As Double
        Dim sMessage As String

100     Set SMS = New Booster

102     SMS.Hosts = App.Path & "\smshosts.txt"
104     SMS.UserKey = frmSMSMain.txtUserkey.Text
106     SMS.Password = frmSMSMain.txtPassword.Text

108     Screen.MousePointer = vbHourglass

110     rCredits = SMS.Credits

112     frmSMSMain.caption = gsApplicationName & " - " & LoadLanguageSpecificString(gnLanguage, 186) & " " & Trim(Str$(rCredits)) & " " & LoadLanguageSpecificString(gnLanguage, 488)

114     Screen.MousePointer = vbDefault

116     Set SMS = Nothing

118     If bEnableMessagebox Then
120         sMessage = LoadLanguageSpecificString(gnLanguage, 450) & " " & Trim(Str$(rCredits)) & " " & LoadLanguageSpecificString(gnLanguage, 451)
122         MsgBox sMessage, vbInformation, gsApplicationName
        End If
    
124     If bShowLabel Then
        
        End If

        '<EhFooter>
        Exit Sub

ShowCredits_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ShowCredits " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Function SelectedSMSType(nLanguage As Integer) As String
        '<EhHeader>
        On Error GoTo SelectedSMSType_Err
        '</EhHeader>
        Dim nIndex As Integer
        Dim sTemp As String

100     Select Case True
  
            Case frmSMSMain.optSMSType(0).Value = True 'Text SMS
102             nIndex = 0
  
104         Case frmSMSMain.optSMSType(1).Value = True 'Operatorlogo
106             nIndex = 1
  
108         Case frmSMSMain.optSMSType(2).Value = True 'Grouplogo
110             nIndex = 2
  
112         Case frmSMSMain.optSMSType(3).Value = True 'Ringtone
114             nIndex = 3
  
116         Case frmSMSMain.optSMSType(4).Value = True 'Picturemessage
118             nIndex = 4
  
120         Case frmSMSMain.optSMSType(5).Value = True 'VCard
122             nIndex = 5
  
124         Case frmSMSMain.optSMSType(6).Value = True 'Unicode
126             nIndex = 6
  
128         Case frmSMSMain.optSMSType(7).Value = True 'WAP Push SMS
130             nIndex = 7
  
132         Case frmSMSMain.optSMSType(8).Value = True 'Binary Data
134             nIndex = 8
      
136         Case frmSMSMain.optSMSType(9).Value = True
138             nIndex = 9
    
140         Case frmSMSMain.optSMSType(10).Value = True
142             nIndex = 10
    
144         Case frmSMSMain.optSMSType(11).Value = True
146             nIndex = 11
    
        End Select

148     Select Case nIndex

            Case 0
150             SelectedSMSType = LoadLanguageSpecificString(nLanguage, 161)
  
152         Case 1
154             SelectedSMSType = LoadLanguageSpecificString(nLanguage, 162)
  
156         Case 2
158             SelectedSMSType = LoadLanguageSpecificString(nLanguage, 163)
  
160         Case 3
162             SelectedSMSType = LoadLanguageSpecificString(nLanguage, 164)
  
164         Case 4
166             SelectedSMSType = LoadLanguageSpecificString(nLanguage, 165)
  
168         Case 5
170             SelectedSMSType = LoadLanguageSpecificString(nLanguage, 166)
  
172         Case 6
174             SelectedSMSType = LoadLanguageSpecificString(nLanguage, 480)
  
176         Case 7
178             SelectedSMSType = LoadLanguageSpecificString(nLanguage, 478)
  
180         Case 8
182             SelectedSMSType = LoadLanguageSpecificString(nLanguage, 167)
  
184         Case 9
186             VersionSpecificAction 25, nLanguage, , sTemp
188             SelectedSMSType = sTemp
  
190         Case 10
192             VersionSpecificAction 26, nLanguage, , sTemp
194             SelectedSMSType = sTemp
  
196         Case 11
198             VersionSpecificAction 27, nLanguage, , sTemp
200             SelectedSMSType = sTemp
    
202         Case Else
204             SelectedSMSType = ""
        End Select
  
        '<EhFooter>
        Exit Function

SelectedSMSType_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.SelectedSMSType " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Function TimeZoneFromListBoxSetting(nListIndex As Integer) As Integer
        '<EhHeader>
        On Error GoTo TimeZoneFromListBoxSetting_Err
        '</EhHeader>
100     TimeZoneFromListBoxSetting = nListIndex - 12
        '<EhFooter>
        Exit Function

TimeZoneFromListBoxSetting_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.TimeZoneFromListBoxSetting " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub UpdateJobList(nLanguage As Integer)
        '<EhHeader>
        On Error GoTo UpdateJobList_Err
        '</EhHeader>
        On Error GoTo ErrorTrap
        Dim lRecipient As Long
        Dim k As Long
        Dim dWork As Date
        Dim sRecipient As String
        Dim lRecipientMultiplier As Long
        Dim lCountMessages As Long
        Dim lCountRecipients As Long
        Dim sJobInfo As String

100     frmSMSMain.lstCurrentJob.Clear
102     Screen.MousePointer = vbHourglass

104     lCountRecipients = frmSMSMain.grdRecipients.Rows - 1
106     lRecipientMultiplier = 1

108     For lRecipient = 1 To lCountRecipients

110         sRecipient = frmSMSMain.grdRecipients.TextMatrix(lRecipient, 0) & " " & frmSMSMain.grdRecipients.TextMatrix(lRecipient, 1)   'Number of the Recipient

112         If gtDeferredDeliveryTimeSettings.bUseDeferredDeliveryTime = False Then
                'Do nothing
114             frmSMSMain.lstCurrentJob.AddItem SelectedSMSType(nLanguage) & " " & LoadLanguageSpecificString(nLanguage, 172) & " " & sRecipient
            Else

116             If gtDeferredDeliveryTimeSettings.bSingleSMS = True Then
118                 frmSMSMain.lstCurrentJob.AddItem gtDeferredDeliveryTimeSettings.sDeliveryDate & " / " & SelectedSMSType(nLanguage) & " " & LoadLanguageSpecificString(nLanguage, 172) & " " & sRecipient
                Else

120                 If IsDate(gtDeferredDeliveryTimeSettings.sStartingDate) Then
122                     dWork = CVDate(gtDeferredDeliveryTimeSettings.sStartingDate)
124                     lRecipientMultiplier = Val(gtDeferredDeliveryTimeSettings.sNumberOfMessages)

126                     For k = 1 To Val(gtDeferredDeliveryTimeSettings.sNumberOfMessages)
128                         frmSMSMain.lstCurrentJob.AddItem dWork & " / " & SelectedSMSType(nLanguage) & Str$(k) & " " & LoadLanguageSpecificString(nLanguage, 172) & " " & sRecipient

130                         Select Case gtDeferredDeliveryTimeSettings.nWaitingPeriod

                                Case 0 'Seconds
132                                 dWork = DateAdd("s", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

134                             Case 1 'Minutes
136                                 dWork = DateAdd("n", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

138                             Case 2 'Hours
140                                 dWork = DateAdd("h", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

142                             Case 3 'Days
144                                 dWork = DateAdd("d", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

146                             Case 4 'Weeks
148                                 dWork = DateAdd("ww", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)

150                             Case 5 'Months
152                                 dWork = DateAdd("m", gtDeferredDeliveryTimeSettings.sWaitingTime, dWork)
            
154                             Case Else
                                    'Unexecpted, do nothing, probably an application bug
            
                            End Select

                        Next

                    End If
                End If
            End If

        Next

156     Screen.MousePointer = vbDefault

158     frmSMSMain.lblCurrentRecipients = LoadLanguageSpecificString(nLanguage, 120) & " " & Trim(Str$(frmSMSMain.grdRecipients.Rows - 1))

160     lCountMessages = lCountRecipients * lRecipientMultiplier

162     If lRecipientMultiplier > 1 Then
164         sJobInfo = "Work List:" & " " & Trim(Str$(lCountRecipients)) & " " & LoadLanguageSpecificString(nLanguage, 91) & " / " & Trim(Str$(lCountMessages)) & " " & LoadLanguageSpecificString(nLanguage, 452)
            'sJobInfo = LoadLanguageSpecificString(nLanguage, 117) & " " & Trim(Str$(lCountRecipients)) & " " & LoadLanguageSpecificString(nLanguage, 91) & " / " & Trim(Str$(lCountMessages)) & " " & LoadLanguageSpecificString(nLanguage, 452)
        Else
166         sJobInfo = "Work List:" & " " & Trim(Str$(lCountRecipients)) & " " & LoadLanguageSpecificString(nLanguage, 91)
            'sJobInfo = LoadLanguageSpecificString(nLanguage, 117) & " " & Trim(Str$(lCountRecipients)) & " " & LoadLanguageSpecificString(nLanguage, 91)
        End If

168     frmSMSMain.fraCurrentJob = sJobInfo

170     gtJobInfo.lMessages = 0
172     gtJobInfo.lRecipients = 0
174     gtJobInfo.tDeferredDeliveryTimeSettings = gtDeferredDeliveryTimeSettings

176     VersionSpecificAction 57, nLanguage

178     UpdateJobInfoScreen nLanguage
        Exit Sub
ErrorTrap:
180     Screen.MousePointer = vbDefault
        '<EhFooter>
        Exit Sub

UpdateJobList_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.UpdateJobList " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
Function ValidityPeriodWithSpecificSettings(tValidityPeriodSettings As ValidityPeriodSettings, tDeferredDeliveryTimeSettings As DeferredDeliveryTimeSettings, bPeriodic As Boolean) As ValidityPeriod
        '<EhHeader>
        On Error GoTo ValidityPeriodWithSpecificSettings_Err
        '</EhHeader>
        Dim tValidityPeriod As ValidityPeriod

100     If bPeriodic = False Then
102         tValidityPeriod.bUserDefinedSettingsOnlyForSpecificSMSTypes = tValidityPeriodSettings.bSingleSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes
    
104         Select Case tValidityPeriodSettings.nSingleSMSValidityPeriodMode

                Case ValidityPeriodMode.UseSpecificSettingsAsLifeTime
106                 tValidityPeriod.lLifeTime = tValidityPeriodSettings.lSingleSMSSpecificSettingLifeTime
108                 tValidityPeriod.nLifeTimeUnit = tValidityPeriodSettings.nSingleSMSSpecificSettingLifeTimeUnit
110                 tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
        
112             Case ValidityPeriodMode.UseSingleshotAsLifeTime
114                 tValidityPeriod.lLifeTime = 0
116                 tValidityPeriod.nLifeTimeUnit = 0
118                 tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
   
120             Case ValidityPeriodMode.UseWaitingTimeAsLifeTime
                    'Do nothing, not possible
    
            End Select

        Else
122         tValidityPeriod.bUserDefinedSettingsOnlyForSpecificSMSTypes = tValidityPeriodSettings.bPeriodicSMSUseUserDefinedSettingsOnlyForSpecificSMSTypes
  
124         Select Case tValidityPeriodSettings.nPeriodicSMSValidityPeriodMode

                Case ValidityPeriodMode.UseSpecificSettingsAsLifeTime
126                 tValidityPeriod.lLifeTime = tValidityPeriodSettings.lPeriodicSMSSpecificSettingLifeTime
128                 tValidityPeriod.nLifeTimeUnit = tValidityPeriodSettings.nPeriodicSMSSpecificSettingLifeTimeUnit
130                 tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
        
132             Case ValidityPeriodMode.UseSingleshotAsLifeTime
134                 tValidityPeriod.lLifeTime = 0
136                 tValidityPeriod.nLifeTimeUnit = 0
138                 tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
   
140             Case ValidityPeriodMode.UseWaitingTimeAsLifeTime

142                 Select Case tDeferredDeliveryTimeSettings.nWaitingPeriod

                        Case 0 'Seconds
144                         tValidityPeriod.lLifeTime = Val(tDeferredDeliveryTimeSettings.sWaitingTime) / 60
146                         tValidityPeriod.nLifeTimeUnit = 0
148                         tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
            
150                     Case 1 'Minutes
152                         tValidityPeriod.nLifeTimeUnit = 0
154                         tValidityPeriod.lLifeTime = Val(tDeferredDeliveryTimeSettings.sWaitingTime)
156                         tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
            
158                     Case 2 'Hours
160                         tValidityPeriod.nLifeTimeUnit = 1
162                         tValidityPeriod.lLifeTime = Val(tDeferredDeliveryTimeSettings.sWaitingTime)
164                         tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
      
166                     Case 3 'Days
168                         tValidityPeriod.nLifeTimeUnit = 1
170                         tValidityPeriod.lLifeTime = 24
172                         tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
      
174                     Case 4 'Weeks
176                         tValidityPeriod.nLifeTimeUnit = 1
178                         tValidityPeriod.lLifeTime = 24
180                         tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
      
182                     Case 5 'Months
184                         tValidityPeriod.nLifeTimeUnit = 1
186                         tValidityPeriod.lLifeTime = 24
188                         tValidityPeriod = TrimValidityPeriodToCorrectValues(tValidityPeriod)
      
                    End Select
    
            End Select

        End If

190     ValidityPeriodWithSpecificSettings = tValidityPeriod
        '<EhFooter>
        Exit Function

ValidityPeriodWithSpecificSettings_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.Programs.ValidityPeriodWithSpecificSettings " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
