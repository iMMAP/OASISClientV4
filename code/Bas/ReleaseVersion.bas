Attribute VB_Name = "ReleaseVersion"
Option Explicit

Type MenuControl
    bTextSMSEnabled As Boolean
    bOperatorLogoEnabled As Boolean
    bGroupLogoEnabled As Boolean
    bRingtoneEnabled As Boolean
    bPictureMessageEnabled As Boolean
    bVCardEnabled As Boolean
    bUnicodeEnabled As Boolean
    bWAPPushSMSEnabled As Boolean
    bBinaryDataEnabled As Boolean
End Type

Sub VersionSpecificAction(nCode As Integer, _
                          Optional nInput As Integer, _
                          Optional lInput As Long, _
                          Optional sInput As String, _
                          Optional bInput As Boolean)
        '<EhHeader>
        On Error GoTo VersionSpecificAction_Err
        '</EhHeader>
        Dim i As Integer
        Dim sTemp As String
        Dim sCaption As String
        Dim sApplicationPlaceHolder As String
        Dim lRet As Long

100     Select Case nCode

            Case 1
102             gsApplicationName = "OASIS SMS Messenger V " & App.Major & "." & App.Minor & "." & App.Revision
    
    '            frmSMSMain.mnuFile.Visible = True
    '            frmSMSMain.mnuOpenDatabase.Visible = True
    '            frmSMSMain.mnuSeparator01.Visible = True
    '            frmSMSMain.mnuImportTextfiles.Visible = True
    '            frmSMSMain.mnuImportLegacy.Visible = True
    '            frmSMSMain.mnuSeparator02.Visible = True
    '            frmSMSMain.mnuExit.Visible = True
    '
    '            frmSMSMain.mnuAccount.Visible = True
    '            frmSMSMain.mnuRegistration.Visible = True
    '            frmSMSMain.mnuSeparator03.Visible = True
    '            frmSMSMain.mnuBuyCredits.Visible = True
    '            frmSMSMain.mnuShowCredits.Visible = True
    '            frmSMSMain.mnuSeparator04.Visible = True
    '            frmSMSMain.mnuJoblog.Visible = True
    '            frmSMSMain.mnuSendlog.Visible = True
    '
    '            frmSMSMain.mnuSettings.Visible = True
    '            frmSMSMain.mnuLanguage(1).Visible = True
    '            frmSMSMain.mnuLanguage(2).Visible = True
    '            frmSMSMain.mnuSeparator05.Visible = True
    '            frmSMSMain.mnuGeneralSettings.Visible = True
    '
    '            frmSMSMain.mnuInfo.Visible = True
    '            frmSMSMain.mnuInfo001.Visible = True
    '            frmSMSMain.mnuSeparator06.Visible = True
    '            frmSMSMain.mnuInfo002.Visible = True
    '            frmSMSMain.mnuInfo003.Visible = True
    '            frmSMSMain.mnuInfo004.Visible = True
    '            frmSMSMain.mnuInfo005.Visible = True
    '            frmSMSMain.mnuInfo006.Visible = True
    '            frmSMSMain.mnuInfo007.Visible = True
    '            frmSMSMain.mnuInfo008.Visible = True
    '            frmSMSMain.mnuSeparator07.Visible = True
    '            frmSMSMain.mnuInfo009.Visible = True
    '
104             frmSMSMain.Caption = gsApplicationName
106             frmSMSMain.mnuInfo001.Caption = LoadLanguageSpecificString(nInput, 240)
108             frmSMSMain.mnuFile.Caption = LoadLanguageSpecificString(nInput, 81)
110             frmSMSMain.mnuRegistration.Caption = LoadLanguageSpecificString(nInput, 82)
112             frmSMSMain.mnuInfo003.Caption = LoadLanguageSpecificString(nInput, 84)

114             VersionSpecificAction 37, , , sCaption
116             frmSMSMain.mnuInfo002.Caption = sCaption

118             frmSMSMain.mnuInfo004.Caption = LoadLanguageSpecificString(nInput, 86)
120             frmSMSMain.mnuInfo005.Caption = LoadLanguageSpecificString(nInput, 87)
122             frmSMSMain.mnuInfo006.Caption = LoadLanguageSpecificString(nInput, 88)

124             VersionSpecificAction 44, , , sApplicationPlaceHolder
126             sTemp = LoadLanguageSpecificString(nInput, 410)
128             sTemp = Replace(sTemp, gcPlaceHolder, sApplicationPlaceHolder)
130             frmSMSMain.mnuOpenDatabase.Caption = sTemp
  
132             frmSMSMain.mnuOpenDatabaseWithAccess2000.Caption = LoadLanguageSpecificString(nInput, 411)

134             VersionSpecificAction 44, , , sApplicationPlaceHolder
136             sTemp = LoadLanguageSpecificString(nInput, 412)
138             sTemp = Replace(sTemp, gcPlaceHolder, sApplicationPlaceHolder)
140             frmSMSMain.mnuImportLegacy.Caption = sTemp

142             frmSMSMain.mnuExit.Caption = LoadLanguageSpecificString(nInput, 89)
144             frmSMSMain.mnuSettings.Caption = LoadLanguageSpecificString(nInput, 334)
146             frmSMSMain.mnuGeneralSettings.Caption = LoadLanguageSpecificString(nInput, 115)
148             frmSMSMain.mnuSendlog.Caption = LoadLanguageSpecificString(nInput, 113)
150             frmSMSMain.mnuJoblog.Caption = LoadLanguageSpecificString(nInput, 589)

152             frmSMSMain.mnuLanguage(1).Caption = LoadLanguageSpecificString(nInput, 3)
154             frmSMSMain.mnuLanguage(2).Caption = LoadLanguageSpecificString(nInput, 4)

156             frmSMSMain.mnuBuyCredits.Caption = LoadLanguageSpecificString(nInput, 487)
158             frmSMSMain.mnuShowCredits.Caption = LoadLanguageSpecificString(nInput, 114)
160             frmSMSMain.mnuOriginators.Caption = LoadLanguageSpecificString(nInput, 751)

162             frmSMSMain.mnuInfo009.Caption = LoadLanguageSpecificString(nInput, 83)
164             frmSMSMain.mnuInfo007.Caption = LoadLanguageSpecificString(nInput, 484)
166             frmSMSMain.mnuInfo008.Caption = LoadLanguageSpecificString(nInput, 483)
  
168         Case 2

170             For i = 0 To 8

172                 If ControlIsLoaded(frmSMSMain.optSMSType(i)) = False Then
174                     Load frmSMSMain.optSMSType(i)
                    End If

176                 frmSMSMain.optSMSType(i).Left = 120
178                 frmSMSMain.optSMSType(i).Top = 250 + (i * 340)
180                 frmSMSMain.optSMSType(i).Visible = True
                Next
  
182         Case 3
                'Do nothing
    
184         Case 4
                'Do nothing
  
186         Case 5
                'Do nothing
  
188         Case 6
     
190         Case 7
                'Do nothing
   
192         Case 8

194             For i = 0 To 8

196                 If ControlIsLoaded(frmSettings.chkSMSTypeEnabled(i)) = False Then
198                     Load frmSettings.chkSMSTypeEnabled(i)
                    End If

                Next

200             frmSettings.chkSMSTypeEnabled(0).Top = 700
202             frmSettings.chkSMSTypeEnabled(0).Left = 120

204             frmSettings.chkSMSTypeEnabled(1).Top = 1100
206             frmSettings.chkSMSTypeEnabled(1).Left = 120

208             frmSettings.chkSMSTypeEnabled(2).Top = 1500
210             frmSettings.chkSMSTypeEnabled(2).Left = 120

212             frmSettings.chkSMSTypeEnabled(3).Top = 1900
214             frmSettings.chkSMSTypeEnabled(3).Left = 120

216             frmSettings.chkSMSTypeEnabled(4).Top = 2300
218             frmSettings.chkSMSTypeEnabled(4).Left = 120

220             frmSettings.chkSMSTypeEnabled(5).Top = 2700
222             frmSettings.chkSMSTypeEnabled(5).Left = 120
  
224             frmSettings.chkSMSTypeEnabled(6).Width = 5000
226             frmSettings.chkSMSTypeEnabled(6).Top = 3100
228             frmSettings.chkSMSTypeEnabled(6).Left = 120

230             frmSettings.chkSMSTypeEnabled(7).Top = 3500
232             frmSettings.chkSMSTypeEnabled(7).Left = 120
  
234             frmSettings.chkSMSTypeEnabled(8).Top = 3900
236             frmSettings.chkSMSTypeEnabled(8).Left = 120

238         Case 9
                'Do nothing
  
240         Case 10
                'Do nothing
  
242         Case 11
                'Do nothing
  
244         Case 12
                'Do nothing

246         Case 13
                'Do nothing

248         Case 14
250             bInput = False
  
252         Case 15
254             bInput = False
  
256         Case 16
258             bInput = False
  
260         Case 17
262             sInput = "DUMMYNOTUSED"
  
264         Case 18
266             sInput = "DUMMYNOTUSED"
  
268         Case 19
270             sInput = "DUMMYNOTUSED"
  
272         Case 20
274             sInput = "2"
  
276         Case 21
278             sInput = "2"
  
280         Case 22
282             sInput = "2"
  
284         Case 23
286             VersionSpecificAction 43, , , sTemp
288             sInput = LoadLanguageSpecificString(nInput, 241) & vbCrLf & LoadLanguageSpecificString(nInput, 242) & vbCrLf & LoadLanguageSpecificString(nInput, 243) & vbCrLf & LoadLanguageSpecificString(nInput, 244) & vbCrLf & LoadLanguageSpecificString(nInput, 245) & vbCrLf & LoadLanguageSpecificString(nInput, 246) & vbCrLf & LoadLanguageSpecificString(nInput, 481) & vbCrLf & LoadLanguageSpecificString(nInput, 479) & vbCrLf & LoadLanguageSpecificString(nInput, 247) & vbCrLf & LoadLanguageSpecificString(nInput, 477) & vbCrLf & LoadLanguageSpecificString(nInput, 250) & vbCrLf & LoadLanguageSpecificString(nInput, 490) & vbCrLf & vbCrLf & LoadLanguageSpecificString(nInput, 251) & vbCrLf & LoadLanguageSpecificString(nInput, 252) & " " & sTemp & " " & LoadLanguageSpecificString(nInput, 253)

290         Case 24
                'Do nothing
   
292         Case 25
                'Do nothing
   
294         Case 26
                'Do nothing
   
296         Case 27
                'Do nothing
   
298         Case 28
300             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/home.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
302         Case 29
304             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/instruction/prices.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
306         Case 30
308             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/news/home.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
310         Case 31
312             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/supportednetworks.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)

314         Case 32
316             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/documentation/home.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)

318         Case 33
320             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/instruction/faq.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
322         Case 34
324             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/instruction/support.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)

326         Case 35
328             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/instruction/aboutus.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)

330         Case 36
332             nInput = 13
   
334         Case 37
336             sInput = "aspsms.com Website"
   
338         Case 38
340             sInput = "SMSBlaster 1.2"
   
342         Case 39
344             sInput = "/img/RandomLogos/Logos.txt"
   
346         Case 40
348             sInput = "www.aspsms.com"
   
350         Case 41
352             sInput = "/img/RandomLogos/"
   
354         Case 42

356             If gsPathAndDatabaseName = "" Then
358                 gsPathAndDatabaseName = App.Path & "\OASISSMS.mdb"
                End If
   
360         Case 43
362             sInput = "www.aspsms.com"
   
364         Case 44
366             sInput = "SMS Blaster"
   
368         Case 45
370             sInput = "650"
   
372         Case 46
                'Do nothing
   
374         Case 47
                'Do nothing
   
376         Case 48
378             sInput = LoadLanguageSpecificString(nInput, 531)
  
380         Case 49
382             sInput = LoadLanguageSpecificString(nInput, 532)
  
384         Case 50
386             sInput = LoadLanguageSpecificString(nInput, 533)
  
388         Case 51
390             sInput = LoadLanguageSpecificString(nInput, 534)
  
392         Case 52
394             sInput = LoadLanguageSpecificString(nInput, 535)
  
396         Case 53
398             sInput = LoadLanguageSpecificString(nInput, 536)
  
400         Case 54
402             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/balance.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
404         Case 55
406             lRet = ShellExecute(Screen.ActiveForm.hWnd, "Open", "http://www.aspsms.com/registration.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
408         Case 56
410             VersionSpecificAction 23, gnLanguage, , sTemp
412             sTemp = sTemp & vbCrLf & vbCrLf & "Versioninfo: " & Trim(Str$(App.Major)) & "." & Trim(Str$(App.Minor)) & " Build: " & Right("000" & Trim(Str$(App.Revision)), 4)

414             MsgBox sTemp, vbInformation, gsApplicationName
   
416         Case 57
                Dim tValidityPeriod As ValidityPeriod
  
418             tValidityPeriod = ActualValidityPeriod()

420             If tValidityPeriod.bStandardPeriod = True Then
                    'Do nothing
                Else

422                 Select Case tValidityPeriod.nLifeTimeUnit

                        Case 0

424                         If tValidityPeriod.lLifeTime = 1 Then
426                             sTemp = "(" & LoadLanguageSpecificString(nInput, 696) & " " & Trim(Str$(tValidityPeriod.lLifeTime)) & " " & LoadLanguageSpecificString(nInput, 699) & ")"
                            Else
428                             sTemp = "(" & LoadLanguageSpecificString(nInput, 696) & " " & Trim(Str$(tValidityPeriod.lLifeTime)) & " " & LoadLanguageSpecificString(nInput, 700) & ")"
                            End If
      
430                     Case 1

432                         If tValidityPeriod.lLifeTime = 1 Then
434                             sTemp = "(" & LoadLanguageSpecificString(nInput, 696) & " " & Trim(Str$(tValidityPeriod.lLifeTime)) & " " & LoadLanguageSpecificString(nInput, 701) & ")"
                            Else
436                             sTemp = "(" & LoadLanguageSpecificString(nInput, 696) & " " & Trim(Str$(tValidityPeriod.lLifeTime)) & " " & LoadLanguageSpecificString(nInput, 702) & ")"
                            End If

                    End Select

                End If
    
438         Case 58
  
440         Case 59
     
442         Case 60
444             lInput = 39
  
446         Case Else
448             MsgBox "Case else " & nCode
  
        End Select

        '<EhFooter>
        Exit Sub

VersionSpecificAction_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASIS_SMS_Messenger.ReleaseVersion.VersionSpecificAction " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

