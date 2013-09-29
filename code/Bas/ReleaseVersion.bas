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
102             gsApplicationName = "OASIS SMS Messenger V " & App.major & "." & App.minor & "." & App.Revision
    
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
104             frmSMSMain.caption = gsApplicationName
106             frmSMSMain.mnuInfo001.caption = LoadLanguageSpecificString(nInput, 240)
108             frmSMSMain.mnuFile.caption = LoadLanguageSpecificString(nInput, 81)
110             frmSMSMain.mnuRegistration.caption = LoadLanguageSpecificString(nInput, 82)
112             frmSMSMain.mnuInfo003.caption = LoadLanguageSpecificString(nInput, 84)

114             VersionSpecificAction 37, , , sCaption
116             frmSMSMain.mnuInfo002.caption = sCaption

118             frmSMSMain.mnuInfo004.caption = LoadLanguageSpecificString(nInput, 86)
120             frmSMSMain.mnuInfo005.caption = LoadLanguageSpecificString(nInput, 87)
122             frmSMSMain.mnuInfo006.caption = LoadLanguageSpecificString(nInput, 88)

124             VersionSpecificAction 44, , , sApplicationPlaceHolder
126             sTemp = LoadLanguageSpecificString(nInput, 410)
128             sTemp = Replace(sTemp, gcPlaceHolder, sApplicationPlaceHolder)
130             frmSMSMain.mnuOpenDatabase.caption = sTemp
  
132             frmSMSMain.mnuOpenDatabaseWithAccess2000.caption = LoadLanguageSpecificString(nInput, 411)

134             VersionSpecificAction 44, , , sApplicationPlaceHolder
136             sTemp = LoadLanguageSpecificString(nInput, 412)
138             sTemp = Replace(sTemp, gcPlaceHolder, sApplicationPlaceHolder)
140             frmSMSMain.mnuImportLegacy.caption = sTemp

142             frmSMSMain.mnuExit.caption = LoadLanguageSpecificString(nInput, 89)
144             frmSMSMain.mnuSettings.caption = LoadLanguageSpecificString(nInput, 334)
146             frmSMSMain.mnuGeneralSettings.caption = LoadLanguageSpecificString(nInput, 115)
148             frmSMSMain.mnuSendlog.caption = LoadLanguageSpecificString(nInput, 113)
150             frmSMSMain.mnuJoblog.caption = LoadLanguageSpecificString(nInput, 589)

152             frmSMSMain.mnuLanguage(1).caption = LoadLanguageSpecificString(nInput, 3)
154             frmSMSMain.mnuLanguage(2).caption = LoadLanguageSpecificString(nInput, 4)

156             frmSMSMain.mnuBuyCredits.caption = LoadLanguageSpecificString(nInput, 487)
158             frmSMSMain.mnuShowCredits.caption = LoadLanguageSpecificString(nInput, 114)
160             frmSMSMain.mnuOriginators.caption = LoadLanguageSpecificString(nInput, 751)

162             frmSMSMain.mnuInfo009.caption = LoadLanguageSpecificString(nInput, 83)
164             frmSMSMain.mnuInfo007.caption = LoadLanguageSpecificString(nInput, 484)
166             frmSMSMain.mnuInfo008.caption = LoadLanguageSpecificString(nInput, 483)
  
168         Case 2

170             For i = 0 To 8

172                 If ControlIsLoaded(frmSMSMain.optSMSType(i)) = False Then
174                     Load frmSMSMain.optSMSType(i)
                    End If

176                 frmSMSMain.optSMSType(i).left = 120
178                 frmSMSMain.optSMSType(i).top = 250 + (i * 340)
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

200             frmSettings.chkSMSTypeEnabled(0).top = 700
202             frmSettings.chkSMSTypeEnabled(0).left = 120

204             frmSettings.chkSMSTypeEnabled(1).top = 1100
206             frmSettings.chkSMSTypeEnabled(1).left = 120

208             frmSettings.chkSMSTypeEnabled(2).top = 1500
210             frmSettings.chkSMSTypeEnabled(2).left = 120

212             frmSettings.chkSMSTypeEnabled(3).top = 1900
214             frmSettings.chkSMSTypeEnabled(3).left = 120

216             frmSettings.chkSMSTypeEnabled(4).top = 2300
218             frmSettings.chkSMSTypeEnabled(4).left = 120

220             frmSettings.chkSMSTypeEnabled(5).top = 2700
222             frmSettings.chkSMSTypeEnabled(5).left = 120
  
224             frmSettings.chkSMSTypeEnabled(6).Width = 5000
226             frmSettings.chkSMSTypeEnabled(6).top = 3100
228             frmSettings.chkSMSTypeEnabled(6).left = 120

230             frmSettings.chkSMSTypeEnabled(7).top = 3500
232             frmSettings.chkSMSTypeEnabled(7).left = 120
  
234             frmSettings.chkSMSTypeEnabled(8).top = 3900
236             frmSettings.chkSMSTypeEnabled(8).left = 120

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
300             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/home.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
302         Case 29
304             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/instruction/prices.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
306         Case 30
308             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/news/home.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
310         Case 31
312             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/supportednetworks.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)

314         Case 32
316             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/documentation/home.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)

318         Case 33
320             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/instruction/faq.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
322         Case 34
324             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/instruction/support.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)

326         Case 35
328             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/instruction/aboutus.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)

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
402             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/balance.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
404         Case 55
406             lRet = ShellExecute(Screen.ActiveForm.hwnd, "Open", "http://www.aspsms.com/registration.asp?REF=39", vbNullString, vbNullString, vbNormalFocus)
   
408         Case 56
410             VersionSpecificAction 23, gnLanguage, , sTemp
412             sTemp = sTemp & vbCrLf & vbCrLf & "Versioninfo: " & Trim(Str$(App.major)) & "." & Trim(Str$(App.minor)) & " Build: " & right("000" & Trim(Str$(App.Revision)), 4)

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


Public Sub DoXLSX(objExcel As Object, RS As ADODB.Recordset, sPath As String)
    Dim objWrkSht As Object 'Excel.Worksheet
    Dim iCol As Integer
    Dim recArray As Variant
    Dim recCount As Long
    Dim iRow As Integer
    
    objExcel.Workbooks.Add

    Set objWrkSht = objExcel.Worksheets(1)
    objWrkSht.Activate
    objWrkSht.Name = "OASIS Export"
    objWrkSht.Range(objWrkSht.Cells(1, 1), objWrkSht.Cells(1, RS.Fields.Count)).Font.Bold = True
    
    objExcel.Visible = True
    objExcel.UserControl = True
        
    For iCol = 1 To RS.Fields.Count
        objWrkSht.Cells(1, iCol).value = RS.Fields(iCol - 1).Name
    Next
    
    On Error Resume Next

    If Val(Mid(objExcel.Version, 1, InStr(1, objExcel.Version, ".") - 1)) > 8 Then
        objWrkSht.Range("A2").CopyFromRecordset RS
    Else
        
        recArray = RS.GetRows
        recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array
        

        For iCol = 0 To RS.Fields.Count - 1
            For iRow = 0 To recCount - 1
                ' Take care of Date fields
                If IsDate(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = Format(recArray(iCol, iRow))
                ' Take care of OLE object fields or array fields
                ElseIf IsArray(recArray(iCol, iRow)) Then
                    recArray(iCol, iRow) = "Array Field"
                End If
            Next iRow 'next record
        Next iCol 'next field
            
        objWrkSht.Cells(2, 1).Resize(recCount, RS.Fields.Count).value = TransposeDim(recArray)
    End If
    
    'objWrkSht.SaveA sPath
    objWrkSht.SaveAs sPath, , , , False
    objExcel.Selection.CurrentRegion.Columns.AutoFit
    objExcel.Selection.CurrentRegion.Rows.AutoFit
    objWrkSht.Close
    Set objWrkSht = Nothing
'    Set objExcel = Nothing
'
'    ShellExecute 0, "open", sPath, "", "", 0
    
End Sub

Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X
    
    TransposeDim = tempArray


End Function

Public Function ExcelVers(oEXLS As Object) As String
    If oEXLS.Version = "12.0" Then
        ExcelVers = "2007"
    ElseIf oEXLS.Version = "11.0" Then
        ExcelVers = "2003"
    ElseIf oEXLS.Version = "10.0" Then
        ExcelVers = "2002"
    ElseIf oEXLS.Version = "9.0" Then
        ExcelVers = "2000"
    ElseIf oEXLS.Version = "8.0" Then
        ExcelVers = "97"
    ElseIf oEXLS.Version = "7.0" Then
        ExcelVers = "95"
    Else
        ExcelVers = "OK" ' Should be ok for all OASIS versions....
    End If
End Function
