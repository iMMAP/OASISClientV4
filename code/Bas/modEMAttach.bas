Attribute VB_Name = "modEMAttach"
Option Explicit

Public Function SendOutLookMail(EM_TO, Em_CC, EM_BCC, EM_Subject, EM_Body, EM_Attachment As String, EM_Attachment1 As String, Display As Boolean, Optional EM_HTMLBody As String)
        '<EhHeader>
        On Error GoTo SendOutLookMail_Err
        '</EhHeader>

        Dim objOA As Object 'Outlook.Application
        Dim objMI As Object 'Outlook.MailItem
        Dim obgAtt As Object 'Outlook.Attachments
        Dim olMailItem As Object
        Dim olNs As Object
    
        'Set objOA = New Outlook.Application
        On Error Resume Next
    
        Set objOA = GetObject(, "Outlook.Application")
    
        If objOA Is Nothing Then
100         Set objOA = CreateObject("Outlook.Application")   'New Outlook.Application
        End If
    
102     If objOA Is Nothing Then Exit Function
        
        'objOA.application.Visible = True
        
           Set olNs = objOA.GetNamespace("MAPI")
           olNs.Logon

        
104     Set objMI = objOA.CreateItem(0) 'olMailItem)
    
106     If EM_TO <> "" Then objMI.To = EM_TO
108     If Em_CC <> "" Then objMI.CC = Em_CC
110     If EM_BCC <> "" Then objMI.BCC = EM_BCC
112     If EM_Subject <> "" Then objMI.Subject = EM_Subject
114     If EM_Body <> "" Then objMI.Body = EM_Body
116     If EM_Attachment <> "" Then objMI.Attachments.Add EM_Attachment, 1, , EM_Attachment
118     If EM_Attachment1 <> "" Then objMI.Attachments.Add EM_Attachment1, 1, , EM_Attachment1

        If EM_HTMLBody <> "" Then objMI.HTMLBody = EM_HTMLBody

120     If Display Then
122         objMI.Display
        Else
124         objMI.Send
        End If

126     Set objOA = Nothing
128     Set objMI = Nothing
        '<EhFooter>
        Exit Function

SendOutLookMail_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.modEMAttach.SendOutLookMail " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function SendMail_Extended(EM_TO, Em_CC, EM_BCC, EM_Subject, EM_Body, EM_Attachment As String, Display As Boolean)

'    Dim objOA As Outlook.Application
'    Dim objMI As Outlook.MailItem
'    Dim obgAtt As Outlook.Attachments
'    Set objOA = New Outlook.Application
'    Set objMI = objOA.CreateItem(olMailItem)
'    If EM_TO <> "" Then objMI.To = EM_TO
'    If Em_CC <> "" Then objMI.CC = Em_CC
'    If EM_BCC <> "" Then objMI.BCC = EM_BCC
'    If EM_Subject <> "" Then objMI.Subject = EM_Subject
'    If EM_Body <> "" Then objMI.Body = EM_Body
'    If EM_Attachment <> "" Then objMI.Attachments.Add EM_Attachment, 1, , EM_Attachment
'
'
'    If Display Then
'        objMI.Display
'    Else
'        objMI.Send
'    End If
'
'    Set objOA = Nothing
'    Set objMI = Nothing
End Function


Public Function CreateNotesAttachment(DBName As String, DocId As String, AttachmentPath As String) As Boolean

'    Dim session As Object
'    Dim db As Object
'    Dim doc As Object
'    Dim Ob As Object
'    Dim rtitem As Object 'NotesRichTextItem
'    Set session = CreateObject("Notes.NotesSession")
'    Set db = session.GetDatabase("", DBName)
'    Set doc = db.GetDocumentByUNID(DocId)
'    Set rtitem = doc.CreateRichTextItem("Attachment")
'    Set Ob = rtitem.EmbedObject _
'    (EMBED_ATTACHMENT, "", AttachmentPath)
'    Call doc.Save(True, False)
'    Dim ws As Object
'    Set ws = CreateObject("Notes.NotesUIWorkspace")
'    Call ws.EditDocument(True, doc)
    
End Function

'
'Function SendNotesMail(Message As String, _
'
'
'
'Subject As String, _
'
'    Recipients() As String, _
'    Logo As Long, _
'    Attachment As String, _
'    MailServer As String, _
'    MailDB As String, _
'    MailPassword As String) As Boolean
'    On Error GoTo Close_Error
'    Const ShowProperties = True
'    Dim lnSession As NotesSession
'    Dim lnDatabase As NotesDatabase
'    Dim lnDocument As Object, lnRTStyle As Object
'    Dim lnRTItem As Object, lnAttachment As Object
'    Dim lnLogo As Long
'    SendNotesMail = False
'    Set lnSession = CreateObject("Lotus.NotesSession")
'    Call lnSession.Initialize(MailPassword)
'    Set lnRTStyle = lnSession.CreateRichTextStyle
'    Set lnDatabase = lnSession.GetDatabase(MailServer, MailDB)
'
'
'    If Not lnDatabase.IsOpen Then
'        lnDatabase.Open
'
'
'        If Not lnDatabase.IsOpen Then
'            MsgBox "Can't open mail file: " & lnDatabase.Server & " " & _
'            lnDatabase.FilePath, vbCritical
'
'
'            DoEvents
'                Exit Function
'            End If
'
'        End If
'
'
'
'        If ShowProperties Then
'            MsgBox "Title: " & lnDatabase.Title & Chr(10) _
'            & "File name: " & lnDatabase.Filename & Chr(10) _
'            & "Path name: " & lnDatabase.FilePath & Chr(10) _
'            & "Replica ID: " & lnDatabase.ReplicaId & Chr(10) _
'            & "Size: " & lnDatabase.Size & Chr(10) _
'            & "Created: " & Chr(10) _
'            & "Last modified: " & lnDatabase.LastModified
'            MsgBox "Current access level: " & lnDatabase.CurrentAccessLevel & Chr(10) _
'            & "Percent used: " & lnDatabase.PercentUsed & Chr(10) _
'            & "Server name: " & lnDatabase.Server & Chr(10) _
'            & "Size limit: " & lnDatabase.SizeQuota
'        End If
'
'
'
'        DoEvents
'            Set lnDocument = lnDatabase.CreateDocument
'            Set lnRTItem = lnDocument.CreateRichTextItem("Body")
'
'
'            If Attachment <> "" Then
'                Set lnAttachment = lnRTItem.EMBEDOBJECT(1454, "", Attachment, "Sample")
'            End If
'
'            lnRTStyle.NotesFont = 4 'Courier
'            lnRTStyle.NotesColor = 2
'            Call lnRTItem.AppendStyle(lnRTStyle)
'            Call lnRTItem.AppendText(Message)
'            'logo values are between 0 and 31
'            lnLogo = Logo
'
'
'            If lnLogo < 0 Or lnLogo > 31 Then
'                lnLogo = 0
'            End If
'
'
'
'            With lnDocument
'                .ReplaceItemValue "SendTo", Recipients
'                .ReplaceItemValue "Subject", Subject
'                .ReplaceItemValue "Logo", "StdNotesLtr" & Trim$(str$(lnLogo))
'                .Send False
'            End With
'
'            Set lnRTItem = Nothing
'            Set lnRTStyle = Nothing
'            Set lnDocument = Nothing
'            Set lnDatabase = Nothing
'            Set lnSession = Nothing
'            SendNotesMail = True
'            Exit Function
'Close_Error:
'            MsgBox "Error: " & Err & " " & Err.Description & " In SendNotesMail.", vbCritical
'
'
'            DoEvents
'            End Function
'
'
'
'Sub TestNotesMail()
'
'    '
'    'Mail Server path and Mail Database file
'    '     can be found in you Notes.Ini file
'    ' under the variable names under Locatio
'    '     n and MailFile respectively
'    '
'    'Change the number of targets for the in
'    '     tended recipients
'    '
'    On Error GoTo Close_Error
'    Const MailServer = "YourMailRoot/YourMailRoute/YourServiceMailDir"
'    Const MailDB = "se-mail\username.nsf"
'    Const MailPassword = "password"
'    Const MsgText = "This is a test of the notes mail subroutine." & vbCrLf & vbCrLf & _
'    "Here is some more text." & vbCrLf & vbCrLf
'    Const MsgSubject = "Test Lotus Notes Mail Using COM Model With Attachment "
'    Const Letterhead = 0
'    Const Attachment = "C:\Documents and Settings\Username\Test.jpg"
'    Dim Targets() As String
'    Dim Response As Boolean
'    Screen.MousePointer = vbHourglass
'    ReDim Targets(1)
'    Targets(0) = ("you@yourservice.com")
'    Targets(1) = ("me@myservice.com")
'    Response = SendNotesMail(MsgText, MsgSubject, Targets, Letterhead, _
'    Attachment, MailServer, MailDB, MailPassword)
'    Screen.MousePointer = vbNormal
'
'
'    If Response Then
'        MsgBox "Mail sent successfully!", vbExclamation
'    Else
'        MsgBox "Mail problems", vbCritical
'    End If
'
'    Exit Sub
'Close_Error:
'    MsgBox "Error: " & Err & " " & Err.Description & " In TestNotesMail.", vbCritical
'
'
'    DoEvents
'    End Sub
'
'Option Explicit
'Option Base 1
'
'
'Sub Mailing()
'
'    Dim oSess 'As MAPI.Session
'    Dim oMsg ' As MAPI.Message
'    Dim oAttach ' As MAPI.attachment
'    Dim oRec
'    Dim FICHIER(1)
'    Dim REPERTOIRE
'    Dim TEMP
'    Dim Destinataire(1)
'    Dim I
'    FICHIER(1) = "File To send"
'    Destinataire(1) = "Recipients" 'Must be separated by " ; " (space ; space)
'    'Create an object of Session
'    Set oSess = CreateObject("Mapi.Session")
'    'Logon to the Session
'    oSess.Logon 'SETTINGS
'    Dim Start
'    Dim POS
'    'create a message and fill in its proper
'    '     ties
'    Set oMsg = oSess.Outbox.Messages.Add
'    oMsg.Subject = "Subject"
'    POS = 1
'    Start = 1
'
'
'    Do While POS <> 0
'        POS = InStr(Start, Destinataire(I), ";") 'To see If there are multiple recipients
'
'
'        If POS > 0 Then
'            TEMP = Mid(Destinataire(I), Start, POS - Start)
'            Set oRec = oMsg.Recipients.Add
'            oRec.Name = TEMP
'            oRec.Type = 1
'            oRec.Resolve
'        End If
'
'        Start = POS + 1
'    Loop
'
'    TEMP = "Body Text"
'
'    oMsg.Text = TEMP
'    Set oAttach = oMsg.Attachments.Add
'    oAttach.Name = "Name To be diplayed In Outlook"
'    oAttach.Position = 1
'    oAttach.Source = FICHIER(I)
'    oMsg.UpDate
'    oMsg.Send
'    'Clear all the objects before exiting th
'    '     e procedure
'    Set oRec = Nothing
'    Set oAttach = Nothing
'    Set oMsg = Nothing
'    oSess.Logoff
'End Sub
'
'wORKS With W2000 AND OUTLOOK2000, DrWATSON COMES WHEN YOU STOP THE PROGRAM DURING EXECUTION, If YOU PAUSE IT ALL WORKS CORRECTLY.
'HAVE FUN.



'**************************************
' Name: Send mail with Attachments using
'     Lotus Notes
' Description:apidude posted this last w
'     eek, I've added the ability to send atta
'     chments as well and resubmitted it. Pers
'     onally, I've been looking for some code
'     like this for a long, long time... Thank
'     s Apidude....
'The idea application For this is To build a bulk email program/Access DB that allows bulk email to be sent With Each one personalised or carrying information specific to an individual. ie. Sending out customer statements by email, etc...
' By: Peter Cawdron
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=32857&lngWId=1'for details.'**************************************

'**************************************
' Name: Use Lotus Notes to send email
' Description:Creates a Lotus Notes sess
'
' ion and use it to send an email
' By: apidude
' attachments added by pcawdron
'
' Inputs:strMessage: The message
'strSubject: the subject
'strSendTo: the recipient 's email addre
'     ss
'lngLogo:Specifies the letter head To us
'     e (Lotus Notes specific)
'
' Assumes:The Font & Color values for th
'
' e NotesRichTextItem class I'm not too
'     su
' re of because I don't have the DevKit
'     or
' the headers
'
'This code is copyrighted and has' limit
'     ed warranties.Please see http://w
' ww.Planet-Source-Code.com/xq/ASP/txtCo
'     de
' Id.32603/lngWId.1/qx/vb/scripts/ShowCo
'     de
' .htm'for details.'********************
'     ******************

'
'Function SendNotesMail(strMessage As String, _
'
'    strSubject As String, _
'    strSendTo As String, _
'    lngLogo As Long, strAttachment As String)
'    On Error GoTo NotesMail_Err
'    Dim lnSession As Object
'    Dim lnDatabase As Object
'    Dim lnDocument As Object
'    Dim lnRTStyle As Object
'    Dim lRTItem As Object
'    Dim lnATTACHMENT As Object
'    Dim sMessage As String
'    Dim lLogo As Long
'    ''start a notes session...
'    Set lnSession = CreateObject("Notes.Notessession")
'    ''create a new style object to control t
'    '
'    ' he appearance of the message
'    Set lnRTStyle = lnSession.CreateRichTextStyle
'    ''get the current database...
'    Set lnDatabase = lnSession.GetDatabase("", "")
'    lnDatabase.OpenMail
'    ''create a new document
'    Set lnDocument = lnDatabase.CreateDocument
'    ''create a new NotesRichTextItem object
'    ' in which we can store,
'    ''and format the main message body in Ri
'    '
'    ' chText format
'    Set lnRTItem = lnDocument.CreateRichTextItem("Body")
'
'
'    If strAttachment <> "" Then
'        Set lnATTACHMENT = lnRTItem.EMBEDOBJECT _
'        (1454, "", strAttachment, "Sample")
'    End If
'
'    sMessage = "Mail sent: " & date & " " & Time & vbCrLf & vbCrLf & _
'    strMessage
'    ''format the message
'    lnRTStyle.NotesFont = 4 ''Courier
'    lnRTStyle.Bold = True
'    lnRTStyle.NotesColor = 2 ''red
'    Call lnRTItem.AppendStyle(lnRTStyle)
'    Call lnRTItem.AppendText(sMessage)
'    'Call lnRTItem.AddNewLine(1)
'    ''logo values are between 0 and 31
'    lLogo = lngLogo
'
'
'    If lLogo < 0 Or lLogo > 31 Then
'        lLogo = 0
'    End If
'
'    ''replace some of the fields that we nee
'    '
'    ' d...
'
'
'    With lnDocument
'        ''who we want to send to...
'        ''recipient
'        .ReplaceItemValue "SendTo", strSendTo
'        ''subject
'        .ReplaceItemValue "Subject", strSubject
'        ''body - non RichText
'        '.ReplaceItemValue "Body", "The body of
'        ' the message!"
'        ''set the logo! (letter head)
'        .ReplaceItemValue "Logo", "StdNotesLtr" & Trim$(str$(lLogo))
'        ''send the message
'        .Send False
'    End With
'
'    Set lRTItem = Nothing
'    Set lnRTStyle = Nothing
'    Set lnDocument = Nothing
'    Set lnDatabase = Nothing
'    Set lnSession = Nothing
'    MsgBox "Mail was sent!", vbInformation, _
'    strSendTo
'    Exit Function
'NotesMail_Err:
'    MsgBox Err.Description, _
'    vbExclamation, _
'    "Send mail error! (" & Trim$(str$(Err)) & ")"
'End Function
'
'
'
'Function Test_note()
'
'    SendNotesMail "Hello! This is an email message! With an attachment", _
'    "Test Lotus Notes Email - Attachment test", _
'    "youraddress@work", 0, "C:\autoexec.bat"
'End Function


