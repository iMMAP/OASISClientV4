Attribute VB_Name = "general"

Public m_PTUID As Long

Public Declare Function Ellipse _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal X1 As Long, _
                            ByVal Y1 As Long, _
                            ByVal X2 As Long, _
                            ByVal Y2 As Long) As Long
Public Declare Function SetROP2 _
               Lib "gdi32" (ByVal hdc As Long, _
                            ByVal nDrawMode As Long) As Long

Public oldPos As New TatukGIS_XDK10.xPoint
Public oldRadius As Integer

'Public Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

Public Declare Function InvalidateRect _
               Lib "user32" (ByVal hwnd As Long, _
                             lpRect As RECT, _
                             ByVal bErase As Long) As Long
Public Declare Function SendMessage _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hwnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     lParam As Any) As Long
Public Declare Function SetWindowPos _
               Lib "user32" (ByVal hwnd As Long, _
                             ByVal hWndInsertAfter As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal cx As Long, _
                             ByVal cy As Long, _
                             ByVal wFlags As Long) As Long
Public Const WM_SETREDRAW = &HB
Public Const SWP_Invalidate = &H27

Public Const ERR_BAD_PARAMETER = "Array parameter required"
Public Const ERR_BAD_TYPE = "Invalid Type"
Public Const ERR_BP_NUMBER = 20000
Public Const ERR_BT_NUMBER = 20001
Public Const ERR_INVALID_ELEMENT = 20000
Public Const ERR_INVALID_ELEMENT_MSG = "Invalid List Element"
Public Const ERR_NONUNIQUE_KEYS = 20000
Public Const ERR_INVALID_LIST = 20001
Public Const ERR_NONUNIQUE_KEYS_MSG = "Keys list does not contain unique values"
Public Const ERR_INVALID_LIST_MSG = "Expected: Array or Collection"

'Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'**************************************
' Name: Check for string existance, Add
'     to either Listbox or Combobox
' Inputs:Input: Addstring
'Text to Find BoxType
'Combobox or ListBox AddIt
'Optional BooleanCompareType
'vbCompareType Enum
'
' Returns:Boolean
'

'For the Listbox
Public Const LB_ADDSTRING = &H180
Public Const LB_FINDSTRINGEXACT = &H1A2
Public Const LB_FINDSTRING = &H18F
'For the ComboBox
Public Const CB_ADDSTRING = &H143
Public Const CB_FINDSTRING = &H14C
Public Const CB_FINDSTRINGEXACT = &H158
Public Const CB_SETITEMDATA = &H151

Private Const CB_SETDROPPEDWIDTH = &H160
Private Const CB_GETDROPPEDWIDTH = &H15F
Private Const DT_CALCRECT = &H400

Private Declare Function SendMessageLong _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

'Private Declare Function DrawText Lib "user32" Alias '    "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, '    ByVal nCount As Long, lpRect As RECT, ByVal wFormat ''    As Long) As Long

Public Enum MsgBox_Flags 'Define for message box form
    '########### Button Combinations ###############
    vbOKOnly = 0 '          &H0&  OK button only (default)
    vbOKCancel = 1 '        &H1&  OK and Cancel buttons
    vbAbortRetryIgnore = 2 '&H2&  Abort, Retry, and Ignore buttons
    vbYesNoCancel = 3 '     &H3&  Yes, No, and Cancel buttons
    vbYesNo = 4 '           &H4&  Yes and No buttons
    vbRetryCancel = 5 '     &H5&  Retry and Cancel buttons
    vbCustomButtons = &H6& 'THIS IS YOUR OWN CHOICE OF BUTTON CAPTIONS
    '########### Icon Types available '##########
    vbCritical = 16 '&H10&        Critical message
    vbQuestion = 32 '&H20&        Warning query
    vbExclamation = 48 '&H30&     Warning message
    vbInformation = 64 '&H40&     Information message
    vbUserIcon = &H50&
    vbSecurityIcon = &H60&
    vbFindIcon = &H70&
    '###### Default Selected button #####
    vbDefaultButton1 = 0 'First button is default (default)
    vbDefaultButton2 = 256 ' Second button is default
    vbDefaultButton3 = 512 'Third button is default
End Enum

Public Declare Function InternetGetConnectedState _
               Lib "wininet" (lpdwFlags As Long, _
                              ByVal dwReserved As Long) As Boolean

Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8
Public Const INTERNET_CONNECTION_CONFIGURED = &H40
Public Const INTERNET_CONNECTION_OFFLINE = &H20
Public Const INTERNET_RAS_INSTALLED = &H10

Private Const MAX_PATH                   As Long = 260
Private Const ERROR_SUCCESS              As Long = 0

'Treat entire URL param as one URL segment
Private Const URL_ESCAPE_SEGMENT_ONLY    As Long = &H2000
Private Const URL_ESCAPE_PERCENT         As Long = &H1000
Private Const URL_UNESCAPE_INPLACE       As Long = &H100000

'escape #'s in paths
Private Const URL_INTERNAL_PATH          As Long = &H800000
Private Const URL_DONT_ESCAPE_EXTRA_INFO As Long = &H2000000
Private Const URL_ESCAPE_SPACES_ONLY     As Long = &H4000000
Private Const URL_DONT_SIMPLIFY          As Long = &H8000000

'Converts unsafe characters,
'such as spaces, into their
'corresponding escape sequences.
Private Declare Function UrlEscape _
                Lib "shlwapi" _
                Alias "UrlEscapeA" (ByVal pszURL As String, _
                                    ByVal pszEscaped As String, _
                                    pcchEscaped As Long, _
                                    ByVal dwFlags As Long) As Long

'Converts escape sequences back into
'ordinary characters.
Private Declare Function UrlUnescape _
                Lib "shlwapi" _
                Alias "UrlUnescapeA" (ByVal pszURL As String, _
                                      ByVal pszUnescaped As String, _
                                      pcchUnescaped As Long, _
                                      ByVal dwFlags As Long) As Long
      
Private Const CB_GETLBTEXTLEN = &H149

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

'Public g_ColAlerts As Collection

Public Function ReadTextFile(strPath As String) As String
    Dim fso As New FileSystemObject
    Dim TS As TextStream
    Dim strOutput As String
    Set TS = fso.OpenTextFile(strPath)

    Do Until TS.AtEndOfStream
        strOutput = strOutput + TS.ReadLine
    Loop
       
    TS.Close
    ReadTextFile = strOutput
End Function

Public Function AutoSizeBox(cboSize As Object, _
                            blSizeBox As Boolean) As Boolean 'ComboBox

    'Dimension variables
    Dim lngRet As Long
    Dim rectCboText As ColorPickerCommon.RECT
    Dim lngParentHDC As Long
    Dim lngListCount As Long
    Dim lngCounter As Long
    Dim lngTempWidth As Long
    Dim lngWidth As Long
    Dim strSavedFont As String
    Dim sngSavedSize As Single
    Dim blnSavedBold As Boolean
    Dim blnSavedItalic As Boolean
    Dim blnSavedUnderline As Boolean
    Dim blnFontSaved As Boolean
    Dim iSize As Integer
    Dim sText As String
    On Error GoTo ErrorHandler
    
    lngParentHDC = cboSize.Parent.hdc
    lngListCount = cboSize.ListCount

    If lngParentHDC = 0 Or lngListCount = 0 Then Exit Function

    With cboSize.Parent
        strSavedFont = .FontName
        sngSavedSize = .FontSize
        blnSavedBold = .FontBold
        blnSavedItalic = .FontItalic
        blnSavedUnderline = .FontUnderline
        .FontName = cboSize.FontName
        .FontSize = cboSize.FontSize
        .FontBold = cboSize.FontBold
        .FontItalic = cboSize.FontItalic
        .FontUnderline = cboSize.FontItalic
    End With
        
    blnFontSaved = True
  
    For lngCounter = 0 To lngListCount
        DrawText lngParentHDC, cboSize.List(lngCounter), -1, rectCboText, DT_CALCRECT
        'Add 10 to the the number as a margin
        lngTempWidth = rectCboText.right - rectCboText.left + 10

        If (lngTempWidth > lngWidth) Then
            lngWidth = lngTempWidth
        End If

    Next

    lngWidth = (lngWidth) * Screen.TwipsPerPixelX

    If lngWidth > Screen.Width - 20 Then lngWidth = Screen.Width - 20
        
    'Set the width of our combobox

    If blSizeBox Then
        'resize entire box add 225 to move the drop control out of the way of the text
        cboSize.Width = lngWidth + 225
    Else
        'resize dropdown only
        SendMessage cboSize.hwnd, CB_SETDROPPEDWIDTH, (lngWidth / Screen.TwipsPerPixelX), 0&
    End If

    AutosizeListbox = True
    Exit Function
ErrorHandler:
    On Error Resume Next
    DebugPrint Err.number

    If blnFontSaved Then

        With cboSize.Parent
            .FontName = strSavedFont
            .FontSize = sngSavedSize
            .FontUnderline = blnSavedUnderline
            .FontBold = blnSavedBold
            .FontItalic = blnSavedItalic
        End With

    End If

End Function
   
Public Function OutlookExists() As Integer
    On Error Resume Next
    
    Dim objOutlook As Object
    Dim intExists As Integer
    
    Set objOutlook = CreateObject("Outlook.Application")
    
    If objOutlook Is Nothing Then
        intExists = 0
    Else
        intExists = 1
    End If
    
    Set objOutlook = Nothing
    OutlookExists = intExists
End Function

Public Function WordExists() As Integer
    On Error Resume Next
    
    Dim objWord As Object
    Dim intExists As Integer
    
    Set objWord = CreateObject("Word.Application")
    
    If objWord Is Nothing Then
        intExists = 0
    Else
        intExists = 1
    End If
    
    Set objWord = Nothing
    
    WordExists = intExists
End Function

Public Function ExcelExists() As Integer
    On Error Resume Next
    
    Dim objExcel As Object
    Dim intExists As Integer
    
    Set objExcel = CreateObject("Excel.Application")
    
    If objExcel Is Nothing Then
        intExists = 0
    Else
        intExists = 1
    End If
    
    Set objExcel = Nothing
    
    ExcelExists = intExists
End Function

Public Function PowerPointExists() As Integer
    On Error Resume Next
    
    Dim objPowerPoint As Object
    Dim intExists As Integer
    
    Set objPowerPoint = CreateObject("PowerPoint.Application")
    
    If objPowerPoint Is Nothing Then
        intExists = 0
    Else
        intExists = 1
    End If
    
    Set objPowerPoint = Nothing
    PowerPointExists = intExists
End Function

Public Function AccessExists() As Integer
    On Error Resume Next
    
    Dim objAccess As Object
    Dim intExists As Integer
    
    Set objAccess = CreateObject("Access.Application")
    
    If objAccess Is Nothing Then
        intExists = 0
    Else
        intExists = 1
    End If
    
    Set objAccess = Nothing
    AccessExists = intExists

End Function

Public Sub FormOnTopEx(hWindow As Long, _
                       bTopMost As Boolean)
    ' Example: Call FormOnTop(me.hWnd, True)
    Const SWP_NOSIZE = &H1
    Const SWP_NOMOVE = &H2
    Const SWP_NOACTIVATE = &H10
    Const SWP_SHOWWINDOW = &H40
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
    
    wFlags = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
    
    Select Case bTopMost

        Case True
            Placement = HWND_TOPMOST

        Case False
            Placement = HWND_NOTOPMOST
    End Select

    SetWindowPos hWindow, Placement, 0, 0, 0, 0, wFlags
End Sub

Public Function EncodeUrl(ByVal sURL As String) As String

    Dim buff As String
    Dim dwSize As Long
    Dim dwFlags As Long
   
    If Len(sURL) > 0 Then
      
        buff = Space$(MAX_PATH)
        dwSize = Len(buff)
        dwFlags = URL_DONT_SIMPLIFY
      
        If UrlEscape(sURL, buff, dwSize, dwFlags) = ERROR_SUCCESS Then
                   
            EncodeUrl = left$(buff, dwSize)
      
        End If  'UrlEscape
    End If  'Len(sUrl)

End Function

Public Function DecodeUrl(ByVal sURL As String) As String

    Dim buff As String
    Dim dwSize As Long
    Dim dwFlags As Long
   
    If Len(sURL) > 0 Then
      
        buff = Space$(MAX_PATH)
        dwSize = Len(buff)
        dwFlags = URL_DONT_SIMPLIFY
      
        If UrlUnescape(sURL, buff, dwSize, dwFlags) = ERROR_SUCCESS Then
                   
            DecodeUrl = left$(buff, dwSize)
      
        End If  'UrlUnescape
    End If  'Len(sUrl)

End Function

Public Sub WriteAttachmentHTML()
    Dim i As Integer

    'arAttachments As Variant

    For i = 0 To 10000
        DebugPrint "MultiThread:" & i
    Next

    '    For i = LBound(arAttachments) To UBound(arAttachments)
    '        DebugPrint
    '    Next

End Sub

Public Sub UploadAttachments(txtURL As String, _
                             sUploadFileName As String)
        '<EhHeader>
        On Error GoTo UploadAttachments_Err
        '</EhHeader>
        Dim objURL As CURL

100     Set objURL = New CURL

102     With objURL
104         Call .ParseURL(txtURL, GetProxy())
        End With

106     Call UploadFiles(objURL, sUploadFileName)
        '<EhFooter>
        Exit Sub

UploadAttachments_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.UploadAttachments " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function GetProxy() As String
        '<EhHeader>
        On Error GoTo GetProxy_Err
        '</EhHeader>

        Dim strProxy As String
        Dim blnProxyEnabled As Boolean
        Dim intBeg As Integer
        Dim intEnd As Integer

100     strProxy = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer", REG_SZ)
102     blnProxyEnabled = CBool(QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable", REG_DWORD))

104     intBeg = InStr(1, strProxy, "http=", vbTextCompare)

106     If intBeg > 0 And blnProxyEnabled Then
108         intEnd = InStr(intBeg + 1, strProxy, ";")

110         If intEnd = 0 Then intEnd = Len(strProxy)
112         strProxy = Mid$(strProxy, intBeg + 5, intEnd - intBeg - 5)
        End If

        'Return
114     GetProxy = strProxy

        '<EhFooter>
        Exit Function

GetProxy_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GetProxy " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function GetFileExtension(strFileName As String) As String
        '<EhHeader>
        On Error GoTo GetFileExtension_Err
        '</EhHeader>

        'Error check

100     If Len(strFileName) < 3 Or InStr(1, strFileName, ".") = 0 Then
102         GetFileExtension = ""
            Exit Function
        End If

        'Return
104     GetFileExtension = Mid$(strFileName, InStrRev(strFileName, ".") + 1)

        '<EhFooter>
        Exit Function

GetFileExtension_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GetFileExtension " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function GetMimeType(strFileName As String) As String
        '<EhHeader>
        On Error GoTo GetMimeType_Err
        '</EhHeader>

        Dim strExtension As String

100     strExtension = LCase$(GetFileExtension(strFileName))

        'Error check
102     If strExtension = "" Then
104         GetMimeType = "text/plain"
            Exit Function
        End If

106     Select Case strExtension

            Case "bmp"
108             GetMimeType = "image/bmp"

110         Case "gif"
112             GetMimeType = "image/gif"

114         Case "jpg", "jpeg"
116             GetMimeType = "image/jpeg"

118         Case "swf"
120             GetMimeType = "application/x-shockwave-flash"

122         Case "mpg", "mpeg"
124             GetMimeType = "video/mpeg"

126         Case "wmv"
128             GetMimeType = "video/x-ms-wmv"

130         Case "avi"
132             GetMimeType = "video/avi"

134         Case "shp"
136             GetMimeType = "application/octet-stream"

138         Case "shx"
140             GetMimeType = "application/octet-stream"

142         Case "dbf"
144             GetMimeType = "application/octet-stream"
    
146         Case "fld"
148             GetMimeType = "application/octet-stream"
    
150         Case "sbn"
152             GetMimeType = "application/octet-stream"
    
154         Case "sbx"
156             GetMimeType = "application/octet-stream"
    
158         Case Else
160             GetMimeType = "text/plain"
        End Select

        '<EhFooter>
        Exit Function

GetMimeType_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.GetMimeType " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DegToRad(degrees As Double)
    DegToRad = degrees * (3.14 / 180)
End Function

Public Sub UploadFiles(objURL As CURL, _
                       txtFile1 As String)
        '<EhHeader>
        On Error GoTo UploadFiles_Err
        '</EhHeader>

        Dim lngStart As Long
        Dim strQuery As String
        Dim strFileContents1 As String
        Dim strFileContents2 As String
        Dim objHTTPRequest As CHTTPRequest
        Dim intSock As Integer

        Const intSecondsToWait = 10 'Seconds to wait = 10

100     DebugPrint "Connecting to " & objURL.Host

102     DoEvents

104     intSock = ConnectSock(objURL.Host, objURL.Port, frmMain.hwnd)

106     If intSock = SOCKET_ERROR Then
108         DebugPrint "Could not connect to " & objURL.Host
            Exit Sub
        End If

110     Set objHTTPRequest = New CHTTPRequest

        'File 1
112     If txtFile1 <> "" Then
114         If Dir(txtFile1) <> "" Then  'File exists
116             strFileContents1 = GetFileQuick(txtFile1)
            End If
        End If

118     With objHTTPRequest
120         .Host = objURL.Host
122         .Proxy = objURL.Proxy
124         .Path = objURL.Path
126         .UserAgent = App.Title

128         .MimeBoundary = "LcEnTeRpRiSeS"

            'Form fields
130         Call .AddFormData("Field1", "test field")

            'Files
132         Call .AddFile("File1", txtFile1, strFileContents1, GetMimeType(txtFile1))

            'MsgBox .GetGETQuery
134         strQuery = .GetPOSTQuery
        End With

136     DebugPrint "Sending request..."

138     DoEvents

        'Send request
140     Call SendData(intSock, strQuery)

        'Wait for page to be downloaded
142     lngStart = timeGetTime
144     While intSecondsToWait - Int((timeGetTime - lngStart) / 1000) > 0 And intSock > 0
146         DebugPrint "Waiting for response from " & objURL.Host & "... " & intSecondsToWait - Int((timeGetTime - lngStart) / 1000)

148         DoEvents

            'You can put a routine that will check if a boolean variable is True here
            'This could indicate that the request has been canceled
            'If CancelFlag = True Then
            '   lblStatus.Caption = "Cancelled request"
            '   Exit Sub
            'End If
        Wend

150     DebugPrint "Ready"

        '<EhFooter>
        Exit Sub

UploadFiles_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.UploadFiles " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function CheckInternConnection(ByRef sConnectionType As String) As Boolean
    Dim flags As Long
    Dim Result As Boolean

    sConnectionType = ""

    Result = InternetGetConnectedState(flags, 0)

    If Result Then
        CheckInternConnection = True
    Else
        CheckInternConnection = False
    End If
     
    If flags And INTERNET_CONNECTION_MODEM Then sConnectionType = "Connection Via Modem"
    If flags And INTERNET_CONNECTION_LAN Then sConnectionType = "Connection Via LAN"
    If flags And INTERNET_CONNECTION_PROXY Then sConnectionType = "Connection uses a Proxy"
    If flags And INTERNET_CONNECTION_MODEM_BUSY Then sConnectionType = "Connection Via Modem but modem is busy"
    If flags And INTERNET_CONNECTION_CONFIGURED Then sConnectionType = "Local system has a valid connection to the Internet, but it may or may not be currently connected."
    If flags And INTERNET_CONNECTION_OFFLINE Then sConnectionType = "Local system is in offline mode."
    
    'If FLAGS And INTERNET_RAS_INSTALLED Then sConnectionType = "Local system has RAS installed."

End Function

Public Function AutosizeCombo(Combo As ComboBox) As Boolean

    'Automatically sizes a combo box to
    'hold the longest item within it
    Dim lngRet As Long
    Dim lngCurrentWidth As Single
    Dim rectCboText As ColorPickerCommon.RECT
    Dim lngParentHDC As Long
    Dim lngListCount As Long
    Dim lngCounter As Long
    Dim lngTempWidth As Long
    Dim lngWidth As Long
    Dim strSavedFont As String
    Dim sngSavedSize As Single
    Dim blnSavedBold As Boolean
    Dim blnSavedItalic As Boolean
    Dim blnSavedUnderline As Boolean
    Dim blnFontSaved As Boolean
    On Error GoTo ErrorHandler
    'Grab the combo handle and list count
    lngParentHDC = Combo.Parent.hdc
    lngListCount = Combo.ListCount

    If lngParentHDC = 0 Or lngListCount = 0 Then Exit Function
    'Save combo box fonts, etc. to the paren
    '     t
    'object (form), for testing lengths with
    '     the API

    With Combo.Parent
        strSavedFont = .FontName
        sngSavedSize = .FontSize
        blnSavedBold = .FontBold
        blnSavedItalic = .FontItalic
        blnSavedUnderline = .FontUnderline
        .FontName = Combo.FontName
        .FontSize = Combo.FontSize
        .FontBold = Combo.FontBold
        .FontItalic = Combo.FontItalic
        .FontUnderline = Combo.FontItalic
    End With

    blnFontSaved = True
    'Get the width of the widest item

    For lngCounter = 0 To lngListCount
        DrawText lngParentHDC, Combo.List(lngCounter), -1, rectCboText, DT_CALCRECT
        'Add twenty to the the number as a margi
        '     n
        lngTempWidth = rectCboText.right - rectCboText.left + 20

        If (lngTempWidth > lngWidth) Then
            lngWidth = lngTempWidth
        End If

    Next
    
    'Get current width of combo
    
    lngCurrentWidth = SendMessageLong(Combo.hwnd, CB_GETDROPPEDWIDTH, 0, 0)
    'If big enough then that's all A-OK

    If lngCurrentWidth > lngWidth Then
        AutosizeCombo = True
        GoTo ErrorHandler
        Exit Function
    End If
    
    '... but if not big enough, first calcul
    '     ate
    'the screen width to ensure we don't exc
    '     eed it!
    
    If lngWidth > Screen.Width \ Screen.TwipsPerPixelX - 20 Then lngWidth = Screen.Width \ Screen.TwipsPerPixelX - 20
    'Set the width of our combo
    lngRet = SendMessageLong(Combo.hwnd, CB_SETDROPPEDWIDTH, lngWidth, 0)
    'Set the function to True/False dependin
    '     g on API success
    AutosizeCombo = lngRet > 0
ErrorHandler:
    'If anything goes wrong, revert back!
    On Error Resume Next

    If blnFontSaved Then

        With Combo.Parent
            .FontName = strSavedFont
            .FontSize = sngSavedSize
            .FontUnderline = blnSavedUnderline
            .FontBold = blnSavedBold
            .FontItalic = blnSavedItalic
        End With

    End If

End Function

'**************************************
' Name: A Better, Faster,Remove Duplicat
'     es Function (Kill Dupes, Remove Dupes, K
'     ill Duplicates)
' Description:Please check "http://chad.
'     offroadextremes.com/forums" for all my s
'     ource code and projects. This is a very
'     fast, very efficient RemoveDupes Functio
'     n. Every single RemoveDupes Sub/Function
'     I have ever seen on PSCode has been very
'     slow and never removes all the duplicate
'     list items. If the List/Combo contained
'     more than one duplicate, only the first
'     duplicate was removed. My Function, howe
'     ver, removes every single duplicate no m
'     atter how many there are. It also does i
'     t very quickly without going back throug
'     h the list. It only sweeps through the l
'     ist once. This is the fastest and most e
'     fficient way of removing dupes in VB. Al
'     so, I have added the options to remove a
'     ll matches no matter the Case or if it h
'     as spaces or not, remove matches that on
'     ly match exactly or any combination of m
'     atching options. As an added bonus, I wr
'     ote it as a Function to return the numbe
'     r of duplicates removed. You can call th
'     is Function like you would a sub for sim
'     plicity, or use anything to display the
'     number of dupes removed. EXAMPLES: [1] C
'     all RemoveDupes(List1, True, True) [2] C
'     all MsgBox(RemoveDupes(List1, True, True
'     )) [3] lngDupes& = Call RemoveDupes(
'     List1, True, True)
' By: Chad Roe
'
' Inputs:lstList = The List you are remo
'     ving dupes from. blnLowerCase = Wether y
'     ou want to match LCase or not. blnNoSpac
'     es = Wether you want to match with or wi
'     thout spaces. Setting both to True will
'     remove all duplicates that are the same
'     no matter what Case they are or if they
'     contain spaces or not. EXAMPLE: "aA a" i
'     s the same as "aaa" if both are set to T
'     rue. "aA a" and "aaa" are different if b
'     oth are set to False.
'
' Returns:Returns a ListBox or ComboBox

'     with no duplicate items.
'
' Assumes:As you can see in the code, af
'     ter an item is removed, I set the curren
'     t Check Item back one to make up for the
'     item just removed. This is why this Func
'     tion removes ALL duplicate items and not
'     just the first one. Also, I do not use D
'     oEvents because this drastically slows d
'     own the process. To me, this is the grea
'     test Dupe Killer that can be written in
'     Visual Basic.
'
' Side Effects:None.
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=62390&lngWId=1'for details.'**************************************

Public Function RemoveDupes(lstList As Control, _
                            blnLowerCase As Boolean, _
                            blnNoSpaces As Boolean) As Long

    Dim strItem As String, strCheck As String
    Dim intItem As Integer, intCheck As Integer
    RemoveDupes& = lstList.ListCount

    For intItem% = 0 To lstList.ListCount - 1

        For intCheck% = intItem% + 1 To lstList.ListCount - 1
            strItem$ = lstList.List(intItem%)

            If blnLowerCase = True Then strItem$ = LCase(strItem$)
            If blnNoSpaces = True Then strItem$ = Replace(strItem$, " ", "")
            strCheck$ = lstList.List(intCheck%)

            If blnLowerCase = True Then strCheck$ = LCase(strCheck$)
            If blnNoSpaces = True Then strCheck$ = Replace(strCheck$, " ", "")

            If strItem$ = strCheck$ Then
                Call lstList.RemoveItem(intCheck%)
                intCheck% = intCheck% - 1
            End If

        Next intCheck%

    Next intItem%

    RemoveDupes& = RemoveDupes& - lstList.Li
    '     stCount
End Function

Public Function ClearsTEXTCOM_IN_Form(frm As Form)

    Dim ctl As Control

    For Each ctl In frm.Controls

        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
            On Error Resume Next
            ctl.ListIndex = 0
        End If

    Next

End Function

Public Sub ComboAddItemEx(cbo As ComboBox, _
                        strItemText As String, _
                        lngItemData As Long)

    Dim lngAddIndex     As Long
   
    'Adding items strings and items data with Win32 API
    
    'Add the new string to the ComboBox.
    lngAddIndex = SendMessage(cbo.hwnd, CB_ADDSTRING, 0, ByVal strItemText)

    'Set the item data of the new item.
    SendMessage cmbTest.hwnd, CB_SETITEMDATA, lngAddIndex, ByVal lngItemData

End Sub

Public Sub ComboAddItem(cbo As ComboBox, _
                        strItemText As String)

   SendMessage cbo.hwnd, CB_ADDSTRING, 0, ByVal strItemText

End Sub

Public Function ItemInBox(AddText As String, _
                          BoxType As Control, _
                          Optional AddIt As Boolean, _
                          Optional CompareMethod As VbCompareMethod = vbBinaryCompare) As Boolean

    Dim RetVal As Long
    Dim FindMessage As Long
    Dim AddMessage As Long

    If TypeOf BoxType Is ListBox Then
        'use the listbox functions
        FindMessage = LB_FINDSTRINGEXACT
        AddMessage = LB_ADDSTRING
    ElseIf TypeOf BoxType Is ComboBox Then
        'use the combobox functions
        FindMessage = CB_FINDSTRINGEXACT
        AddMessage = CB_ADDSTRING
    Else
        'cannot do it
        Exit Function
    End If

    RetVal = SendMessage(hwnd:=BoxType.hwnd, wMsg:=FindMessage, wParam:=-1, lParam:=ByVal AddText) 'Note strings must be passed byval
    'verify using compare type as well

    If RetVal = -1 Or StrComp(AddText, BoxType.List(RetVal), CompareMethod) <> 0 Then
        'not found

        If AddIt Then
            Call SendMessage(hwnd:=BoxType.hwnd, wMsg:=AddMessage, wParam:=0, lParam:=ByVal AddText) 'Note strings must be passed byval
        End If

        ItemInBox = False
    Else
        BoxType.ListIndex = RetVal
        ItemInBox = True
    End If

End Function

Public Function DeCryptString(Phrase As String) As String
    Dim outputText As String
    Dim Position As Integer, Asc1 As Long, Char1 As String
    
    For Position = Len(Phrase) To 1 Step -1
        Char1 = Mid$(Phrase, Position, 1)
        
        Asc1 = Asc(Char1)

        Asc1 = (((Asc1 * Asc1) / 2) / 2)
        Asc1 = Sqr(Asc1)

        Char1 = Chr$(Asc1)
                
        outputText = outputText & Char1
    Next
    
    DeCryptString = outputText

End Function

Public Function EncryptString(Phrase As String) As String
    Dim Encrypted As String
    Dim Position As Integer, Asc1 As Long, Char1 As String

    For Position = Len(Phrase) To 1 Step -1
        Char1 = Mid$(Phrase, Position, 1)
        
        Asc1 = Asc(Char1)
        
        Asc1 = (Asc1 * Asc1) / (Asc1 / 2)
        
        Char1 = Chr$(Asc1)
                
        Encrypted = Encrypted & Char1
    Next
    
    EncryptString = Encrypted

End Function

'PURPOSE: Obtain distinct values from an array
'PARAMETER:  OrigArray: any array
'RETURNS: A variant array with only the unique values of
'OrigArray E.g., pass in an array containing elements 2, 1, 2;
'return value is an array with elements 2, 1

Public Function UniqueValues(ByVal OrigArray As Variant) As Variant

    Dim vAns() As Variant
    Dim lStartPoint As Long
    Dim lEndPoint As Long
    Dim lCtr As Long, lCount As Long
    Dim iCtr As Integer
    Dim Col As New Collection
    Dim sIndex As String

    Dim vTest As Variant, vItem As Variant
    Dim iBadVarTypes(4) As Integer

    'Function does not work with if array element is one of the
    'following types
    iBadVarTypes(0) = vbObject
    iBadVarTypes(1) = vbError
    iBadVarTypes(2) = vbDataObject
    iBadVarTypes(3) = vbUserDefinedType
    iBadVarTypes(4) = vbArray

    'Check to see if the parameter is an array
    If Not IsArray(OrigArray) Then
        Err.Raise ERR_BP_NUMBER, , ERR_BAD_PARAMETER
        Exit Function
    End If

    lStartPoint = LBound(OrigArray)
    lEndPoint = UBound(OrigArray)

    For lCtr = lStartPoint To lEndPoint
        vItem = OrigArray(lCtr)
    
        'First check to see if variable type is acceptable
        For iCtr = 0 To UBound(iBadVarTypes)

            If VarType(vItem) = iBadVarTypes(iCtr) Or VarType(vItem) = iBadVarTypes(iCtr) + vbVariant Then

                Err.Raise ERR_BT_NUMBER, , ERR_BAD_TYPE
                Exit Function

            End If

        Next iCtr

        'Add element to a collection, using it as the index
        'if an error occurs, the element already exists
 
        sIndex = CStr(vItem)

        'first element, add automatically
        If lCtr = lStartPoint Then
            Col.Add vItem, sIndex
            ReDim vAns(lStartPoint To lStartPoint) As Variant
            vAns(lStartPoint) = vItem
        Else
            On Error Resume Next
            Col.Add vItem, sIndex
        
            If Err.number = 0 Then
                lCount = UBound(vAns) + 1
                ReDim Preserve vAns(lStartPoint To lCount)
                vAns(lCount) = vItem
            End If
        End If

        Err.Clear
    Next lCtr
    
    UniqueValues = vAns

End Function

Public Function ContainsUniqueValues(vList As Variant) As Boolean

    Dim vTemp() As Variant
    Dim lCount As Long
    Dim lStartPoint As Long
    Dim bAns As Boolean
    Dim lCtr As Long, lTmpCtr As Long
    Dim lTempCount As Long
    Dim bCollection As Boolean
    Dim vTest As Variant, v As Variant
    Dim vTemp2() As Variant

    'Function won't work if elements in vList can't
    'be used in an equality test (e.g., arrays,
    'collection
    If Not ValidateValues(vList) Then
        Err.Raise ERR_INVALID_ELEMENT, , ERR_INVALID_ELEMENT_MSG
    End If

    bAns = True
    ReDim vTemp(0) As Variant

    bCollection = IsObject(vList)

    If bCollection Then
        lCtr = 1

        For Each v In vList

            If lCtr = 1 Then
                vTemp(0) = v
            Else

                For lTmpCtr = 0 To UBound(vTemp)

                    If vTemp(lTmpCtr) = vList(lCtr) Then 'bingo
                        bAns = False
                        Exit For
                    End If

                Next lTmpCtr
        
                If bAns = False Then
                    Exit For
                Else
                    ReDim Preserve vTemp(UBound(vTemp) + 1) As Variant
                    vTemp(UBound(vTemp)) = vList(lCtr)
                End If
            
            End If

            lCtr = lCtr + 1
        Next

    Else

        lCount = UBound(vList)
        lStartPoint = LBound(vList)

        For lCtr = lStartPoint To lCount
    
            If lCtr = lStartPoint Then
                vTemp(0) = vList(lCtr)
            Else

                For lTmpCtr = 0 To UBound(vTemp)

                    If vTemp(lTmpCtr) = vList(lCtr) Then 'bingo
                        bAns = False
                        Exit For
                    End If

                Next lTmpCtr
        
                If bAns = False Then
                    Exit For
                Else
                    ReDim Preserve vTemp(UBound(vTemp) + 1) As Variant
                    vTemp(UBound(vTemp)) = vList(lCtr)
                End If
            
            End If
        
        Next lCtr

    End If

    ContainsUniqueValues = bAns

End Function

Private Function ValidateValues(values As Variant) As Boolean

    'Purpose: Determines if all the values of a collection or
    'variant array has a value that can be converted into a string
    'and/or can be used in an equality test

    Dim bCollection As Boolean
    Dim iBadVarTypes(3) As Integer
    Dim v As Variant
    Dim i As Integer
    Dim lCtr As Long, lListCount As Long
    Dim lStartPoint As Long
    Dim iCount As Integer

    Dim bAns As Boolean

    iBadVarTypes(0) = vbError
    iBadVarTypes(1) = vbDataObject
    iBadVarTypes(2) = vbUserDefinedType
    iBadVarTypes(3) = vbArray

    bAns = True
    iCount = UBound(iBadVarTypes)

    If IsObject(values) Then
        If Not TypeOf values Is Collection Then
            ValidateValues = False
            Exit Function
        End If

    Else

        If Not IsArray(values) Then 'single value
    
            For i = 0 To iCount

                If VarType(values) = iBadVarTypes(i) Then
                    bAns = False
                    Exit For
                End If

            Next
        
            ValidateValues = bAns
            Exit Function
        End If
    End If

    bCollection = IsObject(values) 'has to be collection

    If bCollection Then

        For Each v In values
            For i = 0 To iCount

                If VarType(v) = iBadVarTypes(i) Or VarType(v) = iBadVarTypes(i) + vbVariant Or IsObject(v) Then
                    bAns = False
                    Exit For
                End If

            Next

            If bAns = False Then Exit For
        Next

    Else
        lListCount = UBound(values)
        lStartPoint = LBound(values)

        For lCtr = lStartPoint To lListCount
          
            For i = 0 To iCount

                If VarType(values(lCtr)) = iBadVarTypes(i) Or IsArray(values(lCtr)) Or IsObject(values(lCtr)) Then
                    bAns = False
                    Exit For
                End If

            Next

            If bAns = False Then Exit For
        Next

    End If
    
    ValidateValues = bAns

End Function

Public Function FontExists(FontName As String) As Boolean
    '***********************************************
    'PURPOSE:   Determine if a font exists on a system
    'PARAMETER: FontName = NameofFont
    'RETURNS:  True if font exists, false otherwise
    'EXAMPLE:
    '          If FontExists("Verdana") then
    '            Text1.FontName = "Verdana"
    '          End If
    '***********************************************
    Dim oFont As New StdFont
    Dim bAns As Boolean
    
    oFont.Name = FontName
    bAns = StrComp(FontName, oFont.Name, vbTextCompare) = 0
    FontExists = bAns

End Function

Public Function MakeDictionary(Keys As Variant, _
                               Data As Variant) As Object

    'arrays or collections permitted

    Dim l As Long, lCtr As Long

    Dim bValid As Boolean
    Dim sSplit() As String
    Dim bKeyCollection As Boolean
    Dim bDataCollection As Boolean

    Dim lDataCount As Long
    Dim lKeyCount As Long
    Dim lTotal As Long
    Dim lDataStartPoint As Long
    Dim lKeyStartPoint As Long
    Dim vTest As Variant
    Dim iDataKeyDifference As Integer

    Dim oDict As New Scripting.Dictionary

    If IsObject(Keys) Then
        If Not TypeOf Keys Is Collection Then
            Err.Raise ERR_INVALID_LIST, , ERR_INVALID_LIST_MSG
        Else
            lKeyStartPoint = 1
            bKeyCollection = True
        End If
   
    ElseIf Not IsArray(Keys) Then
        Err.Raise ERR_INVALID_LIST, , ERR_INVALID_LIST_MSG
    End If

    If IsObject(Data) Then
        If Not TypeOf Data Is Collection Then
            Err.Raise ERR_INVALID_LIST, , ERR_INVALID_LIST_MSG
        Else
            lDataStartPoint = 1
            bDataCollection = True
        End If

    ElseIf Not IsArray(Data) Then
        Err.Raise ERR_INVALID_LIST, , ERR_INVALID_LIST_MSG
    End If

    'In dictionary,keys must be unique
    'if not unique, we raise a custom error

    'However, the containsuniquevalues test
    'won't work with objects or arrays
    'even though  these can be keys

    'Therefore, if there is a non-unique key of one of these
    'types, VB will raise an error for you, instead of it
    'happening here
    If ValidateValues(Keys) Then

        If Not ContainsUniqueValues(Keys) Then
            Err.Raise ERR_NONUNIQUE_KEYS, , ERR_NONUNIQUE_KEYS_MSG
        End If
    End If

    If bDataCollection Then
        lDataCount = Data.Count
    Else
        lDataCount = UBound(Data)
        'determine if it's a 0 or one-based array
        On Error Resume Next
        vTest = Data(0)

        If Err.number = 0 Then
            lDataStartPoint = 0
            lDataCount = UBound(Data)
        Else
            lDataStartPoint = 1
            lDataCount = UBound(Data)
        End If

        Err.Clear
        On Error GoTo 0
    End If

    If bKeyCollection Then
        lKeyCount = Keys.Count
    Else
        lKeyCount = UBound(Keys)
        'determine if it's a 0 or one-based array
        On Error Resume Next
        vTest = Keys(0)

        If Err.number = 0 Then
            lKeyStartPoint = 0
            lKeyCount = UBound(Keys)
        Else
            lKeyStartPoint = 1
            lKeyCount = UBound(Keys)
        End If

        Err.Clear
        On Error GoTo 0
    End If

    iDataKeyDifference = lDataStartPoint - lKeyStartPoint

    For l = lKeyStartPoint To lKeyCount
        oDict.Add Keys(l), Data(l + iDataKeyDifference)
    Next

    Set MakeDictionary = oDict
 
End Function

Private Function ContainsUniqueValues1(vList As Variant) As Boolean

    Dim vTemp() As Variant
    Dim lCount As Long
    Dim lStartPoint As Long
    Dim bAns As Boolean
    Dim lCtr As Long, lTmpCtr As Long
    Dim lTempCount As Long
    Dim bCollection As Boolean
    Dim vTest As Variant, v As Variant
    Dim vTemp2() As Variant

    'Function won't work if elements in vList can't
    'be used in an equality test (e.g., arrays,
    'collection

    bAns = True
    ReDim vTemp(0) As Variant

    bCollection = IsObject(vList)

    If bCollection Then
        lCtr = 1

        For Each v In vList
    
            If lCtr = 1 Then
                vTemp(0) = v
            Else

                For lTmpCtr = 0 To UBound(vTemp)

                    If vTemp(lTmpCtr) = vList(lCtr) Then 'bingo
                        bAns = False
                        Exit For
                    End If

                Next lTmpCtr
        
                If bAns = False Then
                    Exit For
                Else
                    ReDim Preserve vTemp(UBound(vTemp) + 1) As Variant
                    vTemp(UBound(vTemp)) = vList(lCtr)
                End If
            
            End If

            lCtr = lCtr + 1
        Next

    Else

        lCount = UBound(vList)
        On Error Resume Next
        vTest = vList(0)
        lStartPoint = IIf(Err.number = 0, 0, 1)
        On Error GoTo 0
    
        For lCtr = lStartPoint To lCount
    
            If lCtr = lStartPoint Then
                vTemp(0) = vList(lCtr)
            Else

                For lTmpCtr = 0 To UBound(vTemp)

                    If vTemp(lTmpCtr) = vList(lCtr) Then 'bingo
                        bAns = False
                        Exit For
                    End If

                Next lTmpCtr
        
                If bAns = False Then
                    Exit For
                Else
                    ReDim Preserve vTemp(UBound(vTemp) + 1) As Variant
                    vTemp(UBound(vTemp)) = vList(lCtr)
                End If
            
            End If
        
        Next lCtr

    End If

    ContainsUniqueValues1 = bAns

End Function

Public Function ConvertDateToSerial(dDate As Date) As Long

    ConvertDateToSerial = CLng(Format(dDate, "yyyymmdd"))

End Function

Public Function ConvertSerialToDate(lSerialDate As Long) As Date

    Dim sDate As String
    Dim sDateReversed As String
    
    sDate = CStr(lSerialDate)
    sDateReversed = right$(sDate, 2) & "-" & Mid$(sDate, 5, 2) & "-" & left$(sDate, 4)

    ConvertSerialToDate = Format(sDateReversed, "dd-mm-yyyy")
    ConvertSerialToDate = Format(ConvertSerialToDate, "Medium Date")

End Function

Function strReplaceSQL(strTxt) As String

    Dim sbString As Object
    'New StringBuilder
    sbString.Append strTxt
    
    sbString.Replace Chr(34), "' & chr(34) & '"
    sbString.Replace Chr(37), "' & chr(37) & '"
    sbString.Replace Chr(42), "' & chr(42) & '"
    sbString.Replace Chr(39), "' & chr(39) & '"
    
    strReplaceSQL = sbString.ToString

End Function

Function GetConnectionString(sDatabaseFullPath As String) As String

    Dim ofs As New FileSystemObject
    Dim oFile As TextStream
    
    On Error Resume Next
    
    If g_sManualSQLServerPath = "" Then
    
        Dim i As Integer
        i = 1
        g_sManualSQLServerPath = "localhost"
    
        While Environ$(i) <> ""

            If Mid(Environ$(i), 1, InStr(1, Environ(i), "=") - 1) = "oasissqlmanualpath" Then
                g_sManualSQLServerPath = Trim(right(Environ$(i), Len(Environ$(i)) - 1 - Len(Mid(Environ$(i), 1, InStr(1, Environ(i), "=") - 1))))
            End If
    
            i = i + 1
        Wend
    
    End If
    
    If ofs.FileExists(g_sAppPath & "\" & "data\user\ODBC.dat") Then
        GetConnectionString = oFile.ReadAll
    ElseIf bSQLServerInUse Then
    
        g_sSQLServerDatabaseName = Replace(frmLogin.ComServer.Text, "/", "-")
  
        If ofs.FileExists(g_sAppPath & "\" & "data\user\MSSQL-UseWindowsAuth.dat") Then
            g_sGlobalConnectionString = "Provider=SQLNCLI10;Server=" & g_sManualSQLServerPath & "\oasissql;Database=" & g_sSQLServerDatabaseName & ";Integrated Security=SSPI;"
        Else
            'g_sGlobalConnectionString = "Driver={SQL Server Native Client 10.0};Server=localhost\oasissql;Database=OasisV3;Uid=sa;Pwd=!MM@P2O1O"
            g_sGlobalConnectionString = "Provider=SQLNCLI10;Server=" & g_sManualSQLServerPath & "\oasissql;Database=" & g_sSQLServerDatabaseName & ";Uid=sa;Pwd=!MM@P2O1O;"
        End If
        
        GetConnectionString = g_sGlobalConnectionString
        g_sGlobalDialect = "MSSQL"
        bSQLServerInUse = True
        g_sGlobalCursorLocation = adUseClient

    Else
        g_sGlobalDialect = "MSJET"
            If sDatabaseFullPath = "" Then
            sDatabaseFullPath = g_sAppPath & "\data\db\oasisclient.mdb"
            End If
        If g_ClientDBPassword = "none" Then
            GetConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabaseFullPath & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False;"
        Else
            GetConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDatabaseFullPath & ";Mode=ReadWrite|Share Deny None;Jet OLEDB:Database Password=" & g_ClientDBPassword & ";"
        End If

        g_sGlobalConnectionString = GetConnectionString
        g_sGlobalCursorLocation = adUseClient
    End If

    Set ofs = Nothing

End Function

'Function GetClientDatabasePassword(sWebsiteAddress As String) As String
'
'    Dim PassWordRS As adodb.Recordset
'    'Set PassWordRS = OpenSilentHttpCommsRS(sWebsiteAddress & "/oasis4.asp?id=" & CheckEncrypt("SELECT SettingValue2 FROM AppSettings WHERE SettingValue1 = True and SettingName = 'ClientDBPassword'"), True)
'    Set PassWordRS = OpenServerRSCompressed(sWebsiteAddress & "/oasis4.asp", "id", "SELECT SettingValue2 FROM AppSettings WHERE SettingValue1 = True and SettingName = 'ClientDBPassword'")
'
'    If PassWordRS Is Nothing Then
'
'        g_ClientDBPassword = "none"
'
'    Else
'
'        If PassWordRS.EOF Then
'
'            g_ClientDBPassword = "none"
'
'        Else
'
'            If PassWordRS.Fields(0).Value = "" Then
'                g_ClientDBPassword = "none"
'            Else
'                g_ClientDBPassword = PassWordRS.Fields(0).Value
'            End If
'        End If
'    End If
'
'    Set PassWordRS = Nothing
'
'End Function

Public Function GetCountOfDaysInMonth(testYear As Integer, _
                                       testMonth As Integer) As Integer

    GetCountOfDaysInMonth = Day(DateSerial(testYear, testMonth + 1, 0))
    
End Function


'''''' DXDBFILTER FUNCTIONS:

Public Sub ParseFilter(sFilterText As String, _
                       dxColumns As DXDBGRIDLibCtl.IdxGridColumns, _
                       dxFilter As DXDBGRIDLibCtl.IdxDBGridFilter)
        ''
        '  This Function returns a idxGridFilter object based on the
        '  input filter text and the columns of the grid.
        '
        '<EhHeader>
        On Error GoTo ParseFilter_Err
        '</EhHeader>
        Dim sTmp As String
        Dim sFilterFirst As String
        Dim sFilterSecond As String

        Dim icol As Integer
        Dim i As Integer
        Dim lStart As Integer
        Dim lEnd As Long

        Dim iColumnIndex As Integer
        Dim iOperator As Integer
        Dim bIsNot As Boolean
        Dim sValue As String
        Dim sDisplay As String
        Dim bBoolOperator As Boolean
        Dim sColumnFilters() As String

100     sTmp = sFilterText
102     Set mdxColumns = dxColumns

104     icol = 0
        '
        'Step 1: Parse all the columns with (possible) two filteroptions
        '
106     lStart = InStr(1, sTmp, "((", vbTextCompare)

108     Do While (lStart > 0)
110         lEnd = InStr(1, sTmp, "))", vbTextCompare)
112         icol = icol + 1
114         ReDim Preserve sColumnFilters(icol)
116         sColumnFilters(icol) = Mid(sTmp, lStart, lEnd - lStart + 2)
118         sTmp = Replace(sTmp, sColumnFilters(icol), "")
120         lStart = InStr(1, sTmp, "((", vbTextCompare)
        Loop

        '
        'Step 2: Parse all the columns with only one filteroption
        '
122     lStart = InStr(1, sTmp, "(", vbTextCompare)

124     Do While (lStart > 0)
126         lEnd = InStr(1, sTmp, ")", vbTextCompare)
128         icol = icol + 1
130         ReDim Preserve sColumnFilters(icol)
132         sColumnFilters(icol) = Mid(sTmp, lStart, lEnd - lStart + 1)
134         sTmp = Replace(sTmp, sColumnFilters(icol), "")
136         lStart = InStr(1, sTmp, "(", vbTextCompare)
        Loop

        '
        ' Step 3: Translate the filter text to filterobject
        '
138     For i = 1 To icol Step 1

140         If InStr(1, sColumnFilters(i), "((", vbTextCompare) > 0 Then
                'AddFirst + AddSecond for this column
142             sTmp = sColumnFilters(i)
144             sTmp = Replace(sTmp, "((", "(")
146             sTmp = Replace(sTmp, "))", ")")
148             lStart = InStr(1, sTmp, "(", vbTextCompare)
150             lEnd = InStr(1, sTmp, ")", vbTextCompare)
152             sFilterFirst = Mid(sTmp, lStart, lEnd - lStart + 1)
154             sTmp = Replace(sTmp, sFilterFirst, "")

156             Call GetColumnIndex(sFilterFirst, iColumnIndex, dxColumns)
158             Call GetOperator(sFilterFirst, iOperator, bIsNot)
160             Call GetValue(sFilterFirst, sValue, sDisplay)

162             dxFilter.AddFirst iColumnIndex, iOperator, sValue, sDisplay, bIsNot

                'Check if there IS a second: ((<field>='dddd')) is possible
164             If sTmp <> "" Then
166                 lStart = InStr(1, sTmp, "(", vbTextCompare)
168                 lEnd = InStr(1, sTmp, ")", vbTextCompare)

170                 sFilterSecond = Mid(sTmp, lStart, lEnd - lStart + 1)
172                 sTmp = Trim(Replace(sTmp, sFilterSecond, ""))

174                 Call GetColumnIndex(sFilterSecond, iColumnIndex, dxColumns)
176                 Call GetOperator(sFilterSecond, iOperator, bIsNot)
178                 Call GetValue(sFilterSecond, sValue, sDisplay)

180                 If (sTmp = "OR") Then
182                     dxFilter.AddSecond iColumnIndex, iOperator, sValue, sDisplay, bIsNot, boOr
                    Else
184                     dxFilter.AddSecond iColumnIndex, iOperator, sValue, sDisplay, bIsNot, boAnd
                    End If
                End If

            Else
                'AddFirst for this column
186             sTmp = sColumnFilters(i)
188             lStart = InStr(1, sTmp, "(", vbTextCompare)
190             lEnd = InStr(1, sTmp, ")", vbTextCompare)
192             sFilterFirst = Mid(sTmp, lStart, lEnd - lStart + 1)

194             Call GetColumnIndex(sFilterFirst, iColumnIndex, dxColumns)
196             Call GetOperator(sFilterFirst, iOperator, bIsNot)
198             Call GetValue(sFilterFirst, sValue, sDisplay)

200             dxFilter.AddFirst iColumnIndex, iOperator, sValue, sDisplay, bIsNot

            End If

        Next

        '<EhFooter>
        Exit Sub

ParseFilter_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.general.ParseFilter " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub GetColumnIndex(sFilter As String, _
                           iColumnIndex As Integer, mdxColumns As DXDBGRIDLibCtl.IdxGridColumns)
    ''
    '  Determine the index of the column
    '
    Dim sColumnName As String
    Dim lStart As Long
    Dim lEnd As Long

    lStart = InStr(1, sFilter, "[", vbTextCompare)
    lEnd = InStr(1, sFilter, "]", vbTextCompare)
    sColumnName = Mid(sFilter, lStart + 1, lEnd - lStart - 1)

    iColumnIndex = mdxColumns.ColumnByFieldName(sColumnName).Index

End Sub

Private Sub GetOperator(sFilter As String, _
                        iOperator As Integer, _
                        bIsNot As Boolean)
    ''
    '  Return the operator and the is not value of the filtercondition
    '

    iOperator = otIsNull
    bIsNot = True

    If (InStr(1, sFilter, "= Null")) <> 0 Then
        iOperator = otIsNull
        bIsNot = False
    ElseIf (InStr(1, sFilter, "<> Null")) <> 0 Then
        iOperator = otIsNull
        bIsNot = True
    ElseIf (InStr(1, sFilter, ">=")) <> 0 Then
        iOperator = otGreaterEqual
        bIsNot = False
    ElseIf (InStr(1, sFilter, "<=")) <> 0 Then
        iOperator = otLessEqual
        bIsNot = False
    ElseIf (InStr(1, sFilter, "<>")) <> 0 Then
        iOperator = otEqual
        bIsNot = True
    ElseIf (InStr(1, sFilter, "=")) <> 0 Then
        iOperator = otEqual
        bIsNot = False
    ElseIf (InStr(1, sFilter, "<")) <> 0 Then
        iOperator = otLess
        bIsNot = False
    ElseIf (InStr(1, sFilter, ">")) <> 0 Then
        iOperator = otGreater
        bIsNot = False
    End If

End Sub

Private Sub GetValue(sFilter As String, _
                     sValue As String, _
                     sDisplay)
    ''
    '  Return string of the search value
    '
    Dim lStart As Long
    Dim lEnd As Long

    If (InStr(1, sFilter, "Null")) <> 0 Then
        sValue = ""
    ElseIf right(sFilter, 2) = "')" Then
        lStart = InStr(1, sFilter, "'", vbTextCompare)
        lEnd = InStr(lStart + 1, sFilter, "'", vbTextCompare)
        sValue = Mid(sFilter, lStart + 1, lEnd - lStart - 1)
    Else
        'number value
        sValue = Trim(Mid(sFilter, InStrRev(sFilter, " "), InStrRev(sFilter, ")") - InStrRev(sFilter, " ")))
    
    End If

    If sValue = "" Then
        sDisplay = "Null"
    Else
        sDisplay = sValue
    End If

End Sub

'''''' END DXDBFILTER FUNCTIONS:



