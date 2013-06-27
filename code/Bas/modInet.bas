Attribute VB_Name = "modInet"
Option Explicit

Public Const MAX_PATH  As Long = 260
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20

Public Function GetInetError(ByVal lErrorCode As Long) As String

    Dim sBuffer As String
    Dim nBuffer As Long

    Select Case lErrorCode

        Case 12001: GetInetError = "No more handles could be generated at this time"

        Case 12002: GetInetError = "The request has timed out."

        Case 12003:
            'extended error. Retrieve the details using
            'the InternetGetLastResponseInfo API.
      
            sBuffer = Space$(256)
            nBuffer = Len(sBuffer)
      
            If InternetGetLastResponseInfo(lErrorCode, sBuffer, nBuffer) Then
                GetInetError = StripNull(sBuffer)
            Else
                GetInetError = "Extended error returned from server."
            End If
      
        Case 12004: GetInetError = "An internal error has occurred."

        Case 12005: GetInetError = "URL is invalid."

        Case 12006: GetInetError = "URL scheme could not be recognized, or is not supported."

        Case 12007: GetInetError = "Server name could not be resolved."

        Case 12008: GetInetError = "Requested protocol could not be located."

        Case 12009: GetInetError = "Request to InternetQueryOption or InternetSetOption" & " specified an invalid option value."

        Case 12010: GetInetError = "Length of an option supplied to InternetQueryOption or" & " InternetSetOption is incorrect for the type of" & " option specified."

        Case 12011: GetInetError = "Request option can not be set, only queried. "

        Case 12012: GetInetError = "Win32 Internet support is being shutdown or unloaded."

        Case 12013: GetInetError = "Request to connect and login to an FTP server could not" & " be completed because the supplied user name is incorrect."

        Case 12014: GetInetError = "Request to connect and login to an FTP server could not" & " be completed because the supplied password is incorrect. "

        Case 12015: GetInetError = "Request to connect to and login to an FTP server failed."

        Case 12016: GetInetError = "Requested operation is invalid. "

        Case 12017: GetInetError = "Operation was canceled, usually because the handle on" & " which the request was operating was closed before the" & " operation completed."

        Case 12018: GetInetError = "Type of handle supplied is incorrect for this operation."

        Case 12019: GetInetError = "Requested operation can not be carried out because the" & " handle supplied is not in the correct state."

        Case 12020: GetInetError = "Request can not be made via a proxy."

        Case 12021: GetInetError = "Required registry value could not be located. "

        Case 12022: GetInetError = "Required registry value was located but is an incorrect" & " type or has an invalid value."

        Case 12023: GetInetError = "Direct network access cannot be made at this time. "

        Case 12024: GetInetError = "Asynchronous request could not be made because a zero" & " context value was supplied."

        Case 12025: GetInetError = "Asynchronous request could not be made because a" & " callback function has not been set."

        Case 12026: GetInetError = "Required operation could not be completed because" & " one or more requests are pending."

        Case 12027: GetInetError = "Format of the request is invalid."

        Case 12028: GetInetError = "Requested item could not be located."

        Case 12029: GetInetError = "Attempt to connect to the server failed."

        Case 12030: GetInetError = "Connection with the server has been terminated."

        Case 12031: GetInetError = "Connection with the server has been reset."

        Case 12036: GetInetError = "Request failed because the handle already exists."

        Case Else: GetInetError = "Error details not available."
    End Select

End Function

Function StripNull(item As String)

    'Return a string without the chr$(0) terminator.
    Dim pos As Integer
    pos = InStr(item, Chr$(0))

    If pos Then
        StripNull = Left$(item, pos - 1)
        Else: StripNull = item
    End If

End Function
   
Public Function HttpPost(ByVal URL, _
                         ByVal PostData As String) As String
    'The API functions used here are more generic then those in HttpGet and
    'can handle all kinds of internet protocols and verb. However, we use it
    'for HTTP POST here.

    Dim hInternetOpen As Long
    Dim hInternetConnect As Long
    Dim hHttpOpenRequest As Long
    Dim bRet As Boolean
    Dim strServer As String             'the URL's server
    Dim intPort As Integer              'the URL's port
    Dim strPath As String               'the URL's document path
    Dim oUrl As CURL                    'the URL helper object

    Const INTERNET_DEFAULT_HTTP_PORT = 80

    #If 0 Then 'for testing
        strServer = "cgi3.ebay.de"
        strPath = "/aw-cgi/eBayISAPI.dll?TimeShow"
        intPort = 80
    #End If

    'split URL because we need separate pieces here
    Set oUrl = New CURL
    oUrl.Href = URL
    strServer = oUrl.Host
    intPort = Val(oUrl.Port)
    strPath = oUrl.PagePath

    'prepare WinInet
    hInternetOpen = 0
    hInternetConnect = 0
    hHttpOpenRequest = 0

    'se registry access settings.
    Const INTERNET_OPEN_TYPE_PRECONFIG = 0
    hInternetOpen = InternetOpen("http generic", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

    If hInternetOpen <> 0 Then
        'Type of service to access.
        Const INTERNET_SERVICE_HTTP = 3
        'Change the server to your server name
        hInternetConnect = InternetConnect(hInternetOpen, strServer, intPort, vbNullString, "HTTP/1.0", INTERNET_SERVICE_HTTP, 0, 0)

        If hInternetConnect <> 0 Then
            'Brings the data across the wire even if it locally cached.
            Const INTERNET_FLAG_RELOAD = &H80000000
            hHttpOpenRequest = HttpOpenRequest(hInternetConnect, "POST", strPath, "HTTP/1.0", vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

            If hHttpOpenRequest <> 0 Then
                Dim sHeader As String
                Const HTTP_ADDREQ_FLAG_ADD = &H20000000
                Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000
                sHeader = "Content-Type: application/x-www-form-urlencoded" & vbCrLf
                bRet = HttpAddRequestHeaders(hHttpOpenRequest, sHeader, Len(sHeader), HTTP_ADDREQ_FLAG_REPLACE Or HTTP_ADDREQ_FLAG_ADD)

                Dim lPostDataLen As Long

                lPostDataLen = Len(PostData)
                bRet = HttpSendRequest(hHttpOpenRequest, vbNullString, 0, PostData, lPostDataLen)

                Dim bDoLoop             As Boolean
                Dim sReadBuffer         As String * 2048
                Dim lNumberOfBytesRead  As Long
                Dim sBuffer             As String
                bDoLoop = True
                While bDoLoop
                    sReadBuffer = vbNullString
                    bDoLoop = InternetReadFile(hHttpOpenRequest, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                    sBuffer = sBuffer & Left(sReadBuffer, lNumberOfBytesRead)

                    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
                Wend
                HttpPost = sBuffer
                bRet = InternetCloseHandle(hHttpOpenRequest)
            End If

            bRet = InternetCloseHandle(hInternetConnect)
        End If

        bRet = InternetCloseHandle(hInternetOpen)
    End If

End Function

Public Function HttpGet(strURL) As String
    'The API functions used here can only to an HTTP GET
    On Error GoTo Trap

    Dim hInternetSession As Long
    Dim hURLFile As Long
    Dim sReadBuffer            As String * 4096     ' Grab 4k at a time
    Dim sBuffer                As String
    Dim lNumberOfBytesRead     As Long
    Dim bDoLoop As Boolean
    Dim hNewFile As Long
    Dim lngTotalBytes As Long

    hInternetSession = InternetOpen("HttpGet", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)

    If CBool(hInternetSession) Then
        hURLFile = InternetOpenUrl(hInternetSession, strURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

        If CBool(hURLFile) Then
            bDoLoop = True
            While bDoLoop
                sReadBuffer = ""
                bDoLoop = InternetReadFile(hURLFile, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
                sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)

                If Not CBool(lNumberOfBytesRead) Then bDoLoop = False

                DoEvents
                lngTotalBytes = lngTotalBytes + lNumberOfBytesRead
                Debug.Print lngTotalBytes / 1024
            Wend
            HttpGet = sBuffer
        End If
    End If

    InternetCloseHandle (hURLFile)
    InternetCloseHandle (hInternetSession)

    Exit Function

Trap:
    MsgBox Err & " " & Err.Description, vbCritical
    Exit Function

End Function

