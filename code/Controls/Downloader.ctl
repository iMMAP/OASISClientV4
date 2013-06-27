VERSION 5.00
Begin VB.UserControl Downloader 
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Downloader.ctx":0000
   ScaleHeight     =   2385
   ScaleWidth      =   3480
End
Attribute VB_Name = "Downloader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
Event DownloadError(SaveFile As String)
Event DownloadComplete(MaxBytes As Long, SaveFile As String)
Event DownloadAllComplete(FileNotDownload() As String)

Private m_Files As New Collection

Private AsyncPropertyName() As String
Private AsyncStatusCode() As Byte

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)

    On Error Resume Next

        If AsyncProp.BytesMax <> 0 Then
            RaiseEvent DownloadProgress(CLng(AsyncProp.BytesRead), CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
        End If

        Select Case AsyncProp.StatusCode
          Case vbAsyncStatusCodeSendingRequest
            DebugPrint "Attempting to connect " & AsyncProp.Target
          Case vbAsyncStatusCodeConnecting
            DebugPrint "Connecting " & AsyncProp.Status 'show target IP
          Case vbAsyncStatusCodeBeginDownloadData
            DebugPrint "Begin downloading " & AsyncProp.Status 'show temporary saving path
            'Case vbAsyncStatusCodeDownloadingData
            '  DebugPrint "Downloading", AsyncProp.Status 'show target URL
          Case vbAsyncStatusCodeRedirecting
            DebugPrint "Redirecting " & AsyncProp.Status 'show redirected URL
          Case vbAsyncStatusCodeEndDownloadData
            DebugPrint "Download complete " & AsyncProp.Status
          Case vbAsyncStatusCodeError
            DebugPrint "Error...aborting transfer " & AsyncProp.Status
            CancelAsyncRead AsyncProp.PropertyName
        End Select

End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)

  Dim f() As Byte, FN As Long
  Dim i As Integer

    On Error Resume Next

        Select Case AsyncProp.StatusCode
          Case vbAsyncStatusCodeEndDownloadData
            FN = FreeFile
            f = AsyncProp.Value
            DebugPrint "Writing to file " & AsyncProp.PropertyName
            Open AsyncProp.PropertyName For Binary Access Write As #FN
            Put #FN, , f
            Close #FN

            RaiseEvent DownloadComplete(CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)

          Case vbAsyncStatusCodeError
            CancelAsyncRead AsyncProp.PropertyName
            RaiseEvent DownloadError(AsyncProp.PropertyName & " " & AsyncProp.StatusCode & " " & AsyncProp.Status)
        End Select

        For i = 1 To UBound(AsyncPropertyName)
            If AsyncPropertyName(i) = AsyncProp.PropertyName Then
                AsyncStatusCode(i) = AsyncProp.StatusCode
                Exit For
            End If
        Next i

        CheckAllDownloadComplete

End Sub

Private Sub UserControl_Initialize()

    SizeIt
    ReDim AsyncPropertyName(0)
    ReDim AsyncStatusCode(0)

End Sub

Private Sub UserControl_Resize()

    SizeIt

End Sub

Private Sub UserControl_Terminate()

    If UBound(AsyncPropertyName) > 0 Then CancelAllDownload

End Sub

Private Sub SizeIt()

    On Error GoTo ErrorSizeIt
    With UserControl
        .Width = ScaleX(32, vbPixels, vbTwips)
        .Height = ScaleY(32, vbPixels, vbTwips)
    End With

Exit Sub

ErrorSizeIt:
    MsgBox Err & ":Error in call to SizeIt()." _
           & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"

Exit Sub

End Sub

Public Property Get DownloadedFilesAvailable() As Collection
    Set DownloadedFilesAvailable = m_Files
End Property

Public Sub ResetPreviousFiles()
    Set m_Files = New Collection
End Sub

Public Sub BeginDownload(URL As String, SaveFile As String, Optional AsyncReadOptions = vbAsyncReadForceUpdate)
        '<EhHeader>
        On Error GoTo BeginDownload_Err
        '</EhHeader>
        
        m_Files.Add SaveFile

100     UserControl.AsyncRead URL, vbAsyncTypeByteArray, SaveFile, AsyncReadOptions

102     ReDim Preserve AsyncPropertyName(UBound(AsyncPropertyName) + 1)
104     AsyncPropertyName(UBound(AsyncPropertyName)) = SaveFile
106     ReDim Preserve AsyncStatusCode(UBound(AsyncStatusCode) + 1)
108     AsyncStatusCode(UBound(AsyncStatusCode)) = 255


        '<EhFooter>
        Exit Sub

BeginDownload_Err:
MsgBox Err & ":Error in call to BeginDownload()." _
               & vbCrLf & vbCrLf & "Error Description: " & Err.Description & " Error on line: " & Erl, vbCritical, "Warning"

        '</EhFooter>
End Sub

Public Function CancelAllDownload() As Boolean

  Dim i As Integer

    On Error Resume Next

        For i = 1 To UBound(AsyncPropertyName)
            CancelAsyncRead AsyncPropertyName(i)
            DebugPrint "Killing download " & AsyncPropertyName(i)
        Next i

        ReDim AsyncPropertyName(0)
        ReDim AsyncStatusCode(0)

        CancelAllDownload = True

End Function

Private Function CheckAllDownloadComplete()

  Dim i As Integer
  Dim FileNotDownload() As String
  Dim AllDownloadComplete As Boolean

    ReDim FileNotDownload(0)

    AllDownloadComplete = True
    For i = 1 To UBound(AsyncStatusCode)
        If AsyncStatusCode(i) = vbAsyncStatusCodeError Then
            ReDim Preserve FileNotDownload(UBound(FileNotDownload) + 1)
            FileNotDownload(UBound(FileNotDownload)) = AsyncPropertyName(i)
          ElseIf AsyncStatusCode(i) <> vbAsyncStatusCodeEndDownloadData Then
            AllDownloadComplete = False
            Exit For
        End If
    Next i

    If AllDownloadComplete Then
        CancelAllDownload
        RaiseEvent DownloadAllComplete(FileNotDownload)
    End If

End Function
