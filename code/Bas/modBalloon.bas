Attribute VB_Name = "modBalloon"
Option Explicit

Public Sub EasyMove(frm As Form)
  If frm.WindowState <> vbMaximized Then
    ReleaseCapture
    SendMessage frm.hWnd, &HA1, 2, 0&
  End If
End Sub

