VERSION 5.00
Begin VB.Form msgFrm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   3375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   225
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame F1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2310
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   495
      Visible         =   0   'False
      Width           =   4680
      Begin VB.ComboBox ComOptions 
         Height          =   315
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1470
         Visible         =   0   'False
         Width           =   4560
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check2"
         Height          =   465
         Left            =   60
         TabIndex        =   15
         Top             =   1830
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.CommandButton iCmd 
         Caption         =   "OK"
         Height          =   360
         Index           =   1
         Left            =   3690
         TabIndex        =   14
         Top             =   1860
         Width           =   915
      End
      Begin VB.CommandButton iCmd 
         Caption         =   "Cancel"
         Height          =   360
         Index           =   0
         Left            =   2730
         TabIndex        =   13
         Top             =   1860
         Width           =   915
      End
      Begin VB.TextBox iTxt 
         Height          =   315
         Left            =   75
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   1485
         Visible         =   0   'False
         Width           =   4545
      End
      Begin VB.PictureBox Picture2 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   5
         Top             =   0
         Width           =   0
      End
      Begin VB.Label iMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is the message box text. "
         Height          =   195
         Left            =   75
         TabIndex        =   7
         Top             =   465
         Width           =   4530
         WordWrap        =   -1  'True
      End
      Begin VB.Image iPic 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   45
         Top             =   30
         Width           =   240
      End
      Begin VB.Label iTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "This is the Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   375
         TabIndex        =   6
         Top             =   45
         Width           =   4215
      End
   End
   Begin VB.PictureBox pIcons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1050
      Picture         =   "formr.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   2880
      TabIndex        =   21
      Top             =   2940
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.Timer Timer1 
      Left            =   4890
      Top             =   2835
   End
   Begin VB.Frame F1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2310
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   4680
      Begin VB.PictureBox Picture3 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   10
         Top             =   0
         Width           =   0
      End
      Begin VB.Image command1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   4380
         Top             =   45
         Width           =   240
      End
      Begin VB.Label aTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Timed Msgbox"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   420
         TabIndex        =   12
         Top             =   45
         Width           =   3855
      End
      Begin VB.Image aPic 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   45
         Top             =   30
         Width           =   240
      End
      Begin VB.Label aMsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This message will automatcally close in "
         Height          =   195
         Left            =   75
         TabIndex        =   11
         Top             =   465
         Width           =   4530
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame F1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2310
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   510
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Check1"
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   1635
         Visible         =   0   'False
         Width           =   4470
      End
      Begin VB.CommandButton cmd1 
         Height          =   375
         Index           =   2
         Left            =   3495
         TabIndex        =   18
         Top             =   1935
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmd1 
         Height          =   375
         Index           =   1
         Left            =   1815
         TabIndex        =   17
         Top             =   1935
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmd1 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   1920
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox Picture1 
         Height          =   0
         Left            =   0
         ScaleHeight     =   0
         ScaleWidth      =   0
         TabIndex        =   3
         Top             =   0
         Width           =   0
      End
      Begin VB.Label LblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "This is the Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   375
         TabIndex        =   2
         Top             =   45
         Width           =   4215
      End
      Begin VB.Image P1 
         Appearance      =   0  'Flat
         Height          =   240
         Left            =   45
         Top             =   30
         Width           =   240
      End
      Begin VB.Label lblmsg 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is the message box text. "
         Height          =   195
         Left            =   75
         TabIndex        =   1
         Top             =   450
         Width           =   4125
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "msgFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'constants and variables for this form
Private hFRgn As Long 'main msg region
Private hTRgn As Long 'pointer region

Private nAction As Integer
Private DefaultBtn As Integer
Private bColor As OLE_COLOR
Private Const MASK_BUTTONS  As Long = &H7   ' 0000000111 (7)
Private Const MASK_ICONS      As Long = &H70  ' 0001110000 (112)
Private Const MASK_DEFAULTS As Long = &H300 ' 1100000000
Private TwipsX As Integer
Private TwipsY As Integer

'###### used for alphablend  ##############
Private Declare Function GetDC _
                Lib "USER32" (ByVal HWND As Long) As Long
Private Declare Function GetWindowRect _
                Lib "USER32" (ByVal HWND As Long, _
                              lpRect As RECT) As Long
Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long) As Long
Private Declare Function SetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X As Long, _
                             ByVal Y As Long, _
                             ByVal crColor As Long) As Long
Private Const RGN_OR = 2

Private Declare Function CreateRoundRectRgn _
                Lib "gdi32" (ByVal X1 As Long, _
                             ByVal Y1 As Long, _
                             ByVal X2 As Long, _
                             ByVal Y2 As Long, _
                             ByVal X3 As Long, _
                             ByVal Y3 As Long) As Long
Private Declare Function FrameRgn _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hRgn As Long, _
                             ByVal hBrush As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long) As Long
Private Declare Function RoundRect _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal X1 As Long, _
                             ByVal Y1 As Long, _
                             ByVal X2 As Long, _
                             ByVal Y2 As Long, _
                             ByVal X3 As Long, _
                             ByVal Y3 As Long) As Long
Private Declare Function CreateRectRgn _
                Lib "gdi32" (ByVal X1 As Long, _
                             ByVal Y1 As Long, _
                             ByVal X2 As Long, _
                             ByVal Y2 As Long) As Long
Private Declare Function CombineRgn _
                Lib "gdi32" (ByVal hDestRgn As Long, _
                             ByVal hSrcRgn1 As Long, _
                             ByVal hSrcRgn2 As Long, _
                             ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn _
                Lib "USER32" (ByVal HWND As Long, _
                              ByVal hRgn As Long, _
                              ByVal bRedraw As Long) As Long
Private Declare Function Polygon _
                Lib "gdi32" (ByVal hDC As Long, _
                             lpPoint As Any, _
                             ByVal nCount As Long) As Long
Private Declare Function FillRgn _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hRgn As Long, _
                             ByVal hBrush As Long) As Long
Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush _
                Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePolygonRgn _
                Lib "gdi32" (lpPoint As POINTAPI, _
                             ByVal nCount As Long, _
                             ByVal nPolyFillMode As Long) As Long
Private Const ALTERNATE = 1  ' ALTERNATE and WINDING are
Private Const WINDING = 2  ' constants for FillMode.
Private Const BLACKBRUSH = 8  ' Constant for brush type.

Private Declare Function GetCursorPos _
                Lib "USER32" (lpPoint As POINTAPI) As Long
Private Declare Function GetStockObject _
                Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function DeleteObject _
                Lib "gdi32" (ByVal hObject As Long) As Long

'####################
Private Declare Function GetDeviceCaps _
                Lib "gdi32" (ByVal hDC&, _
                             ByVal iCapabilitiy&) As Long
Private Declare Function SendMessageLong _
                Lib "USER32" _
                Alias "SendMessageA" (ByVal HWND As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
Private Declare Function ReleaseCapture _
                Lib "USER32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

' used with the msgbox timeout SetWindowPos Flags to set the balloon ontop
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE As Long = &H1&
Private Const SWP_NOMOVE As Long = &H2&
Private Declare Function SetWindowPos _
                Lib "USER32" (ByVal HWND As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal X As Long, _
                              ByVal Y As Long, _
                              ByVal cX As Long, _
                              ByVal cY As Long, _
                              ByVal wFlags As Long) As Long

'used in LoadIconID
Private Declare Function FillRect _
                Lib "USER32" (ByVal hDC As Long, _
                              lpRect As RECT, _
                              ByVal hBrush As Long) As Long

Private Const API_TRUE As Long = 1&
Private Const RASTERCAPS As Long = 38&
Private Const SIZEPALETTE As Long = 104&
Private Const RC_PALETTE As Long = &H100&
Private Const SRCCOPY As Long = &HCC0020
Private Const SRCPAINT As Long = &HEE0086
Private Const SRCAND As Long = &H8800C6
Private Const NOTSRCCOPY As Long = &H330008

Private Type PALETTEENTRY
    peRed As Byte
    peGreen As Byte
    peBlue As Byte
    peFlags As Byte
End Type
  
Private Type LOGPALETTE256
    palVersion As Integer
    palNumEntries As Integer
    palPalEntry(255) As PALETTEENTRY
End Type
  
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
  
Private Type PICTDESC_BMP
    Size As Long
    Type As Long
        hBmp As Long
        hPal As Long
        reserved As Long
    End Type

    Private Declare Function BitBlt _
                    Lib "gdi32" (ByVal hDestDC&, _
                                 ByVal X&, _
                                 ByVal Y&, _
                                 ByVal nWidth&, _
                                 ByVal nHeight&, _
                                 ByVal hSrcDC&, _
                                 ByVal xSrc&, _
                                 ByVal ySrc&, _
                                 ByVal dwRop&) As Long
    Private Declare Function CreateBitmap _
                    Lib "gdi32" (ByVal nWidth&, _
                                 ByVal nHeight&, _
                                 ByVal nPlanes&, _
                                 ByVal nBitCount&, _
                                 lpBits As Any) As Long
    Private Declare Function CreateCompatibleBitmap _
                    Lib "gdi32" (ByVal hDC&, _
                                 ByVal nWidth&, _
                                 ByVal nHeight&) As Long
    Private Declare Function CreateCompatibleDC _
                    Lib "gdi32" (ByVal hDC&) As Long
    Private Declare Function CreatePalette _
                    Lib "gdi32" (lpLogPalette As LOGPALETTE256) As Long
    Private Declare Function SetBkColor _
                    Lib "gdi32" (ByVal hDC&, _
                                 ByVal crColor&) As Long
    Private Declare Function DeleteDC _
                    Lib "gdi32" (ByVal hDC&) As Long
    Private Declare Function GetSystemPaletteEntries _
                    Lib "gdi32" (ByVal hDC&, _
                                 ByVal wStartIndex&, _
                                 ByVal wNumEntries&, _
                                 lpPaletteEntries As PALETTEENTRY) As Long
    Private Declare Function OleCreatePictureIndirect _
                    Lib "olepro32.dll" (PicDesc As PICTDESC_BMP, _
                                        RefIID As GUID, _
                                        ByVal fPictureOwnsHandle&, _
                                        iPic As IPicture) As Long
    Private Declare Function RealizePalette _
                    Lib "gdi32" (ByVal hDC&) As Long
    Private Declare Function SelectPalette _
                    Lib "gdi32" (ByVal hDC&, _
                                 ByVal hPalette&, _
                                 ByVal bForceBackground&) As Long
  
'Private Declare Function LoadImage Lib "USER32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal dwImageType As Long, ByVal dwDesiredWidth As Long, ByVal dwDesiredHeight As Long, ByVal dwFlags As Long) As Long

Private Function PictureFromDC(ByVal hdcSrc&, _
                               ByVal nLeft&, _
                               ByVal nTop&, _
                               ByVal nWidth&, _
                               ByVal nHeight&) As StdPicture
    '############################################################################
    'creates a standard picture from the passed HDC and returns it to the caller
    '############################################################################

    Dim hDCMemory&, hBmp&, hBmpPrev&, hPal&, hPalPrev&
    Dim fHasPalette&, nPaletteEntries&, LogPal As LOGPALETTE256

    ' create the DC and bitmap we will use
    hDCMemory = CreateCompatibleDC(hdcSrc)
    hBmp = CreateCompatibleBitmap(hdcSrc, nWidth, nHeight)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    ' determine whether or not the video supports 256 color palettes
    nPaletteEntries = GetDeviceCaps(hdcSrc, SIZEPALETTE)
    fHasPalette = GetDeviceCaps(hdcSrc, RASTERCAPS) And RC_PALETTE

    ' if this is 256 color video, we need to create and add a palette to our DC
    If fHasPalette And (nPaletteEntries = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        GetSystemPaletteEntries hdcSrc, 0, 256, LogPal.palPalEntry(0)
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        RealizePalette hDCMemory
    End If

    ' copy the passed image into the local DC
    BitBlt hDCMemory, 0, 0, nWidth, nHeight, hdcSrc, nLeft, nTop, vbSrcCopy

    ' get the bitmap from the DC
    hBmp = SelectObject(hDCMemory, hBmpPrev)

    ' if we created a palette, get it from the DC
    If fHasPalette And (nPaletteEntries = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    ' clean up
    DeleteDC hDCMemory

    ' create a picture from the bitmap and return it to the caller
    Set PictureFromDC = PictureFromBitmap(hBmp, hPal)
    
End Function

Private Function PictureFromBitmap(ByVal hBmp&, _
                                   ByVal hPal&) As StdPicture

    Dim IPictureIID As GUID, iPic As IPicture, tagPic As PICTDESC_BMP
    Dim lpGUID&
  
    ' fill in the IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
    With IPictureIID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B
        .Data4(1) = &HBB
        .Data4(2) = &H0
        .Data4(3) = &HAA
        .Data4(4) = &H0
        .Data4(5) = &H30
        .Data4(6) = &HC
        .Data4(7) = &HAB
    End With
  
    ' set the properties on the picture object
    With tagPic
        .Size = Len(tagPic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With

    ' create a picture that will delete it's bitmap when it is finished with it
    OleCreatePictureIndirect tagPic, IPictureIID, API_TRUE, iPic

    ' return the picture to the caller
    Set PictureFromBitmap = iPic
    
End Function

Function DrawBalloon(NoPointer As Boolean, Optional X As Single, Optional Y As Single, Optional Inty As Integer)
    Dim k As Integer
    Dim hShad As Long, hRect As Long, hPointer As Long

    If NoPointer = False Then
        k = GetPointerStyle(X, Y) 'also sets the position of the form
        hShad = n_GetShadowRect(Inty)
        hRect = n_GetMainRect
        hPointer = n_SetPointer(k) 'decide which pointer to display
    Else 'cetre the form

        With Me
            .Top = (Screen.Height - .Height) / 2
            .Left = (Screen.Width - .Width) / 2
        End With

        hShad = n_GetShadowRect(Inty)
        hRect = n_GetMainRect
        hPointer = CreateRectRgn(0, 0, 0, 0)
    End If

    CombineRgn hRect, hRect, hShad, RGN_OR 'combine shadow and main region
    CombineRgn hRect, hRect, hPointer, RGN_OR 'combine pointer with main and shadow
    'and set all other regions not in the combined region, to be transparent
    Call SetWindowRgn(Me.HWND, hRect, True)

    'cleanup
    DeleteObject hFRgn 'this is the main rectangle region
    DeleteObject hShad
    DeleteObject hPointer
    DeleteObject hRect

End Function

Function msg_Box(ByVal mPrompt As String, ByVal mFlags As MsgBox_Flags, ByVal mCaption As String, Optional NoPointer As Boolean, Optional CheckboxTxt As String, Optional Btn1 As String, Optional Btn2 As String, Optional Btn3 As String, Optional mColor As OLE_COLOR, Optional X As Single, Optional Y As Single, Optional Intensity As Integer) As Integer
                 
    '########### setup the form from the parameters ################
    LblTitle.caption = mCaption
    lblmsg.caption = mPrompt

    If mColor = 0 Then bColor = RGB(255, 255, 204) Else bColor = mColor
    If Intensity <> 0 Then Me.Tag = Intensity

    '######## determine control positions ###########
    With F1(0)
        .BackColor = bColor
        .Top = 30
        .Left = 8
        .Width = Me.ScaleWidth - 30
        .Visible = True
    End With

    lblmsg.Width = (F1(0).Width * TwipsX) - 130 'changing it's width will change it's height also
    F1(0).Height = (lblmsg.Top + lblmsg.Height + 900) / TwipsY 'set the frame height now the msg height has changed

    If CheckboxTxt <> vbNullString Then

        With Check1
            .BackColor = bColor
            .Top = lblmsg.Top + lblmsg.Height + 70
            .Left = lblmsg.Left
            .caption = CheckboxTxt
            .Visible = True
        End With

    End If

    Me.Height = (F1(0).Height + 65) * TwipsY

    '############### Draw the balloon region ###############
    DrawBalloon NoPointer, X, Y, Intensity
    '#########################################################

    'setup the desired message type
    SetupButtons mFlags, Btn1, Btn2, Btn3 'setup the buttons and attach the desired result to the button tags
    SetDisplayIcon mFlags, 0  'decide which icon to display
    SetDefaultButton mFlags 'decide which button is the default

    Me.Show 1

    'return the value of the button clicked + 100 if the checkbox was also selected
    If Check1.Value = 1 Then nAction = nAction + 100
    msg_Box = nAction 'the return value is the buttons tag value
    Unload Me

End Function

Private Function AlphaBlend1(RGB1 As Long, _
                             RGB2 As Long, _
                             Alpha As Integer, _
                             Optional Intensity As Integer) As Long

    Dim Alpha1 As Byte, Alpha2 As Byte
    Dim lngSrcRed&, lngDestRed&, lngSrcGreen&, lngDestGreen&, lngSrcBlue&, lngDestBlue&

    Const nHex As Long = &HFF
    
    If Intensity > 75 Then Intensity = 75
    Alpha1 = Alpha + Intensity
    Alpha2 = 256 - Alpha
    
    lngSrcRed& = (RGB1 And nHex) * Alpha1
    lngDestRed& = (RGB2 And nHex) * Alpha2
    lngSrcGreen& = ((RGB1 \ 256) And nHex) * Alpha1
    lngDestGreen& = ((RGB2 \ 256) And nHex) * Alpha2
    lngSrcBlue& = ((RGB1 \ 65536) And nHex) * Alpha1
    lngDestBlue& = ((RGB2 \ 65536) And nHex) * Alpha2
    
    AlphaBlend1 = RGB((lngDestRed& + lngSrcRed&) \ 256, (lngDestGreen& + lngSrcGreen&) \ 256, (lngDestBlue& + lngSrcBlue&) \ 256)

End Function

Function n_GetMainRect() As Long
    'creates a rounded rectangle and passes back the region for this rectangle
    
    Dim hFFBrush As Long, hrgnTmp1&, hRgnTmp2&
    'Clear the form
    Const CN As Long = 20 ' this value changes the rounded corners
    
    'create a region to enhance the corners
    hFRgn = CreateRoundRectRgn(0, 20, Me.ScaleWidth - 14, Me.ScaleHeight - 32, CN, CN)
    hFFBrush = CreateSolidBrush(RGB(70, 70, 70)) 'create a dark grey brush
    SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
    FrameRgn Me.hDC, hFRgn, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
    
    'draw the yellow rounded rectangle with black border
    hFFBrush = CreateSolidBrush(bColor) 'Create a new solid brush
    SelectObject Me.hDC, hFFBrush 'Select this  colored brush into our form's device context
    'draw a filled rounded rectangle which is slightly offset to the left
    RoundRect Me.hDC, 0, 20, Me.ScaleWidth - 15, Me.ScaleHeight - 33, CN + 4, CN + 4
    
    n_GetMainRect = hFRgn
    
    'Clean up
    DeleteObject hFFBrush
    DeleteObject hrgnTmp1
    DeleteObject hRgnTmp2
    
    'the hfregion is deleted after is has been combined and set
End Function
Function MsgAuto_box(ByVal mPrompt As String, ByVal mFlags As MsgBox_Flags, ByVal mCaption As String, ByVal nTime As Integer, Optional NoPointer As Boolean, Optional mColor As OLE_COLOR, Optional X As Single, Optional Y As Single, Optional Intensity As Integer) As String

    Dim hPointer As Long
    Dim hRect As Long 'handle to the rectangle region
    Dim hRgn1 As Long 'handle to the pointer region
    Dim k%

    '########### setup the form from the parameters ################
    aTitle.caption = mCaption
    aMsg.caption = mPrompt

    If IsEmpty(mColor) Then bColor = RGB(255, 255, 204) Else bColor = mColor
    If Intensity <> 0 Then Me.Tag = Intensity

    '######## determine control positions ###########
    With F1(2)
        .BackColor = bColor
        .Top = 30
        .Left = 8
        .Width = Me.ScaleWidth - 30
        .Visible = True
    End With

    lblmsg.Width = (F1(2).Width * TwipsX) - 130 'changing it's width will change it's height also
    F1(2).Height = (aMsg.Top + aMsg.Height + 300) / TwipsY 'set the frame height now the msg height has changed
    'position the close icon and display the icon
    command1.Top = aPic.Top
    command1.Left = (F1(2).Width * TwipsX) - command1.Width - 50

    Me.Height = (F1(2).Height + F1(2).Top + 40) * TwipsY

    '############### Create the balloon region ###############
    DrawBalloon NoPointer, X, Y, Intensity

    'setup the desired message type
    SetDisplayIcon mFlags, 2  'decide which icon to display
    LoadIconID command1, 9, 16, bColor

    Call SetWindowPos(Me.HWND, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE)
    Me.Show

    'start the timer
    'SetTimer Me.HWND, &H5004, nTime * 1000, AddressOf TimerProc
    With Timer1
        .Interval = nTime * 1000
        .Enabled = True
    End With

End Function

Function n_GetShadowRect(Optional Inty As Integer) As Long

    'creates the region with a drop shadow around the edges
    Dim rc As RECT, MyRec As RECT, TDC As Long, ScrDC&, AlphaLevel As Integer, hShad&
    'get the screen coords of the picturebox
    Call GetWindowRect(Me.HWND, rc) 'this gives the rectangle of the window in screen pixels

    With MyRec
        MyRec.Right = rc.Right - rc.Left - 5
        MyRec.Bottom = rc.Bottom - rc.Top - 23
        MyRec.Top = 23
        MyRec.Left = 5
    End With

    hShad = CreateRectRgn(MyRec.Left, MyRec.Top, MyRec.Right, MyRec.Bottom)

    ScrDC = GetDC(0)
    TDC = Me.hDC
    Dim n%, k%, i%
    AlphaLevel = 0

    For n = 17 To 0 Step -1

        If (n + 1) Mod 2 = 0 Then AlphaLevel = AlphaLevel + 20

        For i = MyRec.Left + n To MyRec.Right - n
            SetPixel TDC, i, MyRec.Bottom - n, AlphaBlend1(GetPixel(ScrDC, rc.Left + i, rc.Bottom - n - 23), &H0, AlphaLevel, Inty)
        Next i

        For i = MyRec.Top + n To MyRec.Bottom - n
            SetPixel TDC, MyRec.Right - n, i, AlphaBlend1(GetPixel(ScrDC, rc.Right - n - 5, rc.Top + i), &H0, AlphaLevel, Inty)
        Next i
    Next n
    
    hShad = CreateRoundRectRgn(MyRec.Left + 6, MyRec.Top + 6, MyRec.Right, MyRec.Bottom, 45, 35)
    n_GetShadowRect = hShad
End Function

Sub n_RefreshShadow(Inty As Integer)

    Dim hRect As Long 'handle to the rectangle region
    Dim hShad As Long

    hShad = n_GetShadowRect(Inty)
    hRect = n_GetMainRect
    CombineRgn hRect, hRect, hShad, RGN_OR 'combine shadow and main region
    Call SetWindowRgn(Me.HWND, hRect, True) 'set the region

End Sub

Sub SetupButtons(ByVal MF As MsgBox_Flags, _
                 Optional b1$, _
                 Optional b2$, _
                 Optional b3$)
    'selects which bits are set and configures the buttons
    'should check bits in order of value
    'MB_OK = &H0&
    'MB_OKCANCEL = &H1&
    'MB_YESNO = &H4&
    'MB_YESNOCANCEL = &H3&
    'MB_ABORTRETRYIGNORE = &H2&
    'MB_RETRYCANCEL = &H5&
    'MB_BUTTONSNOTUSED = &H6& '#######
   
    Dim mButtonsType As Long, n%

    For n = 0 To 2
        cmd1(n).Top = (F1(0).Height * TwipsY) - cmd1(n).Height - 100 'set the command button top position
    Next n

    mButtonsType = (MF And MASK_BUTTONS)

    Select Case mButtonsType

        Case vbOKCancel

            With cmd1(0):   .caption = "OK":     .Visible = True:  .Tag = vbOK:     End With
                With cmd1(1):   .caption = "Cancel": .Visible = True:  .Tag = vbCancel: End With
                n = 2

            Case vbAbortRetryIgnore

                With cmd1(0):   .caption = "Abort":  .Visible = True:  .Tag = vbAbort:  End With
                    With cmd1(1):   .caption = "Retry":  .Visible = True:  .Tag = vbRetry:  End With
                        With cmd1(2):   .caption = "Ignore": .Visible = True:  .Tag = vbIgnore: End With
                        n = 3

                    Case vbYesNoCancel

                        With cmd1(0):   .caption = "Yes":    .Visible = True:  .Tag = vbYes:    End With
                            With cmd1(1):   .caption = "No":     .Visible = True:  .Tag = vbNo:     End With
                                With cmd1(2):   .caption = "Cancel": .Visible = True:  .Tag = vbCancel: End With
                                n = 3

                            Case vbYesNo

                                With cmd1(0):   .caption = "Yes":    .Visible = True:  .Tag = vbYes:    End With
                                    With cmd1(1):   .caption = "No":     .Visible = True:  .Tag = vbNo:     End With
                                        n = 2

                                    Case vbRetryCancel

                                        With cmd1(0):   .caption = "Retry":  .Visible = True:  .Tag = vbRetry:  End With
                                            With cmd1(1):   .caption = "Cancel": .Visible = True:  .Tag = vbCancel: End With
                                            n = 2

                                        Case vbCustomButtons 'custom text on the buttons

                                            'note the return value will depend on the button selected 1,2 or 3
                                            If b1 <> "" Then

                                                With cmd1(0):   .caption = b1:   .Visible = True:  .Tag = 1:        End With
                                                    n = 1
                                                End If

                                                If b2 <> "" Then

                                                    With cmd1(1):   .caption = b2:   .Visible = True:  .Tag = 2:        End With
                                                        n = 2
                                                    End If

                                                    If b3 <> "" Then

                                                        With cmd1(2):   .caption = b3:   .Visible = True:  .Tag = 3:        End With
                                                            n = 3
                                                        End If
        
                                                    Case Else 'defaults to one button

                                                        With cmd1(0)
                                                            .caption = "OK"
                                                            .Visible = True
                                                            .Tag = vbOK
                                                        End With

                                                        n = 1
                                                End Select

                                                Select Case n

                                                    Case 3
                                                        cmd1(0).Left = (cmd1(1).Left - cmd1(0).Width) / 2
                                                        cmd1(1).Left = ((F1(0).Width * TwipsX) - cmd1(1).Width) / 2
                                                        cmd1(2).Left = (F1(0).Width * TwipsX) - cmd1(1).Width - cmd1(0).Left

                                                    Case 2
                                                        cmd1(0).Left = (F1(0).Width * TwipsX) / 3 - cmd1(0).Width / 2
                                                        cmd1(1).Left = (F1(0).Width * TwipsX) - cmd1(0).Left - cmd1(0).Width

                                                    Case 1
                                                        cmd1(0).Left = ((F1(0).Width * TwipsX) - cmd1(1).Width) / 2
                                                End Select

                                            End Sub

Function GetPointerStyle(Optional X As Single, Optional Y As Single) As Integer
    Dim k%, Pt As POINTAPI, sPt As POINTAPI

    If (X <> 0 And Y <> 0) Then
        Pt.X = X
        Pt.Y = Y
    Else
        GetCursorPos Pt 'get the cursor position in screen pixels
    End If

    sPt = GetMaxScreenSize 'get the max screen size in pixels

    If (Pt.X <= sPt.X / 2) And (Pt.Y <= sPt.Y / 2) Then
        k = 4:      Pt.X = Pt.X - 42
    ElseIf (Pt.X > sPt.X / 2) And (Pt.Y <= sPt.Y / 2) Then
        k = 2:      Pt.X = Pt.X + 42 - Me.ScaleWidth
    ElseIf (Pt.X > sPt.X / 2) And (Pt.Y > sPt.Y / 2) Then
        k = 1:      Pt.X = Pt.X + 42 - Me.ScaleWidth:      Pt.Y = Pt.Y - Me.ScaleHeight
    Else
        k = 3:      Pt.X = Pt.X - 42:                      Pt.Y = Pt.Y - Me.ScaleHeight
    End If

    Me.Top = Pt.Y * TwipsY
    Me.Left = Pt.X * TwipsX

    GetPointerStyle = k

End Function

Sub SetDefaultButton(MF As MsgBox_Flags)
    Dim mDefButtonType As Long
    mDefButtonType = (MF And MASK_DEFAULTS)

    Select Case mDefButtonType

        Case vbDefaultButton1
            DefaultBtn = 1

        Case vbDefaultButton2
            DefaultBtn = 2

        Case vbDefaultButton3
            DefaultBtn = 3
    End Select

End Sub

Private Sub cmd1_Click(Index As Integer)

    nAction = CInt(cmd1(Index).Tag) 'integer return determines which button was pressed
    Me.Hide 'allows the msg_box function to continue from the show command

End Sub

Private Sub Command1_Click()

    'if we close it early then stop the timer
    Timer1.Enabled = False
    Unload Me 'close the auto msgbox

End Sub

Private Sub command1_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    LoadIconID command1, 11, 16, bColor
End Sub

Private Sub command1_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    'this is not a good method but will do for demonstration purposes!
    If command1.Tag <> "1" Then
        LoadIconID command1, 10, 16, bColor
        command1.Tag = "1"
    End If

End Sub

Private Sub command1_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)
    LoadIconID command1, 10, 16, bColor
End Sub

Private Sub F1_MouseMove(Index As Integer, _
                         Button As Integer, _
                         Shift As Integer, _
                         X As Single, _
                         Y As Single)

    'allow the form mouse move event to decide if draging is enabled
    Form_MouseDown Button, Shift, X, Y

    If Index = 2 Then
        If command1.Tag = "1" Then
            LoadIconID command1, 9, 16, bColor
            command1.Tag = ""
        End If
    End If

End Sub

Private Sub Form_Activate()

    'set a default focus button for the messsage box form
    If F1(0).Visible Then
        If DefaultBtn > 0 Then cmd1(DefaultBtn - 1).SetFocus
    End If

End Sub

Private Sub TransparentBlt(DstDC, _
                           SrcDC, _
                           SrcRect As RECT, _
                           DstX, _
                           DstY, _
                           TransColor As Long)
    'eg
    'TransparentBlt (HDC of destination) , (HDC of current bmp), (Size of BMP), (X-coord), (y-coord), (Transparent Colour)
  
    'DstDC=Device context into which image must be drawn transparently
    'OutDstDC=Device context into image is actually drawn, even though it is made transparent in terms of DstDC
    'Src=Device context of source to be made transparent in color TransColor
    'SrcRect=rectangular region within SrcDC to be made transparent in terms of DstDC, and drawn to OutDstDC
    'DstX, DstY =coordinates in OutDstDC (and DstDC) where tranparent bitmap must go
  
    Rem In most cases, OutDstDC and DstDC will be the same
  
    Dim nRet As Long, w As Integer, h As Integer
    Dim MonoMaskDC As Long, hMonoMask As Long
    Dim MonoInvDC As Long, hMonoInv As Long
    Dim ResultDstDC As Long, hResultDst As Long
    Dim ResultSrcDC As Long, hResultSrc As Long
    Dim hPrevMask As Long, hPrevInv As Long, hPrevSrc As Long, hPrevDst As Long
    w = SrcRect.Right - SrcRect.Left + 1
    h = SrcRect.Bottom - SrcRect.Top + 1
  
    'create monochrome mask and inverse masks
    MonoMaskDC = CreateCompatibleDC(DstDC)
    MonoInvDC = CreateCompatibleDC(DstDC)
    hMonoMask = CreateBitmap(w, h, 1, 1, ByVal 0&)
    hMonoInv = CreateBitmap(w, h, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    hPrevInv = SelectObject(MonoInvDC, hMonoInv)
  
    'create keeper DCs and bitmaps
    ResultDstDC = CreateCompatibleDC(DstDC)
    ResultSrcDC = CreateCompatibleDC(DstDC)
    hResultDst = CreateCompatibleBitmap(DstDC, w, h)
    hResultSrc = CreateCompatibleBitmap(DstDC, w, h)
    hPrevDst = SelectObject(ResultDstDC, hResultDst)
    hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
   
    'copy src to monochrome mask
    Dim OldBC As Long
    OldBC = SetBkColor(SrcDC, TransColor)
    nRet = BitBlt(MonoMaskDC, 0, 0, w, h, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
    TransColor = SetBkColor(SrcDC, OldBC)
  
    'create inverse of mask
    nRet = BitBlt(MonoInvDC, 0, 0, w, h, MonoMaskDC, 0, 0, vbNotSrcCopy)
    'get background
    nRet = BitBlt(ResultDstDC, 0, 0, w, h, DstDC, DstX, DstY, vbSrcCopy)
    'AND with Monochrome mask
    nRet = BitBlt(ResultDstDC, 0, 0, w, h, MonoMaskDC, 0, 0, vbSrcAnd)
    'get overlapper
    nRet = BitBlt(ResultSrcDC, 0, 0, w, h, SrcDC, SrcRect.Left, SrcRect.Top, vbSrcCopy)
    'AND with inverse monochrome mask
    nRet = BitBlt(ResultSrcDC, 0, 0, w, h, MonoInvDC, 0, 0, vbSrcAnd)
    'XOR these two
    nRet = BitBlt(ResultDstDC, 0, 0, w, h, ResultSrcDC, 0, 0, vbSrcInvert)
   
    'output results normal or stretched
    nRet = BitBlt(DstDC, DstX, DstY, w, h, ResultDstDC, 0, 0, vbSrcCopy)
  
    'clean up
    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask
    hMonoInv = SelectObject(MonoInvDC, hPrevInv)
    DeleteObject hMonoInv
    hResultDst = SelectObject(ResultDstDC, hPrevDst)
    DeleteObject hResultDst
    hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
    DeleteObject hResultSrc
    DeleteDC MonoMaskDC
    DeleteDC MonoInvDC
    DeleteDC ResultDstDC
    DeleteDC ResultSrcDC
 
End Sub

Private Sub LoadIconID(Obj1 As Object, _
                       startcell As Integer, _
                       nWidth%, _
                       BKcol&)

    'Obj1   :   the objects that require the BMP's to be extracted into their picture property
    'nwidth :   the width of the  cells
    'Bkcol  :   the backcolor required

    Dim REShdc&, RESbm&, NEWbm&, NEWhdc&, bmR As RECT
    Dim hBrush&

    'get the dimensions of the bmp file (icon cells)
    REShdc = pIcons.hDC

    With bmR
        .Top = 0:   .Left = 0:    .Bottom = pIcons.ScaleHeight / TwipsY:  .Right = pIcons.ScaleWidth / TwipsX
    End With
    
    'create a dc to hold the icon pictures when made transparent
    NEWhdc = CreateCompatibleDC(REShdc)
    NEWbm& = SelectObject(NEWhdc, CreateCompatibleBitmap(REShdc, bmR.Right, bmR.Bottom))
    
    '  make the background of the NEWbm
    hBrush = CreateSolidBrush(BKcol)
    Call FillRect(NEWhdc, bmR, hBrush)
    
    'create the transparent picture
    TransparentBlt NEWhdc, REShdc, bmR, 0, 0, RGB(0, 255, 0)

    'Now the whole strip is transparaently created in NEWhdc
    'extract the cell as a vb picture and place in the ojects picture property
    Obj1.Picture = PictureFromDC(NEWhdc, (startcell - 1) * nWidth, 0, nWidth, nWidth)
    
    'MUST CleanUP!!
    Call DeleteObject(hBrush)
    Call DeleteObject(NEWbm)
    Call DeleteDC(NEWhdc)

End Sub

Private Function GetMaxScreenSize() As POINTAPI

    Dim ScrRect As POINTAPI

    With ScrRect
        .X = GetDeviceCaps(GetDC(0), 8)
        .Y = GetDeviceCaps(GetDC(0), 10)
    End With

    GetMaxScreenSize = ScrRect

End Function

Public Function input_box(ByVal mPrompt As String, _
                          ByVal mFlags As MsgBox_Flags, _
                          ByVal mCaption As String, _
                          Optional DefaultText As String, _
                          Optional NoPointer As Boolean, _
                          Optional CheckboxTxt As String, _
                          Optional mColor As OLE_COLOR, _
                          Optional X As Single, _
                          Optional Y As Single, _
                          Optional Intensity As Integer, _
                          Optional sComboString As String) As String

    Dim hPointer As Long
    Dim hRect As Long 'handle to the rectangle region
    Dim hRgn1 As Long 'handle to the pointer region
    Dim k As Long
    Dim i As Integer
    Dim sArr() As String
    Dim lTop As Long
    Dim lWidth As Long
    Dim lLeft As Long
    
    '########### setup the form from the parameters ################
    iTitle.caption = mCaption
    iMsg.caption = mPrompt

    If (IsEmpty(mColor) Or mColor = 0) Then bColor = RGB(255, 255, 204) Else bColor = mColor
    If Intensity <> 0 Then Me.Tag = Intensity

    '######## determine control positions ###########
    With F1(1)
        .BackColor = bColor
        .Top = 30
        .Left = 8
        .Width = Me.ScaleWidth - 30
        .Visible = True
    End With

    iMsg.Width = (F1(1).Width * TwipsX) - 130 'changing it's width will change it's height also

    lTop = iMsg.Top + iMsg.Height + 300
    lLeft = iMsg.Left
    lWidth = iMsg.Width
            
    If Len(sComboString) > 0 Then
        sArr = Split(sComboString, ",")
        
        With ComOptions
        
            .Clear
        
            For i = LBound(sArr) To UBound(sArr)
        
                .AddItem sArr(i)
        
            Next
        
            If .ListCount > 0 Then .ListIndex = 0
        
            .Top = lTop
            .Left = lLeft
            .Width = lWidth
            .Visible = True
            iCmd(0).Top = .Top + .Height + 70
            iCmd(1).Top = .Top + .Height + 70
            iCmd(1).Left = .Left + .Width - iCmd(1).Width
            iCmd(0).Left = iCmd(1).Left - iCmd(1).Width - 70
        End With

    Else

        With iTxt
            .Top = lTop
            .Left = lLeft
            .Width = lWidth
            .Visible = True
        End With

        iCmd(0).Top = iTxt.Top + iTxt.Height + 70
        iCmd(1).Top = iTxt.Top + iTxt.Height + 70
        iCmd(1).Left = iTxt.Left + iTxt.Width - iCmd(1).Width
        iCmd(0).Left = iCmd(1).Left - iCmd(1).Width - 70

    End If

    If CheckboxTxt <> vbNullString Then

        With Check2
            .BackColor = bColor
            .Top = lTop + iTxt.Height + 70
            .Left = lLeft
            .Width = iCmd(0).Left - .Left - 10
            .caption = CheckboxTxt
            .Visible = True
        End With

    End If

    F1(1).Height = (iCmd(0).Top + iCmd(0).Height + 115) / TwipsY 'set the frame height now the msg height has changed

    Me.Height = (F1(1).Height + 65) * TwipsY
    iTxt.Text = DefaultText
    '############### Create the balloon region ###############
    DrawBalloon NoPointer, X, Y, Intensity

    'setup the desired message type

    SetDisplayIcon mFlags, 1 'decide which icon to display

    Me.Show 1
    

    'return the value of the button clicked + 100 if the checkbox was also selected
    If Check2.Value = 1 Then
        iTxt = iTxt & "¬"
    End If
    
    If nAction = -1 Then
        input_box = ""
    Else
        If Len(sComboString) > 0 Then
            input_box = ComOptions.List(ComOptions.ListIndex)
        Else
            input_box = iTxt.Text
        End If
    End If
    
    Unload Me

End Function

Private Sub Form_Load()
    'initiate the sizing constants
    TwipsX = Screen.TwipsPerPixelX
    TwipsY = Screen.TwipsPerPixelY

End Sub

Private Sub Form_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
    Dim hRgn As Long

    If Button = vbLeftButton Then
        If F1(2).Visible = False Then
                       
            hRgn = n_GetMainRect
            Call SetWindowRgn(Me.HWND, hRgn, True)
            DeleteObject hRgn
            Me.Line (5, 20)-(Me.ScaleWidth - 5, 20)
            Me.Line (5, Me.ScaleHeight - 25)-(Me.ScaleWidth - 5, Me.ScaleHeight - 25)
            'DoEvents
            ' ######## Fake a mouse down on the titlebar so form can be moved...########
            ReleaseCapture
            SendMessageLong Me.HWND, WM_NCLBUTTONDOWN, HTCAPTION, 0&

            '###########################################################################
            DoEvents
            n_RefreshShadow CInt(Val(Me.Tag))
      
        End If
    End If
    
End Sub

Sub SetDisplayIcon(MF As MsgBox_Flags, _
                   ID As Integer)

    Dim MyObj As Object
    Dim mIconType As Long

    Select Case ID

        Case 0:  Set MyObj = P1 'msgbox display

        Case 1:  Set MyObj = iPic 'inputbox display

        Case 2: Set MyObj = aPic
    End Select

    mIconType = (MF And MASK_ICONS)

    Select Case mIconType
   
        Case vbCritical
            LoadIconID MyObj, 1, 16, bColor

        Case vbQuestion
            LoadIconID MyObj, 4, 16, bColor
        
        Case vbExclamation
            LoadIconID MyObj, 2, 16, bColor
 
        Case vbInformation
            LoadIconID MyObj, 3, 16, bColor

        Case vbUserIcon
            LoadIconID MyObj, 5, 16, bColor

        Case vbSecurityIcon
            LoadIconID MyObj, 6, 16, bColor

        Case vbFindIcon
            LoadIconID MyObj, 7, 16, bColor

        Case Else 'no icon so adjust the position of the title
            MyObj.Visible = False

            Select Case ID

                Case 0:    LblTitle.Left = lblmsg.Left

                Case 1:    iTitle.Left = iMsg.Left

                Case 2:    aTitle.Left = aMsg.Left
            End Select
    End Select

    Set MyObj = Nothing
End Sub

Function SetPointer(idx As Integer) As Long
    Dim poly(1 To 3) As POINTAPI

    'this sets the which pointer should be displayed depending on the mouse POINTAPIinates
    'idx = 1 to 4
    'hopefully this function will return the region for this triangle which can be
    'combined with the rounded rectangle region for faster transparency purposes.
    Dim MyPic As StdPicture, hFFBrush&, hRgnTmp&, hrgnTmp1&

    'Set MyPic = LoadResPicture(72, 0) 'load the pointer picture
    Select Case idx

        Case 1 'bottom right
            hFFBrush = CreateSolidBrush(RGB(140, 140, 140)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            hRgnTmp = CreateRoundRectRgn(Me.ScaleWidth - 39, Me.ScaleHeight - 23, Me.ScaleWidth - 35, Me.ScaleHeight - 1, 3, 3) 'the shaded region
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
            hFFBrush = CreateSolidBrush(RGB(100, 100, 100)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            hrgnTmp1 = CreateRoundRectRgn(Me.ScaleWidth - 39, Me.ScaleHeight - 24, Me.ScaleWidth - 36, Me.ScaleHeight, 5, 5)  'the shaded region
            FrameRgn Me.hDC, hrgnTmp1&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
            hFFBrush = CreateSolidBrush(RGB(0, 0, 0)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            poly(1).X = Me.ScaleWidth - 61:     poly(1).Y = Me.ScaleHeight - 25
            poly(2).X = Me.ScaleWidth - 38:     poly(2).Y = Me.ScaleHeight - 25
            poly(3).X = Me.ScaleWidth - 38:     poly(3).Y = Me.ScaleHeight - 1
            hTRgn = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            FrameRgn Me.hDC, hTRgn, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
            'combine the region
            CombineRgn hTRgn, hTRgn, hRgnTmp, RGN_OR
            
            'draw the yellow area
            poly(1).X = Me.ScaleWidth - 62:     poly(1).Y = Me.ScaleHeight - 27
            poly(2).X = Me.ScaleWidth - 39:     poly(2).Y = Me.ScaleHeight - 27
            poly(3).X = Me.ScaleWidth - 39:     poly(3).Y = Me.ScaleHeight - 3
            hRgnTmp& = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            hFFBrush = CreateSolidBrush(bColor) 'create a yellow brush
            SelectObject Me.hDC, hFFBrush 'Select this brush
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
        Case 2 'top right
        
            hFFBrush = CreateSolidBrush(RGB(140, 140, 140))
            SelectObject Me.hDC, hFFBrush
            hRgnTmp = CreateRoundRectRgn(Me.ScaleWidth - 38, 2, Me.ScaleWidth - 35, 21, 0, 0)  'the shaded region
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
                
            hFFBrush = CreateSolidBrush(RGB(0, 0, 0)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            poly(1).X = Me.ScaleWidth - 61:     poly(1).Y = 20
            poly(2).X = Me.ScaleWidth - 38:     poly(2).Y = 20
            poly(3).X = Me.ScaleWidth - 38:     poly(3).Y = -2
            hTRgn = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            FrameRgn Me.hDC, hTRgn, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
            CombineRgn hTRgn, hTRgn, hRgnTmp, RGN_OR 'combine the region
            
            hFFBrush = CreateSolidBrush(RGB(100, 100, 100)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            hrgnTmp1 = CreateRoundRectRgn(Me.ScaleWidth - 38, 0, Me.ScaleWidth - 36, 21, 0, 0)  'the shaded region
            FrameRgn Me.hDC, hrgnTmp1&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
                    
            CombineRgn hTRgn, hTRgn, hrgnTmp1, RGN_OR
            
            'draw the yellow region on top of this region
            poly(1).X = Me.ScaleWidth - 62:     poly(1).Y = 22
            poly(2).X = Me.ScaleWidth - 39:     poly(2).Y = 22
            poly(3).X = Me.ScaleWidth - 39:     poly(3).Y = 0
            hRgnTmp& = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            hFFBrush = CreateSolidBrush(bColor) 'create a yellow brush
            SelectObject Me.hDC, hFFBrush 'Select this brush
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
        Case 3 'bottom left
        
            hFFBrush = CreateSolidBrush(RGB(0, 0, 0)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            poly(1).X = 41:     poly(1).Y = Me.ScaleHeight - 25
            poly(2).X = 65:     poly(2).Y = Me.ScaleHeight - 25
            poly(3).X = 41:     poly(3).Y = Me.ScaleHeight - 1
            hTRgn = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            FrameRgn Me.hDC, hTRgn, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
        
            'draw the yellow region on top of this region
            poly(1).X = 42:     poly(1).Y = Me.ScaleHeight - 27
            poly(2).X = 66:     poly(2).Y = Me.ScaleHeight - 27
            poly(3).X = 42:     poly(3).Y = Me.ScaleHeight - 3
            hRgnTmp& = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            hFFBrush = CreateSolidBrush(bColor) 'create a yellow brush
            SelectObject Me.hDC, hFFBrush 'Select this brush
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
        
        Case 4 'top left
            hFFBrush = CreateSolidBrush(RGB(0, 0, 0)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            poly(1).X = 41:     poly(1).Y = 20
            poly(2).X = 62:     poly(2).Y = 20
            poly(3).X = 41:     poly(3).Y = 0
            hTRgn = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            FrameRgn Me.hDC, hTRgn, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
            'draw the yellow region on top of this region
            poly(1).X = 42:     poly(1).Y = 22
            poly(2).X = 61:     poly(2).Y = 22
            poly(3).X = 42:     poly(3).Y = 2
            hRgnTmp& = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            hFFBrush = CreateSolidBrush(bColor) 'create a yellow brush
            SelectObject Me.hDC, hFFBrush 'Select this brush
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
        
    End Select
    
    Set MyPic = Nothing
    SetPointer = hTRgn
    
    DeleteObject hRgnTmp
    DeleteObject hrgnTmp1
    DeleteObject hFFBrush&

End Function

Function n_SetPointer(idx As Integer) As Long
    Dim poly(1 To 3) As POINTAPI

    'this sets the which pointer should be displayed depending on the mouse POINTAPIinates
    'idx = 1 to 4
    'hopefully this function will return the region for this triangle which can be
    'combined with the rounded rectangle region for faster transparency purposes.
    Dim MyPic As StdPicture, hFFBrush&, hRgnTmp&, hrgnTmp1&

    'Set MyPic = LoadResPicture(72, 0) 'load the pointer picture
    Select Case idx

        Case 1 'bottom right
            hFFBrush = CreateSolidBrush(RGB(0, 0, 0)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            poly(1).X = Me.ScaleWidth - 61:     poly(1).Y = Me.ScaleHeight - 34
            poly(2).X = Me.ScaleWidth - 36:     poly(2).Y = Me.ScaleHeight - 34
            poly(3).X = Me.ScaleWidth - 36:     poly(3).Y = Me.ScaleHeight - 9
            hTRgn = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            FrameRgn Me.hDC, hTRgn, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
                        
            'draw the yellow area
            poly(1).X = Me.ScaleWidth - 62:     poly(1).Y = Me.ScaleHeight - 35
            poly(2).X = Me.ScaleWidth - 37:     poly(2).Y = Me.ScaleHeight - 35
            poly(3).X = Me.ScaleWidth - 37:     poly(3).Y = Me.ScaleHeight - 11
            hRgnTmp& = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            hFFBrush = CreateSolidBrush(bColor) 'create a yellow brush
            SelectObject Me.hDC, hFFBrush 'Select this brush
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
        Case 2 'top right
                        
            hFFBrush = CreateSolidBrush(RGB(0, 0, 0)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            poly(1).X = Me.ScaleWidth - 58:     poly(1).Y = 20
            poly(2).X = Me.ScaleWidth - 38:     poly(2).Y = 20
            poly(3).X = Me.ScaleWidth - 38:     poly(3).Y = 0
            hTRgn = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            FrameRgn Me.hDC, hTRgn, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
            'draw the yellow region on top of this region
            poly(1).X = Me.ScaleWidth - 59:     poly(1).Y = 22
            poly(2).X = Me.ScaleWidth - 39:     poly(2).Y = 22
            poly(3).X = Me.ScaleWidth - 39:     poly(3).Y = 2
            hRgnTmp& = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            hFFBrush = CreateSolidBrush(bColor) 'create a yellow brush
            SelectObject Me.hDC, hFFBrush 'Select this brush
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            
        Case 3 'bottom left
        
            hFFBrush = CreateSolidBrush(RGB(0, 0, 0)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            poly(1).X = 41:     poly(1).Y = Me.ScaleHeight - 34
            poly(2).X = 66:     poly(2).Y = Me.ScaleHeight - 34
            poly(3).X = 41:     poly(3).Y = Me.ScaleHeight - 9
            hTRgn = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            FrameRgn Me.hDC, hTRgn, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
        
            'draw the yellow region on top of this region
            poly(1).X = 42:     poly(1).Y = Me.ScaleHeight - 35
            poly(2).X = 66:     poly(2).Y = Me.ScaleHeight - 35
            poly(3).X = 42:     poly(3).Y = Me.ScaleHeight - 11
            hRgnTmp& = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            hFFBrush = CreateSolidBrush(bColor) 'create a yellow brush
            SelectObject Me.hDC, hFFBrush 'Select this brush
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            'FillRgn Me.hDC, hRgnTmp&, hFFBrush
        
        Case 4 'top left
            '(0, 20, Me.ScaleWidth - 14, Me.ScaleHeight - 32, CN, CN)
        
            hFFBrush = CreateSolidBrush(RGB(0, 0, 0)) 'create a black brush
            SelectObject Me.hDC, hFFBrush 'Select this brush into our form's device context
            poly(1).X = 40:     poly(1).Y = 20
            poly(2).X = 60:     poly(2).Y = 20
            poly(3).X = 40:     poly(3).Y = 0
            hTRgn = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            FrameRgn Me.hDC, hTRgn, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            'FillRgn Me.hDC, hTRgn, hFFBrush
            
            'draw the yellow region on top of this region
            poly(1).X = 41:     poly(1).Y = 21
            poly(2).X = 59:     poly(2).Y = 21
            poly(3).X = 41:     poly(3).Y = 3
            hRgnTmp& = CreatePolygonRgn(poly(1), 3, ALTERNATE)
            hFFBrush = CreateSolidBrush(bColor) 'create a yellow brush
            SelectObject Me.hDC, hFFBrush 'Select this brush
            FrameRgn Me.hDC, hRgnTmp&, hFFBrush, Me.ScaleWidth, Me.ScaleHeight 'Draw a frame
            'FillRgn Me.hDC, hRgnTmp&, hFFBrush
        
    End Select
    
    Set MyPic = Nothing
    n_SetPointer = hTRgn
    
    DeleteObject hRgnTmp
    DeleteObject hrgnTmp1
    DeleteObject hFFBrush&

End Function

Private Sub iCmd_Click(Index As Integer)

    Select Case Index

        Case 1 'OK
            nAction = 0

            If ComOptions.ListCount > 0 Then
            
            Else
                If iTxt.Text = "" Then Exit Sub
            End If
        Case 0 'cancel
            nAction = -1
    End Select

    Me.Hide

End Sub

Private Sub iMsg_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)

    'allow the form mouse move event to decide if draging is enabled
    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub lblmsg_MouseMove(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    'allow the form mouse move event to decide if draging is enabled
    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub Timer1_Timer()

    'turn off the timer
    Timer1.Enabled = False
    Unload msgFrm
    
End Sub

