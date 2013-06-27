VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.Form frmStart 
   Caption         =   "OASIS Print testing"
   ClientHeight    =   6705
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9195
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TatukGIS_XDK10.XGIS_ViewerWnd GIS 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5535
      BigExtentMargin =   -10
      RestrictedDrag  =   -1  'True
      CachedPaint     =   -1  'True
      IncrementalPaint=   -1  'True
      FullPaint       =   -1  'True
      CodePage        =   1250
      OutCodePage     =   1250
      CharSet         =   238
      UseRTree        =   0   'False
      PrinterTileSize =   512
      PrintTitle      =   ""
      PrintSubtitle   =   ""
      PrintFooter     =   ""
      BeginProperty PrintTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PrintSubtitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PrintFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PrintTitleFontColor=   -16777208
      PrintSubtitleFontColor=   -16777208
      PrintFooterFontColor=   -16777208
      SelectionColor  =   16777215
      SelectionPattern=   "frmStart.frx":0000
      SelectionTransparency=   100
      SelectionWidth  =   100
      SelectionOutlineOnly=   0   'False
      OldCachedPaint  =   0   'False
      PrinterModeDraft=   0   'False
      PrinterModeForceBitmap=   0   'False
      GDIType         =   0
      ScaleAsFloat    =   1
      Mode            =   0
      BorderStyle     =   1
      CursorForDrag   =   0
      CursorForSelect =   0
      CursorForZoom   =   0
      CursorForEdit   =   0
      MinZoomSize     =   -5
      ScrollBars      =   0
      AutoCenter      =   0   'False
      Align           =   5
      Ctl3D           =   0   'False
      Object.Visible         =   -1  'True
      Cursor          =   16
      DoubleBuffered  =   0   'False
      ModeMouseButton =   0
      CursorForUserDefined=   0
      View3D          =   0   'False
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   3165
      Left            =   495
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   258
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   630
      Width           =   3930
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   6450
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fullextent"
            Object.ToolTipText     =   "Full Extent"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "zoomin"
            Object.ToolTipText     =   "Zoom in"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "zoomout"
            Object.ToolTipText     =   "Zoom Out"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin VB.CheckBox CheckDrag 
         Appearance      =   0  'Flat
         Caption         =   "Dragging"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   35
         Width           =   2415
      End
      Begin VB.CommandButton cmdDoit 
         Caption         =   "Print"
         Height          =   360
         Left            =   3960
         TabIndex        =   4
         Top             =   0
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6360
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":0082
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":03D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStart.frx":0728
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim GisUtils As New XGIS_Utils

'This example needs a Picture box (Picture1)
'with an picture loaded in it
Private Const IMAGE_BITMAP = 0
Private Const LR_COPYRETURNORG = &H4
Private Const CF_BITMAP = 2
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long



Private Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Private Sub GetPictureBytes()
  Dim PicBits() As Byte, PicInfo As BITMAP

  GetObject Picture1.Picture, Len(PicInfo), PicInfo

  ReDim PicBits((PicInfo.bmWidth * PicInfo.bmHeight * 3) - 1) As Byte

  GetBitmapBits Picture1.Picture, UBound(PicBits), PicBits(0)
End Sub

Sub saveImage(pbImage As PictureBox, sFile As String)
    pbImage.Picture = pbImage.Image
    SavePicture pbImage.Picture, sFile
End Sub

Private Sub Doit5()

    Me.AutoRedraw = True

        PrintWindow GIS.Viewer.hWnd, Picture1.hdc, 0
 
End Sub

Private Sub DoIt4()

    Dim hNew As Long
    Dim oRect As New XRect
    
    oRect.Prepare 0, 0, 100, 100
    
    GIS.PrintDC Picture1.hdc, 300, oRect, GIS.VisibleExtent, GIS.ScaleAsFloat
    Picture1.Refresh
    Exit Sub
    
    'create an exact copy of the picture
    'hNew = CopyImage(GIS.Viewer.PrintDC(, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG)
    'open the clipboard
    OpenClipboard Me.hWnd
    'clear the clipboard
    EmptyClipboard
    'put the picture on the clipboard
    SetClipboardData CF_BITMAP, hNew
    'close the clipboard
    CloseClipboard
    'note that we don't have to call DeleteObject(hNew)
    'from now on, the clipboard takes care of the bitmap
End Sub

Private Sub ButtonFullExtent_Click()
   GIS.FullExtent
End Sub

Private Sub ButtonZoomIn_Click()
  GIS.Zoom = GIS.Zoom * 2
End Sub

Private Sub ButtonZoomOut_Click()
  GIS.Zoom = GIS.Zoom / 2
End Sub

Private Sub CheckDrag_Click()
  If CheckDrag.value Then
     GIS.Mode = XgisDrag
     GIS.ScrollBars = XssNone
  Else
     GIS.Mode = XgisSelect
     GIS.ScrollBars = XssBoth
  End If

End Sub

Private Sub cmdDoit_Click()
Dim okl As XGIS_ControlLegend

    Dim sPath As String
    Dim xmin, xmax, ymin, ymax As Double
    On Error Resume Next
            
    'sPath = App.Path & "\" & "Now" & ".ttkgp"
    'Kill sPath
    'GIS.SaveProjectAs sPath, False
    'GIS.SaveAll
    
    sPath = "C:\Users\OASIS\Documents\iMMAP - OASIS\OASIS Client\data\user\Maps\DefaultMap.TTKGP"
    
    'frmMainPrint.Show
    
    With GIS.VisibleExtent
        Debug.Print .xmin
        Debug.Print .xmax
        Debug.Print .ymin
        Debug.Print .ymax
        frmMainPrint.InitPrint sPath, .xmin, .xmax, .ymin, .ymax
    End With
    
    frmMainPrint.Show
    Exit Sub
    
    Dim oRect As New XRect
    
    With Picture1
        .Move .left, .top, GIS.Width, GIS.Height
        oRect.Prepare 0, 0, ScaleX(.Width, vbTwips, vbPixels), ScaleY(.Height, vbTwips, vbPixels)
        .ZOrder 0
        .Cls
    End With
    
    GIS.PrintDC Picture1.hdc, 5, oRect, GIS.Viewer.VisibleExtent, 0

    GIS.Draw
  
    Clipboard.Clear
    Clipboard.SetData Picture1.Image, vbCFDIB
  
End Sub

Private Sub Form_Load()
    Dim ll As XGIS_LayerSHP
  
    If 1 = 2 Then
        Set ll = New XGIS_LayerSHP
  
        ll.Path = GisUtils.GisSamplesDataDir + "\World\Countries\Poland\DCW\country.shp"
        ll.Name = "states"
  
        GIS.Add ll

        Set ll = New XGIS_LayerSHP
        ll.Path = GisUtils.GisSamplesDataDir + "\World\Countries\Poland\DCW\lwaters.shp"
        ll.Name = "rivers"
        ll.UseConfig = False
        ll.Params.Line.OutlineWidth = 0
        ll.Params.Line.Width = 3
        ll.Params.Line.color = RGB(0, 0, 255)
        GIS.Add ll
        GIS.FullExtent
    End If

    GIS.Open App.Path & "\" & "Now" & ".ttkgp"
    
    Doit5
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
   Case Is = "fullextent"
     GIS.FullExtent
   Case Is = "zoomin"
     GIS.Zoom = GIS.Zoom * 2
   Case Is = "zoomout"
     GIS.Zoom = GIS.Zoom / 2
   End Select
End Sub

Private Sub Form_Resize()
  If ScaleWidth = 0 Then Exit Sub
  GIS.Move 0, Toolbar.Height, ScaleWidth, ScaleHeight - Toolbar.Height - StatusBar.Height

End Sub


