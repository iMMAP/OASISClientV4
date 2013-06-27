VERSION 5.00
Begin VB.Form frmOASISCharts 
   Caption         =   "OASIS Dynamic Charting"
   ClientHeight    =   6105
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu Clip 
      Caption         =   "Copy to Clipboard"
   End
   Begin VB.Menu mnuCreateReport 
      Caption         =   "Create Report"
      Visible         =   0   'False
   End
End
Attribute VB_Name = "frmOASISCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Const SRCCOPY = &HCC0020 ' (DWORD) destination = source

Private Type DEVMODE
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

'API
Private Declare Function ReleaseDC _
                Lib "user32.dll" (ByVal Hwnd As Long, _
                                  ByVal hdc As Long) As Long
Private Declare Function OpenClipboard _
                Lib "user32.dll" (ByVal Hwnd As Long) As Long
Private Declare Function EmptyClipboard _
                Lib "user32.dll" () As Long
Private Declare Function SetClipboardData _
                Lib "user32.dll" (ByVal wFormat As Long, _
                                  ByVal hMem As Long) As Long
'if you have problems with this function add the Alias "SetClipboardDataA"
Private Declare Function CloseClipboard _
                Lib "user32.dll" () As Long
Private Declare Function SelectObject _
                Lib "gdi32.dll" (ByVal hdc As Long, _
                                 ByVal hObject As Long) As Long
Private Declare Function DeleteDC _
                Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function BitBlt _
                Lib "gdi32.dll" (ByVal hDestDC As Long, _
                                 ByVal x As Long, _
                                 ByVal y As Long, _
                                 ByVal nWidth As Long, _
                                 ByVal nHeight As Long, _
                                 ByVal hSrcDC As Long, _
                                 ByVal xSrc As Long, _
                                 ByVal ySrc As Long, _
                                 ByVal dwRop As Long) As Long
Private Declare Function CreateDC _
                Lib "gdi32.dll" _
                Alias "CreateDCA" (ByVal lpDriverName As String, _
                                   ByVal lpDeviceName As String, _
                                   ByVal lpOutput As String, _
                                   lpInitData As DEVMODE) As Long
Private Declare Function CreateCompatibleBitmap _
                Lib "gdi32.dll" (ByVal hdc As Long, _
                                 ByVal nWidth As Long, _
                                 ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC _
                Lib "gdi32.dll" (ByVal hdc As Long) As Long

Private Declare Function SetParent _
                Lib "user32" (ByVal hWndChild As Long, _
                              ByVal hWndNewParent As Long) As Long
                              
Private Declare Function SetWindowPos _
                Lib "user32" (ByVal Hwnd As Long, _
                              ByVal hWndInsertAfter As Long, _
                              ByVal x As Long, _
                              ByVal y As Long, _
                              ByVal cx As Long, _
                              ByVal cy As Long, _
                              ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_BOTTOM = 1
Private Const HWND_TOP = 0
Private ChartHandle As Long
Private m_RSChart As ADODB.Recordset
Dim bToolbar1 As Boolean
Dim bToolbar2 As Boolean
Public Event ExportDone()

Public Sub CaptureScreen(Left As Long, _
                         Top As Long, _
                         Width As Long, _
                         Height As Long)
        '<EhHeader>
        On Error GoTo CaptureScreen_Err
        '</EhHeader>
        Dim srcDC As Long
        Dim trgDC As Long
        Dim BMPHandle As Long
        Dim dm As DEVMODE

100     srcDC = CreateDC("DISPLAY", "", "", dm)
102     trgDC = CreateCompatibleDC(srcDC)
104     BMPHandle = CreateCompatibleBitmap(srcDC, Width, Height)
106     SelectObject trgDC, BMPHandle
108     BitBlt trgDC, 0, 0, Width, Height, srcDC, Left, Top, SRCCOPY
110     OpenClipboard Screen.ActiveForm.Hwnd
112     EmptyClipboard
114     SetClipboardData 2, BMPHandle
116     CloseClipboard
118     DeleteDC trgDC
120     ReleaseDC BMPHandle, srcDC
        '<EhFooter>
        Exit Sub

CaptureScreen_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISCharts.CaptureScreen " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetChart(objChart As OASISChartObj)
        '<EhHeader>
        On Error GoTo SetChart_Err
        '</EhHeader>
    
        Dim ochart As Object
100     Set ochart = CreateObject("OASISCharting.ChartProvider")
        
102     ochart.AllowAutoResize = True
        'ochart.AllowAutoCloseDBLClick = True
104     ChartHandle = ochart.InitChart(objChart, False, , Me.Hwnd)
106     SetWindowPos ChartHandle, HWND_TOP, 0, 0, ScaleX(Me.Width - 235, vbTwips, vbPixels), ScaleY(Me.Height - 545, vbTwips, vbPixels), 0
 
        '<EhFooter>
        Exit Sub

SetChart_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISCharts.SetChart " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetChartWithRS(objChart As OASISChartObj, _
                          RSChart As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo SetChartWithRS_Err
        '</EhHeader>
    
        Dim ochart As Object
       
100     If Not RSChart Is Nothing Then
102         Set m_RSChart = RSChart
104         mnuCreateReport.Visible = True
        End If
       
106     Set ochart = CreateObject("OASISCharting.ChartProvider")
108     ochart.AllowAutoResize = True
110     ochart.UpdateDataSet RSChart
112     bToolbar1 = objChart.bAnnoTBR
114     bToolbar2 = objChart.bChartTBR
116     ChartHandle = ochart.InitChart(objChart, False, , Me.Hwnd)
118     SetWindowPos ChartHandle, HWND_TOP, 0, 0, ScaleX(Me.Width - 235, vbTwips, vbPixels), ScaleY(Me.Height - 545, vbTwips, vbPixels), 0
 
        '<EhFooter>
        Exit Sub

SetChartWithRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISCharts.SetChartWithRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function OpenChartTemplate(sPath As String) As OASISChartObj
        '<EhHeader>
        On Error GoTo OpenChartTemplate_Err
        '</EhHeader>
    
        Dim ochart As Object 'New OASISCharting.ChartProvider
        Dim oOASISChartObj As OASISChartObj
   
        Set ochart = CreateObject("OASISCharting.ChartProvider")
100     ochart.AllowAutoResize = True
102     oOASISChartObj = ochart.GetOASISObjSettings(sPath)
104     ChartHandle = oOASISChartObj.udtGeneric.lParentHwnd
106     SetParent ChartHandle, Me.Hwnd
108     SetWindowPos ChartHandle, HWND_TOP, 0, 0, ScaleX(Me.Width - 235, vbTwips, vbPixels), ScaleY(Me.Height - 545 - 300, vbTwips, vbPixels), 0
110     OpenChartTemplate = oOASISChartObj

        Set ochart = Nothing

        '<EhFooter>
        Exit Function

OpenChartTemplate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISCharts.OpenChartTemplate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub Clip_Click()
        '<EhHeader>
        On Error GoTo Clip_Click_Err
        '</EhHeader>
        
        Dim iToolbarOffset As Integer
        
100     iToolbarOffset = 0 + IIf(bToolbar1, 400, 0)
102     iToolbarOffset = iToolbarOffset + IIf(bToolbar2, 400, 0)
104     CaptureScreen (Me.Left + 150) \ Screen.TwipsPerPixelX, (Me.Top + 800 + iToolbarOffset) \ Screen.TwipsPerPixelY, (Me.Width - 250) \ Screen.TwipsPerPixelX, (Me.Height - 900 - iToolbarOffset) \ Screen.TwipsPerPixelY
        'Picture1 = Clipboard.GetData()
106     MsgBox "Copied to clipboard"
        '<EhFooter>
        Exit Sub

Clip_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISCharts.Clip_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        
100     mnuCreateReport.Visible = False

102     If Not g_sLanguage = "" Then
104         If Not m_Cnn.State = adStateClosed Then
106             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISCharts.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    If Not ChartHandle = 0 Then
        SetWindowPos ChartHandle, HWND_TOP, 0, 0, ScaleX(Me.Width - 235, vbTwips, vbPixels), ScaleY(Me.Height - 545 - 300, vbTwips, vbPixels), 0
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_RSChart = Nothing
End Sub

Private Sub mnuCreateReport_Click()
        '<EhHeader>
        On Error GoTo mnuCreateReport_Click_Err
        '</EhHeader>

        Dim iToolbarOffset As Integer
        
100     If Not m_RSChart Is Nothing Then
102         iToolbarOffset = 0 + IIf(bToolbar1, 400, 0)
104         iToolbarOffset = iToolbarOffset + IIf(bToolbar2, 400, 0)
106         CaptureScreen (Me.Left + 150) \ Screen.TwipsPerPixelX, (Me.Top + 800 + iToolbarOffset) \ Screen.TwipsPerPixelY, (Me.Width - 250) \ Screen.TwipsPerPixelX, (Me.Height - 900 - iToolbarOffset) \ Screen.TwipsPerPixelY
108         frmReportsFromRS.SetReportRS "OASIS Charts Reports", m_RSChart, "", Clipboard.GetData(vbCFBitmap), "OASIS Chart", ""
110         frmReportsFromRS.ShowReport
112         frmReportsFromRS.Show vbModal, Me
        End If

        '<EhFooter>
        Exit Sub

mnuCreateReport_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmOASISCharts.mnuCreateReport_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
