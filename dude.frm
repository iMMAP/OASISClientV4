VERSION 5.00
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin TatukGIS_XDK10.XGIS_ViewerWnd GIS 
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   180
      Width           =   5055
      BigExtentMargin =   -10
      RestrictedDrag  =   -1  'True
      CachedPaint     =   -1  'True
      IncrementalPaint=   -1  'True
      FullPaint       =   -1  'True
      CodePage        =   0
      OutCodePage     =   0
      CharSet         =   0
      UseRTree        =   0   'False
      PrinterTileSize =   2700
      PrintTitle      =   ""
      PrintSubtitle   =   ""
      PrintFooter     =   ""
      BeginProperty PrintTitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PrintSubtitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty PrintFooterFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PrintTitleFontColor=   -16777208
      PrintSubtitleFontColor=   -16777208
      PrintFooterFontColor=   -16777208
      SelectionColor  =   16777215
      SelectionPattern=   "dude.frx":0000
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
      Align           =   0
      Ctl3D           =   -1  'True
      Object.Visible         =   -1  'True
      Cursor          =   16
      DoubleBuffered  =   0   'False
      ModeMouseButton =   0
      CursorForUserDefined=   0
      View3D          =   0   'False
   End
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3780
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCommand1_Click()
    AddLayer
End Sub

Public Sub AddLayer()
        '<EhHeader>
        On Error GoTo AddLayer_Err
        '</EhHeader>
        Dim oAbsLayer As TatukGIS_XDK10.XGIS_LayerAbstract
        Dim sName As String
        Dim oLayer As Object
            
        Set oLayer = New TatukGIS_XDK10.XGIS_LayerMrSID
                        
180     oLayer.Path = "C:\Users\OASIS\Desktop\1m_kabul_geo.sid"
991       oLayer.Open
                
184     GIS10.Add oLayer
            
196     GIS10.Update
        
        '<EhFooter>
        Exit Sub

AddLayer_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmMain.AddLayer " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

