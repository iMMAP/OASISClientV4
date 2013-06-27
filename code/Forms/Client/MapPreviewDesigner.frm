VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{AA70C0C0-2D16-11D4-99DD-AAD8C1F45126}#1.0#0"; "vp3270.ocx"
Begin VB.Form frmMapPreviewDesigner 
   Caption         =   "HexaTech Report Designer"
   ClientHeight    =   12765
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   18120
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   851
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1208
   WindowState     =   2  'Maximized
   Begin vbalIml6.vbalImageList vbalImageList1 
      Left            =   1320
      Top             =   6600
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.PictureBox PictureArrow 
      Height          =   2175
      Left            =   15480
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox PictureScale 
      Height          =   615
      Left            =   9840
      ScaleHeight     =   555
      ScaleWidth      =   7515
      TabIndex        =   10
      Top             =   11640
      Visible         =   0   'False
      Width           =   7575
   End
   Begin VB.PictureBox PictureLegend 
      Height          =   10935
      Left            =   600
      ScaleHeight     =   10875
      ScaleWidth      =   3915
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox PictureMap 
      Height          =   12495
      Left            =   120
      Picture         =   "MapPreviewDesigner.frx":0000
      ScaleHeight     =   12435
      ScaleWidth      =   17595
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   17655
   End
   Begin VB.ComboBox cboZoom 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   660
      Width           =   1635
   End
   Begin VB.ComboBox cboPage 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   150
      Width           =   1605
   End
   Begin VB.ListBox lstObject 
      Height          =   4350
      Left            =   90
      TabIndex        =   0
      Top             =   1710
      Width           =   2115
   End
   Begin VIEWPR32Lib.ViewPro ViewPro1 
      Height          =   7125
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   8265
      _Version        =   65536
      _ExtentX        =   14579
      _ExtentY        =   12568
      _StockProps     =   57
   End
   Begin VB.Label lblEdit 
      Caption         =   "Double click object on preview page to edit."
      Height          =   765
      Left            =   75
      TabIndex        =   7
      Top             =   6360
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Page"
      Height          =   195
      Left            =   30
      TabIndex        =   6
      Top             =   210
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "Drag/Drop Object:"
      Height          =   405
      Left            =   90
      TabIndex        =   5
      Top             =   1260
      Width           =   2085
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Zoom"
      Height          =   195
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileSaveAsObjectDocument 
         Caption         =   "Save As VPA &Object Document"
      End
      Begin VB.Menu mnuFileSaveAsObjectScript 
         Caption         =   "Save As VPA &Object Script"
      End
      Begin VB.Menu mnuFileLine12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveAsScript 
         Caption         =   "Save As &ViewPro Script"
      End
      Begin VB.Menu mnuFileSaveAsMetafile 
         Caption         =   "Save Current Page As &Metafile"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditLine22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAddPage 
         Caption         =   "&Add Page"
      End
      Begin VB.Menu mnuEditDeletePage 
         Caption         =   "De&lete Page"
      End
      Begin VB.Menu mnuEditLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFront 
         Caption         =   "&Bring to Front"
         Shortcut        =   ^J
      End
      Begin VB.Menu mnuEditBack 
         Caption         =   "&Send to Back"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsSetup 
         Caption         =   "&Page Settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About HexaTech Report Designer"
      End
   End
End
Attribute VB_Name = "frmMapPreviewDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const OBJECT_TEXT = 1
Const OBJECT_PARAGRAPH = 2
Const OBJECT_TEXTRTF0 = 3
Const OBJECT_TEXTRTF = 4
Const OBJECT_TABLERTF0 = 5
Const OBJECT_TABLERTF = 6
Const OBJECT_PICTURE = 7
Const OBJECT_LINE = 8
Const OBJECT_RECTANGLE = 9
Const OBJECT_CIRCLE = 10
Const OBJECT_ELLIPSE = 11
Const OBJECT_CUSTOM = 12
Const OBJECT_ERASER = 15
Const OBJECT_POLYOBJECT = 16
Const OBJECT_EDITBOX = 17
Const OBJECT_FILE = 18
Const OBJECT_META = 19
Const OBJECT_SCRIPT = 20
Const OBJECT_MAP = 21
Const OBJECT_LEGEND = 22
Const OBJECT_ARROW = 23
Const OBJECT_SCALE = 24

Const OBJECT_SELECT_NONE = 0
Const OBJECT_SELECT_SINGLE = 1
Const OBJECT_SELECT_MULTIPLE = 2
Const OBJECT_SELECT_CUSTOM = 3

Const EDITBOX_OBJECT_NAME = "editboxObject"
Const EXT_OBJECT_DOC = "vpd"
Const EXT_SCRIPT = "txt"
Const program_name = "HexaTech Report Designer"

Dim INCH As Long
Dim sFilter As String
Dim bInitialResize As Boolean
Dim ReportFile As String
Dim NormalZoom As Integer
Public ShowGrid As Integer

Public sLegendPath As String

Public mPicScale As StdPicture
Public mPicLegend As StdPicture
Public mPicMap As StdPicture
Public mPicArrow As StdPicture



Dim DefaultFontName As Integer
Dim DefaultFontSize As Integer
Dim DefaultFontBold As Integer


Sub DoResize()
    
    On Error Resume Next

    'Resize the control
    ViewPro1.Width = Me.ScaleWidth - lstObject.Width - 30
    ViewPro1.Height = Me.ScaleHeight - 15
    
    'First time resize
    If bInitialResize = False Then
        bInitialResize = True
        ViewPro1.CenterPage
        NormalZoom = ViewPro1.zoom
    'Not first time resize
    Else
        cboZoom_Click
    End If

End Sub

Function GetSamplePolyData(vp As Control, Filename As String, subAttrib As String)
    On Error Resume Next
    
    Dim sub_y As String, sub_g As String, sub_b As String, sub_r As String
    Dim s As String, a As String

    vp.ObjectPath = Filename & ";" & "object_frog_yellow"
    sub_y = ViewPro1.GetObject
    vp.ObjectPath = Filename & ";" & "object_frog_green"
    sub_g = ViewPro1.GetObject
    vp.ObjectPath = Filename & ";" & "object_frog_black"
    sub_b = ViewPro1.GetObject
    vp.ObjectPath = Filename & ";" & "object_frog_red"
    sub_r = ViewPro1.GetObject
               
    'data
    s = ""
    s = s & sub_y & "|"
    s = s & sub_g & "|"
    s = s & sub_b & "|"
    s = s & sub_r
    GetSamplePolyData = s

    'attribute
    a = ""
    a = a & "SubType=0;FillStyle=1;FillColor = RGB(255, 255, 0)" & "|"
    a = a & "SubType=0;FillStyle=1;FillColor = RGB(0, 255, 0)" & "|"
    a = a & "SubType=0;FillStyle=1;FillColor = RGB(0, 0, 0)" & "|"
    a = a & "SubType=0;FillStyle=1;FillColor = RGB(255, 0, 0)"
    subAttrib = a
    

End Function

Function GetUniqueObjectName(vp As Control, nType As Integer) As String
    On Error Resume Next
    
    Dim s As String, n As Integer, Prefix As String
            
    Select Case nType
        Case OBJECT_TEXT:       Prefix = "txt"
                
        Case OBJECT_PARAGRAPH:  Prefix = "par"
        
        Case OBJECT_TEXTRTF0:   Prefix = "pa0"
        
        Case OBJECT_TEXTRTF:    Prefix = "rtf"
                
        Case OBJECT_TABLERTF0:  Prefix = "tbl"
        
        Case OBJECT_TABLERTF:   Prefix = "tbr"

        Case OBJECT_PICTURE:    Prefix = "pic"
        
        Case OBJECT_LINE:       Prefix = "lin"
        
        Case OBJECT_RECTANGLE:  Prefix = "rct"
        
        Case OBJECT_CIRCLE:     Prefix = "cir"
        
        Case OBJECT_ELLIPSE:    Prefix = "elp"
        
        Case OBJECT_CUSTOM:     Prefix = "cus"

        Case OBJECT_POLYOBJECT:     Prefix = "ply"

        Case OBJECT_ERASER:     Prefix = "era"
        
         Case OBJECT_MAP:      Prefix = "map"

        Case OBJECT_LEGEND:     Prefix = "leg"
        
        Case OBJECT_ARROW:     Prefix = "arr"
Case OBJECT_SCALE:     Prefix = "sca"

    End Select
    
    n = vp.GetObjectCount + 1
    s = Prefix & "Object" & n
    While vp.GetObjectIndex(s) <> 0
        n = n + 1
        s = Prefix & "Object" & n
    Wend
        
    GetUniqueObjectName = s

End Function
Sub Initialize()
    On Error Resume Next

    Me.caption = program_name & " - [Untitled]"
    INCH = 1440

    '-----Set control size (fire Form_Resize event)
    bInitialResize = False
    Me.ScaleMode = 3
    Me.WindowState = 2
    
    '-----Must set ObjectEdit to True for drag and drop
    ViewPro1.ObjectEdit = True
        
    'Custom edit box settings
    Dim s As String
    s = "" & 32 Or 128
    ViewPro1.SetEditBoxInfo s & ",0,0," & 0.5 * INCH & "," & 0.5 * INCH
        
    '-----Shif distance for selected objects by arrow keys
    'ViewPro1.SetShiftDistance INCH / 32, INCH / 32
    
    'Shif distance for selected objects by custom keys (in KeyEvent routine)
    ViewPro1.SetShiftDistance 0, 0
                
    '-----Do not automatically advance page
    ViewPro1.PageAdvance = 0

    '-----Mouse Edit settings
    ViewPro1.AutoSize = True
    ViewPro1.KeepAspect = True
    ViewPro1.MouseZoom = False
    ViewPro1.MouseScroll = False
    ViewPro1.EditBoxColor = RGB(128, 128, 128)
    
    '-----Attributes
    ViewPro1.ObjectBorderStyle = 0
    ViewPro1.ObjectBorderColor = RGB(0, 0, 0)
    ViewPro1.ObjectBKColor = RGB(255, 255, 255)
    ViewPro1.BackgroundMode = 1
    ViewPro1.PenStyle = 0  'solid
    ViewPro1.FillStyle = 0
    ViewPro1.FilePath = App.Path
        
    '-----Grid
    ViewPro1.GridHorz = 15
    ViewPro1.GridVert = 20
    ShowGrid = 1
    
    '-----Default font parameters
    DefaultFontName = 1
    DefaultFontSize = 10
    DefaultFontBold = 0
    ViewPro1.FontName = DefaultFontName
    ViewPro1.FontSize = DefaultFontSize
    ViewPro1.FontBold = DefaultFontBold
    
        
    '-----Picture file filter
    sFilter = ""
    sFilter = sFilter & "All Picture Files (*.bmp;*.jpg;*.gif;*.wmf;*.emf) | *.bmp;*.jpg;*.gif;*.wmf;*.emf;"
    sFilter = sFilter & "|BMP Files (*.bmp) | *.bmp;"
    sFilter = sFilter & "|JPG Files (*.jpg) | *.jpg;"
    sFilter = sFilter & "|GIF Files (*.gif) | *.gif;"
    sFilter = sFilter & "|WMF Files (*.wmf) | *.wmf;"
    sFilter = sFilter & "|EMF Files (*.emf) | *.emf;"
    'sFilter = sFilter & "|All Picture Files (*.bmp;*.wmf;*.emf) | *.bmp;*.wmf;*.emf;"
    sFilter = sFilter & "||"
    
    ReportFile = ""
        
    '-----Page combo box
    cboPage.Clear
    cboPage.AddItem "1"
    cboPage.ListIndex = 0
        
    '-----Update doc
    ViewPro1.UpdateDoc
        
    '-----Zoom combo box
    'ViewPro1.CenterPage
    'NormalZoom = ViewPro1.Zoom
    cboZoom.Clear
    cboZoom.AddItem "60"
    cboZoom.AddItem "80"
    cboZoom.AddItem "100"
    cboZoom.AddItem "150"
    cboZoom.AddItem "200"
    cboZoom.AddItem "250"
    cboZoom.ListIndex = 2
    
End Sub



Private Sub cboPage_Click()
    On Error Resume Next

    Screen.MousePointer = 11
    
    ViewPro1.StoreCurrentPageObjects
    ViewPro1.ObjectPageIndex = Val(cboPage.List(cboPage.ListIndex))
    ViewPro1.UpdateDoc
                
    Screen.MousePointer = 0
    
End Sub


Private Sub cboZoom_Click()
    On Error Resume Next

    ViewPro1.zoom = NormalZoom * cboZoom.List(cboZoom.ListIndex) / 100
    ViewPro1.UpdateDoc
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    
 '   Dim s As String
 '   s = "This Report Designer is developed with ViewPro Advanced (VPA) Edition." & vbCrLf & vbCrLf
 '   s = s & "Its source code is in the directory: " & App.Path & "."
 '   MsgBox s


    
    lstObject.AddItem "Text"
    lstObject.AddItem "Paragraph"
    lstObject.AddItem "TextRTF"
    lstObject.AddItem "Table"
    lstObject.AddItem "TableRTF"
    lstObject.AddItem "Picture"
    lstObject.AddItem "Horz Line"
    lstObject.AddItem "Vert Line"
    lstObject.AddItem "Rectangle"
    lstObject.AddItem "Circle"
    lstObject.AddItem "Ellipse"
    lstObject.AddItem "PolyObject"
    lstObject.AddItem "Custom"
    lstObject.AddItem "Eraser"
    
    lstObject.AddItem "Map"
    lstObject.AddItem "Legend"
    lstObject.AddItem "Arrow"
    lstObject.AddItem "Scale"
    
    Initialize
    
End Sub


Private Sub Form_Resize()
    On Error Resume Next

    DoResize
End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    mnuFileExit_Click
End Sub


Private Sub lstObject_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    
    ViewPro1.ObjectEdit = True
    
    ViewPro1.StartDragAndDrop 0, ""
    
    'ViewPro1.StartDragAndDrop App.hInstance, "HAND"
        
End Sub


Private Sub mnuEditAddPage_Click()
    On Error Resume Next

    Screen.MousePointer = 11
    
    ViewPro1.StoreCurrentPageObjects
    
    '<hexa_vp70>
    'ViewPro1.ObjectPageIndex = ViewPro1.ObjectPageIndex + 1
    ViewPro1.ObjectPageIndex = ViewPro1.GetObjectPageCount + 1
    ViewPro1.UpdateDoc
    
    
    cboPage.AddItem "" & ViewPro1.ObjectPageIndex
    cboPage.ListIndex = cboPage.ListCount - 1
    
    
    Screen.MousePointer = 0

End Sub

Private Sub mnuEditBack_Click()
    On Error Resume Next
    
    'No more than one object selected
    If ViewPro1.GetSelectedObjectCount() > 1 Then
        MsgBox "Please select one object at a time for this operation."
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    ViewPro1.SetObjectOrder ViewPro1.GetSelectedObject, 1
    Screen.MousePointer = 0
    
End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next

    Dim i As Integer, Count As Integer

    'Store selected objects to memory arrays
    ViewPro1.RemoveAllMemObjects
    Count = ViewPro1.GetSelectedObjectCount
    For i = 1 To Count
        ViewPro1.StoreMemObject ViewPro1.GetArrayItem(i), "storedObj_" & i
    Next

    
End Sub

Private Sub mnuEditDelete_Click()
    On Error Resume Next

    If ViewPro1.GetSelectedObject <> "" Then
        'ViewPro1.RemoveObject ViewPro1.GetSelectedObject
        ViewPro1.RemoveObject ""
        ViewPro1.UpdateDoc
    End If

End Sub

Private Sub mnuEditDeletePage_Click()
    On Error Resume Next

Dim Index As Integer, i As Integer
Dim nResponse As Integer

    nResponse = MsgBox("Do you want to delete the page?", vbYesNo + vbCritical + vbDefaultButton2)
    If nResponse = vbNo Then Exit Sub
    
    
    Index = cboPage.ListIndex

    ViewPro1.RemoveCurrentPageObjects
    ViewPro1.RemoveObjectPage Index + 1
    cboPage.Clear
    
    cboPage.AddItem "" & 1
    For i = 2 To ViewPro1.GetObjectPageCount
        cboPage.AddItem "" & i
    Next
    
    If Index + 1 <= ViewPro1.GetObjectPageCount Then ViewPro1.ObjectPageIndex = Index + 1
    
    cboPage.ListIndex = 0
    

End Sub

Private Sub mnuEditFront_Click()
    On Error Resume Next


    'No more than one object selected
    If ViewPro1.GetSelectedObjectCount() > 1 Then
        MsgBox "Please select one object at a time for this operation."
        Exit Sub
    End If


    Screen.MousePointer = 11
    ViewPro1.SetObjectOrder ViewPro1.GetSelectedObject, ViewPro1.GetObjectCount
    Screen.MousePointer = 0

End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next

    Dim i As Integer, Count As Integer
    Dim sName As String, sName2 As String, nType As Integer, x As Long, y As Long


    'Load stored objects to current page
    ViewPro1.SelectObject ""
    Count = ViewPro1.GetMemObjectCount
    For i = 1 To Count
        sName = ViewPro1.GetMemObjectName(i)
        nType = ViewPro1.GetMemObjectInfo(sName)
        sName2 = GetUniqueObjectName(ViewPro1, nType)
        x = ViewPro1.x
        y = ViewPro1.y
        ViewPro1.LoadMemObject sName, sName2
        ViewPro1.SetObjectX sName2, x + 0.1 * INCH
        ViewPro1.SetObjectY sName2, y + 0.1 * INCH
        ViewPro1.SelectObject sName2
    Next
    ViewPro1.UpdateDoc


End Sub

Private Sub mnuEditSelectAll_Click()
    On Error Resume Next

    ViewPro1.SelectAll
    
   'ViewPro1.SetFocus
    
End Sub

Private Sub mnuEditUndo_Click()
    On Error Resume Next

    Screen.MousePointer = 11
    ViewPro1.Undo
    Screen.MousePointer = 0
End Sub

Private Sub mnuFileExit_Click()
    On Error Resume Next
    
    mnuFileNew_Click
'    End
End Sub

Private Sub mnuFileNew_Click()
    On Error Resume Next
    
    Dim nResponse As Integer

    If ViewPro1.GetObjectCount > 0 Or ViewPro1.GetObjectPageCount > 1 Then
        nResponse = MsgBox("Save Changes?", vbYesNo + vbCritical + vbDefaultButton2)
        If nResponse = vbYes Then
            mnuFileSave_Click
        End If
    End If

    ViewPro1.RemoveAllPageObjects
    cboPage.Clear
    cboPage.AddItem "1"
    cboPage.ListIndex = 0
    
    ReportFile = ""
    frmMapPreviewDesigner.caption = program_name & " - [Untitled]"
     
End Sub

Private Sub mnuFileOpen_Click()
    On Error Resume Next
    
    Dim nResponse As Integer
    Dim Filter As String, file_name As String, i As Integer

    'Save changes
    If (ViewPro1.GetObjectCount > 0 Or ViewPro1.GetObjectPageCount > 1) And ReportFile <> "" Then
        nResponse = MsgBox("Save Changes first?", vbYesNo + vbCritical + vbDefaultButton2)
        If nResponse = vbYes Then
            mnuFileSave_Click
        End If
    End If

    'Filter
    Filter = ""
    'filter = filter & "Report Layout Files (*.rlf) | *.rlf;"
    'filter = filter & "|VPA Object Script (*.txt;*.bas) | *.txt;*.bas"
    'filter = filter & "VPA Object Script (*.txt;*.bas) | *.txt;*.bas"
    'filter = filter & "|Report Layout Files (*.rlf) | *.rlf"
    Filter = Filter & "VPA Object Document (*.vpd) | *.vpd"
    Filter = Filter & "|VPA Object Script (*.txt;*.bas) | *.txt;*.bas"
    Filter = Filter & "||"
  
    'Show file dialog box
    If ViewPro1.ShowFileDialog(True, EXT_OBJECT_DOC, "", Filter) = 0 Then Exit Sub
    DoEvents
    ReportFile = ViewPro1.String
    frmMapPreviewDesigner.caption = program_name & " - " & ReportFile
    
    
    Screen.MousePointer = 11
    
    
    file_name = LCase(ReportFile)
    If InStr(file_name, ".vpd") > 0 Then
        ViewPro1.LoadObjectDocument ReportFile
    Else
        ViewPro1.RemoveAllPageObjects
        ViewPro1.LoadScript ReportFile
    End If
    
    
    
    ViewPro1.ObjectPageIndex = 1
    'ViewPro1.SetCurrentPageObjects 1
    ViewPro1.UpdateDoc
        
    'Fill page list box
    cboPage.Clear
    If ViewPro1.GetObjectPageCount = 0 Then
        cboPage.AddItem "1"
    Else
        For i = 1 To ViewPro1.GetObjectPageCount
            cboPage.AddItem "" & i
        Next
    End If
    cboPage.ListIndex = 0
    
        
    Screen.MousePointer = 0

End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next

    '--------------
    'Print current pages
    '
    'ViewPro1.PrintDialog = 1
    'ViewPro1.PrintPages 1, ViewPro1.TotalPages
    
    
'--------------
'Print all pages
Dim i As Integer, PageIndex As Integer

    PageIndex = ViewPro1.ObjectPageIndex
    ViewPro1.StoreCurrentPageObjects
    
    Screen.MousePointer = 11
    ViewPro1.StartDoc
    
        'Draw pages
        For i = 1 To ViewPro1.GetObjectPageCount
        
            If i <> 1 Then ViewPro1.NewPage
            ViewPro1.ObjectPageIndex = i
            ViewPro1.DrawCurrentPageObjects
            
            'Draw grid
            If ShowGrid Then
                ViewPro1.PenStyle = 2
                'ViewPro1.DrawBoundGrid 1
                ViewPro1.PenStyle = 0
            End If
        Next

        
    ViewPro1.EndDoc
    'ViewPro1.Preview
    Screen.MousePointer = 0
    
    
    'PrintSetup dialog
    ViewPro1.PrintDialog = 1
    ViewPro1.PrintPages 1, ViewPro1.TotalPages


    'Restore display
    Screen.MousePointer = 11
    ViewPro1.ObjectPageIndex = PageIndex
    ViewPro1.UpdateDoc
    Screen.MousePointer = 0


End Sub

Private Sub mnuFileSave_Click()
    On Error Resume Next
    
    Dim file_name As String
    
    If ReportFile <> "" Then
        Screen.MousePointer = 11
        ViewPro1.StoreCurrentPageObjects
                
        ViewPro1.SaveObjectDocument ReportFile
        
        Screen.MousePointer = 0
    Else
        mnuFileSaveAsObjectDocument_Click
    End If

End Sub



Private Sub mnuFileSaveAsMetafile_Click()
    On Error Resume Next

    Dim Filter As String, file_name As String

    'File filter
    Filter = ""
    Filter = Filter & "Windows Metafiles (*.wmf;*.emf) | *.wmf;*.emf;"
    Filter = Filter & "||"
    
    'Initial file name
    file_name = "page" & ViewPro1.ObjectPageIndex & ".emf"
        
    'Show file dialog box
    If ViewPro1.ShowFileDialog(False, "emf", file_name, Filter) = 0 Then Exit Sub
    DoEvents

    'Get file name
    file_name = ViewPro1.String
    file_name = LCase(file_name)
    If Not (InStr(file_name, ".emf") > 0 Or InStr(file_name, ".wmf") > 0) Then
        file_name = file_name & ".emf"
    End If
    
    'Save current page objects
    ViewPro1.StoreCurrentPageObjects
        
        
    Screen.MousePointer = 11
        
    'Save metafile without showing grid
    If ShowGrid = 0 Then
        ViewPro1.ExportPage file_name, 1
    Else
        'Remove grid
        ViewPro1.StartDoc
        ViewPro1.DrawCurrentPageObjects
        ViewPro1.EndDoc
        ViewPro1.ExportPage file_name, 1
    
        'Restore display
        ViewPro1.UpdateDoc
    End If
    
    Screen.MousePointer = 0
    
End Sub



Private Sub mnuFileSaveAsObjectDocument_Click()
    On Error Resume Next

    Dim Filter As String, file_name As String

    Filter = ""
    Filter = Filter & "VPA Object Document Files (*.vpd) | *.vpd;"
    Filter = Filter & "|All Files (*.*)|*.*;"
    Filter = Filter & "||"
  
    'Show file dialog box
    If ViewPro1.ShowFileDialog(False, EXT_OBJECT_DOC, "", Filter) = 0 Then
        Exit Sub
    End If
    DoEvents

    file_name = ViewPro1.String
    
    Screen.MousePointer = 11
    ViewPro1.StoreCurrentPageObjects
    ViewPro1.SaveObjectDocument file_name
    Screen.MousePointer = 0
End Sub




Private Sub mnuFileSaveAsObjectScript_Click()
    On Error Resume Next
    Dim Filter As String, file_name As String

    Filter = ""
    Filter = Filter & "VPA Object Script Files (*.txt) | *.txt;"
    Filter = Filter & "|All Files (*.*)|*.*;"
    Filter = Filter & "||"
  
    'Show file dialog box
    If ViewPro1.ShowFileDialog(False, EXT_SCRIPT, "", Filter) = 0 Then Exit Sub
    DoEvents

    file_name = ViewPro1.String
    
    Screen.MousePointer = 11
    ViewPro1.StoreCurrentPageObjects
    ViewPro1.SaveScript2 file_name
    Screen.MousePointer = 0

End Sub

Private Sub mnuFileSaveAsScript_Click()
    On Error Resume Next

    Dim Filter As String, file_name As String

    Filter = ""
    Filter = Filter & "ViewPro Script Files (*.txt) | *.txt;"
    Filter = Filter & "|All Files (*.*)|*.*;"
    Filter = Filter & "||"
  
    'Show file dialog box
    If ViewPro1.ShowFileDialog(False, EXT_SCRIPT, "", Filter) = 0 Then Exit Sub
    DoEvents

    file_name = ViewPro1.String
    
    Screen.MousePointer = 11
    ViewPro1.StoreCurrentPageObjects
    ViewPro1.SaveScript file_name
    Screen.MousePointer = 0

End Sub

Private Sub mnuHelpAbout_Click()
    On Error Resume Next

    Dim s As String

    s = program_name & vbCrLf & vbCrLf
    s = s & "Developed with ViewPro Advanced (VPA) Edition" & vbCrLf & vbCrLf
    s = s & "Copyright (c) 2004, HexaTech" & vbCrLf
    s = s & "All Rights Reserved" & vbCrLf

    MsgBox s

End Sub

Private Sub mnuOptionsSetup_Click()
    On Error Resume Next
    
    frmOptions.Show 1
End Sub

Private Sub ViewPro1_DragAndDropMouseUp(ByVal flag As Integer, ByVal x As Long, ByVal y As Long)
    On Error Resume Next

    Dim sName As String, sData As String, sFmt As String, sAttrib As String, sAttribAfter As String
    Dim nType As Integer, nIndex As Integer
    Dim w As Long, h As Long

    

    'MouseUp outside preview page
    If flag = 0 Then Exit Sub
    
    'Deselect all objects
    ViewPro1.SelectObject ""
   ' Picture1.Image = mPic2
    
    
    nIndex = lstObject.ListIndex + 1
    nType = Choose(nIndex, OBJECT_TEXT, OBJECT_PARAGRAPH, OBJECT_TEXTRTF, OBJECT_TABLERTF0, OBJECT_TABLERTF, OBJECT_PICTURE, OBJECT_LINE, OBJECT_LINE, OBJECT_RECTANGLE, OBJECT_CIRCLE, OBJECT_ELLIPSE, OBJECT_POLYOBJECT, OBJECT_CUSTOM, OBJECT_ERASER, OBJECT_MAP, OBJECT_LEGEND, OBJECT_ARROW, OBJECT_SCALE)

    sName = GetUniqueObjectName(ViewPro1, nType)


    'Prepare object
    w = 1 * INCH
    h = 1 * INCH
    If nIndex = 7 Then h = 0
    If nIndex = 8 Then w = 0
    

    sData = ""
    sFmt = ""
    sAttrib = "FontSize=" & DefaultFontSize & ";FontBold=" & DefaultFontBold & "; ForeColor=RGB(0,0,0)"
    sAttribAfter = sAttrib
    
    Select Case nType
        Case OBJECT_TEXT:
                sData = "Your text here"
                
        Case OBJECT_PARAGRAPH:
                sData = "Your paragraph text here"
        
        Case OBJECT_TEXTRTF:
                sData = "Your {\b RTF text} here"
                
        Case OBJECT_TABLERTF0:
                sData = "Header1|Header2|Header3;col1|col2|col3;"
                sFmt = "1440|1440|1440"
                sAttrib = "ObjectBorderStyle=1"
        
        Case OBJECT_TABLERTF:
                sData = "\qc\ul Header1|\qc\ul Header2|\qc\ul Header3;col1|col2|col3;"
                sFmt = "1440|1440|1440"
                sAttrib = "ObjectBorderStyle=1"

        Case OBJECT_PICTURE:
          w = 2 * INCH
            h = 2 * INCH
            
            'Get file name
            If ViewPro1.ShowFileDialog(True, "bmp", "", sFilter) = 0 Then Exit Sub
            DoEvents
            sData = ViewPro1.String
            sData = LCase(sData)
            
            Screen.MousePointer = 11
            
            'Use Visual Basic's LoadPicture to display JPG and GIF file
            If InStr(sData, ".jpg") > 0 Or InStr(sData, ".gif") > 0 Then
                ViewPro1.Picture = LoadPicture(sData)
                
            'Use the original aspect ratio for initial metafile picture size
            ElseIf InStr(sData, ".emf") > 0 Or InStr(sData, ".wmf") > 0 Then
                ViewPro1.GetMetaFileInfo sData
                If ViewPro1.ObjectWidth <> 0 Then h = w * ViewPro1.ObjectHeight / ViewPro1.ObjectWidth
            End If
            nType = OBJECT_PICTURE
            Screen.MousePointer = 0
    
    Case OBJECT_LEGEND:
    
            DoEvents
            Screen.MousePointer = 11
            Set ViewPro1.Picture = PictureLegend.Image ' mPicLegend 'LoadPicture("c:\dude3.emf") 'mPic2 'LoadPicture(sData)
            Screen.MousePointer = 0
            nType = OBJECT_PICTURE
            
        Case OBJECT_ARROW:

            DoEvents
            Screen.MousePointer = 11
            ViewPro1.Picture = PictureArrow.Image ' mPicLegend ' LoadPicture(sLegendPath)
            Screen.MousePointer = 0
            nType = OBJECT_PICTURE
            
         Case OBJECT_MAP:

            DoEvents
            'sData = "c:\dude.bmp"
            'sData = LCase(sData)
            Screen.MousePointer = 11
            ViewPro1.Picture = PictureMap.Image ' mPicMap ' LoadPicture(sData)
            Screen.MousePointer = 0
            nType = OBJECT_PICTURE
           
        Case OBJECT_SCALE:
        
            DoEvents
            'sData = "c:\dude4.bmp"
            'sData = LCase(sData)
            Screen.MousePointer = 11
            ViewPro1.Picture = PictureScale.Image ' mPicScale ' LoadPicture(sData)
            Screen.MousePointer = 0
            nType = OBJECT_PICTURE
           
        Case OBJECT_ELLIPSE:
            w = 2 * INCH

        Case OBJECT_POLYOBJECT:
            w = 1 * INCH
            sData = GetSamplePolyData(ViewPro1, App.Path & "\polydata.txt", sFmt)
            
            
            
    End Select
    
    
    
    If nType = OBJECT_CUSTOM Then
    ViewPro1.AddObject sName, nType, x, y, w, h, sData, sFmt, sAttrib, sAttribAfter
    Else

    'Add object
    ViewPro1.AddObject sName, nType, x, y, w, h, sData, sFmt, sAttrib, sAttribAfter
    
    End If
    ViewPro1.UpdateDoc
    ViewPro1.SelectObject sName


End Sub


Private Sub ViewPro1_KeyEvent(ByVal EventType As Integer, ByVal KeyChar As Integer, ByVal flag As Integer, ByVal RepeatCount As Integer)
    On Error Resume Next
    
Const SHIFT_D = 22.5
Const KEY_UP = 2

'Define movement keys
Const ARROW_L1 = 37   'left arrow
Const ARROW_L2 = 100  'left arrow (num pad)
Const ARROW_L3 = 76   'char L
Const ARROW_R1 = 39   'right arrow
Const ARROW_R2 = 102  'right arrow (num pad)
Const ARROW_R3 = 82   'char R
Const ARROW_U1 = 38   'up arrow
Const ARROW_U2 = 104  'up arrow (num pad)
Const ARROW_U3 = 85   'char U
Const ARROW_D1 = 40   'down arrow
Const ARROW_D2 = 98  'down arrow (num pad)
Const ARROW_D3 = 68   'char D


    '-----Shift selected objects
    If KEY_UP = 2 Then
        Select Case KeyChar
            Case ARROW_L1, ARROW_L2, ARROW_L3:
                ViewPro1.ShiftSelectedObjects -SHIFT_D, 0
            
            Case ARROW_R1, ARROW_R2, ARROW_R3:
                ViewPro1.ShiftSelectedObjects SHIFT_D, 0
            
            Case ARROW_U1, ARROW_U2, ARROW_U3:
                ViewPro1.ShiftSelectedObjects 0, -SHIFT_D
            
            Case ARROW_D1, ARROW_D2, ARROW_D3:
                ViewPro1.ShiftSelectedObjects 0, SHIFT_D
        End Select
    End If
    

End Sub


Private Sub ViewPro1_MouseEvent(ByVal EventType As Integer, ByVal flag As Integer, ByVal x As Long, ByVal y As Long)
    On Error Resume Next
   
     Const MK_LBUTTON = &H1&
     Const MK_RBUTTON = &H2&
     Const MK_SHIFT = &H4&
     Const MK_CONTROL = &H8&

     Const MOUSE_MOVE = 0
     Const MOUSE_LDOWN = 1
     Const MOUSE_RDOWN = 2
     Const MOUSE_LUP = 3
     Const MOUSE_RUP = 4
     Const MOUSE_LDBLCLK = 5
     Const MOUSE_RDBLCLK = 6

Dim sName As String


Select Case EventType
    Case MOUSE_MOVE:
    Case MOUSE_LDOWN:
    
        'Move custom editbox out of the view
        ViewPro1.SetObjectX EDITBOX_OBJECT_NAME, ViewPro1.PageWidth
        
        'Set edit box for table
'        If ViewPro1.ObjectSelect = OBJECT_SELECT_CUSTOM Then
'            sName = ViewPro1.GetObjectFromPosition(x, y)
'            If sName <> "" Then
'                SetEditBoxForTable ViewPro1, sName, EDITBOX_OBJECT_NAME, x, y
'            End If
'        End If
    
    
    Case MOUSE_RDOWN:
    Case MOUSE_LUP:
    Case MOUSE_RUP:
    Case MOUSE_LDBLCLK:
    
        'Do nothing if table mode is on
        If ViewPro1.ObjectSelect = OBJECT_SELECT_CUSTOM Then Exit Sub

        If flag = MK_LBUTTON Then
            If ViewPro1.ObjectEdit = True Then
                        
                'Allow only one object at a time for editing
                If ViewPro1.GetSelectedObjectCount() > 1 Then
                    MsgBox "Please select only one object at a time for editing."
                    Exit Sub
                End If
                
                'Show editing form
                sName = ViewPro1.GetSelectedObject
                If sName <> "" Then
                    frmMapPreviewEdit.ObjName = sName
                    frmMapPreviewEdit.Show 1
                     
                End If
                
             End If
     
        End If
        
    Case MOUSE_RDBLCLK:
        If flag = (MK_RBUTTON Or MK_SHIFT) Then
            'MsgBox "Double click right button with Shift key down"
        End If
        
End Select



End Sub

Private Sub ViewPro1_NewFont(ByVal FontName As String)
    On Error Resume Next

    frmMapPreviewEdit.lstFont.AddItem FontName
End Sub

Private Sub ViewPro1_NewObject(ByVal Name As String, ByVal ObjectType As Integer, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal data As String, ByVal fmt As String)
    On Error Resume Next

'    Dim sData As String
'
'    Select Case ObjectType
'        Case OBJECT_PICTURE:
'            sData = LCase(data)
'            If InStr(sData, ".jpg") > 0 Or InStr(sData, ".gif") > 0 Then
'                If ViewPro1.GetObjectPictureHandle(Name) = 0 Then
'                    ViewPro1.Picture = LoadPicture(sData)
'                End If
'            End If
'
'    End Select


    Select Case ObjectType
        
        Case OBJECT_CUSTOM  'Custom
            Select Case Left(Name, 3)
                Case "cus":
                'Draw happy face
                ViewPro1.DrawEllipse2 x, y, w, h
                ViewPro1.DrawCircle2 x + w / 2 - w / 6, y + h / 3, w / 16
                ViewPro1.DrawCircle2 x + w / 2 + w / 6, y + h / 3, w / 16
                ViewPro1.DrawArc2 x, y - h, x + w, y + h - h / 4, x + w / 2 - w / 4, y + h - h / 4, x + w / 2 + w / 4, y + h - h / 4
                                
            End Select
    End Select
    
    
End Sub

Private Sub ViewPro1_NewPicture(ByVal Page As Long, ByVal Name As String, ByVal Filename As String)
    On Error Resume Next
    
    If ViewPro1.GetObjectPictureHandle(Name) = 0 Then
        ViewPro1.Picture = LoadPicture(Filename)
    End If
    
End Sub

Private Sub ViewPro1_ObjectPositionAfter(ByVal Name As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    On Error Resume Next

    If Name = EDITBOX_OBJECT_NAME Then
        'UpdateTable ViewPro1, EDITBOX_OBJECT_NAME, x, y, w, h
    End If

End Sub

Private Sub ViewPro1_UpdateDocument()
    On Error Resume Next

Screen.MousePointer = 11

ViewPro1.StartDoc

ViewPro1.DrawCurrentPageObjects

If ShowGrid Then
    ViewPro1.PenStyle = 2
    ViewPro1.DrawBoundGrid 1
    ViewPro1.PenStyle = 0
End If

ViewPro1.EndDoc
ViewPro1.Preview

Screen.MousePointer = 0
End Sub


