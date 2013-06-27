VERSION 5.00
Begin VB.Form frmMapPreviewEdit 
   Caption         =   "Edit Object"
   ClientHeight    =   4935
   ClientLeft      =   1590
   ClientTop       =   1155
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4935
   ScaleWidth      =   6645
   Begin VB.TextBox txtH 
      Enabled         =   0   'False
      Height          =   300
      Left            =   5265
      TabIndex        =   21
      Top             =   1290
      Width           =   1035
   End
   Begin VB.TextBox txtW 
      Enabled         =   0   'False
      Height          =   300
      Left            =   3885
      TabIndex        =   20
      Top             =   1275
      Width           =   1035
   End
   Begin VB.TextBox txtY 
      Height          =   300
      Left            =   2355
      TabIndex        =   2
      Top             =   1275
      Width           =   1035
   End
   Begin VB.TextBox txtX 
      Height          =   315
      Left            =   1035
      TabIndex        =   1
      Top             =   1245
      Width           =   915
   End
   Begin VB.ListBox lstFont 
      Height          =   840
      Left            =   3885
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   225
      Width           =   2430
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   5340
      TabIndex        =   8
      Top             =   4455
      Width           =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   390
      Left            =   3795
      TabIndex        =   7
      Top             =   4470
      Width           =   1305
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1065
      TabIndex        =   0
      Top             =   270
      Width           =   1680
   End
   Begin VB.TextBox txtFmt 
      Height          =   315
      Left            =   1020
      TabIndex        =   4
      Top             =   3180
      Width           =   5310
   End
   Begin VB.TextBox txtAttrib 
      Height          =   315
      Left            =   1020
      TabIndex        =   5
      Top             =   3585
      Width           =   5325
   End
   Begin VB.TextBox txtAttribAfter 
      Height          =   285
      Left            =   1020
      TabIndex        =   6
      Top             =   3975
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.TextBox txtData 
      Height          =   1350
      Left            =   1035
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   5280
   End
   Begin VB.Label lblH 
      AutoSize        =   -1  'True
      Caption         =   "H:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   5055
      TabIndex        =   19
      Top             =   1320
      Width           =   165
   End
   Begin VB.Label lblW 
      AutoSize        =   -1  'True
      Caption         =   "W:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3630
      TabIndex        =   18
      Top             =   1320
      Width           =   210
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Y:"
      Height          =   195
      Left            =   2160
      TabIndex        =   17
      Top             =   1320
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "X:"
      Height          =   195
      Left            =   780
      TabIndex        =   16
      Top             =   1305
      Width           =   150
   End
   Begin VB.Label lblFont 
      AutoSize        =   -1  'True
      Caption         =   "Font:"
      Height          =   195
      Left            =   3435
      TabIndex        =   15
      Top             =   240
      Width           =   360
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Name:"
      Height          =   195
      Left            =   480
      TabIndex        =   14
      Top             =   330
      Width           =   465
   End
   Begin VB.Label lblAttribAfter 
      AutoSize        =   -1  'True
      Caption         =   "AttributeAfter:"
      Height          =   195
      Left            =   45
      TabIndex        =   13
      Top             =   3990
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblAttrib 
      AutoSize        =   -1  'True
      Caption         =   "Attribute:"
      Height          =   195
      Left            =   315
      TabIndex        =   12
      Top             =   3630
      Width           =   630
   End
   Begin VB.Label lblFmt 
      AutoSize        =   -1  'True
      Caption         =   "Format:"
      Height          =   195
      Left            =   435
      TabIndex        =   11
      Top             =   3210
      Width           =   525
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      Caption         =   "Data:"
      Height          =   195
      Left            =   510
      TabIndex        =   10
      Top             =   1740
      Width           =   390
   End
End
Attribute VB_Name = "frmMapPreviewEdit"
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



Public ObjName As String

Dim vp As Control
Dim nType As Integer, x As Long, y As Long, w As Long, h As Long
Dim data As String, fmt As String, attrib As String, attrib_after As String

Sub gsCenterWindow(Window As Form)
    On Error Resume Next
    
   Window.Left = (Screen.Width / 2) - (Window.Width / 2)
   Window.Top = (Screen.Height / 2) - (Window.Height / 2)
      
      
End Sub



Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    
Const MODIFY_NONE = &H0
Const MODIFY_NAME = &H40
Const MODIFY_POS = &H1
Const MODIFY_SIZE = &H2
Const MODIFY_DATA = &H4
Const MODIFY_FMT = &H8
Const MODIFY_ATTRIB = &H10
Const MODIFY_ATTRIBAFTER = &H20


Const REPLACE_NONE = &H0
Const REPLACE_ATTRIB = &H10
Const REPLACE_ATTRIBAFTER = &H20
Dim Modify As Long, Replace As Long

'Replace = REPLACE_NONE
Modify = MODIFY_NAME Or MODIFY_POS Or MODIFY_DATA Or MODIFY_FMT Or MODIFY_ATTRIB Or REPLACE_ATTRIBAFTER
Replace = REPLACE_ATTRIB Or REPLACE_ATTRIBAFTER
    
    
    If vp.GetObjectIndex(txtName.Text) <> 0 And txtName.Text <> ObjName Then
        MsgBox "Another object already uses this name: " & txtName.Text
        Exit Sub
    End If
                
    vp.ModifyObject txtName.Text, Val(txtX.Text), Val(txtY.Text), Val(txtW.Text), Val(txtH.Text), txtData.Text, txtFmt.Text, txtAttrib.Text, txtAttribAfter.Text, Modify, Replace
    vp.UpdateDoc


    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next

    gsCenterWindow Me

    Set vp = frmMapPreviewDesigner.ViewPro1

    Dim sObjectType As String
    
    Screen.MousePointer = 11

    '-----Get the object info
    nType = vp.GetObjectInfo(ObjName)
    x = vp.x
    y = vp.y
    w = vp.ObjectWidth
    h = vp.ObjectHeight
    data = vp.String
    fmt = vp.TableFormat
    attrib = vp.String2
    attrib_after = vp.String3
        
    '-----Fill list box with font names
    vp.EnumerateFonts 0
            
        
    '-----Fill text boxes
    txtName.Text = ObjName
    txtX.Text = "" & x
    txtY.Text = "" & y
    txtW.Text = "" & w
    txtH.Text = "" & h
    txtData.Text = data
    txtFmt.Text = fmt
    txtAttrib.Text = attrib
    txtAttribAfter.Text = attrib_after
    
    '-----Appearance
    txtData.Enabled = False
    txtFmt.Enabled = False
    lblData.Enabled = False
    lblFmt.Enabled = False
    lblFont.Visible = False
    lstFont.Visible = False
    lblW.Enabled = False
    txtW.Enabled = False
    lblH.Enabled = False
    txtH.Enabled = False
    
    Select Case nType
        Case OBJECT_TEXT:       sObjectType = "Text"
                txtData.Enabled = True
                lblData.Enabled = True
                lblFont.Visible = True
                lstFont.Visible = True
                
        Case OBJECT_PARAGRAPH:  sObjectType = "Paragraph"
                txtData.Enabled = True
                lblData.Enabled = True
                lblFont.Visible = True
                lstFont.Visible = True
                
        Case OBJECT_TEXTRTF:    sObjectType = "TextRTF"
                txtData.Enabled = True
                lblData.Enabled = True
                lblFont.Visible = True
                lstFont.Visible = True
        
        Case OBJECT_TABLERTF0:  sObjectType = "Table"
                txtData.Enabled = True
                txtFmt.Enabled = True
                lblData.Enabled = True
                lblFmt.Enabled = True
                lblFont.Visible = True
                lstFont.Visible = True
                
        Case OBJECT_TABLERTF:   sObjectType = "TableRTF"
                txtData.Enabled = True
                txtFmt.Enabled = True
                lblData.Enabled = True
                lblFmt.Enabled = True
                lblFont.Visible = True
                lstFont.Visible = True
                
        Case OBJECT_PICTURE:    sObjectType = "Picture"
                txtData.Enabled = True
                lblData.Enabled = True
        
        
        Case OBJECT_LINE:       sObjectType = "HorzLine"
        Case OBJECT_RECTANGLE:  sObjectType = "Rectangle"
        Case OBJECT_CIRCLE:     sObjectType = "Circle"
        Case OBJECT_ELLIPSE:    sObjectType = "Ellipse"
                'lblW.Enabled = True
                'txtW.Enabled = True
                'lblH.Enabled = True
                'txtH.Enabled = True
        
        Case OBJECT_POLYOBJECT:    sObjectType = "PolyObject"
                txtData.Enabled = True
                txtFmt.Enabled = True
                lblData.Enabled = True
                lblFmt.Enabled = True
        
        
        Case OBJECT_CUSTOM:    sObjectType = "Custom"
        
        Case OBJECT_ERASER:  sObjectType = "Eraser"
        
    End Select
    If sObjectType = "HorzLine" And w = 0 Then sObjectType = "VertLine"
    Me.caption = "Edit Object - " & sObjectType
    
    Screen.MousePointer = 0
    
    

End Sub


Private Sub lstFont_Click()
    On Error Resume Next
Dim sAttrib As String
Dim sPrevFontName As String, sCurrentFontName As String

    sCurrentFontName = lstFont.List(lstFont.ListIndex)
    
    sAttrib = Trim(txtAttrib.Text)
    sPrevFontName = vp.ParseStr(sAttrib, """", """")
    
    'There is no previous font
    If sPrevFontName = "" Then
        If sAttrib <> "" Then sAttrib = sAttrib & ";"
        'txtAttrib.Text = sAttrib & "FontName=15;SetCustomFontName """ & sCurrentFontName & """"
        txtAttrib.Text = sAttrib & "FontName=15;ObjectFontName=""" & sCurrentFontName & """"
        
    'Previous font is present
    ElseIf sCurrentFontName <> sPrevFontName Then
        'txtAttrib.Text = vp.ReplaceStr(sAttrib, sPrevFontName, sCurrentFontName)  'this does not work for multi-word name
        sAttrib = vp.ReplaceStr(sAttrib, sPrevFontName, "_")
        txtAttrib.Text = vp.ReplaceStr(sAttrib, "_", sCurrentFontName)
    End If
    

End Sub



