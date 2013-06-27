VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmImplementingPartner 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4590
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvImplentingPartners 
      Height          =   3255
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmbOK 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   45
      TabIndex        =   0
      Top             =   4185
      Width           =   1110
   End
   Begin MSAdodcLib.Adodc AdodcOrg 
      Height          =   510
      Left            =   1395
      Top             =   4005
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "AdodcOrg"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmImplementingPartner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_bInitialized
Private m_cn As ADODB.Connection
Dim bc As BindingCollection

Public Sub Load(Xleft As Long, _
                YTop As Long, _
                ByVal st As dxStyleController)
        '<EhHeader>
        On Error GoTo Load_Err
        '</EhHeader>
        On Error Resume Next
    '    Dim dx, dy
    '
    '    Set bc = New BindingCollection
    '    Set bc.DataSource = frmMain.Adodc1
    '
    '    With dxStyleController1
    '        .BackColor = st.BackColor
    '        .BorderColor = st.BorderColor
    '        .BorderStyle = st.BorderStyle
    '        .ButtonStyle = st.ButtonStyle
    '        .Shadow = st.Shadow
    '        .HotTrack = st.HotTrack
    '        Set .Font = st.Font
    '        Set .PopupFont = st.PopupFont
    '        dxTextEdit1.StyleController = .Name
    '        dxTextEdit2.StyleController = .Name
    '        dxPickEdit1.StyleController = .Name
    '        dxMaskEdit1.StyleController = .Name
    '
            Dim Rr!, Rb!
100         Rr = Screen.Width
102         Rb = Screen.Height

104         If Xleft + Width > Rr Then Xleft = Rr - Width
106         If YTop + Height > Rb Then YTop = Rb - Height
108         Move Xleft, YTop
    '
    '        bc.Add dxTextEdit1, "DataBind", "Address"
    '        bc.Add dxTextEdit2, "DataBind", "City"
    '        bc.Add dxPickEdit1, "DataBind", "State"
    '        bc.Add dxMaskEdit1, "DataBind", "ZipCode"
    '
    '    End With
    '
110     Visible = True
        '<EhFooter>
        Exit Sub

Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmImplementingPartner.Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Public Sub init(sConnectionString)
        'Set m_cn = cn
        '<EhHeader>
        On Error GoTo init_Err
        '</EhHeader>
    
100     With AdodcOrg
102         .ConnectionString = sConnectionString
104         .RecordSource = "organisations"
106         .Refresh
        End With
    
108     AdodcOrg.Recordset.MoveFirst
    
110     lvImplentingPartners.ColumnHeaders.Add Text:="Organisation"
    
112     Do While Not AdodcOrg.Recordset.EOF
114         lvImplentingPartners.ListItems.Add Text:=AdodcOrg.Recordset.Fields.Item("name").Value
116         AdodcOrg.Recordset.MoveNext
        Loop
    
118     m_bInitialized = True
        '<EhFooter>
        Exit Sub

init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmImplementingPartner.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Property Get Initialized() As Boolean
        '<EhHeader>
        On Error GoTo Initialized_Err
        '</EhHeader>
100     Initialized = m_bInitialized
        '<EhFooter>
        Exit Property

Initialized_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmImplementingPartner.Initialized " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Property

Private Sub cmbOK_Click()
        '<EhHeader>
        On Error GoTo cmbOK_Click_Err
        '</EhHeader>
100  Unload Me
        '<EhFooter>
        Exit Sub

cmbOK_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmImplementingPartner.cmbOK_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     If Not g_sLanguage = "" Then
102         If Not m_Cnn.State = adStateClosed Then
104             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmImplementingPartner.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
