VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSynchReader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Synch Explorer"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14430
   Icon            =   "frmSynchReader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   962
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHrmmm 
      Caption         =   "Hrmmm"
      Height          =   525
      Left            =   6630
      TabIndex        =   53
      Top             =   3330
      Visible         =   0   'False
      Width           =   1245
   End
   Begin C1SizerLibCtl.C1Tab C1TTab1Tab2 
      Height          =   6765
      Left            =   10500
      TabIndex        =   4
      Top             =   0
      Width           =   3915
      _cx             =   6906
      _cy             =   11933
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "Settings|Values|Tab&3|Synch Table"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Begin VB.Frame Frame3 
         Caption         =   "Settings"
         Height          =   6390
         Left            =   5160
         TabIndex        =   51
         Top             =   330
         Width           =   3825
         Begin VB.Frame FraServer 
            Caption         =   "Server:"
            Height          =   1755
            Left            =   60
            TabIndex        =   66
            Top             =   180
            Width           =   3735
            Begin VB.ComboBox ComServertables 
               Height          =   315
               Left            =   60
               Style           =   2  'Dropdown List
               TabIndex        =   70
               Top             =   900
               Width           =   3615
            End
            Begin VB.CommandButton cmdGetServer 
               Caption         =   "Get Server"
               Height          =   285
               Left            =   2640
               TabIndex        =   69
               Top             =   1320
               Width           =   1005
            End
            Begin VB.TextBox txtSynchURL 
               Height          =   315
               Left            =   30
               TabIndex        =   67
               Text            =   "http://www.immap.org/OAS2/oasis.asp"
               Top             =   510
               Width           =   3645
            End
            Begin VB.Label lblServerAddress1 
               Caption         =   "Server Address:"
               Height          =   315
               Left            =   0
               TabIndex        =   68
               Top             =   240
               Width           =   1185
            End
         End
         Begin VB.Frame FraLocal 
            Caption         =   "Local:"
            Height          =   3825
            Left            =   60
            TabIndex        =   54
            Top             =   1950
            Width           =   3765
            Begin VB.CommandButton cmdConnectToDB 
               Caption         =   "..."
               Height          =   255
               Left            =   3420
               TabIndex        =   62
               Top             =   240
               Width           =   285
            End
            Begin VB.TextBox txtLocalDB 
               BeginProperty Font 
                  Name            =   "MS Serif"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   61
               Text            =   "C:\OASIS\Client\data\Colombia\data\db\Oasisclient.mdb"
               Top             =   240
               Width           =   3405
            End
            Begin VB.CommandButton cmdLoad 
               Caption         =   "Load"
               Height          =   315
               Left            =   3180
               TabIndex        =   60
               Top             =   540
               Width           =   525
            End
            Begin VB.ComboBox ComTables 
               Height          =   315
               Left            =   600
               Style           =   2  'Dropdown List
               TabIndex        =   59
               Top             =   540
               Width           =   2565
            End
            Begin VB.TextBox txtName 
               Height          =   285
               Left            =   1200
               TabIndex        =   58
               Top             =   930
               Width           =   2445
            End
            Begin VB.TextBox txtDescription 
               Height          =   1935
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   57
               Top             =   1470
               Width           =   3525
            End
            Begin VB.CheckBox chkAllowSyncUpdate 
               Caption         =   "Allow Update"
               Height          =   315
               Left            =   120
               TabIndex        =   56
               Top             =   3480
               Width           =   1335
            End
            Begin VB.CheckBox chkReadWrite 
               Caption         =   "Read Write"
               Height          =   255
               Left            =   1650
               TabIndex        =   55
               Top             =   3510
               Width           =   1485
            End
            Begin VB.Label lblTables 
               AutoSize        =   -1  'True
               Caption         =   "Tables:"
               Height          =   195
               Left            =   30
               TabIndex        =   65
               Top             =   600
               Width           =   525
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Synch Name:"
               Height          =   195
               Left            =   120
               TabIndex        =   64
               Top             =   930
               Width           =   960
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Description:"
               Height          =   195
               Left            =   150
               TabIndex        =   63
               Top             =   1230
               Width           =   840
            End
         End
         Begin VB.CommandButton cmdCreateSynch 
            Caption         =   "Create Synch "
            Height          =   315
            Left            =   2670
            TabIndex        =   52
            Top             =   5850
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Frame2"
         Height          =   6390
         Left            =   4860
         TabIndex        =   32
         Top             =   330
         Width           =   3825
         Begin VB.CommandButton cmdCreate1Feed 
            Caption         =   "Create1Feed"
            Height          =   465
            Left            =   2490
            TabIndex        =   33
            Top             =   5550
            Width           =   1230
         End
         Begin MSComctlLib.TreeView tvwMenu 
            Height          =   5865
            Left            =   90
            TabIndex        =   34
            Top             =   240
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   10345
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   265
            LabelEdit       =   1
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            Appearance      =   0
         End
      End
      Begin VB.Frame Frame1 
         ClipControls    =   0   'False
         Height          =   6390
         Left            =   4560
         TabIndex        =   15
         Top             =   330
         Width           =   3825
         Begin VB.CommandButton cmdGUIDGEN 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3270
            TabIndex        =   50
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton cmdControl 
            Caption         =   "Merge"
            Height          =   285
            Left            =   660
            TabIndex        =   45
            Top             =   6060
            Width           =   1485
         End
         Begin VB.CommandButton cmdDb 
            Caption         =   "Update Item"
            Height          =   285
            Index           =   2
            Left            =   2220
            TabIndex        =   44
            Top             =   6060
            Width           =   1485
         End
         Begin VB.CommandButton cmdDb 
            Caption         =   "Insert item"
            Height          =   285
            Index           =   3
            Left            =   660
            TabIndex        =   43
            Top             =   5760
            Width           =   1485
         End
         Begin VB.CommandButton cmdDb 
            Caption         =   "Delete item"
            Height          =   285
            Index           =   4
            Left            =   2220
            TabIndex        =   42
            Top             =   5760
            Width           =   1485
         End
         Begin VB.ComboBox ComItemID 
            Height          =   315
            Left            =   60
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   420
            Width           =   3675
         End
         Begin VB.CheckBox chkISGeo 
            Caption         =   "IS Geo Table"
            Height          =   255
            Left            =   1440
            TabIndex        =   26
            Top             =   3780
            Width           =   1395
         End
         Begin VB.TextBox txtFeedInfo 
            Height          =   330
            Index           =   5
            Left            =   60
            TabIndex        =   25
            Text            =   "DudedTable"
            Top             =   3360
            Width           =   3615
         End
         Begin VB.CommandButton cmdDb 
            Caption         =   "New Full"
            Height          =   300
            Index           =   0
            Left            =   2220
            TabIndex        =   24
            Top             =   5340
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VB.CheckBox chkSave 
            Caption         =   "Save"
            Height          =   240
            Left            =   75
            TabIndex        =   23
            Top             =   3780
            Width           =   1095
         End
         Begin VB.TextBox txtFeedInfo 
            Height          =   330
            Index           =   4
            Left            =   75
            TabIndex        =   22
            Text            =   "Duded"
            Top             =   2760
            Width           =   3615
         End
         Begin VB.TextBox txtFeedInfo 
            Height          =   330
            Index           =   3
            Left            =   75
            TabIndex        =   21
            Text            =   "Description"
            Top             =   2175
            Width           =   3615
         End
         Begin VB.TextBox txtFeedInfo 
            Height          =   330
            Index           =   2
            Left            =   75
            TabIndex        =   20
            Text            =   "MyTitle"
            Top             =   1635
            Width           =   3615
         End
         Begin VB.TextBox txtFeedInfo 
            Height          =   330
            Index           =   1
            Left            =   75
            TabIndex        =   19
            Text            =   "sasafasffaf"
            Top             =   1050
            Width           =   3105
         End
         Begin VB.CommandButton cmdFeeds 
            Caption         =   "Delete"
            Height          =   510
            Index           =   2
            Left            =   2430
            TabIndex        =   18
            Top             =   6435
            Width           =   1140
         End
         Begin VB.CommandButton cmdFeeds 
            Caption         =   "Update"
            Height          =   510
            Index           =   1
            Left            =   1260
            TabIndex        =   17
            Top             =   6435
            Width           =   1140
         End
         Begin VB.CommandButton cmdFeeds 
            Caption         =   "Create"
            Height          =   510
            Index           =   0
            Left            =   90
            TabIndex        =   16
            Top             =   6435
            Width           =   1140
         End
         Begin VB.Label lblAvailableGUIDs 
            Caption         =   "Active item GUID:"
            Height          =   225
            Left            =   90
            TabIndex        =   38
            Top             =   150
            Width           =   2115
         End
         Begin VB.Label Label5 
            Caption         =   "Table"
            Height          =   240
            Left            =   60
            TabIndex        =   31
            Top             =   3090
            Width           =   1545
         End
         Begin VB.Label Label4 
            Caption         =   "By"
            Height          =   240
            Left            =   75
            TabIndex        =   30
            Top             =   2535
            Width           =   1545
         End
         Begin VB.Label Label3 
            Caption         =   "Description"
            Height          =   285
            Left            =   75
            TabIndex        =   29
            Top             =   1950
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Title"
            Height          =   330
            Left            =   30
            TabIndex        =   28
            Top             =   1410
            Width           =   1320
         End
         Begin VB.Label lblLabel2 
            Caption         =   "GUID"
            Height          =   240
            Left            =   75
            TabIndex        =   27
            Top             =   825
            Width           =   1230
         End
      End
      Begin VB.Frame FraSettings 
         ClipControls    =   0   'False
         Height          =   6390
         Left            =   45
         TabIndex        =   5
         Top             =   330
         Width           =   3825
         Begin VB.TextBox txtFeedInfo 
            Height          =   330
            Index           =   0
            Left            =   150
            TabIndex        =   48
            Text            =   "C:\OASIS\Client\XML_FOLDER\xml\dude2.xml"
            Top             =   2190
            Width           =   3615
         End
         Begin VB.TextBox txtLocalDB 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Text            =   "C:\OASIS\Client\data\Colombia\data\db\Oasisclient.mdb"
            Top             =   1620
            Width           =   3645
         End
         Begin VB.Frame FraWorkMode 
            Caption         =   "Work Mode:"
            ClipControls    =   0   'False
            Height          =   525
            Left            =   90
            TabIndex        =   39
            Top             =   120
            Width           =   3645
            Begin VB.OptionButton OptMode 
               Caption         =   "Local"
               Height          =   225
               Index           =   0
               Left            =   60
               TabIndex        =   41
               Top             =   210
               Value           =   -1  'True
               Width           =   795
            End
            Begin VB.OptionButton OptMode 
               Caption         =   "Server"
               Height          =   255
               Index           =   1
               Left            =   900
               TabIndex        =   40
               Top             =   180
               Width           =   765
            End
         End
         Begin VB.TextBox txtServer 
            Height          =   315
            Left            =   120
            TabIndex        =   35
            Text            =   "http://www.immap.org/OASIS/oasis.asp"
            Top             =   930
            Width           =   3645
         End
         Begin VB.CheckBox chkIncludeDeleted 
            Caption         =   "Include Deleted"
            Height          =   255
            Left            =   2340
            TabIndex        =   12
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Frame FraTimeConstraint 
            Caption         =   "Time Constraint:"
            ClipControls    =   0   'False
            Height          =   2955
            Left            =   150
            TabIndex        =   9
            Top             =   2760
            Width           =   3285
            Begin OASISRemoteAdmin.TimeControl TimeControl1 
               Height          =   285
               Left            =   2160
               TabIndex        =   71
               Top             =   2640
               Width           =   990
               _ExtentX        =   1746
               _ExtentY        =   503
            End
            Begin MSComCtl2.MonthView MonthView1 
               Height          =   2370
               Left            =   90
               TabIndex        =   10
               Top             =   210
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   4180
               _Version        =   393216
               ForeColor       =   -2147483630
               BackColor       =   -2147483633
               Appearance      =   1
               ShowWeekNumbers =   -1  'True
               StartOfWeek     =   108331009
               CurrentDate     =   39724
            End
            Begin VB.Label lblTIME 
               AutoSize        =   -1  'True
               Caption         =   "TIME:"
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Top             =   2640
               Width           =   435
            End
         End
         Begin VB.CheckBox chkUseTime 
            Caption         =   "Use Time/Date Constraint:"
            Height          =   315
            Left            =   150
            TabIndex        =   8
            Top             =   2490
            Width           =   2235
         End
         Begin VB.TextBox txtSynchTable 
            Height          =   285
            Index           =   0
            Left            =   1650
            TabIndex        =   7
            Text            =   "111Feed"
            Top             =   5760
            Width           =   2115
         End
         Begin VB.TextBox txtSynchTable 
            Height          =   255
            Index           =   1
            Left            =   1650
            TabIndex        =   6
            Text            =   "111FeedsHistory"
            Top             =   6060
            Width           =   2115
         End
         Begin VB.Label Label1 
            Caption         =   "Path"
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label lblLocalDB 
            Caption         =   "Local DB:"
            Height          =   255
            Left            =   150
            TabIndex        =   47
            Top             =   1320
            Width           =   2205
         End
         Begin VB.Label lblServerAddress 
            Caption         =   "Server Address:"
            Height          =   315
            Left            =   90
            TabIndex        =   36
            Top             =   660
            Width           =   1185
         End
         Begin VB.Label lblSynchTable 
            AutoSize        =   -1  'True
            Caption         =   "Synch Table Name:"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   14
            Top             =   5820
            Width           =   1410
         End
         Begin VB.Label lblSynchTable 
            AutoSize        =   -1  'True
            Caption         =   "Table History Name:"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   13
            Top             =   6090
            Width           =   1440
         End
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   525
      Left            =   8970
      TabIndex        =   3
      Top             =   6270
      Visible         =   0   'False
      Width           =   1245
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7065
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   10395
      ExtentX         =   18336
      ExtentY         =   12462
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.CommandButton cmdDb 
      Caption         =   "Get Synch"
      Height          =   285
      Index           =   1
      Left            =   12810
      TabIndex        =   1
      Top             =   6810
      Width           =   1485
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   510
      Left            =   5895
      Top             =   3105
      Width           =   3165
      _ExtentX        =   5583
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\Client\data\Colombia\data\db\Oasisclient.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\Client\data\Colombia\data\db\Oasisclient.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.TextBox txtXMLOut 
      Height          =   1935
      Left            =   10230
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmSynchReader.frx":0442
      Top             =   30
      Visible         =   0   'False
      Width           =   3240
   End
End
Attribute VB_Name = "frmSynchReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private WithEvents oSynchs As clsSynchronisations
Attribute oSynchs.VB_VarHelpID = -1
Private mdtLastDownload As Date
Private Declare Function CoCreateGuid _
                Lib "ole32" (id As Any) As Long
Private m_RSServerFeed As ADODB.Recordset
Private WithEvents m_RSServerItems As ADODB.Recordset
Attribute m_RSServerItems.VB_VarHelpID = -1
Private m_Col_Local As Collection
Private m_Col_Incoming As Collection
Private m_Col_Merged As Collection
Private m_Col_History As Collection

Public Function CreateGUID() As String
    Dim id(0 To 15) As Byte
    Dim Cnt As Long, GUID As String

    If CoCreateGuid(id(0)) = 0 Then

        For Cnt = 0 To 15
            CreateGUID = CreateGUID + IIf(id(Cnt) < 16, "0", "") + Hex$(id(Cnt))
        Next Cnt

        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
    Else
        MsgBox "Error while creating GUID!"
    End If

End Function


Private Sub cmdControl_Click()
    Dim cn As New ADODB.Connection
    'Dim oRSLatestUpdate As New ADODB.Recordset
       
    cn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtLocalDB(0).Text & ";Persist Security Info=False"
    
    ' oRSLatestUpdate.Open "SELECT Max([when]) FROM " & txtSynchTable(1).Text, cn, adOpenDynamic
    
    CheckNewRemoteItems cn
    OptMode(0).Value = True
    cmdDb_Click 0
    'CreateSynchCollections cn

End Sub

Private Sub CreateSynchCollections(cn As ADODB.Connection)
    Dim oRSLocalItms As New ADODB.Recordset
    Dim colNewItems As New Collection
    Dim colExistingItems As New Collection
    Dim sNewItemSQL As String
    Dim sExistingItemSQL As String
    Dim i As Integer
    
    Set m_RSServerItems = New ADODB.Recordset
    
    'If no local item exists with the same id value as the incoming item, add the incoming item to the merge result collection; we are done processing the incoming item
    m_RSServerItems.CursorLocation = adUseClient
    m_RSServerItems.Properties("Initial Fetch Size") = 2
    m_RSServerItems.Properties("Background Fetch Size") = 4

    m_RSServerItems.open txtServer.Text & "?ID=SELECT sGUID FROM " & txtSynchTable(0).Text, Options:=adAsyncFetch

    'Check if there are any returned records
    If Not m_RSServerItems.EOF And Not m_RSServerItems.Bof Then
        
        m_RSServerItems.MoveFirst
        
        Do While Not m_RSServerItems.EOF
            Set oRSLocalItms = New ADODB.Recordset
            oRSLocalItms.open "SELECT sGUID FROM " & txtSynchTable(0).Text & "WHERE sGUID = '" & m_RSServerItems.fields.Item(0).Value & "'", cn
            
            'Check if the item exists
            If Not oRSLocalItms.EOF And Not oRSLocalItms.Bof Then
                If Not sExistingItemSQL = "" Then
                    sExistingItemSQL = sExistingItemSQL & ",'" & m_RSServerItems.fields.Item(0).Value & "'"
                Else
                    sExistingItemSQL = "'" & m_RSServerItems.fields.Item(0).Value & "'"
                End If

                colExistingItems.Add m_RSServerItems.fields.Item(0).Value
            Else

                If Not sNewItemSQL = "" Then
                    sNewItemSQL = sNewItemSQL & ",'" & m_RSServerItems.fields.Item(0).Value & "'"
                Else
                    sNewItemSQL = "'" & m_RSServerItems.fields.Item(0).Value & "'"
                End If

                m_Col_Merged.Add m_RSServerItems.fields.Item(0).Value
                'TODO import
            End If
            
            m_RSServerItems.MoveNext
            
        Loop

    End If
        
    'OK Now we have added the new items into the Merged Collection
        
    If Not sNewItemSQL = "" Then
        Set m_RSServerItems = New ADODB.Recordset
        'SELECT * FROM " & txtSynchTable(0).Text & "WHERE NOT sGUID IN ('{A9908575-3E8A-4157-A3EC-828D13D1BB2B}','{4F1BBCFE-CF87-4979-BE32-7DE8EABAC8A2}','{658B0035-652A-4096-B9ED-8B85F2BE452C}','{3DA2B022-50D2-42DF-BA61-DAFD4EAC0149}','{7BF1E5B8-F03D-49B1-8E0D-BFDD4CC4D0FC}','{AA3119CE-3A19-459A-BA11-6DE6018D8433}','{983D0B6C-5493-4141-8CC1-024DA6DACDB2}','{8776501C-9FCE-4B9F-AA8C-CB900D6C0C86}','{A6FFCFC4-0A5A-40B8-9E77-6D86757298E1}')

        sNewItemSQL = "SELECT * FROM " & txtSynchTable(0).Text & "WHERE sGUID IN (" & sNewItemSQL & ")"
        m_RSServerItems.open txtServer.Text & "?ID=" & sNewItemSQL
        
        If Not m_RSServerItems.EOF And Not m_RSServerItems.Bof Then
        
            If Not m_RSServerItems.Bof Then m_RSServerItems.MoveFirst
            
            Set oRSLocalItms = New ADODB.Recordset
            oRSLocalItms.open "SELECT * FROM " & txtSynchTable(0).Text, cn, adOpenDynamic, adLockOptimistic
            
            With oRSLocalItms
            
                Do While Not m_RSServerItems.EOF
                    .AddNew
                    .fields.Item("sGUID").Value = m_RSServerItems.fields.Item("sGUID").Value
                    .fields.Item("sLocalID").Value = m_RSServerItems.fields.Item("sLocalID").Value
                    .fields.Item("sBy").Value = m_RSServerItems.fields.Item("sBy").Value 'BY
                    .fields.Item("sDescription").Value = m_RSServerItems.fields.Item("sDescription").Value '"[Description]"
                    .fields.Item("Title").Value = m_RSServerItems.fields.Item("Title").Value
                    .fields.Item("time").Value = m_RSServerItems.fields.Item("time").Value
                    .fields.Item("deleted").Value = m_RSServerItems.fields.Item("deleted").Value
                    .fields.Item("TableName").Value = m_RSServerItems.fields.Item("TableName").Value
                    .fields.Item("isGeoTable").Value = m_RSServerItems.fields.Item("isGeoTable").Value
                    .Update
                    m_RSServerItems.MoveNext
                Loop
            
            End With
        
        End If
    End If
        
    If Not sExistingItemSQL = "" Then
        Set m_RSServerItems = New ADODB.Recordset
        sExistingItemSQL = "SELECT * FROM " & txtSynchTable(0).Text & "WHERE sGUID IN (" & sExistingItemSQL & ")"
        m_RSServerItems.open txtServer.Text & "?ID=" & sExistingItemSQL
        
        If Not m_RSServerItems.EOF And Not m_RSServerItems.Bof Then
        
            If Not m_RSServerItems.Bof Then m_RSServerItems.MoveFirst
            
            Set oRSLocalItms = New ADODB.Recordset
            oRSLocalItms.open "SELECT * FROM " & txtSynchTable(0).Text, cn, adOpenDynamic, adLockOptimistic
            
            With oRSLocalItms

                If Not .EOF And Not .Bof Then
                    
                    Do While Not m_RSServerItems.EOF

                        If Not .Bof Then .MoveFirst
                        .Find "sGUID = '" & m_RSServerItems.fields.Item("sGUID").Value & "'"
                        
                        If Not .EOF Then
                            .fields.Item("sLocalID").Value = m_RSServerItems.fields.Item("sLocalID").Value
                            .fields.Item("sBy").Value = m_RSServerItems.fields.Item("sBy").Value
                            .fields.Item("sDescription").Value = m_RSServerItems.fields.Item("sDescription").Value
                            .fields.Item("Title").Value = m_RSServerItems.fields.Item("Title").Value
                            .fields.Item("time").Value = m_RSServerItems.fields.Item("time").Value
                            .fields.Item("deleted").Value = m_RSServerItems.fields.Item("deleted").Value
                            .fields.Item("TableName").Value = m_RSServerItems.fields.Item("TableName").Value
                            .fields.Item("isGeoTable").Value = m_RSServerItems.fields.Item("isGeoTable").Value
                            .Update
                        End If
                        
                        m_RSServerItems.MoveNext
                    Loop

                End If

            End With
        
        End If
    End If
        
    '    'Create SQL to get the new items
    '    For i = 1 To colNewItems.Count
    '
    '        If Not sNewItemSQL = "" Then
    '            sNewItemSQL = sNewItemSQL & ",'" & colNewItems.Item(i) & "'"
    '        Else
    '            sNewItemSQL = "'" & colNewItems.Item(i) & "'"
    '        End If
    '
    '    Next i
    '
    '    sNewItemSQL = "SELECT * FROM " & txtSynchTable(0).Text & "WHERE sGUID IN (" & sNewItemSQL & ")"
    '
    m_frmDebug.DebugPrint sNewItemSQL
    
    'SELECT "column_name"
    'From "table_name"
    'WHERE "column_name" IN ('value1', 'value2', ...)

End Sub

Private Sub cmdCreate1Feed_Click()

    oSynchs.LoadAllChannels
    
End Sub

Sub FormatXML(oChannel As MSXML2.IXMLDOMElement)
    Dim i As Integer
    Dim oDoc As MSXML2.DOMDocument
    
    oDoc
    
End Sub

Private Function GetHistorySequence(sGUID As String, _
                                    cn As ADODB.Connection) As Long
    Dim RS As New ADODB.Recordset
    
    RS.CursorLocation = adUseClient
    
    RS.open "SELECT sGUID, sequence FROM " & txtSynchTable(1).Text & " WHERE sGUID = '" & sGUID & "' AND deleted <> 'true' ORDER BY sequence DESC", cn, adOpenDynamic, adLockReadOnly
    
    If RS.EOF And RS.Bof Then
        GetHistorySequence = 1
    Else
        
        GetHistorySequence = RS.RecordCount + 1
    End If
    
    RS.Close
    
    Set RS = Nothing
    
End Function

Private Function GetRSSSQL() As String
    Dim sRFC3339DateTime As String
    
    If chkUseTime.Value = vbChecked Then
        With MonthView1
            sRFC3339DateTime = RFC3339DateTimeEX(CInt(.Year), CStr(.Month), CStr(.Day), CStr(TimeControl1.H), CStr(TimeControl1.Min), CStr(TimeControl1.S))
        End With

    End If

    'select * From 111Feed, " & txtSynchTable(1).Text & " WHERE [].sGUID = [History].sGUID AND [History].when = '2008-10-02T12:02:02Z'

    GetRSSSQL = "select * From " & txtSynchTable(0).Text & ", " & txtSynchTable(1).Text & " WHERE [" & txtSynchTable(0).Text & "].sGUID = [" & txtSynchTable(1).Text & "].sGUID" & IIf(sRFC3339DateTime <> "", " AND [" & txtSynchTable(1).Text & "].swhen = '" & sRFC3339DateTime & "'", "")
    
    '"SELECT * FROM " & txtSynchTable(0).Text

    '    If chkIncludeDeleted.Value = vbChecked Then CreateRSSSQL = CreateRSSSQL & " WHERE "
    
End Function

Private Function CompareRFC3339DateTime(sDate1 As String, _
                                        sDate2 As String) As Integer
    Dim aDate1() As String
    Dim aDate2() As String
    Dim aDatum() As String
    Dim bDatum() As String
        
    aDatum = Split(Left$(sDate1, InStr(sDate1, "T") - 1), "-")
    bDatum = Split(Left$(sDate2, InStr(sDate2, "T") - 1), "-")
       
    'Check the Year
    If aDatum(0) <> bDatum(0) Then
        CompareRFC3339DateTime = IIf(CInt(aDatum(0)) > CInt(bDatum(0)), 0, 1)
    ElseIf aDatum(1) <> bDatum(1) Then
        CompareRFC3339DateTime = IIf(CInt(aDatum(1)) > CInt(bDatum(1)), 0, 1)
    ElseIf aDatum(2) <> bDatum(2) Then
        CompareRFC3339DateTime = IIf(CInt(aDatum(2)) > CInt(bDatum(2)), 0, 1)
    Else 'Going down to Time Difference date is the same
        aDate1 = Split(Mid(sDate1, InStr(sDate1, "T") + 1), ":")
        aDate2 = Split(Mid(sDate2, InStr(sDate2, "T") + 1), ":")
    
        If aDate1(0) > aDate2(0) Then
            CompareRFC3339DateTime = IIf(CInt(aDate1(0)) > CInt(aDate2(0)), 0, 1)
        ElseIf aDatum(1) > bDatum(1) Then
            CompareRFC3339DateTime = IIf(CInt(aDate1(1)) > CInt(aDate2(1)), 0, 1)
        Else
            CompareRFC3339DateTime = IIf(CInt(Left$(aDate1(2), 2)) > CInt(Left$(aDate2(2), 2)), 0, 1)
        End If
    
    End If
    
End Function

Private Function CheckNewRemoteItems(cn As ADODB.Connection)
    Dim oRSLocalItms As New ADODB.Recordset
    Dim oRSLocalHistory As New ADODB.Recordset
    Dim RS As New ADODB.Recordset
    Dim colNewItems As New Collection
    Dim colExistingItems As New Collection
    Dim sNewItemSQL As String
    Dim sExistingItemSQL As String
    Dim i As Integer
    Dim udtSXHis As sxHistory
    
    Set m_Col_History = New Collection
    
    'If no local item exists with the same id value as the incoming item, add the incoming item to the merge result collection; we are done processing the incoming item
    RS.open txtServer.Text & "?ID=SELECT sGUID FROM " & txtSynchTable(0).Text
        
    'Check if there are any returned records
    If Not RS.EOF And Not RS.Bof Then
        
        RS.MoveFirst
        
        Do While Not RS.EOF
            Set oRSLocalItms = New ADODB.Recordset
            oRSLocalItms.open "SELECT sGUID FROM " & txtSynchTable(0).Text & "WHERE sGUID = '" & RS.fields.Item(0).Value & "'", cn
            
            'Check if the item exists
            If Not oRSLocalItms.EOF And Not oRSLocalItms.Bof Then

                'm_frmDebug.DebugPrint CompareRFC3339DateTime(oRSLocalItms.fields(0).Value, .fields(0).Value)
                If Not sExistingItemSQL = "" Then
                    sExistingItemSQL = sExistingItemSQL & ",'" & RS.fields.Item(0).Value & "'"
                Else
                    sExistingItemSQL = "'" & RS.fields.Item(0).Value & "'"
                End If

                colExistingItems.Add RS.fields.Item(0).Value
            Else

                If Not sNewItemSQL = "" Then
                    sNewItemSQL = sNewItemSQL & ",'" & RS.fields.Item(0).Value & "'"
                Else
                    sNewItemSQL = "'" & RS.fields.Item(0).Value & "'"
                End If

                colNewItems.Add RS.fields.Item(0).Value
                'TODO import
            End If
            
            RS.MoveNext
            
        Loop

    End If
        
    If Not sNewItemSQL = "" Then
        Set RS = New ADODB.Recordset
        sNewItemSQL = "SELECT * FROM " & txtSynchTable(0).Text & "WHERE sGUID IN (" & sNewItemSQL & ")"
        RS.open txtServer.Text & "?ID=" & sNewItemSQL
        
        If Not RS.EOF And Not RS.Bof Then
        
            If Not RS.Bof Then RS.MoveFirst
            
            Set oRSLocalItms = New ADODB.Recordset
            oRSLocalItms.open "SELECT * FROM " & txtSynchTable(0).Text, cn, adOpenDynamic, adLockOptimistic
            oRSLocalHistory.open "SELECT * FROM " & txtSynchTable(1).Text, cn, adOpenDynamic, adLockOptimistic
            
            Do While Not RS.EOF

                With oRSLocalItms
                    .AddNew
                    .fields.Item("sGUID").Value = RS.fields.Item("sGUID").Value
                    .fields.Item("sLocalID").Value = RS.fields.Item("sLocalID").Value
                    .fields.Item("sBy").Value = RS.fields.Item("sBy").Value
                    .fields.Item("sDescription").Value = RS.fields.Item("sDescription").Value
                    .fields.Item("Title").Value = RS.fields.Item("Title").Value
                    .fields.Item("time").Value = RS.fields.Item("time").Value
                    .fields.Item("deleted").Value = RS.fields.Item("deleted").Value
                    .fields.Item("TableName").Value = RS.fields.Item("TableName").Value
                    .fields.Item("isGeoTable").Value = RS.fields.Item("isGeoTable").Value
                    .Update
                End With
                    
                With oRSLocalHistory
                    .AddNew
                    .fields.Item("sGUID").Value = RS.fields.Item("sGUID").Value
                    .fields.Item("sBy").Value = RS.fields.Item("sBy").Value
                    .fields.Item("sequence").Value = 1
                    .fields.Item("swhen").Value = RS.fields.Item("time").Value 'History When
                    .fields.Item("deleted").Value = IIf(IsNull(RS.fields.Item("deleted").Value), "false", RS.fields.Item("deleted").Value)
                    .fields.Item("noconflicts").Value = "true"
                    .Update
                End With

                RS.MoveNext
            Loop
        
        End If
    End If
        
    If Not sExistingItemSQL = "" Then
        
        Dim rsRemoteHistory As ADODB.Recordset
        Dim rsLocalHistory As ADODB.Recordset
    
        Set RS = New ADODB.Recordset
        sExistingItemSQL = "SELECT * FROM " & txtSynchTable(1).Text & " WHERE sGUID IN (" & sExistingItemSQL & ")"
        RS.open txtServer.Text & "?ID=" & sExistingItemSQL
        
        Set rsRemoteHistory = New ADODB.Recordset
        
        rsRemoteHistory.open txtServer.Text & "?ID=" & sExistingItemSQL
        'rsLocalHistory.Open sExistingItemSQL, cn, adOpenDynamic, adLockOptimistic
        
        For i = 0 To colExistingItems.Count - 1
            rsRemoteHistory.Filter = ""
            Set rsLocalHistory = New ADODB.Recordset
            rsLocalHistory.open "SELECT MAX(swhen) AS Latest, sGUID FROM " & txtSynchTable(1).Text & " WHERE sGUID ='" & colExistingItems.Item(i + 1) & "' GROUP BY sGUID", cn, adOpenDynamic, adLockOptimistic
            rsRemoteHistory.Filter = "sGUID ='" & colExistingItems.Item(i + 1) & "' AND swhen >= '" & rsLocalHistory.fields.Item("Latest").Value & "'"
        
            If rsRemoteHistory.EOF Or rsRemoteHistory.Bof Then
                'Local Item is newer or Equal
                
                    'Local is latest
                    'TODO Update remote
                    m_frmDebug.DebugPrint "Local is latest"
                'End If
                
            Else
                
                'Now final test, check If Equal date or if local is newer
                If rsRemoteHistory.fields.Item("swhen").Value <> rsLocalHistory.fields.Item("Latest").Value Then
                    m_frmDebug.DebugPrint "Remote Item is newer"
                Else
                'Remote Item is newer
                'TODO update local
                m_frmDebug.DebugPrint "Remote Item is Same"
                End If
            End If
        
        Next
        
        If Not RS.EOF And Not RS.Bof Then
        
            If Not RS.Bof Then RS.MoveFirst
            
            Set oRSLocalItms = New ADODB.Recordset
            oRSLocalItms.open "SELECT * FROM " & txtSynchTable(0).Text, cn, adOpenDynamic, adLockOptimistic
            
            With oRSLocalItms

                If Not .EOF And Not .Bof Then
                    
                    Do While Not RS.EOF

                        If Not .Bof Then .MoveFirst
                        .Find "sGUID = '" & RS.fields.Item("sGUID").Value & "'"
                        
                        If Not .EOF Then
                            .fields.Item("sLocalID").Value = RS.fields.Item("sLocalID").Value
                            .fields.Item("sBy").Value = RS.fields.Item("sBy").Value
                            .fields.Item("sDescription").Value = RS.fields.Item("sDescription").Value
                            .fields.Item("Title").Value = RS.fields.Item("Title").Value
                            .fields.Item("time").Value = RS.fields.Item("time").Value
                            .fields.Item("deleted").Value = RS.fields.Item("deleted").Value
                            .fields.Item("TableName").Value = RS.fields.Item("TableName").Value
                            .fields.Item("isGeoTable").Value = RS.fields.Item("isGeoTable").Value
                            .Update
                        End If
                        
                        RS.MoveNext
                    Loop

                End If

            End With
        
        End If
    End If
        
    '    'Create SQL to get the new items
    '    For i = 1 To colNewItems.Count
    '
    '        If Not sNewItemSQL = "" Then
    '            sNewItemSQL = sNewItemSQL & ",'" & colNewItems.Item(i) & "'"
    '        Else
    '            sNewItemSQL = "'" & colNewItems.Item(i) & "'"
    '        End If
    '
    '    Next i
    '
    '    sNewItemSQL = "SELECT * FROM " & txtSynchTable(0).Text & "WHERE sGUID IN (" & sNewItemSQL & ")"
    '
    m_frmDebug.DebugPrint sNewItemSQL
    
    'SELECT "column_name"
    'From "table_name"
    'WHERE "column_name" IN ('value1', 'value2', ...)
        
End Function

Private Function checkIEVersion() As String
    checkIEVersion = ReadVersion("c:\windows\system32\ieframe.dll")
End Function

'Private Sub cmdCreateSynch_Click()
'    Dim sTable As String
'    Dim oCn As ADODB.Connection
'    Dim oRSSynch As New ADODB.Recordset
'    Dim sSQL As String
'
'    Set oCn = New ADODB.Connection
'
'    oCn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtLocalDB(1).Text & ";Persist Security Info=False"
'
'    If Not DoTableExists("SynchTables", oCn) Then
'        Create_SynchTables oCn
'    End If
'
'    oRSSynch.open "SELECT * FROM SynchTables", oCn, adOpenDynamic, adLockBatchOptimistic
'
'    If Not oRSSynch.EOF And Not oRSSynch.Bof Then
'        oRSSynch.MoveFirst
'        oRSSynch.Find "sTableName = '" & ComTables.List(ComTables.ListIndex) & "'"
'
'        If Not oRSSynch.EOF Then
'            MsgBox "It seems like the table " & ComTables.List(ComTables.ListIndex) & " you have defined already exists in the Settings..."
'            Exit Sub
'        End If
'
'    End If
'
'    sSQL = "INSERT INTO SynchTables (sGUID, sTableName, sName, sDescription, OwnerID, AllowWrite, SynchFrequency, AutoUpdate) VALUES "
'
'    sSQL = sSQL & "('" & CreateGUID & "', '" & ComTables.List(ComTables.ListIndex) & "', '" & txtName.Text & "', '" & txtDescription.Text & "', 0, " & IIf(chkAllowSyncUpdate.Value = vbChecked, "True", "False") & ", 0, " & IIf(chkReadWrite.Value = vbChecked, "True", "False") & ")"
'
'    oCn.Execute sSQL
'
'    '    With oRSSynch
'    '
'    '        If .EOF And .BOF Then
'    '            .AddNew
'    '            With .fields
'    ''                .Item("sGUID").Value = CreateGUID
'    '                .Item("sTableName").Value = ComTables.List(ComTables.ListIndex)
'    '                .Item("sName").Value = ComTables.List(ComTables.ListIndex)
'    '                .Item("sDescription").Value = ComTables.List(ComTables.ListIndex)
'    '                .Item("OwnerID").Value = ComTables.List(ComTables.ListIndex)
'    '                .Item("AllowWrite").Value = False
'    '                .Item("SynchFrequency").Value = 5
'    '                .Item("AutoUpdate").Value = False
'    '            End With
'    '            .UpdateBatch adAffectCurrent
'    '        End If
'    '
'    '    End With
'
'End Sub

Private Sub cmdDb_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo cmdDb_Click_Err
        '</EhHeader>
        Dim cn As New ADODB.Connection
        Dim RS As New ADODB.Recordset
        Dim RsHistory As New ADODB.Recordset
        Dim oChannel As MSXML2.IXMLDOMElement
        Dim lSequence As Long
        Dim sGUID As String
        Dim sRFC3339DateTime As String
        Dim sSQL As String

100     'm_frmDebug.DebugPrint CompareRFC3339DateTime(RFC3339DateTime, "2008-10-03T11:03:03Z") 'RFC3339DateTime

        'm_frmDebug.DebugPrint CompareRFC3339DateTime("2008-10-03T11:03:03Z", "2008-10-03T11:03:04Z")
        'm_frmDebug.DebugPrint CompareRFC3339DateTime("2006-10-03T11:03:03Z", "2008-10-03T11:03:03Z")
        'm_frmDebug.DebugPrint CompareRFC3339DateTime("2008-10-03T12:03:03Z", "2008-10-03T11:03:03Z")

        Dim oRSLatestDate As New ADODB.Recordset

102     With RS
            
104         Set oRSLatestDate = New ADODB.Recordset
        
            'oRSLatestDate.Open "SELECT MAX([when]) FROM " & txtSynchTable(1).Text, cn
            '.Open txtServer.Text & "?ID=SELECT MAX([when]) FROM " & txtSynchTable(1).Text
            'm_frmDebug.DebugPrint CompareRFC3339DateTime(oRSLatestDate.fields(0).Value, .fields(0).Value)
        
106         sSQL = GetRSSSQL
        
108         If OptMode(0).Value = True Then
110             cn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtLocalDB(0).Text & ";Persist Security Info=False"
112             .open "select * From " & txtSynchTable(0).Text, cn, adOpenDynamic, adLockOptimistic
            Else
114             Set m_RSServerFeed = New ADODB.Recordset
116             .open txtServer.Text & "?ID=select * From " & txtSynchTable(0).Text '" & sSQL
            End If
    
118         Select Case Index
    
                Case 0
120                 ComItemID.Clear

122                 Do While Not .EOF
                        
124                     With .fields
126                         ComItemID.AddItem .Item("sGUID").Value
128                         txtXMLOut.Text = txtXMLOut.Text & vbCrLf & CreateNewRSSSyncFeed(txtFeedInfo(0).Text, .Item("sGUID").Value, .Item("Title").Value, .Item("sDescription").Value, .Item("sBy").Value, IIf(chkSave.Value = vbChecked, True, False))
                        
                        End With

130                     .MoveNext
                    Loop

132                 If ComItemID.ListCount > 0 Then ComItemID.ListIndex = 0

                    'kr
134             Case 1
                
136                 ComItemID.Clear
                    
138                 If Not .EOF And Not .Bof Then
140                     Set oChannel = GetNewRSSSyncFeed(txtFeedInfo(0).Text, .fields.Item("sGUID").Value, .fields.Item("Title").Value, .fields.Item("sDescription").Value, .fields.Item("sBy").Value, .fields.Item("time").Value, .fields.Item("TableName").Value, .fields.Item("isGeoTable").Value)
142                     ComItemID.AddItem .fields.Item("sGUID").Value
144                     .MoveNext
                    
146                     Do While Not .EOF

148                         With .fields
150                             ComItemID.AddItem .Item("sGUID").Value
152                             InsertNewRSSSyncFeedItem .Item("sGUID").Value, .Item("Title").Value, .Item("sDescription").Value, .Item("sBy").Value, .Item("time").Value, .Item("TableName").Value, .Item("isGeoTable").Value, oChannel, True, cn
                            End With

154                         .MoveNext
                        Loop

156                     ComItemID.ListIndex = 0
                
                        On Error Resume Next
                    
158                     Kill "c:\OASIS\Client\grr.xml"
                    
160                     Open "c:\OASIS\Client\grr.xml" For Output As #1
162                     Print #1, "<rss version=""2.0"" xmlns:sx=""http://feedsync.org/2007/feedsync"">" & oChannel.xml & "</rss>"
164                     Close 1
                    Else
                        On Error Resume Next
                    
166                     Kill "c:\OASIS\Client\grr.xml"
                    
168                     Open "c:\OASIS\Client\grr.xml" For Output As #1
170                     Print #1, "<rss version=""2.0"" xmlns:sx=""http://feedsync.org/2007/feedsync""><channel></channel></rss>"
172                     Close 1
                    End If
                
174                 WebBrowser1.Navigate2 "file:\\c:\OASIS\Client\grr.xml"
                    
176                 DoEvents
178                 WebBrowser1.Refresh
                    
                    'txtXMLOut.Text = oChannel.Xml
                    
180             Case 2

182                 If Not .EOF And Not .Bof Then
184                     RS.MoveFirst
186                     RS.Find "sGUID = '" & ComItemID.List(ComItemID.ListIndex) & "'"
                
                        'If (Not .EOF And Not .BOF) Then
                     
188                     UpdateRssFeedSync "c:\OASIS\Client\grr.xml", ComItemID.List(ComItemID.ListIndex), txtFeedInfo(2).Text, txtFeedInfo(3).Text, txtFeedInfo(4).Text, lSequence, False, True
                     
190                     .fields.Item("Title").Value = txtFeedInfo(2).Text
192                     .fields.Item("sLocalID").Value = ""
194                     .fields.Item("sDescription").Value = txtFeedInfo(3).Text
196                     .fields.Item("sBy").Value = txtFeedInfo(4).Text
198                     .fields.Item("time").Value = RFC3339DateTime
200                     .fields.Item("TableName").Value = txtFeedInfo(4).Text
202                     .fields.Item("isGeoTable").Value = IIf(chkISGeo.Value = vbChecked, True, False)
204                     .Update
                    
206                     RsHistory.open "SELECT * FROM " & txtSynchTable(1).Text & " WHERE deleted <> 'true'", cn, adOpenDynamic, adLockOptimistic
                    
208                     RsHistory.AddNew
210                     RsHistory.fields.Item("sGUID").Value = ComItemID.List(ComItemID.ListIndex)
212                     RsHistory.fields.Item("swhen").Value = RFC3339DateTime
214                     RsHistory.fields.Item("sequence").Value = GetHistorySequence(ComItemID.List(ComItemID.ListIndex), cn)
216                     RsHistory.fields.Item("sBy").Value = txtFeedInfo(4).Text
218                     RsHistory.fields.Item("deleted").Value = "false"
                        
                        'RSHistory.fields.Item("isGeoTable").Value = IIf(chkISGeo.Value = vbChecked, True, False)
220                     RsHistory.Update
                    
222                     WebBrowser1.Refresh
                    End If

                    'End If
224             Case 3 'Insert Local Item
                    
226                 sRFC3339DateTime = RFC3339DateTime
228                 sGUID = CreateGUID
                            
                    'rs.MoveFirst
                    
230                 Set oChannel = LoadRootElement("c:\OASIS\Client\grr.xml")
                    
232                 InsertNewRSSSyncFeedItem sGUID, txtFeedInfo(2).Text, txtFeedInfo(3).Text, txtFeedInfo(4).Text, sRFC3339DateTime, txtFeedInfo(5).Text, IIf(chkISGeo.Value = vbChecked, True, False), oChannel
                                          
234                 RS.AddNew
236                 .fields.Item("Title").Value = txtFeedInfo(2).Text
238                 .fields.Item("sLocalID").Value = ""
240                 .fields.Item("sDescription").Value = txtFeedInfo(3).Text
242                 .fields.Item("sBy").Value = txtFeedInfo(4).Text
244                 .fields.Item("time").Value = RFC3339DateTime
246                 .fields.Item("TableName").Value = txtFeedInfo(4).Text
248                 .fields.Item("isGeoTable").Value = IIf(chkISGeo.Value = vbChecked, True, False)
250                 .Update
                    
252                 RsHistory.open "SELECT * FROM " & txtSynchTable(1).Text & " WHERE deleted <> 'true'", cn, adOpenDynamic, adLockOptimistic
                    
254                 RsHistory.AddNew
256                 RsHistory.fields.Item("sGUID").Value = .fields.Item("sGUID").Value
258                 RsHistory.fields.Item("swhen").Value = RFC3339DateTime
260                 RsHistory.fields.Item("sequence").Value = 1
262                 RsHistory.fields.Item("sBy").Value = txtFeedInfo(4).Text
                    ' RSHistory.fields.Item("time").Value = RFC3339DateTime
                        
                    'RSHistory.fields.Item("isGeoTable").Value = IIf(chkISGeo.Value = vbChecked, True, False)
264                 RsHistory.Update
266                 cmdDb_Click 1

268             Case 4 'Delete Item
                   
270                 If ComItemID.List(ComItemID.ListIndex) = "" Then Exit Sub
                    Dim sDelGUID As String
                
272                 cn.Execute "DELETE * FROM " & txtSynchTable(0).Text & "WHERE sGUID = '" & ComItemID.List(ComItemID.ListIndex) & "'"
                                            
274                 RsHistory.open "SELECT * FROM " & txtSynchTable(1).Text & " WHERE deleted <> 'true'", cn, adOpenDynamic, adLockOptimistic
                    
276                 If Not RsHistory.Bof Then RsHistory.MoveFirst
                    
                    'RsHistory.Filter = "sGUID = '" & ComItemID.List(ComItemID.ListIndex) & "'"
                    
278                 RsHistory.AddNew
280                 RsHistory.fields.Item("sGUID").Value = ComItemID.List(ComItemID.ListIndex)
282                 RsHistory.fields.Item("swhen").Value = RFC3339DateTime
284                 RsHistory.fields.Item("sequence").Value = GetHistorySequence(ComItemID.List(ComItemID.ListIndex), cn)
286                 RsHistory.fields.Item("sBy").Value = txtFeedInfo(4).Text
288                 RsHistory.fields.Item("deleted").Value = "true"
290                 RsHistory.Update
                    
292                 cmdDb_Click 1
                    
            End Select

        End With
    
        '<EhFooter>
        Exit Sub

cmdDb_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Synch_Explorer.frmSynchReader.cmdDb_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub ClearAll()
 Dim i As Integer
 
 On Error Resume Next
 
 For i = txtFeedInfo.LBound To txtFeedInfo.UBound
    txtFeedInfo(i).Text = ""
 Next
 
 txtLocalDB(0).Text = ""
 txtLocalDB(1).Text = ""
 
 txtServer.Text = ""
 
 txtSynchTable(0).Text = ""
 txtSynchTable(1).Text = ""
 txtXMLOut.Text = ""
 
End Sub

Private Function LoadRootElement(sSourcePath As String) As MSXML2.IXMLDOMElement

    Dim oXMLDoc As New MSXML2.DOMDocument
    Dim oXMLElement As MSXML2.IXMLDOMElement
    Dim oChannel As MSXML2.IXMLDOMElement
    
    oXMLDoc.async = False

    If oXMLDoc.Load(sSourcePath) Then
        '//  Get "rss" element
        Set oXMLElement = oXMLDoc.documentElement
    
        '//  Get "channel" element
        Set oChannel = oXMLElement.selectSingleNode("channel")
    
        '//  Validate that "channel" element exists
        If oChannel Is Nothing Then
            MsgBox "Unable to get channel"
            Exit Function
        End If
    
    End If
    
    Set LoadRootElement = oChannel
    
End Function

Private Sub cmdFeeds_Click(Index As Integer)

    Select Case Index
    
        Case 0
            txtXMLOut.Text = CreateRSSSyncFeed(txtFeedInfo(0).Text, txtFeedInfo(1).Text, txtFeedInfo(2).Text, txtFeedInfo(3).Text, txtFeedInfo(4).Text, IIf(chkSave.Value = vbChecked, True, False))
            'kr

        Case 1
            txtXMLOut.Text = UpdateRssFeedSync(txtFeedInfo(0).Text, txtFeedInfo(1).Text, txtFeedInfo(2).Text, txtFeedInfo(3).Text, txtFeedInfo(4).Text, CLng(0), IIf(chkSave.Value = vbChecked, True, False))
        
        Case 3
            txtXMLOut.Text = DeleteRssSyncFeed(txtFeedInfo(0).Text, txtFeedInfo(1).Text, txtFeedInfo(4).Text, IIf(chkSave.Value = vbChecked, True, False))
    
    End Select

End Sub

Private Sub cmdGetServer_Click()
    Dim oRs As New ADODB.Recordset
    oRs.open txtSynchURL.Text & "?ID=select * From SynchTables"


    If oRs.State <> adStateClosed Then
        If Not oRs.Bof And Not oRs.EOF Then oRs.MoveFirst
        ComServertables.Clear
        
        Do While Not oRs.EOF
            ComServertables.AddItem oRs.fields.Item("Name").Value
            oRs.MoveNext
        Loop
    End If
    

End Sub

Private Sub cmdGUIDGEN_Click()
    txtFeedInfo(1).Text = CreateGUID
End Sub

Private Sub cmdHrmmm_Click()
    Dim MsXmlHttp As New MSXML2.ServerXMLHTTP40
    Dim MsXmlDoc As New MSXML2.DOMDocument
    Dim oRSTest As New ADODB.Recordset
    Dim oCn As New ADODB.Connection
    
    oCn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtLocalDB(1).Text & ";Persist Security Info=False"

    

    oRSTest.open "SELECT * FROM 1sector", oCn
    
    MsXmlHttp.open "POST", "http://www.immap.org/OAS2/Oasis.asp?synchval=ashfjasfhfh&synchtable=1sector", 0
    oRSTest.Save MsXmlDoc, 1
    If frmDatabaseConnect.g_bProxyEnabled Then MsXmlHttp.setProxy 2, frmDatabaseConnect.g_sProxy
    MsXmlHttp.send MsXmlDoc
    
    m_frmDebug.DebugPrint MsXmlHttp.responseText
    
'    Dim fld As ADODB.Field
'
'    For Each fld In oRSTest
'        fld.Type
'        fld.DefinedSize
'        fld.Name
'    Next
    
End Sub

Private Sub CreateServerSynchTable(oCn As ADODB.Recordset, _
                                   sTableName As String)
    Dim oRs As New ADODB.Recordset

    oRs.CursorLocation = adUseClient
    oRs.open "SELECT * FROM " & sTableName, oCn, adOpenDynamic, adLockBatchOptimistic
    
    Set oRs.ActiveConnection = Nothing
    
    

End Sub

Private Sub cmdLoad_Click()
    Dim oCn As ADODB.Connection
    
    Set oCn = New ADODB.Connection
    
    oCn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtLocalDB(1).Text & ";Persist Security Info=False"
    
    ShowAllTables oCn, False, ComTables
     
    If ComTables.ListCount > 0 Then ComTables.ListIndex = 0
        
     
End Sub

Private Sub cmdNext_Click()

    If Not m_RSServerItems.EOF Then m_RSServerItems.MoveNext
    
End Sub

Private Sub ComItemID_Click()
        '<EhHeader>
        On Error GoTo ComItemID_Click_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset
        Dim cn As New ADODB.Connection
    
100     If Not Len(ComItemID.List(ComItemID.ListIndex)) = 0 Then
    
102         With RS
           
104             If OptMode(0).Value = True Then
106                 cn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & txtLocalDB(0).Text & ";Persist Security Info=False"
108                 .open "select * From " & txtSynchTable(0).Text & " WHERE sGUID = '" & ComItemID.List(ComItemID.ListIndex) & "'", cn, adOpenDynamic, adLockOptimistic
                Else
110                 .open txtServer.Text & "?ID=select * From " & txtSynchTable(0).Text & " WHERE sGUID = '" & ComItemID.List(ComItemID.ListIndex) & "'"
                End If
           
112             With .fields
           
114                 txtFeedInfo(1).Text = .Item("sGUID").Value
116                 txtFeedInfo(2).Text = .Item("Title").Value
118                 txtFeedInfo(3).Text = .Item("sDescription").Value
120                 txtFeedInfo(4).Text = .Item("sBy").Value
122                 txtFeedInfo(5).Text = .Item("TableName").Value
            
124                 If Not IsNull(.Item("TableName").Value) Then
126                     chkISGeo.Value = vbChecked
                    Else
128                     chkISGeo.Value = vbUnchecked
                    End If
            
                End With
           
            End With

        End If

        '<EhFooter>
        Exit Sub

ComItemID_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Synch_Explorer.frmSynchReader.ComItemID_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
    
100     Set oSynchs = New clsSynchronisations
    
        'LoadSettings
    
        Exit Sub
    
102     oSynchs.LoadFromCacheEx App.Path & "\XML_FOLDER\xml\" & "dude2.xml"
    
104     If Not oSynchs.LoadFromCache(App.Path & "\XML_FOLDER\xml\" & "rss_" & Format(Now, "yyyymmdd") & ".xml") Then

106         DownloadAllSynchs
        Else
108         Me.Caption = "Loaded from cache"
110         PopulateMenu
        End If
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Synch_Explorer.frmSynchReader.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub DownloadAllSynchs()
    
    If Not oSynchs.Busy Then
        oSynchs.LoadAllChannels
        PopulateMenu
        mdtLastDownload = Now
        Me.Caption = "Synch Reader,  Updated " & mdtLastDownload
    End If

End Sub

Private Sub PopulateMenu()
    
    Dim oChannels As Collection
    Dim oChannel As clsSynchChannel
    Dim oSynchItm As clsSynchItem
    Dim oTreeNode As MSComctlLib.Node
    Dim stitle As String
    
    Set oChannels = oSynchs.Channels

    If Not oChannels Is Nothing Then
        tvwMenu.Nodes.Clear

        For Each oChannel In oChannels
            Set oTreeNode = AddItemToTreeMenu(oChannel.Title, oChannel.Link, oChannel.Path)

            For Each oSynchItm In oChannel

                If Len(oSynchItm.Title) > 0 Then
                    stitle = oSynchItm.Title
                Else
                    stitle = Left$(oSynchItm.Description, 25)
                End If

                AddItemToTreeMenu stitle, oSynchItm.URL, oSynchItm.Path, oTreeNode
            Next oSynchItm

            tvwMenu.Nodes(1).Selected = True
        Next oChannel

    Else
        MsgBox "Error loading XML document." & vbCrLf & "The document is not in RSS format.", vbExclamation, Me.Caption
    End If
    
    oSynchs.Save App.Path & "\XML_FOLDER\" & "rss_" & Format(Now, "yyyymmdd") & ".xml"
    
    '   If (mlAutoRefreshMin > 0) And (Not tmrRefresh.Enabled) Then
    '       tmrRefresh.Enabled = True
    '   End If
    
End Sub

Private Function AddItemToTreeMenu(ByVal stitle As String, _
                                   ByVal sLink As String, _
                                   ByVal sPath As String, _
                                   Optional ByVal oParent As MSComctlLib.Node = Nothing) As MSComctlLib.Node
    
    Dim oNode As MSComctlLib.Node
    
    If oParent Is Nothing Then
        Set oNode = tvwMenu.Nodes.Add(Text:=stitle)
        oNode.Expanded = True
        oNode.Bold = True
    Else
        On Error Resume Next
        Set oNode = tvwMenu.Nodes.Add(Key:=sLink, Text:=stitle, relative:=oParent, relationship:=tvwChild)
    End If

    oNode.Key = sPath
    oNode.Tag = sPath
    Set AddItemToTreeMenu = oNode
    
End Function

'//  *********************************************************************************
'//  BIG HONKING NOTE:  The "by" value should be a unique value per user/endpoint -
'//                     this sample uses a random number to generate uniqueness.
'//                     Other applications should considering using a more robust
'//                     and persistant value.

Private Function DeleteRssSyncFeed(sSourcePath As String, _
                                   sGUID As String, _
                                   sBy As String, _
                                   Optional bSave As Boolean) As String
    Dim oXMLDoc As New MSXML2.DOMDocument
    Dim oXMLElement As MSXML2.IXMLDOMElement
    Dim oChannel As MSXML2.IXMLDOMElement
    Dim oRssItem As MSXML2.IXMLDOMElement
    Dim oSynchElement As MSXML2.IXMLDOMElement
    Dim oHisElements As MSXML2.IXMLDOMElement
    Dim sRFC3339DateTime As String
    Dim lngAttr As Long
    
    oXMLDoc.async = False

    If oXMLDoc.Load(sSourcePath) Then
        '//  Get "rss" element
        Set oXMLElement = oXMLDoc.documentElement

        '//  Check if FeedSync namespace exists
        'If oXMLElement.getAttribute("xmlns:sx") Is Nothing Then
        '    MsgBox "Invalid Feed Sync File"
        '    Exit Function
        'End If
    
        '//  Get "channel" element
        Set oChannel = oXMLElement.selectSingleNode("channel")
    
        '//  Validate that "channel" element exists
        If oChannel Is Nothing Then
            MsgBox "Unable to get channel"
            Exit Function
        End If

        '//  Get "sx:sync" element with matching id
        Set oSynchElement = oChannel.selectSingleNode("//item/sx:sync[@id='" + sGUID + "']")

        '//  Validate that "sx:sync" element exists
        If oSynchElement Is Nothing Then
            MsgBox "Unable to find 'sx:sync' element with id='" + sGUID + "'"
            Exit Function
        End If
        
        '//  Get "updates" attribute from "sx:sync" element
        lngAttr = oSynchElement.getAttribute("updates")

        '//  Validate "updates" attribute
        'If lngAttr Is Nothing Then
        '    MsgBox "Unable to get 'updates' attribute from 'sx:sync' element with id='" + sGUID + "'"
        '    Exit Function
        'End If
        
        sRFC3339DateTime = RFC3339DateTime
        
        '//  Increment "updates" attribute
        lngAttr = lngAttr + 1
        
        '//  Create "sx:history" element
        Set oHisElements = oXMLDoc.createElement("sx:history")
        
        '//  Set "sequence" attribute for "sx:history" element
        oHisElements.setAttribute "sequence", lngAttr
        
        '//  Set "when" attribute for "sx:history" element
        oHisElements.setAttribute "when", sRFC3339DateTime
        
        '//  Set "by" attribute for "sx:history" element
        oHisElements.setAttribute "by", sBy
        
        '//  Insert "sx:history" element as topmost sub-element of "sx:sync" element
        oSynchElement.insertBefore oHisElements, oSynchElement.childNodes(0)
        
        '//  Set "updates" attribute for "sx:sync" element
        oSynchElement.setAttribute "updates", lngAttr
        
        '//  Set "deleted" attribute for "sx:sync" element
        oSynchElement.setAttribute "deleted", "true"
        
        If bSave Then
            oXMLDoc.Save sSourcePath
        End If
        
        DeleteRssSyncFeed = oXMLDoc.xml
    End If

End Function

'
'//  *********************************************************************************
'//  BIG HONKING NOTE:  The "by" value should be a unique value per user/endpoint -
'//                     this sample uses a random number to generate uniqueness.
'//                     Other applications should considering using a more robust
'//                     and persistant value.

Private Function CreateRSSSyncFeed(sSourcePath As String, _
                                   sNewGUID As String, _
                                   stitle As String, _
                                   sDescription As String, _
                                   sBy As String, _
                                   Optional bSave As Boolean = False) As String

    Dim oXMLDoc As New MSXML2.DOMDocument
    Dim oXMLElement As MSXML2.IXMLDOMElement
    Dim oChannel As MSXML2.IXMLDOMElement
    Dim oRssItem As MSXML2.IXMLDOMElement
    Dim oSynchElement As MSXML2.IXMLDOMElement
    Dim oHisElements As MSXML2.IXMLDOMElement
    Dim oTitleElements As MSXML2.IXMLDOMElement
    Dim oDescElements As MSXML2.IXMLDOMElement
    Dim sRFC3339DateTime As String
    Dim obj As Object
    
    oXMLDoc.async = False

    If oXMLDoc.Load(sSourcePath) Then
        '//  Get "rss" element
        Set oXMLElement = oXMLDoc.documentElement

        m_frmDebug.DebugPrint oXMLDoc.documentElement.baseName
        
        '//  Check if FeedSync namespace exists
        'Set obj = oXMLElement.getAttribute("xmlns:sx")
        
        ' If oXMLElement.getAttribute("xmlns:sx") Is Nothing Then
        '     MsgBox "Invalid Feed Sync File"
        '     Exit Function
        ' End If
    
        '//  Get "channel" element
        Set oChannel = oXMLElement.selectSingleNode("channel")
    
        '//  Validate that "channel" element exists
        If oChannel Is Nothing Then
            MsgBox "Unable to get channel"
            Exit Function
        End If
    
        '//  Create "item" element
        Set oRssItem = oXMLDoc.createElement("item")
    
        '//  Create "sx:sync" element
        Set oSynchElement = oXMLDoc.createElement("sx:sync")
    
        '//  Set "id" attribute for "sx:sync" element
        oSynchElement.setAttribute "id", sNewGUID
    
        '//  Set "updates" attribute for "sx:sync" element
        oSynchElement.setAttribute "updates", "1"
    
        '//  Set "deleted" attribute for "sx:sync" element
        oSynchElement.setAttribute "deleted", "false"
    
        '//  Set "noconflicts" attribute for "sx:sync" element
        oSynchElement.setAttribute "noconflicts", "false"
    
        '//  Create "history" element
        Set oHisElements = oXMLDoc.createElement("sx:history")
    
        '//  Get the current timedate and format it as RFC 3339
        sRFC3339DateTime = RFC3339DateTime
    
        '//  Set "sequence" attribute for "sx:history" element
        oHisElements.setAttribute "sequence", "1"
        
        '//  Set "when" attribute for "sx:history" element
        oHisElements.setAttribute "when", sRFC3339DateTime
        
        '//  Set "by" attribute for "sx:history" element
        oHisElements.setAttribute "by", sBy
        
        '//  Append "sx:history" element to "sx:sync" element
        oSynchElement.appendChild oHisElements
        
        '//  Append "sx:sync" element to "item" element
        oRssItem.appendChild oSynchElement
        
        '//  Create & populate "title" element
        Set oTitleElements = oXMLDoc.createElement("title")
        oTitleElements.Text = stitle
        
        '//  Append "title" element to "item" element
        oRssItem.appendChild oTitleElements
        
        '//  Create & populate "description" element
        Set oDescElements = oXMLDoc.createElement("description")
        oDescElements.Text = sDescription
        
        '//  Append "description" element to "item" element
        oRssItem.appendChild oDescElements
        
        '//  Append "item" element to "channel" element
        oChannel.appendChild oRssItem
        
        If bSave Then oXMLDoc.Save sSourcePath
        
        CreateRSSSyncFeed = oXMLDoc.xml
    
    End If

End Function

Private Function CreateNewRSSSyncFeed(sSourcePath As String, _
                                      sNewGUID As String, _
                                      stitle As String, _
                                      sDescription As String, _
                                      sBy As String, _
                                      Optional bSave As Boolean = False) As String

    Dim oXMLDoc As New MSXML2.DOMDocument
    Dim oXMLElement As MSXML2.IXMLDOMElement
    Dim oChannel As MSXML2.IXMLDOMElement
    Dim oRssItem As MSXML2.IXMLDOMElement
    Dim oSynchElement As MSXML2.IXMLDOMElement
    Dim oHisElements As MSXML2.IXMLDOMElement
    Dim oTitleElements As MSXML2.IXMLDOMElement
    Dim oDescElements As MSXML2.IXMLDOMElement
    Dim sRFC3339DateTime As String
    Dim obj As Object
    
    'Dim o as MSXML2.
    
    oXMLDoc.async = False

    Set oChannel = oXMLDoc.createElement("channel")

    '//  Get "rss" element
    'Set oXMLElement = oXMLDoc.documentElement

    'm_frmDebug.DebugPrint oXMLDoc.documentElement.baseName
        
    '//  Check if FeedSync namespace exists
    'Set obj = oXMLElement.getAttribute("xmlns:sx")
        
    ' If oXMLElement.getAttribute("xmlns:sx") Is Nothing Then
    '     MsgBox "Invalid Feed Sync File"
    '     Exit Function
    ' End If
    
    '//  Get "channel" element
    '    Set oChannel = oXMLElement.selectSingleNode("channel")
    
    '//  Validate that "channel" element exists
    If oChannel Is Nothing Then
        MsgBox "Unable to get channel"
        Exit Function
    End If
    
    '//  Create "item" element
    Set oRssItem = oXMLDoc.createElement("item")
    
    '//  Create "sx:sync" element
    Set oSynchElement = oXMLDoc.createElement("sx:sync")
    
    '//  Set "id" attribute for "sx:sync" element
    oSynchElement.setAttribute "id", sNewGUID
    
    '//  Set "updates" attribute for "sx:sync" element
    oSynchElement.setAttribute "updates", "1"
    
    '//  Set "deleted" attribute for "sx:sync" element
    oSynchElement.setAttribute "deleted", "false"
    
    '//  Set "noconflicts" attribute for "sx:sync" element
    oSynchElement.setAttribute "noconflicts", "false"
    
    '//  Create "history" element
    Set oHisElements = oXMLDoc.createElement("sx:history")
    
    '//  Get the current timedate and format it as RFC 3339
    sRFC3339DateTime = RFC3339DateTime
    
    '//  Set "sequence" attribute for "sx:history" element
    oHisElements.setAttribute "sequence", "1"
        
    '//  Set "when" attribute for "sx:history" element
    oHisElements.setAttribute "when", sRFC3339DateTime
        
    '//  Set "by" attribute for "sx:history" element
    oHisElements.setAttribute "by", sBy
        
    '//  Append "sx:history" element to "sx:sync" element
    oSynchElement.appendChild oHisElements
        
    '//  Append "sx:sync" element to "item" element
    oRssItem.appendChild oSynchElement
        
    '//  Create & populate "title" element
    Set oTitleElements = oXMLDoc.createElement("title")
    oTitleElements.Text = stitle
        
    '//  Append "title" element to "item" element
    oRssItem.appendChild oTitleElements
        
    '//  Create & populate "description" element
    Set oDescElements = oXMLDoc.createElement("description")
    oDescElements.Text = sDescription
        
    '//  Append "description" element to "item" element
    oRssItem.appendChild oDescElements
        
    '//  Append "item" element to "channel" element
    oChannel.appendChild oRssItem
        
    oXMLDoc.loadXML "<rss version=""""2.0"""" xmlns:sx=""""http://feedsync.org/2007/feedsync"""">" & oChannel.xml & "</rss>"
        
    If bSave Then oXMLDoc.Save sSourcePath
    'oChannel.Normalize
    CreateNewRSSSyncFeed = "<rss version=""""2.0"""" xmlns:sx=""""http://feedsync.org/2007/feedsync"""">" & oChannel.xml & "</rss>"
    'oXMLDoc.Xml

End Function

Private Function GetNewRSSSyncFeed(sSourcePath As String, _
                                   sNewGUID As String, _
                                   stitle As String, _
                                   sDescription As String, _
                                   sBy As String, _
                                   sRFC3339DateTime As String, _
                                   sTableName As String, _
                                   bIsGeoLayer As Boolean) As MSXML2.IXMLDOMElement

    Dim oXMLElement As MSXML2.IXMLDOMElement
    Dim oChannel As MSXML2.IXMLDOMElement
    Dim oRssItem As MSXML2.IXMLDOMElement
    Dim oSynchElement As MSXML2.IXMLDOMElement
    Dim oHisElements As MSXML2.IXMLDOMElement
    Dim oTitleElements As MSXML2.IXMLDOMElement
    Dim oDescElements As MSXML2.IXMLDOMElement
    Dim obj As Object
    Dim oXMLDoc As New MSXML2.DOMDocument

    Set oChannel = oXMLDoc.createElement("channel")
        
    '//  Create "item" element
    Set oRssItem = oXMLDoc.createElement("item")
    
    '//  Create "sx:sync" element
    Set oSynchElement = oXMLDoc.createElement("sx:sync")
    
    '//  Set "id" attribute for "sx:sync" element
    oSynchElement.setAttribute "id", sNewGUID
    
    '//  Set "updates" attribute for "sx:sync" element
    oSynchElement.setAttribute "updates", "1"
    
    '//  Set "deleted" attribute for "sx:sync" element
    oSynchElement.setAttribute "deleted", "false"
    
    '//  Set "noconflicts" attribute for "sx:sync" element
    oSynchElement.setAttribute "noconflicts", "false"
    
    '//  Create "history" element
    Set oHisElements = oXMLDoc.createElement("sx:history")
    
    '//  Get the current timedate and format it as RFC 3339
    'sRFC3339DateTime = RFC3339DateTime
    
    '//  Set "sequence" attribute for "sx:history" element
    oHisElements.setAttribute "sequence", "1"
        
    '//  Set "when" attribute for "sx:history" element
    oHisElements.setAttribute "when", sRFC3339DateTime
        
    '//  Set "by" attribute for "sx:history" element
    oHisElements.setAttribute "by", sBy
        
    '//  Append "sx:history" element to "sx:sync" element
    oSynchElement.appendChild oHisElements
        
    '//  Append "sx:sync" element to "item" element
    oRssItem.appendChild oSynchElement
        
    '//  Create & populate "title" element
    Set oTitleElements = oXMLDoc.createElement("title")
    oTitleElements.Text = stitle
        
    '//  Append "title" element to "item" element
    oRssItem.appendChild oTitleElements
        
    '//  Create & populate "description" element
    Set oDescElements = oXMLDoc.createElement("description")
    oDescElements.Text = sDescription
        
    '//  Append "description" element to "item" element
    oRssItem.appendChild oDescElements
        
    '//  Create & populate "Table" element, Reuse the Description Element
    Set oDescElements = oXMLDoc.createElement("TableName")
    oDescElements.Text = sTableName
        
    oDescElements.setAttribute "isGeoTable", LCase$(CStr(bIsGeoLayer))
        
    '//  Append "tableName" element to "item" element
    oRssItem.appendChild oDescElements
                
    '//  Append "item" element to "channel" element
    oChannel.appendChild oRssItem
        
    Set GetNewRSSSyncFeed = oChannel

End Function

Private Sub InsertNewRSSSyncFeedItem(sNewGUID As String, _
                                     stitle As String, _
                                     sDescription As String, _
                                     sBy As String, _
                                     sRFC3339DateTime As String, _
                                     sTableName As String, _
                                     bIsGeoLayer As Boolean, _
                                     oChannel As MSXML2.IXMLDOMElement, _
                                     Optional bCheckHistory As Boolean = False, _
                                     Optional cn As ADODB.Connection)
    
    Dim oXMLDoc As New MSXML2.DOMDocument
    Dim oXMLElement As MSXML2.IXMLDOMElement
    Dim oRssItem As MSXML2.IXMLDOMElement
    Dim oSynchElement As MSXML2.IXMLDOMElement
    Dim oHisElements As MSXML2.IXMLDOMElement
    Dim oTitleElements As MSXML2.IXMLDOMElement
    Dim oDescElements As MSXML2.IXMLDOMElement
    Dim obj As Object
        
    '//  Create "item" element
    Set oRssItem = oXMLDoc.createElement("item")
    
    '//  Create "sx:sync" element
    Set oSynchElement = oXMLDoc.createElement("sx:sync")
    
    '//  Set "id" attribute for "sx:sync" element
    oSynchElement.setAttribute "id", sNewGUID
    
    '//  Set "updates" attribute for "sx:sync" element
    oSynchElement.setAttribute "updates", "1"
    
    '//  Set "deleted" attribute for "sx:sync" element
    oSynchElement.setAttribute "deleted", "false"
    
    '//  Set "noconflicts" attribute for "sx:sync" element
    oSynchElement.setAttribute "noconflicts", "true"
    
    If bCheckHistory Then
        bCheckHistory = AddHistoryElements(oSynchElement, sNewGUID, cn)
    End If
    
    If Not bCheckHistory Then
    
        '//  Create "history" element
        Set oHisElements = oXMLDoc.createElement("sx:history")
    
        '//  Get the current timedate and format it as RFC 3339
        'sRFC3339DateTime = RFC3339DateTime
    
        '//  Set "sequence" attribute for "sx:history" element
        oHisElements.setAttribute "sequence", "1"
        
        '//  Set "when" attribute for "sx:history" element
        oHisElements.setAttribute "when", sRFC3339DateTime
        
        '//  Set "by" attribute for "sx:history" element
        oHisElements.setAttribute "by", sBy
        
        '//  Append "sx:history" element to "sx:sync" element
        oSynchElement.appendChild oHisElements
        
    End If
        
    '//  Append "sx:sync" element to "item" element
    oRssItem.appendChild oSynchElement
        
    '//  Create & populate "title" element
    Set oTitleElements = oXMLDoc.createElement("title")
    oTitleElements.Text = stitle
        
    '//  Append "title" element to "item" element
    oRssItem.appendChild oTitleElements
        
    '//  Create & populate "description" element
    Set oDescElements = oXMLDoc.createElement("description")
    oDescElements.Text = sDescription
        
    '//  Append "description" element to "item" element
    oRssItem.appendChild oDescElements
        
    '//  Create & populate "Table" element, Reuse the Description Element
    Set oDescElements = oXMLDoc.createElement("TableName")
    oDescElements.Text = sTableName
    oDescElements.setAttribute "isGeoTable", LCase$(CStr(bIsGeoLayer))
    '//  Append "tableName" element to "item" element
    oRssItem.appendChild oDescElements
        
    '//  Append "item" element to "channel" element
    oChannel.appendChild oRssItem

End Sub

Private Function UpdateRssFeedSync(sSourcePath As String, _
                                   sGUID As String, _
                                   stitle As String, _
                                   sDescription As String, _
                                   sBy As String, _
                                   lngSequence As Long, _
                                   bResolveConflicts As Boolean, _
                                   Optional bSave As Boolean = False) As String

    Dim oXMLDoc As New MSXML2.DOMDocument
    Dim oXMLElement As MSXML2.IXMLDOMElement
    Dim oChannel As MSXML2.IXMLDOMElement
    Dim oRssItem As MSXML2.IXMLDOMElement
    Dim oSynchElement As MSXML2.IXMLDOMElement
    Dim oHisElements As MSXML2.IXMLDOMElement
    Dim oTitleElements As MSXML2.IXMLDOMElement
    Dim oDescElements As MSXML2.IXMLDOMElement
    Dim lngAttrUpdt As Long
    
    Dim sRFC3339DateTime As String
    Dim oConflictElement As MSXML2.IXMLDOMElement
    Dim oConflictItm As MSXML2.IXMLDOMElement
    Dim oConflict As MSXML2.IXMLDOMElement
    Dim oConflictSynchElm As MSXML2.IXMLDOMElement
    Dim oConflictHisElm As MSXML2.IXMLDOMElement
    Dim oHisElm As MSXML2.IXMLDOMElement
    Dim oClonedHis As MSXML2.IXMLDOMElement
    Dim bNoConflicts As Boolean
    Dim bConflictHistoryRepresented As Boolean
    
    oXMLDoc.async = False

    If oXMLDoc.Load(sSourcePath) Then
        '//  Get "rss" element
        Set oXMLElement = oXMLDoc.documentElement

        '//  Check if FeedSync namespace exists
        'If oXMLElement.getAttribute("xmlns:sx") Is Nothing Then
        '    MsgBox "Invalid Feed Sync File"
        '    Exit Function
        'End If
    
        '//  Get "channel" element
        Set oChannel = oXMLElement.selectSingleNode("channel")
    
        '//  Validate that "channel" element exists
        If oChannel Is Nothing Then
            MsgBox "Unable to get channel"
            Exit Function
        End If

        '//  Get "sx:sync" element with matching id
        Set oSynchElement = oChannel.selectSingleNode("//item/sx:sync[@id='" + sGUID + "']")

        '//  Validate that "sx:sync" element exists
        If oSynchElement Is Nothing Then
            MsgBox "Unable to find 'sx:sync' element with id='" + sGUID + "'"
            Exit Function
        End If

        '//  Get "updates" attribute from "sx:sync" element
        lngAttrUpdt = oSynchElement.getAttribute("updates")

        '//  Validate "updates" attribute
        ' If lngAttrUpdt Is Nothing Then
        '     MsgBox "Unable to get 'updates' attribute from 'sx:sync' element with id='" + sGUID + "'"
        '     Exit Function
        ' End If

        '//  Get corresponding "item" element
        Set oRssItem = oSynchElement.parentNode

        '//  Validate that "item" element exists
        If oRssItem Is Nothing Then
            MsgBox "Unable to get 'item' element for 'sx:sync' element with id='" + sGUID + "'"
            Exit Function
        End If

        '//  Get "title" element from "item" element
        Set oTitleElements = oRssItem.selectSingleNode("title")

        '//  Validate that "title" element exists
        If oTitleElements Is Nothing Then
            '    //  Create "title" element
            Set oTitleElements = oXMLDoc.createElement("title")

            '    //  Append "title" element to "item" element
            oRssItem.appendChild oTitleElements
    
        End If
        
        '//  Set title for "title" element
        oTitleElements.Text = stitle
    
        '//  Get "description" element from "item" element
        Set oDescElements = oRssItem.selectSingleNode("description")
    
        '//  Validate that "description" element exists
        If oDescElements Is Nothing Then
            '    //  Create & populate "description" element
            oDescElements = oXMLDoc.createElement("description")
    
            '    //  Append "description" element to "item" element
            oRssItem.appendChild oDescElements
    
        End If
    
        '//  Set description for "description" element
        oDescElements.Text = sDescription

        sRFC3339DateTime = RFC3339DateTime

        '//  Increment "updates" attribute
        lngAttrUpdt = lngAttrUpdt + 1
        '
        '//  Create "sx:history" element
        Set oHisElements = oXMLDoc.createElement("sx:history")
    
        '//  Set "sequence" attribute for "sx:history" element
        lngSequence = lngAttrUpdt

        If sBy <> "" Then
            Dim oHisObj As MSXML2.IXMLDOMElement
            Dim sHisBy As String
            Dim lngHisSeq As Long
            
            '    //  Iterate "sx:history" sub-elements of main item's "sx:sync" element
            For Each oHisObj In oSynchElement.selectNodes("sx:history")

                '
                '        //  Get "by" attribute from main item's "sx:history"
                '        //  sub-element
                sHisBy = oHisObj.getAttribute("by")
                '
                '        //  Get "sequence" attribute from main item's
                '        //  "sx:history" sub-element
                lngHisSeq = oHisObj.getAttribute("sequence")
                '
                '        //  If "by" attributes don't match, continue loop
                
                If sHisBy = sBy And (lngHisSeq >= lngSequence) Then
                    lngSequence = lngHisSeq + 1
                End If
                
            Next oHisObj

            '
            '
        End If
    
        oHisElements.setAttribute "sequence", lngSequence
    
        '//  Set "when" attribute for "sx:history" element
        oHisElements.setAttribute "when", sRFC3339DateTime
    
        '//  Set "by" attribute for "sx:history" element
        oHisElements.setAttribute "by", sBy
    
        '//  Insert "sx:history" element as topmost sub-element of "sx:sync" element
        oSynchElement.insertBefore oHisElements, oSynchElement.childNodes(0)
    
        '//  Set "updates" attribute for "sx:sync" element
        oSynchElement.setAttribute "updates", lngAttrUpdt
    
        '
        '//  Get "noconflicts" attribute for "sx:sync" element
        bNoConflicts = CBool(oSynchElement.getAttribute("noconflicts"))
    
        '//  See if conflict resolution should be performed
        If Not bNoConflicts And bResolveConflicts Then
            '    {
            '    //  *********************************************************************************
            '    //  BIG HONKING NOTE:  This sample resolves all conflicts and does not accomodate for
            '    //                     selective conflict resolution
            '    //  *********************************************************************************
            '
            '    //  Get the "sx:conflicts" sub-element
            Set oConflictElement = oSynchElement.selectSingleNode("sx:conflicts")

            '
            '    //  Validate that "sx:conflicts" sub-element exists
            If Not oConflictElement Is Nothing Then
    
                '        //  Construct hashtable for item history
                '        var ItemHistoryHashtable = new Object();
                '
                '        //  Get "sx:history" sub-elements of "sx:sync" element
                Set oHisObj = oSynchElement.selectNodes("sx:history")
                '
                '        //  Get "item" sub-elements of "sx:conflicts" element
                Set oConflictItm = oConflictElement.selectNodes("item")
                '
                '        //  Iterate "item" sub-elements of "sx:conflicts" element
            
                Dim ConflictSequence As String
                Dim ConflictBy As String
                Dim ConflictWhen As String
                Dim oHis As MSXML2.IXMLDOMElement
                Dim sequence As String
                Dim By As String
                Dim when As String

                For Each oConflict In oConflictItm.childNodes
                    '            //  Get the conflict item's "sx:sync" sub-element
                    Set oConflictSynchElm = oConflict.selectSingleNode("sx:sync")
                    '
                    '            //  Get "sx:history" sub-elements of conflict item's "sx:sync" element
                    Set oConflictHisElm = oConflictSynchElm.selectNodes("sx:history")
                
                    '            //  Iterate "sx:history" sub-elements of conflict item's "sx:sync" element
                    'ok Now Loop Through the History Items
                    For Each oHisElm In oConflictHisElm

                        '
                        '                //  Get "sequence" attribute from conflict item's topmost
                        '                //  "sx:history" sub-element
                        ConflictSequence = oHisElm.getAttribute("sequence")
                        '
                        '                //  Get "by" attribute from conflict item's topmost "sx:history"
                        '                //  sub-element
                        ConflictBy = oHisElm.getAttribute("by")
                        '
                        '                //  Get "when" attribute from conflict item's topmost "sx:history"
                        '                //  sub-element
                        ConflictWhen = oHisElm.getAttribute("when")
                        '
                        bConflictHistoryRepresented = False

                        '
                        '                //  Iterate "sx:history" sub-elements of main item's "sx:sync" element
                        For Each oHis In oHisObj
                            '                    //  Get "sequence" attribute from main item's
                            '                    //  "sx:history" sub-element
                            sequence = oHis.getAttribute("sequence")
                            '
                            '                    //  Get "by" attribute from main item's "sx:history"
                            '                    //  sub-element
                            By = oHis.getAttribute("by")
                            '
                            '                    //  Get "when" attribute from main item's "sx:history"
                            '                    //  sub-element
                            when = oHis.getAttribute("when")

                            '
                            '                    //  See if "by" attribute exists for main item's "sx:history"
                            '                    //  element and if it does, see if it's value matches "by"
                            '                    //  attribute value for conflict's "sx:history" element
                            If ((By <> "") And (By = ConflictBy)) Then

                                '                        //  See if "sequence" attribute for the main item's
                                '                        //  "sx:history" element is greater than or equal to the
                                '                        //  "sequence" attribute for the conflict's "sx:history"
                                '                        //  element
                                If sequence >= ConflictSequence Then
                                    '                            //  Indicate conflict history represented
                                    bConflictHistoryRepresented = True
                                End If
    
                                '                        //  Stop iterating main item's "sx:history" elements
                                Exit For
    
                                '                    //  See if "by" attribute does not exist for both main item's
                                '                    //  "sx:history" element and conflict's "sx:history" element
                            ElseIf ((By = "") And (ConflictBy = "")) Then

                                '                        //  See if "when" attribute exists for both main item's
                                '                        //  "sx:history" element and conflict's "sx:history"
                                '                        //  element
                                If ((when <> "") And (ConflictWhen <> "")) Then

                                    '                            //  Compare date values - since we use RFC3339 values, we can use
                                    '                            //  string comparison when comparing datetimes
                                    If (when = ConflictWhen) Then
    
                                        '                                //  Indicate conflict history represented
                                        bConflictHistoryRepresented = True
    
                                        '                                //  Stop iterating "sx:history" elements
                                        Exit For
                                    End If
                                End If
                            End If
                        
                        Next
    
                        ' //  Do Not Continue iterating conflict's "sx:history" elements
                        If Not bConflictHistoryRepresented Then Exit For
     
                        '                //  Create clone of conflict item's "sx:history" sub-element
                        Set oClonedHis = oHisElm.cloneNode(True)
                        '
                        '                //  Insert cloned conflict item's "sx:history" sub-element after
                        '                //  main item's topmost "sx:history" sub-element
                        oSynchElement.insertBefore oClonedHis, oHisElements.nextSibling
                    
                    Next
                
                Next
    
                '
                '        //  Since we have resolved all conflicts, we remove the "sx:conflicts"
                '        //  sub-element from current parent element
                oConflictElement.parentNode.removeChild (oConflictElement)
            End If
        End If
    
        If bSave Then oXMLDoc.Save sSourcePath
    
        UpdateRssFeedSync = oXMLDoc.xml
    
    End If

End Function

Private Function AddHistoryElements(oSynchElement As MSXML2.IXMLDOMElement, _
                                    sGUID As String, _
                                    cn As ADODB.Connection) As Boolean
    Dim oXMLDoc As New MSXML2.DOMDocument
    Dim oHisElements As MSXML2.IXMLDOMElement
    Dim RsHistory As ADODB.Recordset
    Dim oClonedHis As MSXML2.IXMLDOMElement
    Dim bDeleted As Boolean
    Dim dNoConflicts As Boolean

    AddHistoryElements = False
    
    dNoConflicts = True
    
    Set RsHistory = New ADODB.Recordset
    
    RsHistory.CursorLocation = adUseClient
    
    RsHistory.open "SELECT * FROM " & txtSynchTable(1).Text & " WHERE sGUID = '" & sGUID & "' AND deleted <> 'true' ORDER BY sequence DESC", cn, adOpenDynamic, adLockOptimistic

    If RsHistory.EOF And RsHistory.Bof Then
        Exit Function
    End If

    '//  Get "sx:sync" element with matching id
    'Set oSynchElement = oChannel.selectSingleNode("//item/sx:sync[@id='" + sGUID + "']")

    '    '//  Validate that "sx:sync" element exists
    '    If oSynchElement Is Nothing Then
    '        MsgBox "Unable to find 'sx:sync' element with id='" + sGUID + "'"
    '        Exit Sub
    '    End If
    
    RsHistory.MoveFirst
                
    Do While Not RsHistory.EOF

        '//  Create "sx:history" element
        Set oHisElements = oXMLDoc.createElement("sx:history")
    
        oHisElements.setAttribute "by", RsHistory.fields.Item("sBy").Value
        oHisElements.setAttribute "sequence", RsHistory.fields.Item("sequence").Value
        oHisElements.setAttribute "when", RsHistory.fields.Item("swhen").Value

        If CBool(RsHistory.fields.Item("deleted").Value) Then bDeleted = True
        
        If Not (RsHistory.fields.Item("noconflicts").Value) Then dNoConflicts = False

        If oSynchElement.hasChildNodes Then
            oSynchElement.insertBefore oHisElements, oSynchElement.childNodes(0)
        Else
            oSynchElement.appendChild oHisElements
        End If

        RsHistory.MoveNext
    Loop
    
    oSynchElement.setAttribute "updates", RsHistory.RecordCount
    oSynchElement.setAttribute "deleted", IIf(bDeleted, "true", "false")
    oSynchElement.setAttribute "noconflicts", IIf(dNoConflicts, "true", "false")

    AddHistoryElements = True

End Function

Private Function MergeRSSFeed() As String
    '//  *********************************************************************************
    '//  File:   fsRSSMerge.js
    '//  Notes:  You must run this file with cscript.exe (i.e. not wscript.exe)
    '//
    '//  Copyright (c) Microsoft Corporation.  All Rights Reserved.
    '//  *********************************************************************************
    '
    '
    '//  -------------------------- MAIN (BEGIN) --------------------------
    '
    '//  Validate arguments
    'var g_Arguments = WScript.Arguments;
    'if (g_Arguments.length < 2)
    '    {
    '    DisplayUsage();
    '    WScript.Quit();
    '    }
    '
    '//  Get required parameters
    'var g_LocalPath = g_Arguments(0);
    'var g_IncomingPath = g_Arguments(1);
    '
    'var g_pILocalRSSXmlDOMDocument = null;
    '
    'try
    '    {
    '    //  Create instance of XML DOM
    '    g_pILocalRSSXmlDOMDocument = new ActiveXObject("Microsoft.XMLDOM");
    '    g_pILocalRSSXmlDOMDocument.async = false;
    '
    '    //  Load local document
    '    var Success = g_pILocalRSSXmlDOMDocument.load(g_LocalPath);
    '    if (!Success)
    '        throw new Error(0, "IXmlDocument::load failed");
    '    }
    'catch (e)
    '    {
    '    WScript.Echo("Exception while loading '" + g_LocalPath +"': " + e.message);
    '    WScript.Quit();
    '    }
    '
    '//  Get local "rss" element
    'var g_pILocalRSSXmlDOMElement = g_pILocalRSSXmlDOMDocument.documentElement;
    '
    '//  Check if FeedSync namespace exists, if not then display error
    'if (g_pILocalRSSXmlDOMElement.getAttribute("xmlns:sx") == null)
    '    {
    '    WScript.Echo("Can't process local RSS file - it does not contain a 'sx' namespace!");
    '    WScript.Quit();
    '    }
    '
    '//  Get local "channel" element
    'var g_pILocalChannelXmlDOMElement = g_pILocalRSSXmlDOMElement.selectSingleNode("channel");
    '
    '//  Validate that "channel" element exists
    'if (g_pILocalChannelXmlDOMElement == null)
    '    {
    '    WScript.Echo("Unable to get 'channel' element from local 'rss' element");
    '    WScript.Quit(0);
    '    }
    '
    'var g_pIIncomingRSSXmlDOMDocument = null;
    '
    'try
    '    {
    '    //  Create instance of XML DOM
    '    g_pIIncomingRSSXmlDOMDocument = new ActiveXObject("Microsoft.XMLDOM");
    '    g_pIIncomingRSSXmlDOMDocument.async = false;
    '
    '    //  Load incoming document
    '    var Success = g_pIIncomingRSSXmlDOMDocument.load(g_IncomingPath);
    '    if (!Success)
    '        throw new Error(0, "IXmlDocument::load failed");
    '    }
    'catch (e)
    '    {
    '    WScript.Echo("Exception while loading '" + g_IncomingPath +"': " + e.message);
    '    WScript.Quit();
    '    }
    '
    '//  Get incoming "rss" element
    'var g_pIIncomingRSSXmlDOMElement = g_pIIncomingRSSXmlDOMDocument.documentElement;
    '
    '//  Check if FeedSync namespace exists, if not then display error
    'if (g_pIIncomingRSSXmlDOMElement.getAttribute("xmlns:sx") == null)
    '    {
    '    WScript.Echo("Can't process incoming RSS file - it does not contain a 'sx' namespace!");
    '    WScript.Quit();
    '    }
    '
    '//  Get incoming "channel" element
    'var g_pIIncomingChannelXmlDOMElement = g_pIIncomingRSSXmlDOMElement.selectSingleNode("channel");
    '
    '//  Validate that "channel" element exists
    'if (g_pIIncomingChannelXmlDOMElement == null)
    '    {
    '    WScript.Echo("Unable to get 'channel' element from incoming 'rss' element");
    '    WScript.Quit(0);
    '    }
    '
    '//  *********************************************************************************
    '//  BIG HONKING NOTE:  We only deal with "item" elements when merging, so any other
    '//                     changes made in the incoming document are ignored.  Remember that
    '//                     the goal of FeedSync isn't to replicate RSS files, it is to
    '//                     replicate items via RSS.
    '//  *********************************************************************************
    '
    '//  Create hashtable for local FSNodes
    'var g_LocalFSNodeHashtable = new Object();
    '
    '//  Populate local FSNode hashtable
    'PopulateFSNodesFromXmlDOMElement(g_LocalFSNodeHashtable, g_pILocalChannelXmlDOMElement);
    '
    '//  Create hashtable for incoming FSNodes
    'var g_IncomingFSNodeHashtable = new Object();
    '
    '//  Populate incoming FSNode hashtable
    'PopulateFSNodesFromXmlDOMElement(g_IncomingFSNodeHashtable, g_pIIncomingChannelXmlDOMElement);
    '
    'var g_pIOutputRSSXmlDOMDocument =   null;
    '
    'try
    '    {
    '    //  Create instance of XML DOM
    '    g_pIOutputRSSXmlDOMDocument = new ActiveXObject("Microsoft.XMLDOM");
    '    g_pIOutputRSSXmlDOMDocument.async = false;
    '
    '    //  Create output rss document based on local rss document
    '    g_pILocalRSSXmlDOMDocument.save(g_pIOutputRSSXmlDOMDocument);
    '    if (!Success)
    '        throw new Error(0, "IXmlDocument::load failed");
    '    }
    'catch (e)
    '    {
    '    WScript.Echo("Exception while saving local document: " + e.message);
    '    WScript.Quit();
    '    }
    '
    '//  Create output array
    'var g_OutputFSNodeArray = new Array();
    '
    '//  Get output "rss" element
    'var g_pIOutputRSSXmlDOMElement = g_pIOutputRSSXmlDOMDocument.documentElement;
    '
    '//  Get output "channel" element
    'var g_pIOutputChannelXmlDOMElement = g_pIOutputRSSXmlDOMElement.selectSingleNode("channel");
    '
    '//  Get "item" elements of "channel" element
    'var g_pIItemXmlDOMElements = g_pIOutputChannelXmlDOMElement.selectNodes("item");
    '
    '//  *********************************************************************************
    '//  BIG HONKING NOTE:  Iterate all "item" elements of the "channel" element (start
    '//                     with last and progress to first) in order to remove them.  We
    '//                     remove them because there is further processing below that
    '//                     will take the appropriate "item" elements from the local and
    '//                     incoming documents and add them.  Note that we don't just remove
    '//                     the "channel" element and add a new one in it's place because
    '//                     a) there could be attributes on the "channel" element and b)
    '//                     there could be non-"item" elements of the "channel" element.
    '
    '    for (var Index = g_pIItemXmlDOMElements.length - 1; Index >= 0; --Index)
    '        {
    '        //  Get next "item" element
    '        var pIItemXmlDOMElement = g_pIItemXmlDOMElements(Index);
    '
    '        //  Remove "item" element from output "channel" element
    '        g_pIOutputChannelXmlDOMElement.removeChild(pIItemXmlDOMElement);
    '        }
    '
    '//  *********************************************************************************
    '
    'var g_HashtableKey = null;
    '
    '//  Iterate items in local FSNode hashtable
    'for (g_HashtableKey in g_LocalFSNodeHashtable)
    '    {
    '    //  Get FSNode from local hashtable
    '    var LocalFSNode = g_LocalFSNodeHashtable[g_HashtableKey];
    '
    '    //  Get reference to local FSSyncNode
    '    var LocalFSSyncNode = LocalFSNode.m_FSSyncNode;
    '
    '    //  Get FSNode from incoming hashtable
    '    var IncomingFSNode = g_IncomingFSNodeHashtable[g_HashtableKey];
    '
    '    //  Validate incoming FSNode exists, if not then node was added to local
    '    //  document
    '    if (IncomingFSNode == null)
    '        {
    '        //  Create clone of LocalFSNode
    '        var ClonedFSNode = CloneFSNode(LocalFSNode);
    '
    '        //  Add cloned FSNode to output array
    '        g_OutputFSNodeArray[g_OutputFSNodeArray.length] = ClonedFSNode;
    '
    '        //  Continue loop
    '        continue;
    '        }
    '
    '    //  Get merged FSNode
    '    var MergedFSNode = MergeFSNodes(LocalFSNode, IncomingFSNode);
    '
    '    //  Add merged FSNode to output array
    '    g_OutputFSNodeArray[g_OutputFSNodeArray.length] = MergedFSNode;
    '
    '    //  Set incoming hashtable entry to null so we don't process it during
    '    //  second pass below
    '    g_IncomingFSNodeHashtable[g_HashtableKey] = null;
    '    }
    '
    '//  Iterate items in incoming FSNode hashtable - all remaining items are
    '//  guaranteed to be additions to incoming document
    'for (g_HashtableKey in g_IncomingFSNodeHashtable)
    '    {
    '    //  Get FSNode from incoming hashtable
    '    var IncomingFSNode = g_IncomingFSNodeHashtable[g_HashtableKey];
    '
    '    //  Validate FSNode exists, if not then just continue loop because
    '    //  entry was set to null when processing local hashtable
    '    if (IncomingFSNode == null)
    '        continue;
    '
    '    //  Create clone of IncomingFSNode
    '    var ClonedFSNode = CloneFSNode(IncomingFSNode);
    '
    '    //  Add cloned FSNode to output array
    '    g_OutputFSNodeArray[g_OutputFSNodeArray.length] = ClonedFSNode;
    '    }
    '
    '//  Iterate output FSNodes
    'for (var Index = 0; Index < g_OutputFSNodeArray.length; ++Index)
    '    {
    '    //  Get next output FSNode
    '    var OutputFSNode = g_OutputFSNodeArray[Index];
    '
    '    //  Append output FSNode's element to "channel" element
    '    g_pIOutputChannelXmlDOMElement.appendChild(OutputFSNode.m_pIXmlDOMElement);
    '    }
    '
    '//  Save modified contents to standand output stream
    'WScript.StdOut.Write(g_pIOutputRSSXmlDOMDocument.xml);
    '
    '//  -------------------------- MAIN (END) --------------------------
    '
    '
    '//  -------------------------- FSNodeClass (BEGIN) --------------------------
    '
    'Function FSNodeClass(i_pIXmlDOMElement)
    '    {
    '    //  Assign m_pIOXmlDOMElement member variable
    '    this.m_pIXmlDOMElement = i_pIXmlDOMElement;
    '
    '    //  Assign m_FSSyncNode member variable by creating a new instance of
    '    //  FSSyncNode and passing a reference to the current FSNode
    '    this.m_FSSyncNode = new FSSyncNodeClass(this);
    '    }
    '
    '//  -------------------------- FSNodeClass (END) --------------------------
    '
    '
    '
    '//  -------------------------- FSSyncNodeClass (BEGIN) --------------------------
    '
    'Function FSSyncNodeClass(i_FSNode)
    '    {
    '    //  Assign m_FSNode member variable
    '    this.m_FSNode = i_FSNode;
    '
    '    //  Get reference to FSNode's XmlDOMElement
    '    var pIXmlDOMElement = this.m_FSNode.m_pIXmlDOMElement;
    '
    '    //  Assign m_pIXmlDOMElement member variable
    '    this.m_pIXmlDOMElement = pIXmlDOMElement.selectSingleNode("sx:sync");
    '
    '    //  Validate m_pISyncXmlDOMElement member variable
    '    if (this.m_pIXmlDOMElement == null)
    '        {
    '        WScript.Echo("Unable to find 'sx:sync' element where parent id='" + this.m_FSNode.m_ParentID + "'");
    '        WScript.Quit(0);
    '        }
    '
    '    //  Assign m_ID member variable
    '    this.m_ID = this.m_pIXmlDOMElement.getAttribute("id");
    '
    '    //  Validate m_ID member variable
    '    if (this.m_ID == null)
    '        {
    '        WScript.Echo("Unable to find 'id' attribute for 'sx:sync' element where parent id='" + this.m_FSNode.m_ParentID + "'");
    '        WScript.Quit(0);
    '        }
    '
    '    //  Assign m_Updates member variable
    '    this.m_Updates = this.m_pIXmlDOMElement.getAttribute("updates");
    '
    '    //  Validate m_Updates member variable
    '    if (this.m_Updates == null)
    '        {
    '        WScript.Echo("Unable to find 'updates' attribute for 'sx:sync' element where id='" + this.m_ID + "'");
    '        WScript.Quit(0);
    '        }
    '
    '    //  Assign m_NoConflicts member variable
    '    var NoConflicts = this.m_pIXmlDOMElement.getAttribute("noconflicts");
    '
    '    //  Validate m_Conflict member variable
    '    if (NoConflicts == "true")
    '        this.m_NoConflicts = true;
    '    Else
    '        this.m_NoConflicts = false;
    '
    '    //  Assign m_FSConflictNodes member variable by creating a new array
    '    this.m_FSConflictNodes = new Array();
    '
    '    if (!this.m_NoConflicts)
    '        {
    '        //  Get reference to "sx:conflicts" element
    '        this.m_pIConflictsXmlDOMElement = this.m_pIXmlDOMElement.selectSingleNode("sx:conflicts");
    '
    '        //  Validate that "sx:conflicts" element exists
    '        if (this.m_pIConflictsXmlDOMElement != null)
    '            {
    '            //  Get conflict "item" elements
    '            var pIConflictItemXmlDOMElements = this.m_pIConflictsXmlDOMElement.selectNodes("item");
    '
    '            //  Iterate conflict "item" elements
    '            for (var Index = 0; Index < pIConflictItemXmlDOMElements.length; ++Index)
    '                {
    '                //  Get reference to next conflict "item" element
    '                var pIConflictItemXmlDOMElement = pIConflictItemXmlDOMElements(Index);
    '
    '                //  Assign array entry by creating a new instance of FSNode and passing
    '                //  a reference to the current conflict "item" element
    '                this.m_FSConflictNodes[Index] = new FSNodeClass(pIConflictItemXmlDOMElement);
    '                }
    '            }
    '        }
    '
    '    //  Assign m_FSHistoryNodes member variable by creating a new array
    '    this.m_FSHistoryNodes = new Array();
    '
    '    //  Get "sx:history" elements
    '    var pIHistoryXmlDOMElements = this.m_pIXmlDOMElement.selectNodes("sx:history");
    '
    '    //  Iterate "sx:history" elements
    '    for (var Index = 0; Index < pIHistoryXmlDOMElements.length; ++Index)
    '        {
    '        //  Get reference to next "sx:history" element
    '        var pIHistoryXmlDOMElement = pIHistoryXmlDOMElements(Index);
    '
    '        //  Assign array entry by creating a new instance of FSHistoryNode and passing
    '        //  a reference to the current "sx:history" element
    '        this.m_FSHistoryNodes[Index] = new FSHistoryNodeClass(this, pIHistoryXmlDOMElement);
    '        }
    '    }
    '
    '//  -------------------------- FSSyncNodeClass (END) --------------------------
    '
    '
    '//  -------------------------- FSHistoryNodeClass (BEGIN) --------------------------
    '
    'Function FSHistoryNodeClass(i_FSSyncNode, i_pIHistoryXmlDOMElement)
    '    {
    '    //  Assign m_FSSyncNode member variable
    '    this.m_FSSyncNode = i_FSSyncNode;
    '
    '    //  Assign m_pIXmlDOMElement member variable
    '    this.m_pIXmlDOMElement = i_pIHistoryXmlDOMElement;
    '
    '    //  Validate m_pIXmlDOMElement member variable
    '    if (this.m_pIXmlDOMElement == null)
    '        {
    '        WScript.Echo("Unable to find 'sx:history' element for 'sx:sync' element where id='" + this.m_FSSyncNode.m_ID + "'");
    '        return;
    '        }
    '
    '    //  Assign m_Sequence member variable
    '    this.m_Sequence = this.m_pIXmlDOMElement.getAttribute("sequence");
    '
    '    //  Assign m_When member variable
    '    this.m_When = this.m_pIXmlDOMElement.getAttribute("when");
    '
    '    //  Assign m_By member variable
    '    this.m_By = this.m_pIXmlDOMElement.getAttribute("by");
    '
    '    //  Validate that either m_When or m_By member variable have been assigned
    '    if ((this.m_When == null) && (this.m_By == null))
    '        {
    '        WScript.Echo("Unable to find 'when' or 'by' attribute in 'sx:history' element for 'sx:sync' element where id='" + this.m_FSSyncNode.m_ID + "'");
    '        WScript.Quit(0);
    '        }
    '    }
    '
    '//  -------------------------- FSHistoryNodeClass (BEGIN) --------------------------
    '
    '
    'Function PopulateFSNodesFromXmlDOMElement(i_Hashtable, i_pIXmlDOMElement)
    '    {
    '    //  Get "item" elements
    '    var pIItemXmlDOMElements = i_pIXmlDOMElement.selectNodes("item");
    '
    '    //  Iterate "item" elements
    '    for (var Index = 0; Index < pIItemXmlDOMElements.length; ++Index)
    '        {
    '        //  Get reference to next "item" element
    '        var pIItemXmlDOMElement = pIItemXmlDOMElements(Index);
    '
    '        //  Create new instance of FSNodeClass
    '        var FSNode = new FSNodeClass(pIItemXmlDOMElement);
    '
    '        //  Get reference to FSSyncNode
    '        var FSSyncNode = FSNode.m_FSSyncNode;
    '
    '        //  Add FSNode to hashtable using FSSyncNode's id as key
    '        i_Hashtable[FSSyncNode.m_ID] = FSNode;
    '        }
    '    }
    '
    'Function MergeFSNodes(i_LocalFSNode, i_IncomingFSNode)
    '    {
    '    //  Create collection for local item and local item's conflicts
    '    var LocalItemCollection = new Array();
    '
    '    //  Created clone of local FSNode
    '    var ClonedLocalFSNode = CloneFSNode(i_LocalFSNode);
    '
    '    //  Get reference to local FSSyncNode
    '    var ClonedLocalFSSyncNode = ClonedLocalFSNode.m_FSSyncNode;
    '
    '    //  Populate collection with clone of local item's conflicts
    '    for (var Index = 0; Index < ClonedLocalFSSyncNode.m_FSConflictNodes.length; ++Index)
    '        LocalItemCollection[LocalItemCollection.length] = CloneFSNode(ClonedLocalFSSyncNode.m_FSConflictNodes[Index]);
    '
    '    //  See if "sx:conflicts" element exists, if so remove it
    '    if (ClonedLocalFSSyncNode.m_pIConflictsXmlDOMElement != null)
    '        ClonedLocalFSSyncNode.m_pIConflictsXmlDOMElement.parentNode.removeChild(ClonedLocalFSSyncNode.m_pIConflictsXmlDOMElement);
    '
    '    //  Populate collection with clone of local item
    '    LocalItemCollection[LocalItemCollection.length] = ClonedLocalFSNode;
    '
    '    //  Create collection for incoming item and incoming item's conflicts
    '    var IncomingItemCollection = new Array();
    '
    '    //  Created clone of incoming FSNode
    '    var ClonedIncomingFSNode = CloneFSNode(i_IncomingFSNode);
    '
    '    //  Get reference to incoming FSSyncNode
    '    var ClonedIncomingFSSyncNode = ClonedIncomingFSNode.m_FSSyncNode;
    '
    '    //  Populate collection with clone of incoming item's conflicts
    '    for (var Index = 0; Index < ClonedIncomingFSSyncNode.m_FSConflictNodes.length; ++Index)
    '        IncomingItemCollection[IncomingItemCollection.length] = CloneFSNode(ClonedIncomingFSSyncNode.m_FSConflictNodes[Index]);
    '
    '    //  See if "sx:conflicts" element exists, if so remove it
    '    if (ClonedIncomingFSSyncNode.m_pIConflictsXmlDOMElement != null)
    '        ClonedIncomingFSSyncNode.m_pIConflictsXmlDOMElement.parentNode.removeChild(ClonedIncomingFSSyncNode.m_pIConflictsXmlDOMElement);
    '
    '    //  Populate collection with clone of incoming item
    '    IncomingItemCollection[IncomingItemCollection.length] = ClonedIncomingFSNode;
    '
    '    //  Create collection for merge result
    '    var MergeResultItemCollection = new Array();
    '
    '    var WinnerFSNode = null;
    '
    '    //  Process collections using local item collection as outer collection
    '    //  and incoming item collection as inner collection
    '    WinnerFSNode = ProcessCollections(LocalItemCollection, IncomingItemCollection, MergeResultItemCollection, WinnerFSNode);
    '
    '    //  Process collections using incoming item collection as outer collection
    '    //  and local item collection as inner collection
    '    WinnerFSNode = ProcessCollections(IncomingItemCollection, LocalItemCollection, MergeResultItemCollection, WinnerFSNode);
    '
    '    //  Get reference to winner's FSSyncNode
    '    var WinnerFSSyncNode = WinnerFSNode.m_FSSyncNode;
    '
    '    //  If the "noconflicts" attribute is true, or if there is only one
    '    //  item in the merge result collection (i.e. the winner), then we are
    '    //  done processing
    '    if (WinnerFSSyncNode.m_NoConflicts || (MergeResultItemCollection.length == 1))
    '        return WinnerFSNode;
    '
    '    //  Create "sx:conflicts" element for winner
    '    var pIWinnerConflictsXmlDOMElement = g_pIOutputRSSXmlDOMDocument.createElement("sx:conflicts");
    '
    '    //  Append "sx:conflicts" element to winner's "sx:sync" element
    '    WinnerFSSyncNode.m_pIXmlDOMElement.appendChild(pIWinnerConflictsXmlDOMElement);
    '
    '    //  Get reference to winner's conflict nodes
    '    var WinnerFSConflictNodes = WinnerFSSyncNode.m_FSConflictNodes;
    '
    '    //  Create empty array to hold winner's conflict nodes
    '    WinnerFSConflictNodes = new Array();
    '
    '    //  Iterate items in merge result collection
    '    for (var Index = 0; Index < MergeResultItemCollection.length; ++Index)
    '        {
    '        //  Get next item in merge result collection
    '        var MergeResultItem = MergeResultItemCollection[Index];
    '
    '        //  If the merge result item matches the winner item, just
    '        //  continue the loop
    '        if (0 == CompareFSNodes(WinnerFSNode, MergeResultItem))
    '            continue;
    '
    '        //  Get reference to merge result item's element
    '        var pIMergeResultItemXmlDOMElement = MergeResultItemCollection[Index].m_pIXmlDOMElement;
    '
    '        //  Append merge result's element to winner's "sx:conflicts" element
    '        pIWinnerConflictsXmlDOMElement.appendChild(pIMergeResultItemXmlDOMElement);
    '
    '        //  Add new item to winner's conflict nodes
    '        WinnerFSConflictNodes[WinnerFSConflictNodes.length] = new FSNodeClass(pIMergeResultItemXmlDOMElement);
    '        }
    '
    '    return WinnerFSNode;
    '    }
    '

End Function

Function ProcessCollections(i_OuterFSNodeCollection, i_InnerFSNodeCollection, io_MergeFSNodeCollection, i_WinnerFSNode)
    '    {
    '    //  Iterate outer FSNode collection
    '    for (var OuterFSNodeCollectionIndex = 0; OuterFSNodeCollectionIndex < i_OuterFSNodeCollection.length; ++OuterFSNodeCollectionIndex)
    '        {
    '        //  Get next FSNode in outer collection
    '        var OuterFSNode = i_OuterFSNodeCollection[OuterFSNodeCollectionIndex];
    '
    '        //  Get reference to outer FSSyncNode
    '        var OuterFSSyncNode = OuterFSNode.m_FSSyncNode;
    '
    '        var OuterFSNodeSubsumed = false;
    '
    '        //  Iterate inner FSNode collection
    '        for (var InnerFSNodeCollectionIndex = 0; InnerFSNodeCollectionIndex < i_InnerFSNodeCollection.length; ++InnerFSNodeCollectionIndex)
    '            {
    '            //  Get next FSNode in inner collection
    '            var InnerFSNode = i_InnerFSNodeCollection[InnerFSNodeCollectionIndex];
    '
    '            //  Check value of inner FSNode exists - if not then
    '            //  just continue loop
    '            if (InnerFSNode == null)
    '                continue;
    '
    '            //  Get reference to inner FSSyncNode
    '            var InnerFSSyncNode = InnerFSNode.m_FSSyncNode;
    '
    '            //  Get the topmost "sx:history" element for the outer FSSyncNode
    '            var OuterFSHistoryNode = OuterFSNode.m_FSSyncNode.m_FSHistoryNodes[0];
    '
    '            //  Iterate FSHistoryNodes for inner FSSyncNode
    '            for (var HistoryIndex = 0; HistoryIndex < InnerFSSyncNode.m_FSHistoryNodes.length; ++HistoryIndex)
    '                {
    '                //  Get next FSHistoryNode
    '                var InnerFSHistoryNode = InnerFSSyncNode.m_FSHistoryNodes[HistoryIndex];
    '
    '                //  See if "by" attribute exists for outer FSHistoryNode and if
    '                //  it does, see if it's value matches "by" attribute value for
    '                //  inner FSHistoryNode
    '                if ((OuterFSHistoryNode.m_By != null) && (OuterFSHistoryNode.m_By == InnerFSHistoryNode.m_By))
    '                    {
    '                    //  See if "sequence" attribute for the inner FSHistoryNode
    '                    //  is greater than or equal to the "sequence" attribute for
    '                    //  the outer FSHistoryNode
    '                    if (InnerFSHistoryNode.m_Sequence >= OuterFSHistoryNode.m_Sequence)
    '                        {
    '                        //  Indicate subsumption
    '                        OuterFSNodeSubsumed = true;
    '                        }
    '
    '                    //  Stop iterating FSHistoryNodes
    '                    break;
    '                    }
    '
    '                //  See if "by" attribute does not exist for both outer FSHistoryNode
    '                //  and inner FSHistoryNode
    '                else if ((OuterFSHistoryNode.m_By == null) && (InnerFSHistoryNode.m_By == null))
    '                    {
    '                    //  See if "when" attribute exists for both outer FSHistoryNode
    '                    //  and inner FSHistoryNode
    '                    if ((InnerFSHistoryNode.m_When != null) && (OuterFSHistoryNode.m_When != null))
    '                        {
    '                        //  See if normalized dates match - if so then the outer FSNode
    '                        //  is subsumed
    '                        if (InnerFSHistoryNode.m_When == OuterFSHistoryNode.m_When)
    '                            {
    '                            //  Indicate subsumption
    '                            OuterFSNodeSubsumed = true;
    '
    '                            //  Stop iterating FSHistoryNodes
    '                            break;
    '                            }
    '                        }
    '                    }
    '                }
    '
    '            //  Check for subsumption
    '            if (OuterFSNodeSubsumed)
    '                {
    '                //  Stop iterating inner FSNodes
    '                break;
    '                }
    '            }
    '
    '        //  Check for subsumption
    '        if (OuterFSNodeSubsumed)
    '            {
    '            //  Remove outer FSNode from outer FSNode collection
    '            i_OuterFSNodeCollection[OuterFSNodeCollectionIndex] = null;
    '
    '            //  Continue iterating outer FSNodes
    '            continue;
    '            }
    '
    '        //  See if outer FSSyncNode has any FSConflictNodes
    '        if (OuterFSSyncNode.m_FSConflictNodes.length > 0)
    '            {
    '            //  Remove the "sx:conflicts" sub-element for outer
    '            //  FSSyncNode
    '            OuterFSSyncNode.m_pIConflictsXmlDOMElement.parentNode.removeChild(OuterFSSyncNode.m_pIConflictsXmlDOMElement);
    '            }
    '
    '        //  Add the outer FSNode to the merge result collection
    '        io_MergeFSNodeCollection[io_MergeFSNodeCollection.length] = OuterFSNode;
    '
    '        //  See if winner FSNode has not been assigned yet or
    '        //  if the outer FSNode represents a more recent update
    '        //  than that of the current winner FSNode
    '        if ((i_WinnerFSNode == null) || (-1 == CompareFSNodes(i_WinnerFSNode, OuterFSNode)))
    '            {
    '            //  Assign the outer FSNode as the winner FSNode
    '            i_WinnerFSNode = OuterFSNode;
    '            }
    '        }
    '
    '    return i_WinnerFSNode;
    '    }
    '
End Function

Function CompareFSNodes(i_FSNode1 As Variant, i_FSNode2 As Variant) As Boolean
    '    {
    '    //  This function compares the two FSNodes and returns:
    '    //     1 if i_FSNode1 is newer than i_FSNode2
    '    //    -1 if i_FSNode2 is newer than i_FSNode1
    '    //     0 if FSNodes are equal
    '    //     null if FSNodes are equal but conflict data is different
    '    //
    '
    '    //  Get reference to FSSyncNode for i_FSNode1
    '    var FSSyncNode1 = i_FSNode1.m_FSSyncNode;
    '
    '    //  Get reference to FSSyncNode for i_FSNode2
    '    var FSSyncNode2 = i_FSNode2.m_FSSyncNode;
    '
    '    //  Compare "updates" attributes - if they are equal then do subsequent checks
    '    if (FSSyncNode1.m_Updates == FSSyncNode2.m_Updates)
    '        {
    '        //  Get reference to topmost FSHistoryNode for FSSyncNode1
    '        var FSHistoryNode1 = FSSyncNode1.m_FSHistoryNodes[0];
    '
    '        //  Get reference to topmost FSHistoryNode for FSSyncNode2
    '        var FSHistoryNode2 = FSSyncNode2.m_FSHistoryNodes[0];
    '
    '        //  See if "when" attribute exist for either FSHistoryNode
    '        if ((FSHistoryNode1.m_When != null) || (FSHistoryNode2.m_When != null))
    '            {
    '            //  See if "when" attribute exist for both FSHistoryNodes
    '            if ((FSHistoryNode1.m_When != null) && (FSHistoryNode2.m_When != null))
    '                {
    '                //  Compare date values - since we use RFC3339 values, we can use
    '                //  string comparison when comparing datetimes
    '                if (FSHistoryNode1.m_When > FSHistoryNode2.m_When)
    '                    {
    '                    //  FSHistoryNode1 node has a later "when" attribute, so i_FSNode1
    '                    //  is newer
    '                    return 1;
    '                    }
    '                else if (FSHistoryNode2.m_When > FSHistoryNode1.m_When)
    '                    {
    '                    //  FSHistoryNode2 node has a later "when" attribute, so i_FSNode2
    '                    //  is newer
    '                    return -1;
    '                    }
    '                Else
    '                    {
    '                    //  Same "when" attribute value for both FSHistoryNodes - try further
    '                    //  checking below
    '                    }
    '                }
    '            else if (FSHistoryNode1.m_When != null)
    '                {
    '                //  FSHistoryNode1 has a "when" attribute but FSHistoryNode2 does not, so
    '                //  i_FSNode1 is newer
    '                return 1;
    '                }
    '            Else
    '                {
    '                //  FSHistoryNode2 has a "when" attribute but FSHistoryNode1 does not, so
    '                //  i_FSNode2 is newer
    '                return -1;
    '                }
    '            }
    '        Else
    '            {
    '            //  Neither FSHistoryNode has "when" attribute - try further checking below
    '            }
    '
    '        //  See if "by" attribute exist for either FSHistoryNode
    '        if ((FSHistoryNode1.m_By != null) || (FSHistoryNode2.m_By != null))
    '            {
    '            //  See if "by" attribute exist for both FSHistoryNodes
    '            if ((FSHistoryNode1.m_By != null) && (FSHistoryNode2.m_By != null))
    '                {
    '                //  Compare "by" values
    '                if (FSHistoryNode1.m_By > FSHistoryNode2.m_By)
    '                    {
    '                    //  FSHistoryNode1 node has a later "by" attribute, so i_FSNode1
    '                    //  is newer
    '                    return 1;
    '                    }
    '                else if (FSHistoryNode1.m_By < FSHistoryNode2.m_By)
    '                    {
    '                    //  FSHistoryNode2 node has a later "by" attribute, so i_FSNode2
    '                    //  is newer
    '                    return -1;
    '                    }
    '                Else
    '                    {
    '                    //  Same "by" attribute value for both FSHistoryNodes - so we must
    '                    //  compare conflict items
    '
    '                    //  If number of conflict item nodes is different, items are equal
    '                    //  but conflict item data is different
    '                    if (FSSyncNode1.m_FSConflictNodes.length != FSSyncNode2.m_FSConflictNodes.length)
    '                        return null;
    '
    '                    //  Check if conflict item nodes exist
    '                    if (FSSyncNode1.m_FSConflictNodes.length > 0)
    '                        {
    '                        //  Iterate conflict nodes for item 1
    '                        for (var Index1 = 0; Index1 < FSSyncNode1.m_FSConflictNodes.length; ++Index)
    '                            {
    '                            var MatchingConflictItem = false;
    '
    '                            //  Get reference to next conflict node for item 1
    '                            var ConflictNode1 = FSSyncNode1.m_FSConflictNodes[Index1];
    '
    '                            //  Iterate conflict nodes for item 2
    '                            for (var Index2 = 0; Index2 < FSSyncNode2.m_FSConflictNodes.length; ++Index)
    '                                {
    '                                //  Get reference to next conflict node for item 2
    '                                var ConflictNode2 = FSSyncNode2.m_FSConflictNodes[Index2];
    '
    '                                //  Compare conflict nodes
    '                                if (0 == CompareFSNodes(ConflictNode1, ConflictNode2))
    '                                    {
    '                                    MatchingConflictItem = true;
    '                                    break;
    '                                    }
    '                                }
    '                            }
    '
    '                        if (!MatchingConflictItem)
    '                            {
    '                            //  No matching conflict item - so items are equal but conflict
    '                            //  item data is different
    '                            return null;
    '                            }
    '                        }
    '
    '                        //  Items are equal
    '                        return 0;
    '                    }
    '                }
    '            }
    '        else if (FSHistoryNode1.m_By != null)
    '            {
    '            //  FSHistoryNode1 has a "by" attribute but FSHistoryNode2 does not, so
    '            //  i_FSNode1 is newer
    '            return 1;
    '            }
    '        else if (FSHistoryNode2.m_By != null)
    '            {
    '            //  FSHistoryNode2 has a "by" attribute but FSHistoryNode1 does not, so
    '            //  i_FSNode2 is newer
    '            return -1;
    '            }
    '        Else
    '            {
    '            //  Neither FSHistoryNode has "by" attribute - so we can't tell which
    '            //  FSNode is newer
    '            return 0;
    '            }
    '        }
    '    else if (FSSyncNode1.m_Updates> FSSyncNode2.m_Updates)
    '        {
    '        //  FSSyncNode1 has a later "updates" attribute, so i_FSNode1 is newer
    '        return 1;
    '        }
    '    Else
    '        {
    '        //  FSSyncNode2 has a later "updates" attribute, so i_FSNode2 is newer
    '        return -1;
    '        }
    '    }
    '

End Function

Function CloneFSNode(i_FSNode As Variant)
    '    {
    '    //  Get reference to original FSNode's XmlDOMElement
    '    var pIXmlDOMElement = i_FSNode.m_pIXmlDOMElement;
    '
    '    //  Create (deep copy) clone of XmlDOMElement
    '    var pIClonedXmlDOMElement = pIXmlDOMElement.cloneNode(true);
    '
    '    //  Create new instance of FSNode
    '    var ClonedFSNode = new FSNodeClass(pIClonedXmlDOMElement);
    '
    '    //  Return new instance of FSNode
    '    return ClonedFSNode;
    '    }
End Function

Private Sub m_RSServerItems_FetchComplete(ByVal pError As ADODB.Error, _
                                          adStatus As ADODB.EventStatusEnum, _
                                          ByVal pRecordset As ADODB.Recordset)
    m_frmDebug.DebugPrint "progress Done"
End Sub

Private Sub m_RSServerItems_FetchProgress(ByVal Progress As Long, _
                                          ByVal MaxProgress As Long, _
                                          adStatus As ADODB.EventStatusEnum, _
                                          ByVal pRecordset As ADODB.Recordset)
    m_frmDebug.DebugPrint "progress:" & Progress & " maxprogress:" & MaxProgress
End Sub

Private Sub OptMode_Click(Index As Integer)

    Select Case Index
    
        Case 0
        
        Case 1
    
    End Select

End Sub

