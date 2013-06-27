VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmDataToolMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Data Maintainance Tool"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10140
   Icon            =   "frmDataToolMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDataToolMain.frx":6852
   ScaleHeight     =   6390
   ScaleWidth      =   10140
   StartUpPosition =   3  'Windows Default
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   6390
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10140
      _cx             =   17886
      _cy             =   11271
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
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   2
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   0
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmDataToolMain.frx":D0A4
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic elStatusBar 
         Height          =   465
         Left            =   0
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   5925
         Width           =   10140
         _cx             =   17886
         _cy             =   820
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   -1  'True
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         FloodColor      =   49152
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   0   'False
         FloodDirection  =   1
         FloodPercent    =   0
         CaptionPos      =   4
         WordWrap        =   -1  'True
         MaxChildSize    =   0
         MinChildSize    =   0
         TagWidth        =   0
         TagPosition     =   0
         Style           =   0
         TagSplit        =   2
         PicturePos      =   4
         CaptionStyle    =   0
         ResizeFonts     =   0   'False
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   5
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
      End
      Begin C1SizerLibCtl.C1Tab c1TabMain 
         Height          =   5925
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10140
         _cx             =   17886
         _cy             =   10451
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
         Caption         =   "Import|Export"
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
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic elExport 
            Height          =   5550
            Left            =   10785
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   330
            Width           =   10050
            _cx             =   17727
            _cy             =   9790
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
            Appearance      =   0
            MousePointer    =   0
            Version         =   801
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   1
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"frmDataToolMain.frx":D0E4
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame FraExport 
               Caption         =   "Export"
               Height          =   5370
               Left            =   90
               TabIndex        =   3
               Top             =   90
               Width           =   9870
               Begin VB.CommandButton Command1 
                  Caption         =   "Export"
                  Height          =   375
                  Left            =   8775
                  TabIndex        =   28
                  Top             =   4860
                  Width           =   960
               End
               Begin VB.CommandButton cmdReset 
                  Caption         =   "Reset"
                  Height          =   375
                  Left            =   7695
                  TabIndex        =   27
                  Top             =   4860
                  Width           =   960
               End
               Begin C1SizerLibCtl.C1Tab C1TSQLScripts 
                  Height          =   5055
                  Left            =   45
                  TabIndex        =   16
                  Top             =   225
                  Width           =   9780
                  _cx             =   17251
                  _cy             =   8916
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
                  Appearance      =   0
                  MousePointer    =   0
                  Version         =   801
                  BackColor       =   -2147483633
                  ForeColor       =   -2147483630
                  FrontTabColor   =   -2147483633
                  BackTabColor    =   -2147483633
                  TabOutlineColor =   -2147483632
                  FrontTabForeColor=   -2147483630
                  Caption         =   "SQL Scripts|Data"
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
                  TabHeight       =   1
                  TabCaptionPos   =   4
                  TabPicturePos   =   0
                  CaptionEmpty    =   ""
                  Separators      =   0   'False
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   37
                  Flags(1)        =   2
                  Begin C1SizerLibCtl.C1Elastic c1Data 
                     Height          =   5025
                     Left            =   10395
                     TabIndex        =   29
                     TabStop         =   0   'False
                     Top             =   15
                     Width           =   9750
                     _cx             =   17198
                     _cy             =   8864
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
                     Appearance      =   0
                     MousePointer    =   0
                     Version         =   801
                     BackColor       =   -2147483633
                     ForeColor       =   -2147483630
                     FloodColor      =   6553600
                     ForeColorDisabled=   -2147483631
                     Caption         =   ""
                     Align           =   0
                     AutoSizeChildren=   0
                     BorderWidth     =   6
                     ChildSpacing    =   4
                     Splitter        =   0   'False
                     FloodDirection  =   0
                     FloodPercent    =   0
                     CaptionPos      =   1
                     WordWrap        =   -1  'True
                     MaxChildSize    =   0
                     MinChildSize    =   0
                     TagWidth        =   0
                     TagPosition     =   0
                     Style           =   0
                     TagSplit        =   2
                     PicturePos      =   4
                     CaptionStyle    =   0
                     ResizeFonts     =   0   'False
                     GridRows        =   0
                     GridCols        =   0
                     Frame           =   3
                     FrameStyle      =   0
                     FrameWidth      =   1
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   ""
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                  End
                  Begin C1SizerLibCtl.C1Elastic c1SQLScripts 
                     Height          =   5025
                     Left            =   15
                     TabIndex        =   17
                     TabStop         =   0   'False
                     Top             =   15
                     Width           =   9750
                     _cx             =   17198
                     _cy             =   8864
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
                     Appearance      =   0
                     MousePointer    =   0
                     Version         =   801
                     BackColor       =   -2147483633
                     ForeColor       =   -2147483630
                     FloodColor      =   6553600
                     ForeColorDisabled=   -2147483631
                     Caption         =   ""
                     Align           =   0
                     AutoSizeChildren=   0
                     BorderWidth     =   6
                     ChildSpacing    =   4
                     Splitter        =   0   'False
                     FloodDirection  =   0
                     FloodPercent    =   0
                     CaptionPos      =   1
                     WordWrap        =   -1  'True
                     MaxChildSize    =   0
                     MinChildSize    =   0
                     TagWidth        =   0
                     TagPosition     =   0
                     Style           =   0
                     TagSplit        =   2
                     PicturePos      =   4
                     CaptionStyle    =   0
                     ResizeFonts     =   0   'False
                     GridRows        =   0
                     GridCols        =   0
                     Frame           =   3
                     FrameStyle      =   0
                     FrameWidth      =   1
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   ""
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin VB.Frame FraExportSQL 
                        Caption         =   "Client DB Update SQL:"
                        Height          =   4605
                        Left            =   90
                        TabIndex        =   24
                        Top             =   0
                        Width           =   4875
                        Begin VB.CheckBox chkUseClient 
                           Caption         =   "Use Client DB Update SQL Script:"
                           Height          =   375
                           Left            =   135
                           TabIndex        =   26
                           Top             =   225
                           Width           =   2805
                        End
                        Begin VB.TextBox txtExportSQL 
                           ForeColor       =   &H0000C000&
                           Height          =   3885
                           Left            =   135
                           MultiLine       =   -1  'True
                           TabIndex        =   25
                           Text            =   "frmDataToolMain.frx":D11B
                           Top             =   630
                           Width           =   4650
                        End
                     End
                     Begin VB.Frame FraTables 
                        Caption         =   "Tables"
                        Height          =   4605
                        Left            =   4995
                        TabIndex        =   18
                        Top             =   0
                        Width           =   4695
                        Begin DXDBGRIDLibCtl.dxDBGrid gridTables 
                           Height          =   3435
                           Left            =   90
                           OleObjectBlob   =   "frmDataToolMain.frx":D13A
                           TabIndex        =   30
                           Top             =   990
                           Width           =   4560
                        End
                        Begin VB.ListBox lstTAbles 
                           Height          =   3435
                           Left            =   90
                           Sorted          =   -1  'True
                           Style           =   1  'Checkbox
                           TabIndex        =   23
                           Top             =   990
                           Width           =   3615
                        End
                        Begin VB.Frame FraType 
                           Caption         =   "Type:"
                           Height          =   645
                           Left            =   90
                           TabIndex        =   19
                           Top             =   225
                           Width           =   4515
                           Begin VB.OptionButton OptExpType 
                              Caption         =   "Advanced (Query)"
                              Height          =   195
                              Index           =   2
                              Left            =   2745
                              TabIndex        =   22
                              Top             =   270
                              Visible         =   0   'False
                              Width           =   1815
                           End
                           Begin VB.OptionButton OptExpType 
                              Caption         =   "Partial"
                              Height          =   195
                              Index           =   0
                              Left            =   225
                              TabIndex        =   21
                              Top             =   270
                              Value           =   -1  'True
                              Width           =   825
                           End
                           Begin VB.OptionButton OptExpType 
                              Caption         =   "Full"
                              Height          =   195
                              Index           =   1
                              Left            =   1575
                              TabIndex        =   20
                              Top             =   270
                              Width           =   780
                           End
                        End
                     End
                  End
               End
               Begin VB.CommandButton Command2 
                  Caption         =   "Command2"
                  Height          =   195
                  Left            =   765
                  TabIndex        =   9
                  Top             =   4230
                  Visible         =   0   'False
                  Width           =   285
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic elImport 
            Height          =   5550
            Left            =   45
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   330
            Width           =   10050
            _cx             =   17727
            _cy             =   9790
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
            Appearance      =   0
            MousePointer    =   0
            Version         =   801
            BackColor       =   -2147483633
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   1
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   4
            CaptionStyle    =   0
            ResizeFonts     =   0   'False
            GridRows        =   0
            GridCols        =   0
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.Frame FraImport 
               Caption         =   "Import"
               Height          =   5370
               Left            =   0
               TabIndex        =   5
               Top             =   0
               Width           =   10005
               Begin VB.CheckBox chkImportClientDBSQLScripts 
                  Caption         =   "Import Client DB SQL Scripts "
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   135
                  TabIndex        =   15
                  Top             =   4860
                  Width           =   3705
               End
               Begin VB.Frame FraImportDescription 
                  Caption         =   "Import Description:"
                  Height          =   4470
                  Left            =   4005
                  TabIndex        =   12
                  Top             =   225
                  Width           =   5865
                  Begin VB.TextBox txtImportDescription 
                     ForeColor       =   &H008080FF&
                     Height          =   4065
                     Left            =   135
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   13
                     Top             =   270
                     Width           =   5550
                  End
               End
               Begin VB.Frame FraImportData 
                  Caption         =   "Available Import Data"
                  Height          =   4470
                  Left            =   135
                  TabIndex        =   10
                  Top             =   225
                  Width           =   3750
                  Begin VB.ListBox lstImportTables 
                     Height          =   4110
                     Left            =   180
                     Sorted          =   -1  'True
                     Style           =   1  'Checkbox
                     TabIndex        =   11
                     Top             =   270
                     Width           =   3390
                  End
               End
               Begin VB.TextBox txtFile 
                  Height          =   330
                  Left            =   8820
                  TabIndex        =   8
                  Top             =   3330
                  Visible         =   0   'False
                  Width           =   915
               End
               Begin VB.CommandButton cmdBrowse 
                  Caption         =   "Browse"
                  Height          =   330
                  Left            =   8910
                  TabIndex        =   7
                  Top             =   2880
                  Visible         =   0   'False
                  Width           =   735
               End
               Begin VB.CommandButton cmdUpdate 
                  Caption         =   "IMPORT"
                  Height          =   420
                  Left            =   8640
                  TabIndex        =   6
                  Top             =   4770
                  Width           =   1230
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmDataToolMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_bExportMade As Boolean
Private oINIReader As New clIniReader
Private m_StrIniFile As String
Private m_strTables() As String
Private m_strUIDs() As String
Private m_strImportFileNames() As String
Private m_strSQL() As String
Private m_strTimeStamp() As String
Private m_strAction() As String
Private m_strKey() As String

Private oRSFields As New ADODB.Recordset

Private Declare Function CoCreateGuid _
                Lib "ole32.dll" (pguid As GUID) As Long
Private Declare Function StringFromGUID2 _
                Lib "ole32.dll" (rguid As Any, _
                                 ByVal lpstrClsId As Long, _
                                 ByVal cbMax As Long) As Long
Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

'GUID STRUCT
Private Type GUID
    data1 As Long
    data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName _
                Lib "comdlg32.dll" _
                Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function GetSaveFileName _
                Lib "comdlg32.dll" _
                Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Function SaveRSToXML(ConnectionString As String, _
                            SQLString As String, _
                            FullPath As String) As Boolean
        '<EhHeader>
        On Error GoTo SaveRSToXML_Err
        '</EhHeader>

        Dim oCn As New ADODB.Connection
        Dim oCmd As New ADODB.Command
        Dim oRs As ADODB.Recordset

        On Error Resume Next
100     Kill FullPath

        On Error GoTo ErrorHandler:

102     oCn.ConnectionString = ConnectionString
104     oCn.Open
106     Set oCmd.ActiveConnection = oCn
108     oCmd.CommandText = SQLString
110     oCmd.CommandType = adCmdText
112     Set oRs = oCmd.Execute
114     oRs.Save FullPath, adPersistXML
116     SaveRSToXML = True

ErrorHandler:
        On Error Resume Next
118     Set oRs = Nothing
120     Set oCmd = Nothing

122     If oCn.State <> 0 Then oCn.Close
124     Set oCn = Nothing
    
        '<EhFooter>
        Exit Function

SaveRSToXML_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.SaveRSToXML " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function TableExists(sName As String, _
                             cn As ADODB.Connection) As Boolean
        '<EhHeader>
        On Error GoTo TableExists_Err
        '</EhHeader>
        On Error GoTo ErrH
        Dim rs As New ADODB.Recordset
    
100     TableExists = True
    
102     rs.Open "SELECT * FROM " & sName, cn
    
        Exit Function
    
ErrH:
104     TableExists = False
        '<EhFooter>
        Exit Function

TableExists_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.TableExists " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function LoadRsFromXML(FullPath As String, _
                              sTableName As String, _
                              Optional sSQLCommand As String, _
                              Optional sAction As String, _
                              Optional sKey As String) As ADODB.Recordset
        '<EhHeader>
        On Error GoTo LoadRsFromXML_Err
        '</EhHeader>

        Dim rs As New ADODB.Recordset
        Dim oRs As New ADODB.Recordset
        Dim cn As New ADODB.Connection
        Dim dProc As Double
        Dim i As Long
        Dim dTotProc As Double
        'On Error Resume Next

100     If Dir(FullPath) = "" Then
102         If sSQLCommand <> "" Then
104             cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\OASISClient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
106             cn.Open
108             cn.Execute sSQLCommand
110             cn.Close
112             Set cn = Nothing
            End If

            Exit Function
114     ElseIf sSQLCommand <> "" Then
116         cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\OASISClient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
118         cn.Open
120         cn.Execute sSQLCommand
122         cn.Close
124         Set cn = Nothing

            Exit Function
        End If
        
126     oRs.CursorLocation = adUseClient
128     oRs.Open FullPath, "Provider=MSPersist;", adOpenForwardOnly, adLockReadOnly, adCmdFile
        
130     If Not oRs.RecordCount = 0 Then
132         dProc = 100 / oRs.RecordCount
        Else
134         dProc = 100
        End If
        
136     elStatusBar.FloodPercent = 0
        
138     cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\OASISClient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    
140     If sSQLCommand = "" Then
142         cn.Open
            
144         If Not TableExists(sTableName, cn) Then
146             CloneTable sTableName, oRs, cn
            End If
            
148         If sAction = "Replace" Then
150             cn.Execute "DELETE * FROM " & sTableName
            End If
            
152         rs.CursorLocation = adUseClient
154         rs.Open "SELECT * FROM " & sTableName, cn, adOpenDynamic, adLockOptimistic
        
156         If Not oRs.BOF Then
158             oRs.MoveFirst
            End If
        
160         If sAction = "Replace" Or sAction = "Append" Then

162             Do While Not oRs.EOF
164                 dTotProc = dTotProc + dProc
166                 elStatusBar.FloodPercent = dTotProc
168                 elStatusBar.Caption = "Importing Data For: " & sTableName & " " & Round(dTotProc, 2) & "%"
170                 elStatusBar.Refresh
    
172                 With rs
174                     .AddNew
    
176                     For i = 0 To .Fields.Count - 1
178                         .Fields.Item(i).Value = oRs.Fields.Item(i).Value
                        Next
    
180                     .Update
                    End With
    
182                 oRs.MoveNext
                Loop

184         ElseIf sAction = "Update" Then

186             If Not oRs.BOF Then
188                 oRs.MoveFirst
                End If
            
190             Do While Not oRs.EOF
192                 dTotProc = dTotProc + dProc
194                 elStatusBar.FloodPercent = dTotProc
196                 elStatusBar.Caption = "Importing Data For: " & sTableName & " " & Round(dTotProc, 2) & "%"
198                 elStatusBar.Refresh
                    
200                 If sKey = "" Then
                        
202                     sKey = oRs.Fields(0).Name
                    End If
                    
                    Dim sPreFix As String
                    
204                 If oRs.Fields(sKey).Type = adChar Then sPreFix = "'"
                    
206                 With rs
208                     .MoveFirst
210                     .Find sKey & " = " & sPreFix & oRs.Fields.Item(sKey).Value & sPreFix

212                     If Not .EOF Then

214                         For i = 0 To .Fields.Count - 1

216                             If .Fields.Item(i).Name <> sKey Then
218                                 .Fields.Item(i).Value = oRs.Fields.Item(i).Value
                                End If

                            Next
                        
220                         .Update
                        End If

                    End With
    
222                 oRs.MoveNext
                Loop
            
            End If
            
        End If
    
224     If Err.Number = 0 Then
226         Set LoadRsFromXML = oRs
        End If
    
        'Clean Up
        On Error Resume Next
228     oRs.Close
230     Set oRs = Nothing
232     rs.Close
234     Set rs = Nothing
236     Kill FullPath
238     cn.Close
240     Set cn = Nothing

        '<EhFooter>
        Exit Function

LoadRsFromXML_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.LoadRsFromXML " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub CloneTable(sTableName As String, _
                       oRSTemplateTable As ADODB.Recordset, _
                       connect As ADODB.Connection)
        '<EhHeader>
        On Error GoTo CloneTable_Err
        '</EhHeader>
        Dim FileDoc As Object
        Dim SQLString As String
        Dim sConn As String
        Dim cat As ADOx.Catalog
        Dim tblTemplate As ADOx.Table
        Dim col As ADODB.Field
        Dim tbl As ADOx.Table
        Dim CurrentProperty As ADODB.Property

100     Set cat = CreateObject("ADOX.Catalog")
102     Set tblTemplate = CreateObject("ADOX.Table")
104     Set tbl = CreateObject("ADOX.Table")

106     Set cat.ActiveConnection = connect

108     tbl.Name = sTableName
110     Set tbl.ParentCatalog = cat

112     With tbl.Columns

114         For Each col In oRSTemplateTable.Fields
116             .Append col.Name, col.Type, col.DefinedSize
118             .Item(col.Name).Properties("Nullable").Value = True
120             .Item(col.Name).Properties("Jet OLEDB:Allow Zero Length").Value = True
            Next

        End With

122     cat.Tables.Append tbl

124     Set tbl = Nothing
126     Set cat = Nothing
128     Set FileDoc = Nothing
    
        '<EhFooter>
        Exit Sub

CloneTable_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.CloneTable " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub c1TabMain_Switch(OldTab As Integer, _
                             NewTab As Integer, _
                             Cancel As Integer)
        '<EhHeader>
        On Error GoTo c1TabMain_Switch_Err
        '</EhHeader>
    
100     If NewTab = 1 Then
102         Cancel = CreateNewINI
        Else

104         Form_Load
        End If
    
        '<EhFooter>
        Exit Sub

c1TabMain_Switch_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.c1TabMain_Switch " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdBrowse_Click()
        '<EhHeader>
        On Error GoTo cmdBrowse_Click_Err
        '</EhHeader>
100     txtFile.Text = OpenDialog(Me, "OASIS Data Import (*.xml)", "OASIS Data Tool Suites", g_sAppPath)
        '<EhFooter>
        Exit Sub

cmdBrowse_Click_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.cmdBrowse_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
        '<EhHeader>
        On Error GoTo SaveDialog_Err
        '</EhHeader>
    
        Dim ofn As OPENFILENAME
        Dim A As Long
100     ofn.lStructSize = Len(ofn)
102     ofn.hwndOwner = Form1.hwnd
104     ofn.hInstance = App.hInstance

106     If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

108     For A = 1 To Len(Filter)

110         If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
        Next

112     ofn.lpstrFilter = Filter
114     ofn.lpstrFile = Space$(254)
116     ofn.nMaxFile = 255
118     ofn.lpstrFileTitle = Space$(254)
120     ofn.nMaxFileTitle = 255
122     ofn.lpstrInitialDir = InitDir
124     ofn.lpstrTitle = Title
126     ofn.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
128     A = GetSaveFileName(ofn)

130     If (A) Then
132         SaveDialog = Trim$(ofn.lpstrFile)
        Else
134         SaveDialog = ""
        End If

        '<EhFooter>
        Exit Function

SaveDialog_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.SaveDialog " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function OpenDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
        '<EhHeader>
        On Error GoTo OpenDialog_Err
        '</EhHeader>
    
        Dim ofn As OPENFILENAME
        Dim A As Long
100     ofn.lStructSize = Len(ofn)
102     ofn.hwndOwner = Form1.hwnd
104     ofn.hInstance = App.hInstance

106     If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"

108     For A = 1 To Len(Filter)

110         If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
        Next

112     ofn.lpstrFilter = Filter
114     ofn.lpstrFile = Space$(254)
116     ofn.nMaxFile = 255
118     ofn.lpstrFileTitle = Space$(254)
120     ofn.nMaxFileTitle = 255
122     ofn.lpstrInitialDir = InitDir
124     ofn.lpstrTitle = Title
126     ofn.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
128     A = GetOpenFileName(ofn)

130     If (A) Then
132         OpenDialog = Trim$(ofn.lpstrFile)
        Else
134         OpenDialog = ""
        End If

        '<EhFooter>
        Exit Function

OpenDialog_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.OpenDialog " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdReset_Click()
        '<EhHeader>
        On Error GoTo cmdReset_Click_Err
        '</EhHeader>
        Dim i As Integer

100     For i = 0 To lstTAbles.ListCount - 1
            'Debug.Print lstTAbles.List(i)
102         lstTAbles.Selected(i) = False
        Next

        '<EhFooter>
        Exit Sub

cmdReset_Click_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.cmdReset_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub cmdUpdate_Click()
        '<EhHeader>
        On Error GoTo cmdUpdate_Click_Err
        '</EhHeader>
        Dim i As Integer
        Dim ofs As New FileSystemObject
        Dim oFile As TextStream
        Dim sSQL As String
        Dim sAction As String
        Dim sKey As String
        
100     cmdUpdate.Enabled = False
    
102     If chkImportClientDBSQLScripts.Value = vbChecked Then
104         oINIReader.Section = "Default"
106         oINIReader.Key = "ClientDbUpdate"
108         Set oFile = ofs.OpenTextFile(g_sAppPath & "\" & oINIReader.Value)
110         sSQL = oFile.ReadAll
112         LoadRsFromXML "", "", sSQL
114         Set oFile = Nothing
116         ofs.DeleteFile g_sAppPath & "\" & oINIReader.Value, True
        End If
        
118     If Not gbHideGUI Then
120         If MsgBox("Are you sure you want to update?", vbYesNo, "OASIS Data Importer") = vbYes Then gbHideGUI = True
        End If
        
122     If gbHideGUI Then
    
124         For i = 0 To lstImportTables.ListCount - 1
126             Debug.Print lstImportTables.List(i)
            
128             If lstImportTables.Selected(i) Then
130                 oINIReader.Section = lstImportTables.List(i)
132                 oINIReader.Key = "Action"
134                 sAction = oINIReader.Value
136                 oINIReader.Key = "PKey"
138                 sKey = oINIReader.Value

140                 oINIReader.Key = "ImportFileName"
142                 Debug.Print oINIReader.Value
144                 LoadRsFromXML g_sAppPath & "\" & oINIReader.Value, lstImportTables.List(i), , sAction, sKey
                End If

            Next
                        
            On Error Resume Next
146         Kill m_StrIniFile
148         lstImportTables.Clear
150         txtImportDescription.Text = ""
152         elStatusBar.FloodPercent = 0
154         elStatusBar.Caption = "Import Process 100% Done!"
156         chkImportClientDBSQLScripts.Value = vbUnchecked
158         chkImportClientDBSQLScripts.Enabled = False
            
        End If
    
160     cmdUpdate.Enabled = True
        '<EhFooter>
        Exit Sub

cmdUpdate_Click_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.cmdUpdate_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateINI(sFileName As String)
    
End Sub

Public Function ArrayInit(ByVal NotValue As Long) As Boolean
        '<EhHeader>
        On Error GoTo ArrayInit_Err
        '</EhHeader>

100     ArrayInit = Not (NotValue = -1&)

102     If App.LogMode <> 0 Then Exit Function

        On Error Resume Next
104     Debug.Assert 0.1
        On Error GoTo ArrayInit_Err
        '<EhFooter>
        Exit Function

ArrayInit_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.ArrayInit " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub InitializeArrays()
        '<EhHeader>
        On Error GoTo InitializeArrays_Err
        '</EhHeader>
100     ReDim m_strTables(0)
102     ReDim m_strUIDs(0)
104     ReDim m_strImportFileNames(0)
106     ReDim m_strSQL(0)
108     ReDim m_strTimeStamp(0)
110     ReDim m_strAction(0)
112     ReDim m_strKey(0)
        '<EhFooter>
        Exit Sub

InitializeArrays_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.InitializeArrays " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
        Dim strXportName As String
        Dim i As Integer
        Dim bAnySelected As Boolean
        Dim dProc As Double
        Dim dTotProc As Double
        Dim ofs As New FileSystemObject
        Dim oFile As TextStream
        Dim strSQLScriptName As String
    
100     Command1.Enabled = False
        
102     gridTables.SetFocus
104     gridTables.Dataset.FindLast
106     gridTables.Dataset.FindFirst
        
108     oRSFields.Filter = adFilterPendingRecords
        
110     If Not oRSFields.BOF Then
112         oRSFields.MoveFirst
        End If
        
114     elStatusBar.Caption = "Exporting..."
        
116     If oRSFields.RecordCount > 0 Then
118         dProc = 100 / oRSFields.RecordCount
        Else
120         dProc = 100
        End If
        
122     InitializeArrays
        
124     If Not oRSFields.EOF And Not oRSFields.BOF Then
        
126         With oRSFields
        
128             Do While Not .EOF

130                 dTotProc = dTotProc + dProc
132                 elStatusBar.FloodPercent = dTotProc
134                 elStatusBar.Caption = "Exporting... " & Round(dTotProc, 2) & "%"
136                 elStatusBar.Refresh

138                 With .Fields
                 
140                     If (Not m_strTables(0) = "") Then
142                         ReDim Preserve m_strTables(UBound(m_strTables) + 1)
144                         ReDim Preserve m_strUIDs(UBound(m_strUIDs) + 1)
146                         ReDim Preserve m_strImportFileNames(UBound(m_strImportFileNames) + 1)
148                         ReDim Preserve m_strSQL(UBound(m_strSQL) + 1)
150                         ReDim Preserve m_strTimeStamp(UBound(m_strTimeStamp) + 1)
152                         ReDim Preserve m_strAction(UBound(m_strAction) + 1)
154                         ReDim Preserve m_strKey(UBound(m_strKey) + 1)
                     
                        End If
                
156                     m_strTables(UBound(m_strTables)) = Trim(.Item(1).Value)
158                     m_strUIDs(UBound(m_strUIDs)) = GUIDGen
160                     m_strImportFileNames(UBound(m_strImportFileNames)) = Trim$(.Item(1).Value) & ".xml"
162                     m_strSQL(UBound(m_strSQL)) = ""
164                     m_strTimeStamp(UBound(m_strTimeStamp)) = Now
166                     m_strAction(UBound(m_strAction)) = Trim(.Item(2).Value)
168                     m_strKey(UBound(m_strKey)) = Trim(.Item(3).Value)
                
170                     SaveRSToXML "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\OASISClient.mdb;Mode=ReadWrite|Share Deny None;Persist Security Info=False", "SELECT * FROM " & Trim(.Item(1).Value), g_sAppPath & "\" & Trim(.Item(1).Value) & ".xml"

                    End With
            
172                 .MoveNext
            
                Loop
        
            End With

        End If
        
174     With oINIReader

176         For i = LBound(m_strUIDs) To UBound(m_strUIDs)
        
178             .Section = m_strTables(i)
180             .Key = "ImportFileName"
182             .Value = m_strImportFileNames(i)
184             .AddNewSection
            
186             .Key = "UID"
188             .Value = m_strUIDs(i)
190             .AddKeyWithValue
            
192             .Key = "TimeStamp"
194             .Value = m_strTimeStamp(i)
196             .AddKeyWithValue
            
198             .Key = "Action"
200             .Value = m_strAction(i)
202             .AddKeyWithValue

204             .Key = "PKey"
206             .Value = m_strKey(i)
208             .AddKeyWithValue

210             .Key = "ExeSQL"
212             .Value = m_strSQL(i)
214             .AddKeyWithValue
            
216             .Section = "Default"
218             .Key = "ImportTables"
                
220             If i = LBound(m_strUIDs) Then
222                 .Value = m_strTables(i)
                Else
224                 .Value = .Value & "," & m_strTables(i)
                End If

226             .AddKeyWithValue
            Next
    
228         .Section = "Default"
230         .Key = "ClientDbUpdate"
232         .Value = ""
            
234         If chkUseClient.Value = vbChecked Then
236             strSQLScriptName = InputBox("Add a name for the SQL script", "OASIS Data Tools", "OASISSQLScript")
        
238             If ofs.FileExists(g_sAppPath & "\" & strSQLScriptName & ".sql") Then
240                 ofs.DeleteFile g_sAppPath & "\" & strSQLScriptName & ".sql", True
                End If
        
242             Set oFile = ofs.CreateTextFile(g_sAppPath & "\" & strSQLScriptName & ".sql", True)
244             oFile.Write txtExportSQL.Text
            
246             .Value = strSQLScriptName & ".sql"

248             oFile.Close
            End If
    
250         .AddKeyWithValue
    
        End With
     
252     Set oFile = Nothing
254     Set ofs = Nothing
256     elStatusBar.Caption = "Export Done..."
258     Command1.Enabled = True
        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.Command1_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Command2_Click()
        '<EhHeader>
        On Error GoTo Command2_Click_Err
        '</EhHeader>

100     With oINIReader
102         .Section = "Default"
104         .Key = "ImportTables"
106         .Value = ""
108         .AddKeyWithValue
        End With

        '<EhFooter>
        Exit Sub

Command2_Click_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.Command2_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim fs As New FileSystemObject
        Dim oDB As New ADOx.Catalog
        Dim oTable As ADOx.Table
        Dim i As Integer
        
100     Set oRSFields = New ADODB.Recordset
        
102     oRSFields.Fields.Append "fldUse", adBoolean
104     oRSFields.Fields.Append "fldTable", adChar, 255
106     oRSFields.Fields.Append "fldAction", adChar, 128
108     oRSFields.Fields.Append "fldKey", adChar, 128
110     oRSFields.Open

112     oDB.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & g_sAppPath & "\data\db\OASISClient.mdb"
114     lstTAbles.Clear

116     For Each oTable In oDB.Tables

118         If oTable.Type = "TABLE" Then
120             oRSFields.AddNew
122             oRSFields.Fields.Item("fldUse").Value = False
124             oRSFields.Fields.Item("fldTable").Value = oTable.Name
126             oRSFields.Fields.Item("fldAction").Value = "Append"
128             oRSFields.Fields.Item("fldKey").Value = ""
130             oRSFields.UpdateBatch
132             i = i + 1
134             Debug.Print i
136             lstTAbles.AddItem oTable.Name
            End If

        Next

138     Set gridTables.DataSource = oRSFields
140     gridTables.Dataset.Active = True
        'oRSFields.Update
142     m_StrIniFile = g_sAppPath & "\settings.ini"

144     oINIReader.Path = m_StrIniFile
    
146     ReadINI

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function CreateNewINI() As Integer
        '<EhHeader>
        On Error GoTo CreateNewINI_Err
        '</EhHeader>
        Dim fs As New FileSystemObject

100     CreateNewINI = 1
    
102     If fs.FileExists(m_StrIniFile) Then
104         If MsgBox("This will reset & delete previous exports! Do you want to continue?", vbYesNo, "OASIS Data Export") = vbNo Then Exit Function
106         fs.DeleteFile m_StrIniFile, True
        End If

108     fs.CreateTextFile m_StrIniFile
110     Set fs = Nothing

112     With oINIReader
114         .Section = "Default"
116         .Key = "ImportTables"
118         .Value = ""
120         .AddNewSection
        End With
    
122     CreateNewINI = 0
    
        '<EhFooter>
        Exit Function

CreateNewINI_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.CreateNewINI " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub ReadINI()
        '<EhHeader>
        On Error GoTo ReadINI_Err
        '</EhHeader>
        Dim i As Integer
        Dim sIMPTables() As String
        Dim sDesc As String
    
100     txtImportDescription.Text = "----------- IMPORT DETAILS -------------" & vbCrLf
102     lstImportTables.Clear
    
104     With oINIReader
106         .Section = "Default"
108         .Key = "ImportTables"
110         sIMPTables = Split(.Value, ",")
    
112         .Section = "Default"
114         .Key = "ClientDbUpdate"
        
116         If Not .Value = "" Then
118             chkImportClientDBSQLScripts.Enabled = True
120             chkImportClientDBSQLScripts.Value = vbChecked
            Else
122             chkImportClientDBSQLScripts.Enabled = False
124             chkImportClientDBSQLScripts.Value = vbUnchecked
            End If
    
126         If Not UBound(sIMPTables) = LBound(sIMPTables) Then

128             For i = LBound(sIMPTables) To UBound(sIMPTables)
130                 lstImportTables.AddItem sIMPTables(i)
132                 lstImportTables.Selected(i) = True
134                 sDesc = ""
136                 .Section = sIMPTables(i)
138                 .Key = "TimeStamp"
140                 sDesc = sDesc & vbCrLf & "************ DETAILS *************" & vbCrLf
142                 sDesc = sDesc & "Import TABLE:" & sIMPTables(i) & vbCrLf
144                 sDesc = sDesc & "Time of Export:" & .Value & vbCrLf
146                 .Key = "Action"
148                 sDesc = sDesc & "Action:" & .Value & vbCrLf
                    
150                 .Key = "PKey"
152                 sDesc = sDesc & "Key:" & .Value & vbCrLf

154                 .Key = "ImportFileName"
156                 sDesc = sDesc & "File Name:" & .Value
                                        
158                 sDesc = sDesc & vbCrLf & "************ END *************" & vbCrLf
160                 txtImportDescription.Text = txtImportDescription.Text & vbCrLf & sDesc
                Next

            Else
162             lstImportTables.AddItem sIMPTables(0)
164             lstImportTables.Selected(0) = True
166             sDesc = ""
168             .Section = sIMPTables(0)
170             .Key = "TimeStamp"
172             sDesc = sDesc & vbCrLf & "************ DETAILS *************" & vbCrLf
174             sDesc = sDesc & "Import TABLE:" & sIMPTables(0) & vbCrLf
176             sDesc = sDesc & "Time of Export:" & .Value & vbCrLf
178             .Key = "Action"
180             sDesc = sDesc & "Action:" & .Value & vbCrLf
                    
182             .Key = "PKey"
184             sDesc = sDesc & "Key:" & .Value & vbCrLf

186             .Key = "ImportFileName"
188             sDesc = sDesc & "File Name:" & .Value
190             sDesc = sDesc & vbCrLf & "************ END *************" & vbCrLf
192             txtImportDescription.Text = txtImportDescription.Text & vbCrLf & sDesc
            End If

        End With

        '<EhFooter>
        Exit Sub

ReadINI_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.ReadINI " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub OptExpType_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo OptExpType_Click_Err
        '</EhHeader>
        Dim i As Integer

100     If Index = 1 Then

102         For i = 0 To lstTAbles.ListCount - 1
104             Debug.Print lstTAbles.List(i)
106             lstTAbles.Selected(i) = True
            Next

        End If

        '<EhFooter>
        Exit Sub

OptExpType_Click_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.OptExpType_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function GUIDGen() As String
        '<EhHeader>
        On Error GoTo GUIDGen_Err
        '</EhHeader>
        Dim uGUID As GUID
        Dim sGUID As String
        Dim bGUID() As Byte
        Dim lLen As Long
        Dim retval As Long
100     lLen = 40
102     bGUID = String(lLen, 0)
104     CoCreateGuid uGUID
    
106     retval = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
108     sGUID = bGUID

110     If (Asc(Mid$(sGUID, retval, 1)) = 0) Then retval = retval - 1
112     GUIDGen = Left$(sGUID, retval)
        '<EhFooter>
        Exit Function

GUIDGen_Err:
        MsgBox Err.Description & vbCrLf & "in oasis_DataUpdate.frmDataToolMain.GUIDGen " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
