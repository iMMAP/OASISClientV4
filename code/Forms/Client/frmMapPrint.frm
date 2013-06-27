VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{9C989D2F-3596-477B-B719-6DC4E6893AF2}#1.0#0"; "TatukGIS_XDK10_WIN32.ocx"
Begin VB.Form frmMapPrint 
   Caption         =   "OASIS - Print Toolbox"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   11055
   Icon            =   "frmMapPrint.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5025
   ScaleWidth      =   11055
   StartUpPosition =   3  'Windows Default
   Begin TatukGIS_XDK10.XGIS_ControlPrintPreviewSimple SimplePrintPreview 
      Left            =   7020
      Top             =   720
      Caption         =   "Print Preview"
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   640
      WindowHeight    =   480
      DoubleBuffered  =   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   5025
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11055
      _cx             =   19500
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
      Align           =   5
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
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmMapPrint.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin TatukGIS_XDK10.XGIS_ControlPrintPreview PrintPreview 
         Height          =   4845
         Left            =   90
         TabIndex        =   29
         Top             =   90
         Width           =   6525
         Align           =   0
         BevelInner      =   0
         BevelOuter      =   0
         BorderStyle     =   0
         Ctl3D           =   -1  'True
         Color           =   8421504
         Enabled         =   -1  'True
         Object.Visible         =   -1  'True
         DoubleBuffered  =   0   'False
         BevelWidth      =   1
         BorderWidth     =   0
         HelpContextId   =   0
         TabOrder        =   -1
         TabStop         =   0   'False
      End
      Begin C1SizerLibCtl.C1Tab C1TTabPrint 
         Height          =   4845
         Left            =   6675
         TabIndex        =   7
         Top             =   90
         Width           =   4290
         _cx             =   7567
         _cy             =   8546
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
         Caption         =   "Available Templates|Settings"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   0
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   -1  'True
         TabsPerPage     =   2
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
         Begin VB.Frame FraPrntSettngs 
            BorderStyle     =   0  'None
            Height          =   4470
            Left            =   4935
            TabIndex        =   12
            Top             =   330
            Width           =   4200
            Begin VB.CommandButton cmdPrint 
               Caption         =   "Print"
               Height          =   345
               Left            =   2670
               Picture         =   "frmMapPrint.frx":6891
               TabIndex        =   28
               Top             =   4080
               Width           =   1485
            End
            Begin VB.CheckBox chkShowPrinter 
               Caption         =   "Show Printer Settings"
               Height          =   240
               Left            =   90
               TabIndex        =   27
               Top             =   3690
               Width           =   2085
            End
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Height          =   360
               Left            =   2670
               TabIndex        =   26
               Top             =   3645
               Width           =   1485
            End
            Begin XpressEditorsLibCtl.dxDateEdit dxDateEdit1 
               Height          =   315
               Left            =   2745
               OleObjectBlob   =   "frmMapPrint.frx":D0E3
               TabIndex        =   22
               Top             =   3150
               Width           =   1365
            End
            Begin VB.TextBox txtTitle 
               Height          =   330
               Left            =   1080
               TabIndex        =   17
               Text            =   "IRAQ"
               Top             =   45
               Width           =   3030
            End
            Begin VB.TextBox txtCopyRight 
               Height          =   1005
               Left            =   1080
               MultiLine       =   -1  'True
               TabIndex        =   16
               Text            =   "frmMapPrint.frx":D183
               Top             =   765
               Width           =   3030
            End
            Begin VB.TextBox txtSubTitle 
               Height          =   330
               Left            =   1080
               TabIndex        =   15
               Text            =   "Map title"
               Top             =   405
               Width           =   3030
            End
            Begin VB.TextBox txtVeiwer 
               Height          =   1230
               Left            =   1080
               MultiLine       =   -1  'True
               TabIndex        =   14
               Text            =   "frmMapPrint.frx":D1B6
               Top             =   1845
               Width           =   3030
            End
            Begin VB.TextBox txtMapID 
               Height          =   330
               Left            =   1080
               TabIndex        =   13
               Text            =   "MapID"
               Top             =   3150
               Width           =   960
            End
            Begin VB.Label lblDate 
               Caption         =   "Date:"
               Height          =   285
               Left            =   2115
               TabIndex        =   24
               Top             =   3240
               Width           =   510
            End
            Begin VB.Label lblMapID 
               Caption         =   "Map ID:"
               Height          =   285
               Left            =   0
               TabIndex        =   23
               Top             =   3240
               Width           =   870
            End
            Begin VB.Label lblNotes 
               Caption         =   "Notes:"
               Height          =   285
               Left            =   45
               TabIndex        =   21
               Top             =   1845
               Width           =   870
            End
            Begin VB.Label lblCopyright 
               Caption         =   "Copyright:"
               Height          =   420
               Left            =   0
               TabIndex        =   20
               Top             =   765
               Width           =   915
            End
            Begin VB.Label lblMapSub 
               Caption         =   "Map Sub Title:"
               Height          =   195
               Left            =   0
               TabIndex        =   19
               Top             =   450
               Width           =   1050
            End
            Begin VB.Label lblMapTitle 
               Caption         =   "Map Title:"
               Height          =   240
               Left            =   0
               TabIndex        =   18
               Top             =   45
               Width           =   735
            End
         End
         Begin VB.Frame FraAvailablePrint 
            BorderStyle     =   0  'None
            Height          =   4470
            Left            =   45
            TabIndex        =   8
            Top             =   330
            Width           =   4200
            Begin VB.TextBox txtTplDesc 
               Height          =   3300
               Left            =   135
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   10
               Top             =   810
               Width           =   3885
            End
            Begin VB.ComboBox ComTplPrint 
               Height          =   315
               Left            =   1395
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   90
               Width           =   2625
            End
            Begin VB.Label lblTemplateDescription 
               Caption         =   "Template Description:"
               Height          =   240
               Left            =   135
               TabIndex        =   25
               Top             =   585
               Width           =   1950
            End
            Begin VB.Label lblSelectTemplate 
               Caption         =   "Select Template:"
               Height          =   285
               Left            =   135
               TabIndex        =   11
               Top             =   180
               Width           =   1275
            End
         End
      End
      Begin VB.TextBox edPrintFooter 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4845
         Left            =   6675
         TabIndex        =   5
         Text            =   "www.immap.org"
         Top             =   90
         Width           =   4290
      End
      Begin VB.TextBox edPrintTitle 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4845
         Left            =   6675
         TabIndex        =   2
         Text            =   "OASIS Map Print"
         Top             =   90
         Width           =   4290
      End
      Begin VB.TextBox edPrintSubTitle 
         Appearance      =   0  'Flat
         Height          =   4845
         Left            =   6675
         TabIndex        =   1
         Text            =   "Oasis Map"
         Top             =   90
         Width           =   4290
      End
      Begin VB.Label lblPrintFooter 
         AutoSize        =   -1  'True
         Caption         =   "Print Footer:"
         Height          =   4845
         Left            =   6675
         TabIndex        =   6
         Top             =   90
         Width           =   4290
      End
      Begin VB.Label Label1 
         Caption         =   "Print title:"
         Height          =   4845
         Left            =   6675
         TabIndex        =   4
         Top             =   90
         Width           =   4290
      End
      Begin VB.Label Label2 
         Caption         =   "Print subtitle:"
         Height          =   4845
         Left            =   6675
         TabIndex        =   3
         Top             =   90
         Width           =   4290
      End
   End
End
Attribute VB_Name = "frmMapPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_rsPrintTpl As ADODB.Recordset
Private m_Legend As TatukGIS_XDK10.XGIS_Legend
Private m_GIS As TatukGIS_XDK10.XGIS_Viewer

Private Sub GetUpdates(strUserGroupPrefix As String)
        '<EhHeader>
        On Error GoTo GetUpdates_Err
        '</EhHeader>
        Dim rsRemote As ADODB.Recordset
        Dim RS As ADODB.Recordset
        Dim j As Integer
        Dim sString As String
        'Dim RSUpdater As ADODB.Recordset
                    
        'Now Check the Dynamic Content version
100     If Not strUserGroupPrefix = "" Then

102         'sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM " & strUserGroupPrefix & "PrintTemplates")
104         'Set rsRemote = OpenSilentHttpCommsRS(sString, True)
            Set rsRemote = OpenServerRSCompressed(g_sAppServerPath & "/oasis4.asp", "id", "SELECT * FROM " & strUserGroupPrefix & "PrintTemplates")
            
106         If Not rsRemote.State = 0 Then

108             m_Cnn.Execute "delete from PrintTemplates"
                
110             If rsRemote.EOF And rsRemote.Bof Then
                    Exit Sub
                End If
            
112             Set RS = New ADODB.Recordset
                
114             SafeMoveFirst rsRemote
    
116             RS.Open "SELECT * FROM PrintTemplates", m_Cnn, adOpenDynamic, adLockOptimistic
    
118             Do While Not rsRemote.EOF
120                 RS.AddNew
    
122                 For j = 1 To rsRemote.Fields.Count - 2
124                     'RS.Fields.Item(j).Value = rsRemote.Fields.Item(j).Value
                        RS.Fields.Item(j).Value = rsRemote.Fields(RS.Fields.Item(j).Name).Value
                    Next
    
126                 SaveFileToDisk g_sAppPath, rsRemote
128                 rsRemote.MoveNext
                Loop
                
130             RS.UpdateBatch
132             rsRemote.Close
134             RS.Close
    
'136             Set rsRemote = Nothing
'138             Set RS = Nothing
'
'140             sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT SettingValue9 FROM " & strUserGroupPrefix & "AppSettings WHERE SettingName = 'ProfileSettings'")
'142             Set rsRemote = OpenSilentHttpCommsRS(sString, True)
'
'144             If Not rsRemote.State = 0 Then
'
'146                 Set RSUpdater = New ADODB.Recordset
'148                 With RSUpdater
'
'150                     .Open "SELECT * FROM AppSettings", m_Cnn, adOpenDynamic, adLockBatchOptimistic
'152                     .Find "SettingName = 'ProfileSettings'"
'
'154                     If Not .EOF Then
'
'156                         If Not IsNull(rsRemote.Fields.Item("SettingValue9").Value) Then
'158                             .Fields("SettingValue9").Value = rsRemote.Fields.Item("SettingValue9").Value
'                            Else
'160                             .Fields("SettingValue9").Value = 1
'                            End If
'
'162                         .UpdateBatch adAffectCurrent
'164                         .Close
'                        End If
'
'                    End With
'166                 Set RSUpdater = Nothing
'
'168                 rsRemote.Close
'170                 Set rsRemote = Nothing

                    SynchProfileSettingWithServer "SettingValue9", strUserGroupPrefix, m_Cnn

                End If
                
            End If
                
        

        '<EhFooter>
        Exit Sub

GetUpdates_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.GetUpdates " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Init(oViewer As TatukGIS_XDK10.XGIS_Viewer, _
                oLgd As TatukGIS_XDK10.XGIS_Legend, _
                bMapTplUpdate As Boolean, _
                strUserGroupPrefix As String)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>

100     If bMapTplUpdate Then GetUpdates strUserGroupPrefix

102     Set m_GIS = oViewer
104     Set m_Legend = oLgd

106     Set m_rsPrintTpl = New ADODB.Recordset

108     m_rsPrintTpl.Open "SELECT * FROM PrintTemplates", m_Cnn, adOpenDynamic, adLockBatchOptimistic

110     ComTplPrint.Clear
    
112     SafeMoveFirst m_rsPrintTpl
    
114     ComTplPrint.AddItem "--No Template--"
    
116     Do While Not m_rsPrintTpl.EOF
118         ComTplPrint.AddItem m_rsPrintTpl.Fields.Item("Name").Value
120         m_rsPrintTpl.MoveNext
        Loop
    
122     txtTplDesc.Text = ""
124     txtTitle.Text = ""
126     txtSubTitle.Text = ""
128     txtCopyRight.Text = ""
130     txtVeiwer.Text = ""
132     txtMapID.Text = ""

134     PrintPreview.GIS_Viewer = oViewer
        SimplePrintPreview.GIS_Viewer = oViewer
136     'PrintPreviewSimple.GIS_Viewer = oViewer
138     oViewer.PrintTitle = edPrintTitle.Text
140     oViewer.PrintSubTitle = edPrintSubTitle.Text
142     oViewer.PrintFooter = edPrintFooter.Text
        'PrintPreview.Preview (1)
    
144     If ComTplPrint.ListCount >= 0 Then ComTplPrint.ListIndex = 0
    
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.Init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdPrint_Click()
        '<EhHeader>
        On Error GoTo cmdPrint_Click_Err
        '</EhHeader>
100     'PrintPreviewSimple.Preview
        SimplePrintPreview.Preview
102     Me.Hide
        '<EhFooter>
        Exit Sub

cmdPrint_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.cmdPrint_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdUpdate_Click()
        '<EhHeader>
        On Error GoTo cmdUpdate_Click_Err
        '</EhHeader>
100     ApplyPrintTpl True, IIf(chkShowPrinter.Value = vbChecked, True, False)
        '<EhFooter>
        Exit Sub

cmdUpdate_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.cmdUpdate_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComTplPrint_Click()
        '<EhHeader>
        On Error GoTo ComTplPrint_Click_Err
        '</EhHeader>
100     ApplyPrintTpl
        '<EhFooter>
        Exit Sub

ComTplPrint_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.ComTplPrint_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ApplyPrintTpl(Optional bUseIgnoreDbSettings As Boolean = False, _
                          Optional bShowPrinterSettings As Boolean = False)
        '<EhHeader>
        On Error GoTo ApplyPrintTpl_Err
        '</EhHeader>
        Dim tmp As TatukGIS_XDK10.XGIS_TemplatePrint

100     If ComTplPrint.ListIndex = 0 Then
102         PrintPreview.Preview (1)
        Else
        
104         SafeMoveFirst m_rsPrintTpl
106         m_rsPrintTpl.Find "Name = '" & ComTplPrint.List(ComTplPrint.ListIndex) & "'"
            
108         If Not bUseIgnoreDbSettings Then
110             txtTplDesc.Text = m_rsPrintTpl.Fields.Item("description").Value
112             txtTitle.Text = m_rsPrintTpl.Fields.Item("MapTitle").Value
114             txtSubTitle.Text = m_rsPrintTpl.Fields.Item("MapSubTitle").Value
116             txtCopyRight.Text = m_rsPrintTpl.Fields.Item("copyright").Value
118             txtVeiwer.Text = m_rsPrintTpl.Fields.Item("note").Value
120             txtMapID.Text = m_rsPrintTpl.Fields.Item("MapIDPrefix").Value
            End If
        
122         Set tmp = New TatukGIS_XDK10.XGIS_TemplatePrint
    
124         tmp.Create_ m_GIS
126         tmp.Path = g_sAppPath & "\data\gis\other\Graphics\" & m_rsPrintTpl.Fields.Item("FileName").Value
128         tmp.GIS_Legend(1) = m_Legend
    
            'tmp.GIS_Scale(1) = ControlScale1.Scale
130         tmp.GIS_ViewerExtent(1) = m_GIS.VisibleExtent
132         tmp.Text(1) = txtTitle.Text
134         tmp.Text(2) = txtCopyRight.Text
136         tmp.Text(3) = txtSubTitle.Text
138         tmp.Text(4) = txtVeiwer.Text
140         m_GIS.PrintTitle = edPrintTitle.Text
142         m_GIS.PrintSubTitle = edPrintSubTitle.Text
144         m_GIS.PrintFooter = edPrintFooter.Text
            'Printer.Orientation = vbPRORPortrait

146         If bShowPrinterSettings Then SimplePrintPreview.PrinterSetup ' PrintPreviewSimple.PrinterSetup
148         PrintPreview.Preview (1)
        
        End If

        '<EhFooter>
        Exit Sub

ApplyPrintTpl_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.ApplyPrintTpl " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function SaveFileToDisk(sPath As String, _
                                RS As ADODB.Recordset) As String
        '<EhHeader>
        On Error GoTo SaveFileToDisk_Err
        '</EhHeader>
        Dim FileStream As New ADODB.Stream
        Dim Filename As String
                
100     If Not RS.EOF Then
102         If Not IsNull(RS.Fields("blob_tpl").Value) Then
104             Filename = sPath & "\data\gis\other\Graphics\" & RS.Fields.Item("FileName").Value
106             FileStream.Type = adTypeBinary
108             FileStream.Open
110             FileStream.Write RS.Fields("blob_tpl").Value
112             FileStream.SaveToFile Filename, adSaveCreateOverWrite
            End If
        End If
    
114     SaveFileToDisk = Filename

        '<EhFooter>
        Exit Function

SaveFileToDisk_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.SaveFileToDisk " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub edPrintFooter_Change()
        '<EhHeader>
        On Error GoTo edPrintFooter_Change_Err
        '</EhHeader>
100     PrintPreview.GIS_Viewer.PrintFooter = edPrintFooter.Text
        '<EhFooter>
        Exit Sub

edPrintFooter_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.edPrintFooter_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub edPrintSubTitle_Change()
        '<EhHeader>
        On Error GoTo edPrintSubTitle_Change_Err
        '</EhHeader>
  
100     PrintPreview.GIS_Viewer.PrintSubTitle = edPrintSubTitle.Text
        '<EhFooter>
        Exit Sub

edPrintSubTitle_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.edPrintSubTitle_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub edPrintTitle_Change()
        '<EhHeader>
        On Error GoTo edPrintTitle_Change_Err
        '</EhHeader>
  
100     PrintPreview.GIS_Viewer.PrintTitle = edPrintTitle.Text
        '<EhFooter>
        Exit Sub

edPrintTitle_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMapPrint.edPrintTitle_Change " & _
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
               "in OASISClient.frmMapPrint.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

