VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{F8F9FBF9-12B5-11D4-8ED3-00E07D815373}#1.0#0"; "MBScroll.ocx"
Begin VB.Form frmDynamicDataAddEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding and Editing Dynamic Data"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8490
   DrawMode        =   1  'Blackness
   FillColor       =   &H00C0FFC0&
   FillStyle       =   5  'Downward Diagonal
   Icon            =   "frmDynamicDataAddEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleMode       =   0  'User
   ScaleWidth      =   8490
   StartUpPosition =   2  'CenterScreen
   Begin MBScroller.Scroller frmDynam2 
      Height          =   915
      Index           =   0
      Left            =   7350
      TabIndex        =   34
      Top             =   7080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1614
      BorderStyle     =   0
      BackColor       =   5292196
      ScrollBars      =   2
      ScrollBarsColor =   5292196
   End
   Begin MBScroller.Scroller frmDynam 
      Height          =   1095
      Index           =   0
      Left            =   6840
      TabIndex        =   33
      Top             =   6540
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   1931
      BorderStyle     =   0
      BackColor       =   5292196
      ScrollBars      =   2
      ScrollBarsColor =   5292196
   End
   Begin VB.Frame frameSpatial 
      BackColor       =   &H0050C0A4&
      Caption         =   "Spatial Reference:"
      ForeColor       =   &H00FFFFFF&
      Height          =   2205
      Left            =   90
      TabIndex        =   27
      Top             =   150
      Width           =   2775
      Begin VB.CommandButton cmdPickLocation 
         Caption         =   "Pick Location"
         Height          =   315
         Left            =   450
         TabIndex        =   30
         Top             =   1650
         Width           =   1905
      End
      Begin VB.TextBox txtY 
         Height          =   315
         Left            =   330
         TabIndex        =   29
         Text            =   "Text2"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox txtX 
         Height          =   315
         Left            =   300
         TabIndex        =   28
         Text            =   "Text1"
         Top             =   510
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H0050C0A4&
         BackStyle       =   0  'Transparent
         Caption         =   "Latitude"
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   870
         Width           =   2235
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Longitude"
         ForeColor       =   &H00C0FFFF&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   300
         Width           =   2235
      End
   End
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      Caption         =   "Cancel"
      Height          =   285
      Left            =   7440
      TabIndex        =   26
      Top             =   5520
      Width           =   1005
   End
   Begin C1SizerLibCtl.C1Tab C1Tab 
      Height          =   1965
      Left            =   330
      TabIndex        =   0
      Top             =   6150
      Width           =   6585
      _cx             =   11615
      _cy             =   3466
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
      Caption         =   "Table|New Tab|New Tab|New Tab|New Tab"
      Align           =   0
      CurrTab         =   2
      FirstTab        =   0
      Style           =   0
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   0   'False
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   0   'False
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
      Begin VB.Frame Frame 
         Height          =   1590
         Index           =   2
         Left            =   -7440
         TabIndex        =   8
         Top             =   330
         Width           =   6495
      End
      Begin VB.Frame Frame 
         Height          =   1590
         Index           =   0
         Left            =   45
         TabIndex        =   2
         Top             =   330
         Width           =   6495
         Begin VB.TextBox txt1 
            BackColor       =   &H0050C0A4&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   3960
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   1080
            Width           =   495
         End
         Begin VB.CheckBox chk1 
            BackColor       =   &H0050C0A4&
            Height          =   375
            Index           =   0
            Left            =   5160
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
         Begin XpressEditorsLibCtl.dxDateEdit date1 
            Height          =   315
            Index           =   0
            Left            =   960
            OleObjectBlob   =   "frmDynamicDataAddEdit.frx":6852
            TabIndex        =   3
            Top             =   960
            Width           =   1575
         End
         Begin XpressEditorsLibCtl.dxLookUpEdit cmb1 
            Height          =   315
            Index           =   0
            Left            =   1320
            OleObjectBlob   =   "frmDynamicDataAddEdit.frx":69CA
            TabIndex        =   4
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label lbl1 
            BackColor       =   &H00C0FFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "Label1"
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   7
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   1590
         Index           =   1
         Left            =   -7140
         TabIndex        =   1
         Top             =   330
         Width           =   6495
         Begin VB.Frame frameOperation 
            Caption         =   "Select Operation"
            Height          =   2055
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   2895
            Begin VB.OptionButton optOperation 
               Caption         =   "Edit Data Definitions"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   16
               Top             =   1440
               Width           =   2175
            End
            Begin VB.OptionButton optOperation 
               Caption         =   "Edit Record"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   15
               Top             =   480
               Value           =   -1  'True
               Width           =   2175
            End
            Begin VB.OptionButton optOperation 
               Caption         =   "New Record"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   14
               Top             =   960
               Width           =   2175
            End
         End
         Begin VB.Frame frameDynamData 
            Height          =   4270
            Left            =   6960
            TabIndex        =   9
            Top             =   300
            Width           =   5415
            Begin VB.Frame frameDynam 
               ForeColor       =   &H00FF0000&
               Height          =   4305
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   5175
            End
            Begin VB.VScrollBar VScroll1 
               Height          =   4170
               LargeChange     =   1000
               Left            =   5160
               Max             =   100
               SmallChange     =   100
               TabIndex        =   10
               Top             =   90
               Width           =   255
            End
         End
         Begin VB.Label lblActiveTable 
            Alignment       =   2  'Center
            Caption         =   "Active Table: "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   7200
            TabIndex        =   12
            Top             =   100
            Width           =   4815
         End
      End
   End
   Begin VB.CommandButton cmdCommit 
      Appearance      =   0  'Flat
      Caption         =   "Add"
      Height          =   285
      Left            =   6360
      TabIndex        =   25
      Top             =   5520
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   285
      Left            =   4200
      TabIndex        =   23
      Top             =   5520
      Width           =   1005
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   285
      Left            =   5280
      TabIndex        =   22
      Top             =   5520
      Width           =   1005
   End
   Begin VB.Frame frmDynam2OLD 
      BackColor       =   &H0050C0A4&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FF0000&
      Height          =   990
      Index           =   0
      Left            =   3450
      TabIndex        =   17
      Top             =   6690
      Width           =   2715
      Begin VB.TextBox Text 
         Height          =   285
         Left            =   480
         TabIndex        =   18
         Text            =   "Text"
         Top             =   240
         Width           =   1095
      End
   End
   Begin C1SizerLibCtl.C1Elastic frmDynamOLD 
      Height          =   615
      Index           =   0
      Left            =   4710
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1455
      _cx             =   2566
      _cy             =   1085
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
      BackColor       =   5292196
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "UNKILLER"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   0
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
      Frame           =   0
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   0
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Label Label1 
         Caption         =   "belly"
         Height          =   1020
         Left            =   90
         TabIndex        =   21
         Top             =   705
         Width           =   1335
      End
   End
   Begin C1SizerLibCtl.C1Tab tabData 
      Height          =   5295
      Left            =   3000
      TabIndex        =   19
      Top             =   150
      Width           =   5445
      _cx             =   9604
      _cy             =   9340
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
      Caption         =   ""
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   4
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
   End
   Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
      Height          =   2655
      Left            =   120
      OleObjectBlob   =   "frmDynamicDataAddEdit.frx":6C57
      TabIndex        =   24
      Top             =   2760
      Visible         =   0   'False
      Width           =   2775
   End
End
Attribute VB_Name = "frmDynamicDataAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event UpdateMasterTableView(sPassedGUID As String)
Public Event GetSpatialLoc(PassedForm As frmDynamicDataAddEdit)
Private RS As ADODB.Recordset
Private RSActiveTable As ADODB.Recordset
Private RSSchema As ADODB.Recordset
Private bEdit As Boolean
Private m_Conn As ADODB.Connection

Dim WithEvents lblDynamic As Label
Attribute lblDynamic.VB_VarHelpID = -1
Dim ctlControl As Control
Attribute ctlControl.VB_VarHelpID = -1
Dim oldPos As Integer
Dim iframeDynamHeight As Integer
Dim iframeDynamDataHeight As Integer
Dim idxSidebarHeight As Integer
Dim iScrollbarHeight As Integer
Dim bEditDDTable As Boolean

Dim sOLDGUID As String
Dim sGUID As String
Dim RSCollection As Collection

Private RSActiveTables() As ADODB.Recordset
Private RSActiveTablesGEO() As ADODB.Recordset
Dim iTableIndex As Integer

Private Sub SetOnTabChange()
        '<EhHeader>
        On Error GoTo SetOnTabChange_Err
        '</EhHeader>
 
        Dim i As Integer
        Dim sSQL As String

100     If bEditDDTable Or Not tabData = 0 Then
102         i = 0

104         Set dxDBGrid1.DataSource = Nothing
106         dxDBGrid1.Filter.Clear
108         dxDBGrid1.Columns.DestroyColumns
110         Set dxDBGrid1.DataSource = RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Clone
112         dxDBGrid1.Columns.RetrieveFields

114         Do Until i = Me.dxDBGrid1.Columns.Count

116             If Left(Me.dxDBGrid1.Columns(i).FieldName, 2) = "dd" Then
118                 Call SetDropdown(i, Me.dxDBGrid1)
                End If
        
120             Me.dxDBGrid1.Columns(i).caption = getFieldCaption(m_Conn, RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1), dxDBGrid1.Columns(i).FieldName)
 
122             i = i + 1
            Loop

124         Set dxDBGrid1.DataSource = RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1)

126         If Not bEditDDTable Then
            
128             sSQL = "SELECT * FROM " & frmDynam2(tabData.NumTabs - tabData.CurrTab).toolTipText & " WHERE " & RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields(0).Name & " = '" & sGUID & "'"
130             dxDBGrid1.Filter.AddFirst 0, otEqual, sGUID, sGUID, False
132             dxDBGrid1.Filter.FilterStatus = fsNone
134             dxDBGrid1.Filter.Apply

            End If

136         Me.dxDBGrid1.Visible = True
138         Me.cmdCommit.Visible = True
140         dxDBGrid1.Columns(0).Width = 160

142         If Not RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).EOF Then RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).MoveLast

        Else
        
144         Me.dxDBGrid1.Visible = False
146         Me.cmdCommit.Visible = False

        End If

148     If Not RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Source = "none" Then
150         Me.frameSpatial.Visible = True

152         If IsNull(RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID").Value) Then
            
154             RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID").Value = RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID").Value
         
156         ElseIf IsNull(RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID").Value) Then
            
158             RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID").Value = RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID").Value
            End If

        Else
160         Me.frameSpatial.Visible = False
        End If
        
162     If Me.tabData.CurrTab > 0 Then Call dxDBGrid1_OnClick

        '<EhFooter>
        Exit Sub

SetOnTabChange_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.SetOnTabChange " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function getFieldCaption(oConn As ADODB.Connection, _
                                 RSLocalRecordset As ADODB.Recordset, _
                                 sFieldName As String)
        '<EhHeader>
        On Error GoTo getFieldCaption_Err
        '</EhHeader>
 
        Dim oDB As ADOx.Catalog
        Dim itbl As ADOx.Table
        Dim fld As ADOx.Column
 
100     Set oDB = New ADOx.Catalog
102     Set itbl = New ADOx.Table
104     Set oDB.ActiveConnection = oConn
    
106     getFieldCaption = "desc not defined"
        
108     For Each itbl In oDB.Tables
        
110         If itbl.Name = RSLocalRecordset.Fields(0).Properties(1) Then
112             getFieldCaption = itbl.Columns(sFieldName).Properties(2).Value
            End If
        
        Next
        
114     Set eoDB = Nothing
116     Set itbl = Nothing
118     Set fld = Nothing
        '<EhFooter>
        Exit Function

getFieldCaption_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.getFieldCaption " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub cmdBack_Click()
        '<EhHeader>
        On Error GoTo cmdBack_Click_Err
        '</EhHeader>
100     SetGeoRSData
        
102     If Not bEditDDTable And tabData.CurrTab = 1 Then
104         Call ShrinkGrid
        End If
    
106     If tabData.CurrTab > 0 Then
    
108         tabData.TabVisible(tabData.CurrTab - 1) = True
110         tabData.CurrTab = tabData.CurrTab - 1
112         tabData.TabVisible(tabData.CurrTab + 1) = False
114         Call SetOnTabChange
        End If

        '<EhFooter>
        Exit Sub

cmdBack_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.cmdBack_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCancel_Click()
        '<EhHeader>
        On Error GoTo cmdCancel_Click_Err
        '</EhHeader>
100     Unload Me
        '<EhFooter>
        Exit Sub

cmdCancel_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.cmdCancel_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCommit_Click()
        '<EhHeader>
        On Error GoTo cmdCommit_Click_Err
        '</EhHeader>
        
        Dim iOldUID As Integer
    
100     If bEditDDTable Then
            'RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).AddNew
            'RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields(RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields(0).Name).Value = GUIDGen
        
102         RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).AddNew RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields(0).Name, GUIDGen
        
        Else
            'RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).AddNew
            'RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields(RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields(0).Name).Value = sGUID
        
104         RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).AddNew RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields(0).Name, sGUID
        
        End If
        
106     If Not RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Source = "none" Then
        
108         iOldUID = RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID")
110         RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID").Value = iOldUID + 1
112         RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).AddNew "UID", iOldUID + 1

114         If Me.tabData.CurrTab > 0 Then Call dxDBGrid1_OnClick
        End If
    
        '<EhFooter>
        Exit Sub

cmdCommit_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.cmdCommit_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetGeoRSData()

End Sub

Private Sub cmdNext_Click()
        '<EhHeader>
        On Error GoTo cmdNext_Click_Err
        '</EhHeader>
        Dim sString As String
        
100     Me.cmdNext.Enabled = False
102     SetGeoRSData
        
104     If tabData.CurrTab < tabData.NumTabs - 1 Then
        
106         If Not tabData.CurrTab = (tabData.NumTabs - 1) Then 'And Not tabData.CurrTab = 0 Then
108             Call EnlargeGrid
            End If
        
110         tabData.TabVisible(tabData.CurrTab + 1) = True
112         tabData.CurrTab = tabData.CurrTab + 1
114         tabData.TabVisible(tabData.CurrTab - 1) = False
116         Call SetOnTabChange
    
118     ElseIf tabData.CurrTab = tabData.NumTabs - 1 Then
        
120         If MsgBox("Do you want to save this data?", vbYesNo, "Confirm Save") = vbYes Then
            
122             If bEditDDTable Then

124                 Call SaveData
126                 Unload Me
            
128             ElseIf Not bEdit Then
            
130                 tabData.TabVisible(0) = True
132                 Me.tabData.CurrTab = 0
134                 tabData.TabVisible(tabData.NumTabs - 1) = False
136                 Call ShrinkGrid
138                 Call SaveData
140                 sOLDGUID = sGUID
142                 sGUID = GUIDGen
144                 i = iTableIndex

146                 Do Until i = 0

148                     RSActiveTables(i - 1).AddNew RSActiveTables(i - 1).Fields(0).Name, sGUID

150                     sString = Me.frmDynam(tabData.NumTabs - Me.tabData.CurrTab - 1).toolTipText

152                     If Right(sString, 4) = "_FEA" Then sString = Left(sString, Len(sString) - 4)
154                     AddNewToGeoRS i - 1, sString & "_GEO"
156                     i = i - 1
                
                    Loop

158                 Call SetOnTabChange
            
                Else
                
160                 Call SaveData
162                 Unload Me
                    
                End If

            End If
        
        End If
        
164     Me.cmdNext.Enabled = True

        '<EhFooter>
        Exit Sub

cmdNext_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.cmdNext_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SaveData()
        '<EhHeader>
        On Error GoTo SaveData_Err
        '</EhHeader>

        Dim i As Integer
100     i = iTableIndex

102     Do Until i = 0
                
104         If bEditDDTable Or Not i = iTableIndex Then

106             If Not RSActiveTables(i - 1).EOF And Not RSActiveTables(i - 1).BOF Then

108                 RSActiveTables(i - 1).MoveLast
110                 RSActiveTables(i - 1).Delete
112                 SafeMoveFirst RSActiveTables(i - 1)
114                 RSActiveTables(i - 1).Filter = adFilterPendingRecords
116                 RSActiveTables(i - 1).UpdateBatch adAffectAllChapters

118                 If Not RSActiveTablesGEO(i - 1).Source = "none" Then

120                     RSActiveTablesGEO(i - 1).Filter = adFilterPendingRecords
122                     RSActiveTablesGEO(i - 1).UpdateBatch adAffectAllChapters
               
                    End If

                End If
                
            Else
            
124             RSActiveTables(i - 1).Filter = adFilterPendingRecords
126             RSActiveTables(i - 1).UpdateBatch adAffectAllChapters

128             If Not RSActiveTablesGEO(i - 1).Source = "none" Then

130                 RSActiveTablesGEO(i - 1).Filter = adFilterPendingRecords
132                 RSActiveTablesGEO(i - 1).UpdateBatch adAffectAllChapters
                End If
                    
            End If

134         i = i - 1
                
        Loop

        '<EhFooter>
        Exit Sub

SaveData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.SaveData " & _
               "at line " & Erl
        Stop
        Resume Next
        '</EhFooter>
End Sub

Private Sub EnlargeGrid()
        '<EhHeader>
        On Error GoTo EnlargeGrid_Err
        '</EhHeader>
100     If Not tabData.Height < 2800 Then
102         Me.tabData.Height = Me.tabData.Height - 2800
104         'Me.VScroll.Height = Me.VScroll.Height - 2800
106         'Me.VScroll1.Height = Me.VScroll1.Height - 2800
108         Me.dxDBGrid1.Width = dxDBGrid1.Width + 5500
        End If

        '<EhFooter>
        Exit Sub

EnlargeGrid_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.EnlargeGrid " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ShrinkGrid()
        '<EhHeader>
        On Error GoTo ShrinkGrid_Err
        '</EhHeader>
100     If Not dxDBGrid1.Width < 5500 Then
102         Me.tabData.Height = Me.tabData.Height + 2800
104         'Me.VScroll.Height = Me.VScroll.Height + 2800
106         'Me.VScroll1.Height = Me.VScroll1.Height + 2800
108         Me.dxDBGrid1.Width = dxDBGrid1.Width - 5500
        End If

        '<EhFooter>
        Exit Sub

ShrinkGrid_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.ShrinkGrid " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetOperation(bPassedEditDDTable As Boolean, _
                        oConn As ADODB.Connection, _
                        Optional sMasterTableName As String, _
                        Optional bNewRecord As Boolean, _
                        Optional sIndex As String)
        '<EhHeader>
        On Error GoTo SetOperation_Err
        '</EhHeader>
100     m_frmDebug.DebugPrint "Start SetOperation"
        Dim i As Integer
        
102     If bPassedEditDDTable Then
        
104         Call EnlargeGrid
        End If

106     bEditDDTable = bPassedEditDDTable
108     iTableIndex = 0

110     Set m_Conn = oConn
112     Set RSSchema = m_Conn.OpenSchema(adSchemaTables)

114     If bNewRecord Then

116         sGUID = GUIDGen
118         sOLDGUID = sGUID
        Else
120         sOLDGUID = sIndex
122         sGUID = sIndex
        End If

124     While Not RSSchema.EOF

126         If Not Right(RSSchema!TABLE_NAME, 4) = "_GEO" Then

128             If bEditDDTable And Left(RSSchema!TABLE_NAME, 2) = "dd" Then

130                 Call addTabWithData(RSSchema!TABLE_NAME)

132             ElseIf Not bEditDDTable And Left(RSSchema!TABLE_NAME, 4) = "link" Then

134                 Call addTabWithData(RSSchema!TABLE_NAME)
            
                End If
                
            End If
            
136         If RSSchema!TABLE_NAME = sMasterTableName & "_FEA" Then sMasterTableName = sMasterTableName & "_FEA"

138         RSSchema.MoveNext
        Wend

140     If Not bEditDDTable Then Call addTabWithData(sMasterTableName)
142     tabData.CurrTab = 0
144     i = 0

146     Do Until i = tabData.NumTabs
148         tabData.TabCaption(i) = frmDynam2(Abs(tabData.NumTabs - i)).toolTipText
150         i = i + 1
        Loop
    
152     Set dxDBGrid1.DataSource = RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1)
154     Call SetOnTabChange

156     If bEditDDTable Then

158         i = iTableIndex
160         bEdit = False

162         Do Until i = 0
            
164             RSActiveTables(i - 1).AddNew RSActiveTables(i - 1).Fields(0).Name, GUIDGen
166             i = i - 1
            Loop

168     ElseIf bNewRecord Then
        
170         i = iTableIndex
172         bEdit = False

174         Do Until i = 0
            
176             RSActiveTables(i - 1).AddNew RSActiveTables(i - 1).Fields(0).Name, sGUID
178             i = i - 1
            Loop
        
        Else
    
180         i = iTableIndex - 1
182         bEdit = True

184         Do Until i = 0
            
186             RSActiveTables(i - 1).AddNew RSActiveTables(i - 1).Fields(0).Name, sGUID
188             i = i - 1
            Loop
    
        End If
    
190     Call tabData_Click
192     m_frmDebug.DebugPrint "End SetOperation"
        '<EhFooter>
        Exit Sub

SetOperation_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.SetOperation " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub AddNewToGeoRS(i As Integer, _
                          sGeoTableName As String)
        '<EhHeader>
        On Error GoTo AddNewToGeoRS_Err
        '</EhHeader>
    
        Dim RSGeoTable As ADODB.Recordset
        Dim iUID As Integer
    
100     If Not RSActiveTablesGEO(i).Source = "none" Then
    
102         Set RSGeoTable = New ADODB.Recordset
    
104         With RSGeoTable
106             Set .ActiveConnection = m_Conn
108             .CursorType = adOpenDynamic
110             .LockType = adLockBatchOptimistic
112             .Source = "SELECT UID FROM " & sGeoTableName & " ORDER BY UID DESC"
114             .CursorLocation = adUseClient
116             .Open
            End With
            
            
    
118         If Not RSGeoTable.EOF Then
    
120             iUID = RSGeoTable.Fields(0) + 1
                'RSActiveTablesGEO(i).MoveFirst
                'RSActiveTablesGEO (i)
122             RSActiveTablesGEO(i).AddNew "UID", iUID
                
124             m_frmDebug.DebugPrint "Added new UID: " & iUID & " to table: " & sGeoTableName
            Else
                RSActiveTablesGEO(i).AddNew "UID", 100
126             'MsgBox "table: [" & sGeoTableName & "] does not have a UID field!!!"
        
            End If
    
128         Set RSGeoTable = Nothing
        
        End If

        '<EhFooter>
        Exit Sub

AddNewToGeoRS_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.AddNewToGeoRS " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub addTabWithData(sTableName As String)
        '<EhHeader>
        On Error GoTo addTabWithData_Err
        '</EhHeader>

        Dim cControl As Control
        Dim strQuery As String
        Dim i As Integer
        Dim iX As Integer
        Dim iY As Integer
        Dim iFieldIndex As Integer
        Dim bIsFirstFields As Boolean
        Dim sCaption As String
        Dim cControl2 As Control
        Dim sFirstField As String
        
100     m_frmDebug.DebugPrint "addTabWithData: " & sTableName
        
102     bIsFirstFields = True

104     iFieldIndex = 0
106     iX = 300
108     iY = 300

110     If iTableIndex = 0 Then
112         ReDim RSActiveTables(0)
114         ReDim RSActiveTablesGEO(0)
        Else
116         ReDim Preserve RSActiveTables(iTableIndex + 1)
118         ReDim Preserve RSActiveTablesGEO(iTableIndex + 1)
        End If

120     Set RSActiveTables(iTableIndex) = New ADODB.Recordset

122     strQuery = "SELECT * FROM " & sTableName
    
124     With RSActiveTables(iTableIndex)
126         Set .ActiveConnection = m_Conn
128         .CursorType = adOpenDynamic
130         .LockType = adLockBatchOptimistic
132         .Source = strQuery
134         .CursorLocation = adUseClient
136         .Open
        End With
        
138     If sTableName = "mastertable" Then
140         sFirstField = RSActiveTables(iTableIndex).Fields(0).Name
142         RSActiveTables(iTableIndex).Close
144         RSActiveTables(iTableIndex).Source = "SELECT * FROM " & sTableName & " WHERE " & sFirstField & " = '" & sGUID & "'"
146         RSActiveTables(iTableIndex).Open
        End If
        
148     Set RSActiveTablesGEO(iTableIndex) = New ADODB.Recordset
150     RSActiveTablesGEO(iTableIndex).Source = "none"
    
152     If Right$(sTableName, 4) = "_FEA" Then
    
154         strQuery = "SELECT * FROM " & Left(sTableName, Len(sTableName) - 4) & "_GEO"
    
156         With RSActiveTablesGEO(iTableIndex)
158             Set .ActiveConnection = m_Conn
160             .CursorType = adOpenDynamic
162             .LockType = adLockBatchOptimistic
164             .Source = strQuery
166             .CursorLocation = adUseClient
168             .Open
            End With
        
            'If Not bEdit Then
170         AddNewToGeoRS iTableIndex, Left(sTableName, Len(sTableName) - 4) & "_GEO"
            ' Else
            '   AddNewToGeoRS iTableIndex, Left(sTableName, Len(sTableName) - 4) & "_GEO"
            ' End If
        
        End If
    
172     If Not iTableIndex = 0 Then
174         tabData.AddTab sTableName ', tabData.NumTabs ' iTableIndex
176         tabData.TabCaption(tabData.NumTabs - 1) = sTableName
178         tabData.TabVisible(tabData.NumTabs - 1) = False
        Else
180         tabData.TabCaption(0) = sTableName
        End If

182     Load frmDynam(frmDynam.Count)
184     frmDynam(frmDynam.Count - 1).toolTipText = sTableName
186     frmDynam(frmDynam.Count - 1).Refresh
188     frmDynam(frmDynam.Count - 1).Visible = True

        On Error Resume Next
190     tabData.AttachPageToTab frmDynam(frmDynam.Count - 1), CLng(tabData.NumTabs - 1)
        On Error GoTo addTabWithData_Err
        
192     Load frmDynam2(frmDynam2.Count)
194     Set frmDynam2(frmDynam2.Count - 1).Container = frmDynam(frmDynam.Count - 1)
196     frmDynam2(frmDynam2.Count - 1).toolTipText = sTableName
198     frmDynam2(frmDynam2.Count - 1).Refresh
200     frmDynam2(frmDynam2.Count - 1).Visible = True
202     frmDynam2(frmDynam2.Count - 1).Move 1, 1, 5000, 6375

204     Do Until iFieldIndex = RSActiveTables(iTableIndex).Fields.Count
            
206         m_frmDebug.DebugPrint "iFieldIndex = " & iFieldIndex

208         If Not IsThisExcluded(sTableName, RSActiveTables(iTableIndex).Fields(iFieldIndex).Name) Then
        
210             If Me.tabData.CurrTab = 0 And Not iFieldIndex = 0 Then bIsFirstFields = False
212             sCaption = getFieldCaption(m_Conn, RSActiveTables(iTableIndex), RSActiveTables(iTableIndex).Fields(iFieldIndex).Name)
214             Call displayLabel(sCaption, iX, iY, 2500, 250, frmDynam2(frmDynam2.Count - 1))

216             If Left(RSActiveTables(iTableIndex).Fields(iFieldIndex).Name, 2) = "dd" Then
    
218                 Call displayCombo(iX + 2500, iY, 2000, 250, "Numeric", frmDynam2(frmDynam2.Count - 1), RSActiveTables(iTableIndex), RSActiveTables(iTableIndex).Fields(iFieldIndex).Name, bIsFirstFields)

220             ElseIf FieldIsBoolean(RSActiveTables(iTableIndex).Fields(iFieldIndex)) Then
            
222                 Call displayCheckBox(iX + 2500, iY, 2000, 250, frmDynam2(frmDynam2.Count - 1), RSActiveTables(iTableIndex), RSActiveTables(iTableIndex).Fields(iFieldIndex).Name, bIsFirstFields)
                
224             ElseIf FieldIsBinary(RSActiveTables(iTableIndex).Fields(iFieldIndex)) Then
            
226                 MsgBox "Binary field detected - this is not supported by this tool"
                
228             ElseIf FieldIsNumeric(RSActiveTables(iTableIndex).Fields(iFieldIndex)) Then
                    
                    'Call displayTextbox(iX + 2500, iY, 2000, 250, "Numeric", frmDynam2(frmDynam2.Count - 1), RSActiveTables(iTableIndex), RSActiveTables(iTableIndex).Fields(iFieldIndex).Name, bIsFirstFields)

230                 If sCaption = "UID" Then
232                     Call displayTextbox(iX + 2500, iY, 2000, 250, "Numeric", frmDynam2(frmDynam2.Count - 1), RSActiveTables(iTableIndex), RSActiveTables(iTableIndex).Fields(iFieldIndex).Name, bIsFirstFields, False, True)
                    Else
234                     Call displayTextbox(iX + 2500, iY, 2000, 250, "Numeric", frmDynam2(frmDynam2.Count - 1), RSActiveTables(iTableIndex), RSActiveTables(iTableIndex).Fields(iFieldIndex).Name, bIsFirstFields, True, True)

                    End If

236             ElseIf FieldIsString(RSActiveTables(iTableIndex).Fields(iFieldIndex)) Then
            
238                 Call displayTextbox(iX + 2500, iY, 2000, 250, "String", frmDynam2(frmDynam2.Count - 1), RSActiveTables(iTableIndex), RSActiveTables(iTableIndex).Fields(iFieldIndex).Name, bIsFirstFields, True, False)
                
240             ElseIf FieldIsTimeDate(RSActiveTables(iTableIndex).Fields(iFieldIndex)) Then
            
242                 Call displayDateTextbox(iX + 2500, iY, 2000, 250, frmDynam2(frmDynam2.Count - 1), RSActiveTables(iTableIndex), RSActiveTables(iTableIndex).Fields(iFieldIndex).Name, bIsFirstFields)
                
                Else
                
244                 MsgBox "Unknown field detected which is not supported by this tool"
                
                End If
                
246             iY = iY + 300
        
            End If

248         iFieldIndex = iFieldIndex + 1
        Loop
        
250     iY = iY + 400
252     frmDynam2(frmDynam2.Count - 1).Height = iY
254     frmDynam2(frmDynam2.Count - 1).Tag = iY
256     iTableIndex = iTableIndex + 1

        '<EhFooter>
        Exit Sub

addTabWithData_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.addTabWithData " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function IsThisExcluded(sTableName As String, _
                                sFieldName As String)
        '<EhHeader>
        On Error GoTo IsThisExcluded_Err
        '</EhHeader>

        Dim i As Integer
100     i = 0
102     IsThisExcluded = False
        
        On Error Resume Next
104     i = UBound(ExcludeArray)
        On Error GoTo IsThisExcluded_Err

106     Do Until i = 0

108         If ExcludeArray(i).sTableName = sTableName And ExcludeArray(i).sFieldName = sFieldName Then
110             IsThisExcluded = True
            End If

112         i = i - 1
        Loop
    
        '<EhFooter>
        Exit Function

IsThisExcluded_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.IsThisExcluded " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Sub MoveColumn(Name As String)
        '<EhHeader>
        On Error GoTo MoveColumn_Err
        '</EhHeader>

        'If chSym = 0 Then Exit Sub
        Dim j, i
100     j = -1

102     For i = 0 To Me.dxDBGrid1.Columns.Count - 1

104         If dxDBGrid1.Columns(i).FieldName = Name Then j = i
        Next

106     If j <> -1 Then dxDBGrid1.Columns(j).Visible = True '  .Scroll j - dxDBGrid.LeftCol, 0
    
        '<EhFooter>
        Exit Sub

MoveColumn_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.MoveColumn " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub SetDropdown(iColumnIndex As Integer, _
                        dxDBGridLocal As dxDBGrid)
        '<EhHeader>
        
        On Error GoTo SetDropdown_Err
        '</EhHeader>

100     dxDBGridLocal.Columns(iColumnIndex).ColumnType = gedLookupEdit
    
102     With dxDBGridLocal.Columns(iColumnIndex).LookupColumn
        
104         .LookupDatasetType = dtADODataset
    
106         With .LookupDataset.ADODataset
108             .ConnectionString = m_Conn.ConnectionString '"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Documents\MCDB_backend.mdb;Persist Security Info=False"
110             .CursorLocation = clUseServer
112             .CursorType = ctKeyset  'ctStatic
114             .CommandType = cmdTable
116             .CommandText = dxDBGridLocal.Columns(iColumnIndex).FieldName
118             .LockType = ltBatchOptimistic
            End With

120         .LookupDataset.Open
122         .DisplaySize = 20
124         .ListAutoWidth = False
126         .ListColumns = "*"
128         .ListFieldIndex = 0
130         .ListFieldName = "option"
132         .ListWidth = 0
134         .LookupCache = True
136         .LookupDataset.Active = False
138         .LookupDatasetType = dtADODataset
140         .LookupKeyField = "id"
142         .LookupResultField = "option"
144         .ListFieldIndex = 0
146         .ListWidth = 200
148         .ListColumns = ""

        End With

        '<EhFooter>
        Exit Sub

SetDropdown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.SetDropdown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chk1_GotFocus(Index As Integer)
ScrollControl chk1(Index)
End Sub

Private Sub cmb1_GotFocus(Index As Integer)
ScrollControl cmb1(Index)
End Sub

Private Sub cmdPickLocation_Click()
        '<EhHeader>
        On Error GoTo cmdPickLocation_Click_Err
        '</EhHeader>

100     Me.Visible = False
102     RaiseEvent GetSpatialLoc(Me)
        'Me.Visible = True
   
        '<EhFooter>
        Exit Sub

cmdPickLocation_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.cmdPickLocation_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub date1_GotFocus(Index As Integer)
ScrollControl date1(Index)
End Sub

Private Sub dxDBGrid1_OnClick()
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnClick_Err
        '</EhHeader>
            
100     If Not RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Source = "none" Then
            
102         If IsNull(RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID").Value) Then

104             Me.txtX.Text = 0
106             Me.txtY.Text = 0

108         ElseIf Not RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Source = "none" Then
                'If Not RSActiveTablesGEO(tabData.CurrTab).Fields("UID").Value = "" Then
110             SafeMoveFirst RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1)
112             RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Find "UID = " & RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Fields("UID").Value
        
114             Me.txtX.Text = IIf(IsNull(RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("XMIN").Value), 0, RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("XMIN").Value)
116             Me.txtY.Text = IIf(IsNull(RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("YMIN").Value), 0, RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("YMIN").Value)
                'End If
            End If
        
        End If
    
        '<EhFooter>
        Exit Sub

dxDBGrid1_OnClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.dxDBGrid1_OnClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGrid1_OnDblClick()
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnDblClick_Err
        '</EhHeader>
    
        On Error Resume Next
    
100     If Not RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).BOF And Not RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).EOF Then
        
102         RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).MoveNext
        
104         If RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).EOF Then
        
106             RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).MoveLast
108             MsgBox "You cannot delete this row"
        
            Else
            
110             RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).MovePrevious

112             If MsgBox("Are you sure you wanna delete this record?", vbYesNo, "Confirm Deletion") = vbYes Then
        
114                 RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1).Delete adAffectCurrent
116                 Me.dxDBGrid1.Columns.RetrieveFields
118                 SafeMoveFirst RSActiveTables(tabData.NumTabs - tabData.CurrTab - 1)
120                 Call dxDBGrid1_OnClick
                End If
        
            End If
        
        End If
    
        '<EhFooter>
        Exit Sub

dxDBGrid1_OnDblClick_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.dxDBGrid1_OnDblClick " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Set Me.Picture = g_PictureDialogLarge
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub popXY(x As Double, _
                 y As Double)
        '<EhHeader>
        On Error GoTo popXY_Err
        '</EhHeader>

100     Me.txtX = x
102     Me.txtY = y

        '<EhFooter>
        Exit Sub

popXY_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.popXY " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadLangFromClient()
        '<EhHeader>
        On Error GoTo LoadLangFromClient_Err
        '</EhHeader>

100     If Not g_sLanguage = "" Then
102         If Not m_Cnn.State = adStateClosed Then
104             LoadLanguage Me.Name, g_sLanguage, m_Cnn
            End If
        End If

        '<EhFooter>
        Exit Sub

LoadLangFromClient_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.LoadLangFromClient " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub removeDynamicControls()
        '<EhHeader>
        On Error GoTo removeDynamicControls_Err
        '</EhHeader>

        Dim i As Integer
100     i = lbl1.Count

102     Do While i > 1
104         Unload lbl1(i - 1)
106         i = i - 1
        Loop
    
108     i = txt1.Count

110     Do While i > 1
112         Unload txt1(i - 1)
114         i = i - 1
        Loop
    
116     i = cmb1.Count

118     Do While i > 1
120         Unload cmb1(i - 1)
122         i = i - 1
        Loop
    
124     i = chk1.Count

126     Do While i > 1
128         Unload chk1(i - 1)
130         i = i - 1
        Loop
    
132     i = date1.Count

134     Do While i > 1
136         Unload date1(i - 1)
138         i = i - 1
        Loop

        '<EhFooter>
        Exit Sub

removeDynamicControls_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.removeDynamicControls " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function displayLabel(sCaption As String, _
                              x As Integer, _
                              y As Integer, _
                              w As Integer, _
                              H As Integer, _
                              ByVal fFrame As Control) As Control
        '<EhHeader>
        On Error GoTo displayLabel_Err
        '</EhHeader>

100     Load lbl1(lbl1.Count)
102     Set lbl1(lbl1.UBound).Container = fFrame
104     lbl1(lbl1.UBound).Move x, y, w, H
106     lbl1(lbl1.UBound).Font.Size = 1
108     lbl1(lbl1.UBound).caption = sCaption
110     lbl1(lbl1.UBound).Visible = True
112     lbl1(lbl1.UBound).Tag = iControlIndex
114     lbl1(lbl1.UBound).BackStyle = vbTransparent

        '<EhFooter>
        Exit Function

displayLabel_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.displayLabel " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function displayTextbox(x As Integer, _
                                y As Integer, _
                                w As Integer, _
                                H As Integer, _
                                sFormat As String, _
                                ByVal cControl As Control, _
                                dsDataSource As ADODB.Recordset, _
                                sFieldName As String, _
                                bIsFirstFields As Boolean, _
                                bNotUIDField As Boolean, _
                                bSetAsNumeric As Boolean) As Control
        '<EhHeader>
        On Error GoTo displayTextbox_Err
        '</EhHeader>
 
100     Load txt1(txt1.Count)
102     Set txt1(txt1.UBound).Container = cControl
104     txt1(txt1.UBound).Move x, y, w, H
106     txt1(txt1.UBound).Font.Size = 1
108     txt1(txt1.UBound).Text = ""
110     txt1(txt1.UBound).Visible = True
112     Set txt1(txt1.UBound).DataSource = Nothing
114     txt1(txt1.UBound).DataField = sFieldName
116     Set txt1(txt1.UBound).DataSource = dsDataSource

118     If bSetAsNumeric Then
120         txt1(txt1.UBound).DataFormat.Format = 0
122         txt1(txt1.UBound).DataFormat.Type = 1
            
        End If
    
124     If Not bNotUIDField Or bIsFirstFields Then
126         txt1(txt1.UBound).Enabled = False
128         txt1(txt1.UBound).BackColor = &H8000000C
        End If

        '<EhFooter>
        Exit Function

displayTextbox_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.displayTextbox " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function displayDateTextbox(x As Integer, _
                                    y As Integer, _
                                    w As Integer, _
                                    H As Integer, _
                                    ByVal cControl As Control, _
                                    dsDataSource As ADODB.Recordset, _
                                    sFieldName As String, _
                                    bIsFirstFields As Boolean) As Control
        '<EhHeader>
        On Error GoTo displayDateTextbox_Err
        '</EhHeader>

100     Load date1(date1.Count)
102     Set date1(date1.UBound).Container = cControl
104     date1(date1.UBound).Move x, y, w, H
106     date1(date1.UBound).DataFormat.Format = "Medium Date"
108     date1(date1.UBound).Visible = True

110     Set date1(date1.UBound).DataSource = Nothing
112     date1(date1.UBound).DataField = sFieldName
114     Set date1(date1.UBound).DataSource = dsDataSource

116     If bIsFirstFields Then
118         date1(date1.UBound).Enabled = False
120         date1(date1.UBound).BackColor = &H8000000C
        End If

        '<EhFooter>
        Exit Function

displayDateTextbox_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.displayDateTextbox " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function displayCombo(x As Integer, _
                              y As Integer, _
                              w As Integer, _
                              H As Integer, _
                              sFormat As String, _
                              ByVal cControl As Control, _
                              dsDataSource As ADODB.Recordset, _
                              sFieldName As String, _
                              bIsFirstFields As Boolean) As Control
        '<EhHeader>
        On Error GoTo displayCombo_Err
        '</EhHeader>

        Dim strQuery As String
        Dim i As Integer
    
100     Load cmb1(cmb1.Count)
102     Set cmb1(cmb1.UBound).Container = cControl
104     cmb1(cmb1.UBound).Move x, y, w
106     cmb1(cmb1.UBound).Visible = True
    
108     strQuery = "SELECT id, option FROM " & sFieldName & " ORDER BY option"
110     Set RS = New ADODB.Recordset

112     With RS
114         Set .ActiveConnection = m_Conn
116         .CursorType = adOpenDynamic  'adOpenKeyset
118         .LockType = adLockOptimistic
120         .Source = strQuery
122         .Open
        End With
    
124     If Not RS.EOF And Not RS.BOF Then

126         SafeMoveFirst RS
    
128         Set cmb1(cmb1.UBound).LookUpRecordset = RS
130         cmb1(cmb1.UBound).LookUpKeyFieldName = "id"
132         cmb1(cmb1.UBound).LookUpDisplayFieldName = "option"
134         cmb1(cmb1.UBound).ListFieldName = "option"
136         cmb1(cmb1.UBound).ListColumns = ""
    
138         Set cmb1(cmb1.UBound).DataSource = Nothing
140         cmb1(cmb1.UBound).DataField = sFieldName
142         Set cmb1(cmb1.UBound).DataSource = dsDataSource

144         If bIsFirstFields Then
146             cmb1(cmb1.UBound).Enabled = False
            End If
        
        End If

        '<EhFooter>
        Exit Function

displayCombo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.displayCombo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function displayCheckBox(x As Integer, _
                                 y As Integer, _
                                 w As Integer, _
                                 H As Integer, _
                                 ByVal cControl As Control, _
                                 dsDataSource As ADODB.Recordset, _
                                 sFieldName As String, _
                                 bIsFirstFields As Boolean) As Control
        '<EhHeader>
        On Error GoTo displayCheckBox_Err
        '</EhHeader>

100     Load chk1(chk1.Count)
102     Set chk1(chk1.UBound).Container = cControl
104     chk1(chk1.UBound).Move x, y, w
106     chk1(chk1.UBound).Font.Size = 1
108     chk1(chk1.UBound).caption = sCaption
110     chk1(chk1.UBound).Visible = True

112     Set chk1(chk1.UBound).DataSource = Nothing
114     chk1(chk1.UBound).DataField = sFieldName
116     Set chk1(chk1.UBound).DataSource = dsDataSource

118     If bIsFirstFields Then
120         chk1(chk1.UBound).Enabled = False
122         chk1(chk1.UBound).BackColor = &H8000000C
        End If

        '<EhFooter>
        Exit Function

displayCheckBox_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.displayCheckBox " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     Me.Visible = False

102     If Not bEdit And sOLDGUID = sGUID Then sOLDGUID = ""
104     RaiseEvent UpdateMasterTableView(sOLDGUID)
        
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub tabData_Click()
'        '<EhHeader>
'        On Error GoTo tabData_Click_Err
'        '</EhHeader>
'
'100     If frmDynam2(Abs(tabData.NumTabs - tabData.CurrTab)).Tag < tabData.Height Then
'102         VScroll.Max = 0
'104         VScroll.Min = 0
'        Else
'106         VScroll.Max = Abs((frmDynam2(Abs(tabData.NumTabs - tabData.CurrTab)).Tag - (tabData.Height)) / 300)
'108         VScroll.Min = 0
'        End If
'
'110     VScroll.Value = 0
'
'        '<EhFooter>
'        Exit Sub
'
'tabData_Click_Err:
'        MsgBox Err.Description & vbCrLf & _
'               "in OASISClient.frmDynamicDataAddEdit.tabData_Click " & _
'               "at line " & Erl
'        Resume Next
'        '</EhFooter>
End Sub

Private Sub txt1_GotFocus(Index As Integer)

    ScrollControl txt1(Index)

End Sub

Private Sub ScrollControl(cControl As Control)

    Dim i As Integer
    i = (cControl.Top / Int(0 & frmDynam2(Abs(tabData.NumTabs - tabData.CurrTab)).Tag)) * 1000

    If i < 100 Then
        i = 0
    ElseIf i > 900 Then
        i = 1000
    ElseIf i > 100 Then
        i = i
    End If
    
    frmDynam(Abs(tabData.NumTabs - tabData.CurrTab)).vValue = i

End Sub


Private Sub txtX_Change()
        '<EhHeader>
        On Error GoTo txtX_Change_Err
        '</EhHeader>

100     If Not RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Source = "none" Then
102         RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("XMIN").Value = txtX
104         RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("XMAX").Value = txtX
106         RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("SHAPETYPE").Value = 2
        End If
    
        '<EhFooter>
        Exit Sub

txtX_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.txtX_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub txtY_Change()
        '<EhHeader>
        On Error GoTo txtY_Change_Err
        '</EhHeader>

100     If Not RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Source = "none" Then
102         RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("YMIN").Value = txtY
104         RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("YMAX").Value = txtY
106         RSActiveTablesGEO(tabData.NumTabs - tabData.CurrTab - 1).Fields("SHAPETYPE").Value = 2
        End If

        '<EhFooter>
        Exit Sub

txtY_Change_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDynamicDataAddEdit.txtY_Change " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

