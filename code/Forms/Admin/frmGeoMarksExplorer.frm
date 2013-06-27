VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmGeoMarksExplorer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geo Marks Wizard"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "frmGeoMarksExplorer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   285
      Left            =   5070
      TabIndex        =   24
      Top             =   4980
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   6060
      TabIndex        =   23
      Top             =   4980
      Width           =   915
   End
   Begin VB.CommandButton cmdTools 
      Caption         =   "Admin Tools"
      Height          =   315
      Left            =   5910
      TabIndex        =   22
      Top             =   30
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   285
      Left            =   4080
      TabIndex        =   18
      Top             =   4980
      Width           =   915
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   285
      Left            =   3120
      TabIndex        =   17
      Top             =   4980
      Width           =   915
   End
   Begin C1SizerLibCtl.C1Tab C1TGM 
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   7005
      _cx             =   12356
      _cy             =   5741
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
      Caption         =   "Tab&1|Tab&2|New Tab|Tab&3|New Tab|New Tab"
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
      Begin VB.Frame Frame 
         Height          =   3225
         Left            =   15
         TabIndex        =   19
         Top             =   15
         Width           =   6975
         Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
            Height          =   3225
            Left            =   0
            OleObjectBlob   =   "frmGeoMarksExplorer.frx":6852
            TabIndex        =   20
            Top             =   0
            Width           =   7005
         End
      End
      Begin C1SizerLibCtl.C1Elastic el5 
         Height          =   3225
         Left            =   8820
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   15
         Width           =   6975
         _cx             =   12303
         _cy             =   5689
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
         Begin VB.CommandButton cmdSubmitAnd 
            Caption         =   "Submit and AddNew?"
            Height          =   405
            Left            =   4230
            TabIndex        =   16
            Top             =   1080
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.CommandButton cmdTry 
            Caption         =   "Try"
            Height          =   285
            Left            =   5100
            TabIndex        =   15
            Top             =   660
            Width           =   495
         End
         Begin VB.TextBox txtGMUrl 
            Enabled         =   0   'False
            Height          =   285
            Left            =   750
            TabIndex        =   14
            Top             =   660
            Width           =   4305
         End
         Begin VB.CheckBox chkURLMArk 
            Caption         =   "Provide Hyperlink for this GeoBookmark"
            DataField       =   "isURLMark"
            Height          =   255
            Left            =   780
            TabIndex        =   12
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label lblUrl 
            AutoSize        =   -1  'True
            Caption         =   "Url"
            Height          =   195
            Left            =   450
            TabIndex        =   13
            Top             =   660
            Width           =   195
         End
      End
      Begin C1SizerLibCtl.C1Elastic el4 
         Height          =   3225
         Left            =   8520
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   15
         Width           =   6975
         _cx             =   12303
         _cy             =   5689
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
         Begin VB.Frame Frame1 
            Caption         =   "Location:"
            Height          =   1275
            Left            =   690
            TabIndex        =   26
            Top             =   390
            Width           =   5505
            Begin VB.TextBox txtX 
               Height          =   285
               Left            =   375
               TabIndex        =   30
               Top             =   420
               Width           =   1680
            End
            Begin VB.TextBox txtY 
               Height          =   285
               Left            =   390
               TabIndex        =   29
               Top             =   720
               Width           =   1680
            End
            Begin VB.TextBox txtZoom 
               Height          =   315
               Left            =   3120
               TabIndex        =   28
               Top             =   420
               Width           =   1665
            End
            Begin VB.CommandButton cmdGetFromMap 
               Caption         =   "..."
               Height          =   285
               Left            =   4860
               TabIndex        =   27
               Top             =   450
               Width           =   555
            End
            Begin VB.Label lblLBL 
               AutoSize        =   -1  'True
               Caption         =   "X:"
               Height          =   195
               Index           =   8
               Left            =   150
               TabIndex        =   33
               Tag             =   "Nam"
               Top             =   465
               Width           =   150
            End
            Begin VB.Label lblLBL 
               AutoSize        =   -1  'True
               Caption         =   "Y:"
               Height          =   195
               Index           =   9
               Left            =   150
               TabIndex        =   32
               Tag             =   "Nam"
               Top             =   765
               Width           =   150
            End
            Begin VB.Label lblZoom 
               Caption         =   "Zoom:"
               Height          =   255
               Left            =   2550
               TabIndex        =   31
               Top             =   480
               Width           =   525
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic el3 
         Height          =   3225
         Left            =   8220
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   15
         Width           =   6975
         _cx             =   12303
         _cy             =   5689
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
         Appearance      =   4
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
         Begin VB.TextBox txtCategory 
            Height          =   285
            Left            =   2880
            TabIndex        =   21
            Text            =   "Text1"
            Top             =   1320
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ComboBox ComGeoMCategory 
            Height          =   315
            Left            =   210
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   510
            Width           =   4785
         End
         Begin VB.CommandButton cmdNewGMCategory 
            Caption         =   "..."
            Height          =   315
            Left            =   5040
            TabIndex        =   8
            Top             =   510
            Width           =   435
         End
         Begin VB.Label Label1 
            Caption         =   "Category:"
            Height          =   225
            Left            =   210
            TabIndex        =   25
            Top             =   240
            Width           =   1695
         End
      End
      Begin C1SizerLibCtl.C1Elastic el2 
         Height          =   3225
         Left            =   7920
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   6975
         _cx             =   12303
         _cy             =   5689
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
         Begin VB.TextBox txtGeomarks 
            Height          =   915
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   510
            Width           =   5175
         End
         Begin VB.Label lblLBL 
            AutoSize        =   -1  'True
            Caption         =   "Description:"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   7
            Top             =   270
            Width           =   840
         End
      End
      Begin C1SizerLibCtl.C1Elastic el1 
         Height          =   3225
         Left            =   7620
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   15
         Width           =   6975
         _cx             =   12303
         _cy             =   5689
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
         Begin VB.TextBox txtGeomarksName 
            Height          =   285
            Left            =   750
            TabIndex        =   4
            Top             =   300
            Width           =   3660
         End
         Begin VB.Label lblLBL 
            AutoSize        =   -1  'True
            Caption         =   "Name:"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   5
            Tag             =   "Nam"
            Top             =   360
            Width           =   465
         End
      End
   End
End
Attribute VB_Name = "frmGeoMarksExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSLocalUserGroups As New ADODB.Recordset
Dim RSBookmarks As New ADODB.Recordset
Dim RSCategories As New ADODB.Recordset
Dim c As New cCommonDialog
Dim WithEvents m_frmGeoMarksCategory As frmGeoMarksCategory
Attribute m_frmGeoMarksCategory.VB_VarHelpID = -1

Public GisUtils As New XGIS_Utils
Private m_GeoMarksLayer As New XGIS_LayerVector

Public Sub setUserGroupsRS(ByVal PassedRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo setUserGroupsRS_Err
        '</EhHeader>
100     Set RSLocalUserGroups = PassedRS
    
        '<EhFooter>
        Exit Sub

setUserGroupsRS_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.setUserGroupsRS " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub chkURLMArk_Click()
        '<EhHeader>
        On Error GoTo chkURLMArk_Click_Err
        '</EhHeader>

100     If chkURLMArk.Value = vbChecked Then
102         Me.txtGMUrl.Enabled = True
        Else
104         Me.txtGMUrl.Enabled = False
        End If

        '<EhFooter>
        Exit Sub

chkURLMArk_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.chkURLMArk_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdBack_Click()
        '<EhHeader>
        On Error GoTo cmdBack_Click_Err
        '</EhHeader>
    
100     With Me.C1TGM

102         If Not .CurrTab = 0 Then
104             .CurrTab = .CurrTab - 1
            End If

        End With

        '<EhFooter>
        Exit Sub

cmdBack_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.cmdBack_Click " & "at line " & Erl
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
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.cmdCancel_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdGetFromMap_Click()
        '<EhHeader>
        On Error GoTo cmdGetFromMap_Click_Err
        '</EhHeader>
        
        Dim SymbolList As New XGIS_SymbolList
        On Error Resume Next
100     c.DefaultExt = "*.ttkgp"
102     c.DialogTitle = "Open Map Definition File"
104     c.Filter = "Map Definition Files (*.ttkgp;*.prj)|*.ttkgp;*.prj"
106     c.ShowOpen
    
108     If Not c.fileName = "" Then
110         frmCoordPicker.Init c.fileName
112         Set m_GeoMarksLayer = New XGIS_LayerVector

114         With m_GeoMarksLayer
116             .AddField "Name", XgisFieldTypeString, 255, 0
118             .Params.Marker.Color = vbWhite
120             .Params.Marker.OutlineColor = vbBlue
122             .Name = Now
124             .Params.Marker.Symbol = SymbolList.Prepare(Replace("c:\OASIS\Client", "\", "\\\\") & "\\\\Data\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE") '"..\\\\..\\\\GIS\\\\OTHER\\\\CUSTSYMB\\\\PIN1-32.BMP?TRUE") 'SymbolList.Prepare("ESRI Crime Analysis:41:NORMAL")
126             .Params.Marker.Size = 440
128             .Params.Marker.ShowLegend = 1
130             .Params.Legend = "My GeoBookMarks"
132             .Params.Visible = True
134             .Params.Label.Field = "Name"
136             .Params.Label.Visible = True
            
138             frmCoordPicker.AddLyr m_GeoMarksLayer
140             CreateGMLyr
        
142             frmCoordPicker.RefreshLyr .Name
        
            End With

        End If
    
144     frmCoordPicker.Show vbModal, Me
    
146     With frmCoordPicker
148         txtX.Text = .xCoord
150         txtY.Text = .yCoord
152         txtZoom.Text = .ZoomVal
        End With
    
        '<EhFooter>
        Exit Sub

cmdGetFromMap_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.cmdGetFromMap_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdNewGMCategory_Click()
        'MsgBox "TODO Create an add New Category Wizard"
        '<EhHeader>
        On Error GoTo cmdNewGMCategory_Click_Err
        '</EhHeader>
    
100     Set m_frmGeoMarksCategory = New frmGeoMarksCategory
102     m_frmGeoMarksCategory.setUserGroupsRS RSLocalUserGroups
104     m_frmGeoMarksCategory.Show vbModeless, Me
    
        '<EhFooter>
        Exit Sub

cmdNewGMCategory_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.cmdNewGMCategory_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub m_frmGeoMarksCategory_RefreshCategories()
        '<EhHeader>
        On Error GoTo m_frmGeoMarksCategory_RefreshCategories_Err
        '</EhHeader>
        
        Dim sString As String
        
100     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT ID, Name FROM " & RSLocalUserGroups!Name & "GeoBookMarksCategories ORDER BY Name")
102     Set RSCategories = m_frmOASISProgress.OpenHttpCommsRS(sString, True)

104     Me.ComGeoMCategory.Clear

106     If Not RSCategories.EOF Or Not RSCategories.Bof Then

108         RSCategories.MoveFirst
     
110         Do Until RSCategories.EOF
     
112             Me.ComGeoMCategory.AddItem RSCategories!Name
114             Me.ComGeoMCategory.ItemData(Me.ComGeoMCategory.NewIndex) = RSCategories!id
116             RSCategories.MoveNext
     
            Loop
        
        End If

        '<EhFooter>
        Exit Sub

m_frmGeoMarksCategory_RefreshCategories_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.m_frmGeoMarksCategory_RefreshCategories " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function getClientDBPath()
        '<EhHeader>
        On Error GoTo getClientDBPath_Err
        '</EhHeader>
        
100     getClientDBPath = CreateAppPath & "\Data\db"

        '<EhFooter>
        Exit Function

getClientDBPath_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.getClientDBPath " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub CreateGMLyr()
        '<EhHeader>
        On Error GoTo CreateGMLyr_Err
        '</EhHeader>
        Dim RS As New ADODB.Recordset
        Dim WRS As New ADODB.Recordset
        Dim cn As New Connection
        Dim shpInc As XGIS_Shape
        Dim SymbolList As New XGIS_SymbolList

100     cn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & getClientDBPath & "\OasisClient.mdb;"

102     With WRS
104         .open "SELECT * FROM GeoBookMarks", cn, adOpenDynamic, adLockOptimistic
            
106         If Not .State = adStateClosed Then
                ' If Not .EOF Then .MoveFirst

108             If Not .Bof And Not .EOF Then
110                 .MoveFirst

112                 Do While Not .EOF

114                     If Not .fields.Item("X").Value = vbNull Or Not .fields.Item("Y").Value = vbNull Then
116                         Set shpInc = m_GeoMarksLayer.CreateShape(XgisShapeTypePoint)
            
118                         shpInc.Lock XgisLockExtent
120                         shpInc.AddPart
            
122                         With .fields
124                             shpInc.AddPoint GisUtils.GisPoint(CDbl(Replace(.Item("X").Value, ",", ".")), CDbl(Replace(.Item("Y").Value, ",", ".")))
126                             shpInc.SetField "Name", .Item("Name").Value
                                'shpInc.SetField "Z", RS.fields.Item("name").Value
                                'shpInc.SetField "Description", .Item("Description").Value
                                'shpInc.SetField "PCode", "IQ200700" & .Item("pCode").Value
                            End With
                
128                         shpInc.Unlock

130                         m_GeoMarksLayer.AddShape shpInc
                        End If

132                     .MoveNext
                    Loop

                End If
            End If

        End With
        
        '    m_oW3Lyr.ParamsList.Add
        '    m_oW3Lyr.Params.Query = "Type <> 'home'"
        'sFont = .Item("Font_Name").value
        'sFont = sFont & ":" & .Item("Ascii").value & ":NORMAL"
            
134     With m_GeoMarksLayer

136         .Paint
        End With

        '<EhFooter>
        Exit Sub

CreateGMLyr_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.CreateGMLyr " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub addNewRow()
        '<EhHeader>
        On Error GoTo addNewRow_Err
        '</EhHeader>

        'RSBookmarks.AddNew
100     chkURLMArk.Value = vbUnchecked
        '    Me.txtSessionGUID = GUIDGen()
        'ComGeoMCategory.Text = ComGeoMCategory.List(0)
102     Call chkURLMArk_Click

        '<EhFooter>
        Exit Sub

addNewRow_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.addNewRow " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdNext_Click()
        '<EhHeader>
        On Error GoTo cmdNext_Click_Err
        '</EhHeader>

        Dim bVer As Boolean
100     cmdBack.Caption = "Back"

102     With C1TGM
    
104         Select Case .CurrTab
    
                Case 0
106                 bVer = True
108                 Call addNewRow

110             Case 1

112                 If Len(txtGeomarksName.Text) > 2 Then
114                     bVer = True
                    Else
116                     MsgBox "Please enter a valid name!"
                    End If

118             Case 2
               
120                 If Len(txtGeomarks.Text) > 2 Then
122                     bVer = True
                    Else
124                     MsgBox "Please enter a valid description!"
                    End If
               
126             Case 3
                
128                 If Len(Me.cmdNewGMCategory.Caption) > 2 Then
130                     bVer = True
132                     txtCategory.Text = Me.ComGeoMCategory.Text
                    Else
134                     MsgBox "Please select a valid category!"
                    End If

136             Case 4
                
138                 If IsNumeric(Me.txtX) And IsNumeric(Me.txtY) And IsNumeric(Me.txtZoom) Then
140                     bVer = True
                    Else
142                     MsgBox "Please ensure you enter fill all fields!"
                    End If

144             Case 5

146                 bVer = False
                
148                 If Not txtGMUrl.Enabled Or Len(Me.txtGMUrl.Text) > 6 Then

150                     RSBookmarks.AddNew

152                     With RSBookmarks.fields
154                         .Item("Name").Value = Me.txtGeomarksName
156                         .Item("Description").Value = Me.txtGeomarks
158                         .Item("X").Value = Me.txtX
160                         .Item("Y").Value = Me.txtY
162                         .Item("Z").Value = Me.txtZoom
164                         .Item("sURL").Value = Me.txtGMUrl

166                         .Item("UseSymbol").Value = False
168                         .Item("SymbolChar").Value = ""
170                         .Item("SymbolFont").Value = ""
172                         .Item("SymbolSize").Value = ""
174                         .Item("MapName").Value = c.fileName
176                         .Item("BmkrID").Value = Me.ComGeoMCategory.ItemData(Me.ComGeoMCategory.ListIndex)
178                         .Item("sGUID").Value = GUIDGen()
180                         .Item("dTimeStamp").Value = Now
182                         .Item("Deleted").Value = False
184                         .Item("OwnerGUID").Value = RSLocalUserGroups!sGUID
186                         .Item("isURLMark").Value = IIf(Me.chkURLMArk.Value = vbChecked, True, False)
                
                        End With
                        
188                     Me.C1TGM.CurrTab = 0
                        'Set dxDBGrid1.DataSource = RSBookmarks
             
                    Else
190                     MsgBox "Please enter a URL!"
                    End If
                
            End Select

192         If bVer Then
194             If Not .CurrTab = .NumTabs Then
196                 .CurrTab = .CurrTab + 1
                End If
            End If

        End With
    
        '<EhFooter>
        Exit Sub

cmdNext_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.cmdNext_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdSave_Click()
        '<EhHeader>
        On Error GoTo cmdSave_Click_Err
        '</EhHeader>

        Dim bReturnValue As Boolean
100     RSBookmarks.Filter = adFilterPendingRecords

102     If Not RSBookmarks.EOF And Not RSBookmarks.Bof Then
        
104         If MsgBox("Do you wish to save your changes?", vbYesNo, "Confirm Save") = vbYes Then
        
106             bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(RSBookmarks, WebSite & "Oasis.asp", True)

108             If bReturnValue Then
110                 MsgBox "Data saved to server"
                Else
112                 MsgBox "Saving to server failed!"
                End If

            End If

        End If

114     Unload Me

        '<EhFooter>
        Exit Sub

cmdSave_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.cmdSave_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTools_Click()
        '<EhHeader>
        On Error GoTo cmdTools_Click_Err
        '</EhHeader>
        Dim m_frmAdminTools As frmAdminTools
100     Set m_frmAdminTools = New frmAdminTools
102     m_frmAdminTools.Show vbModeless, Me
104     Set m_frmAdminTools = Nothing
        '<EhFooter>
        Exit Sub

cmdTools_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.cmdTools_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdTry_Click()
        '<EhHeader>
        On Error GoTo cmdTry_Click_Err
        '</EhHeader>
100     ShellExecute Me.hwnd, vbNullString, Me.txtGMUrl.Text, vbNullString, vbNullString, 1
102     MsgBox txtGMUrl.Text
        '<EhFooter>
        Exit Sub

cmdTry_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.cmdTry_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub dxDBGrid1_OnDblClick()
        '<EhHeader>
        On Error GoTo dxDBGrid1_OnDblClick_Err
        '</EhHeader>
    
100     If DeleteRecordFromRSAndSave(RSBookmarks) Then Unload Me
    
        '<EhFooter>
        Exit Sub

dxDBGrid1_OnDblClick_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.dxDBGrid1_OnDblClick " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
    
        Dim sString As String
100     Set RSCategories = New ADODB.Recordset
102     Set RSBookmarks = New ADODB.Recordset
104     Me.Picture = g_PictureDialogSmall

106     DoEvents
    
108     sString = WebSite & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & RSLocalUserGroups!Name & "GeoBookMarks")
110     Set RSBookmarks = m_frmOASISProgress.OpenHttpCommsRS(sString, True)
    
112     Set dxDBGrid1.DataSource = RSBookmarks
114     dxDBGrid1.Columns.RetrieveFields

116     Call m_frmGeoMarksCategory_RefreshCategories
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.Form_Load " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>

100     Set RSCategories = Nothing
102     Set RSBookmarks = Nothing
104     Set RSLocalUserGroups = Nothing

        Set m_frmGeoMarksCategory = Nothing

        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.frmGeoMarksExplorer.Form_Unload " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
