VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmChangeTracer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trend Analysis Settings"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "frmChangeTracer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3465
      Top             =   2205
   End
   Begin C1SizerLibCtl.C1Elastic elMain 
      Height          =   3870
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6330
      _cx             =   11165
      _cy             =   6826
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
      Begin VB.Frame FraStatus 
         Caption         =   "Status:"
         Height          =   1185
         Left            =   3150
         TabIndex        =   11
         Top             =   1170
         Width           =   2985
         Begin VB.Label lblCurrentAnimation 
            AutoSize        =   -1  'True
            Caption         =   "Current Animation: N/A"
            Height          =   195
            Left            =   90
            TabIndex        =   12
            Top             =   270
            Width           =   1635
         End
      End
      Begin VB.Frame FraDataSettings 
         Caption         =   "Data Settings:"
         Height          =   3615
         Left            =   90
         TabIndex        =   5
         Top             =   135
         Width           =   2940
         Begin MSComctlLib.ListView lvUniqueValues 
            Height          =   1770
            Left            =   135
            TabIndex        =   10
            Top             =   1755
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   3122
            View            =   2
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.ComboBox ComAnalysisField 
            Height          =   315
            Left            =   135
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   1305
            Width           =   2625
         End
         Begin VB.ComboBox ComLayerName 
            Height          =   315
            Left            =   180
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   585
            Width           =   2580
         End
         Begin VB.Label lblAnalysisField 
            Caption         =   "Analysis field:"
            Height          =   285
            Left            =   135
            TabIndex        =   9
            Top             =   1035
            Width           =   2715
         End
         Begin VB.Label lblDataTo 
            Caption         =   "Data to analyse:"
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   315
            Width           =   2580
         End
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Start"
         Height          =   420
         Left            =   4725
         TabIndex        =   2
         Top             =   3285
         Width           =   1455
      End
      Begin VB.Frame FraGeneralSettings 
         Caption         =   "General Settings"
         Height          =   960
         Left            =   3150
         TabIndex        =   1
         Top             =   135
         Width           =   2985
         Begin VB.TextBox txtInterval 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   330
            Left            =   1935
            TabIndex        =   4
            Text            =   "5"
            Top             =   225
            Width           =   780
         End
         Begin VB.Label lblSequenceInterval 
            Caption         =   "Sequence interval (sec):"
            Height          =   285
            Left            =   135
            TabIndex        =   3
            Top             =   315
            Width           =   1770
         End
      End
   End
End
Attribute VB_Name = "frmChangeTracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_sLastFieldName As String
Private m_sLastLayer As String
Public Event GetFields(sLayer As String)
Public Event GetUniqueValues(sLayer As String, sField As String)
Public Event ChangeScope(sScope As String, sLayer As String)
Private M_sVals As Variant
Private m_iCurListItem As Integer
Private m_SQLPrefix As String

Public Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
    Dim i As Integer

100     SafeMoveFirst g_RSGISGridTableSettings
    
102     ComLayerName.Clear
    
104     Do While Not g_RSGISGridTableSettings.EOF
106         ComLayerName.AddItem g_RSGISGridTableSettings.Fields.Item("alias").value
108         g_RSGISGridTableSettings.MoveNext
        Loop
 
110     If m_oColUserLayers.Count > 0 Then
112         For i = 1 To m_oColUserLayers.Count - 1
114             ComLayerName.AddItem m_oColUserLayers.Item(i)
            Next
        End If

116     If ComLayerName.ListCount > 0 Then ComLayerName.ListIndex = 0
    
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangeTracer.init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdStart_Click()
    m_iCurListItem = 1
    Timer1.Interval = txtInterval.Text * 1000
    If Not Timer1.Enabled Then
        Timer1.Enabled = True
        cmdStart.caption = "Stop"
    Else
        Timer1.Enabled = False
        cmdStart.caption = "Start"
    End If
End Sub

Private Sub ComAnalysisField_Click()
        '<EhHeader>
        On Error GoTo ComAnalysisField_Click_Err
        '</EhHeader>
    Dim sLyr As String
    
        Timer1.Enabled = False
        cmdStart.caption = "Start"
        
        If ComAnalysisField.List(ComAnalysisField.ListIndex) = "--None--" Then Exit Sub

100     SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & ComLayerName.List(ComLayerName.ListIndex) & "'"
    
104     If Not g_RSGISGridTableSettings.EOF Then
106         sLyr = g_RSGISGridTableSettings.Fields.Item("name").value
            
            Select Case ComAnalysisField.ItemData(ComAnalysisField.ListIndex)
                Case 0
                    m_SQLPrefix = "'"
                Case 1, 2, 3
                    m_SQLPrefix = ""
                Case 4
                    m_SQLPrefix = "#"
            End Select
            
108         RaiseEvent GetUniqueValues(sLyr, ComAnalysisField.List(ComAnalysisField.ListIndex))
        End If
    
110     m_sLastFieldName = sLyr
    
        '<EhFooter>
        Exit Sub

ComAnalysisField_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangeTracer.ComAnalysisField_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComLayerName_Click()
        '<EhHeader>
        On Error GoTo ComLayerName_Click_Err
        '</EhHeader>
        Timer1.Enabled = False
        cmdStart.caption = "Start"

100     SafeMoveFirst g_RSGISGridTableSettings
102     g_RSGISGridTableSettings.Find "alias = '" & ComLayerName.List(ComLayerName.ListIndex) & "'"

104     If Not g_RSGISGridTableSettings.EOF Then
106         m_sLastLayer = g_RSGISGridTableSettings.Fields.Item("name").value
108         RaiseEvent GetFields(m_sLastLayer)
        End If

        '<EhFooter>
        Exit Sub

ComLayerName_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangeTracer.ComLayerName_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetTrendValues(vVals As Variant)

End Sub

Public Sub RemoveDuplicates(lst As ListView)
        '<EhHeader>
        On Error GoTo RemoveDuplicates_Err
        '</EhHeader>

        Dim lRet As ListItem
        Dim strTemp As String
        Dim intCnt As Integer
100     intCnt = 0

102     Do While intCnt <= lst.ListItems.Count - 1
        
104         intCnt = intCnt + 1
            'Save the text that was in the listvew i
            '     ndex
106         strTemp = lst.ListItems.Item(intCnt).Text

            On Error Resume Next

            Do
108             lst.ListItems.Item(intCnt).Text = "" 'Remove the text inside the specific index
                'Use the FindItem() call to search for t
                '     he specific item
110             Set lRet = lst.FindItem(strTemp, lvwText, lvwPartial)
                'If the item is found, then it is a dupl
                '     icate and is removed

112             If Not lRet Is Nothing Then
114                 lst.ListItems.Remove (lRet.Index)
                End If

116         Loop While Not lRet Is Nothing 'If no item is found the loop is exited
        
118         lst.ListItems.Item(intCnt).Text = strTemp 'reset the listitem index text back To what it was, and Then continue
120         DebugPrint intCnt

122         DoEvents 'Added To ensure that the application does Not lock up when doing large amounts of data.
            
        Loop

        '<EhFooter>
        Exit Sub

RemoveDuplicates_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmChangeTracer.RemoveDuplicates " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
    If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim sScope As String

    lblCurrentAnimation.caption = "Current Animation: N/A"

    If Not lvUniqueValues.ListItems(m_iCurListItem).Text = "" Then
        sScope = ComAnalysisField.List(ComAnalysisField.ListIndex) & " = " & m_SQLPrefix & lvUniqueValues.ListItems(m_iCurListItem).Text & m_SQLPrefix
        RaiseEvent ChangeScope(sScope, m_sLastLayer)
        lblCurrentAnimation.caption = "Current Animation: " & lvUniqueValues.ListItems(m_iCurListItem).Text
    End If

    If Not m_iCurListItem = lvUniqueValues.ListItems.Count Then
        m_iCurListItem = m_iCurListItem + 1
    Else
        m_iCurListItem = 1
    End If
End Sub
