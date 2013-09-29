VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLayerSelection 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Available Layers"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3540
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   4980
      Width           =   1215
   End
   Begin VB.CheckBox chkSelectUnselect 
      Caption         =   "Select / Unselect all"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   60
      Width           =   2655
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   8070
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Only Column"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmLayerSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
     
Dim oRS As ADODB.Recordset
Dim sLayerName As String
    
Private Sub chkSelectUnselect_Click()
Dim i As Integer
    
    For i = 1 To ListView1.ListItems.Count
        If chkSelectUnselect.value = vbUnchecked Then
            ListView1.ListItems.Item(i).Checked = False
        Else
            ListView1.ListItems.Item(i).Checked = True
        End If
    Next
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
    chkSelectUnselect.value = vbUnchecked
End Sub

Private Sub cmdOK_Click()
        '<EhHeader>
        On Error GoTo cmdOk_Click_Err
        '</EhHeader>
        Dim i As Integer
        Dim sAllLayers As String

100
    
102     If oRS.State = adStateOpen Then
            If ListView1.CheckBoxes = True Then

104             For i = 1 To ListView1.ListItems.Count

106                 If ListView1.ListItems.Item(i).Checked Then
108                     oRS.MoveFirst
110                     oRS.Find "sAlias = '" & ListView1.ListItems.Item(i).Text & "'"
                
112                     If Not oRS.EOF Then
114                         DebugPrint "Selected layer name: " & oRS.Fields("sName").value
116                         DebugPrint "Selected layer alias: " & oRS.Fields("sAlias").value
                            If Len(sLayerName) = 0 Then
                                sLayerName = oRS.Fields("sName").value
                            Else
118                             sLayerName = sLayerName & "::::" & oRS.Fields("sName").value
                            End If
                        End If

                    End If

                Next
                
                'If Not sLayerName = "" Then sLayerName = Mid$(sLayerName, 1, Len(sLayerName) - 4)

            Else
122             oRS.MoveFirst
                'oRS.Find "sAlias = '" & lstLayers.List(lstLayers.ListIndex) & "'"
124             oRS.Find "sAlias = '" & ListView1.SelectedItem.Text & "'"
        
126             If Not oRS.EOF Then
128                 DebugPrint "Selected layer name: " & oRS.Fields("sName").value
130                 DebugPrint "Selected layer alias: " & oRS.Fields("sAlias").value
132                 sLayerName = oRS.Fields("sName").value
                End If
            End If
        End If
        
        Me.Hide
        chkSelectUnselect.value = vbUnchecked
134     oRS.Close
136     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

cmdOk_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmLayerSelection.cmdOK_Click " & "at line " & Erl
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

Public Sub Init(GIS As Object, _
                Optional bAddRaster As Boolean, _
                Optional bOGISFormat As Boolean)
    
    Dim i As Integer
    Set oRS = New ADODB.Recordset
    sLayerName = ""
    oRS.Fields.Append "sName", adVarChar, 255
    oRS.Fields.Append "sAlias", adVarChar, 255
    oRS.Open
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Item(1).Width = ListView1.Width * 0.9
    'lstLayers.Clear
    ReDim sLayerNames(GIS.items.Count)
    SafeMoveFirst g_RSAppSettings
    g_RSAppSettings.Find "SettingName = 'OASIS_Incident_Layer_Name'"
    Dim sIncidentsName As String

    If Not g_RSAppSettings.EOF Then sIncidentsName = g_RSAppSettings.Fields.Item("SettingValue1").value

    For i = 0 To GIS.items.Count - 1

        If InStr(g_sPermanentLyrs, GIS.items.Item(i).Name & ",") < 1 Then
        
            If Not GIS.items.Item(i).Name = "Draw_Layer" Then 'And Not GIS.items.Item(i).Name = sIncidentsName Then
                If GisUtils.IsInherited(GIS.items.Item(i), "XGIS_LayerVector") Then
                    'lstLayers.AddItem GIS.items.Item(i).caption
                    ListView1.ListItems.Add , , GIS.items.Item(i).caption
                    oRS.AddNew
                    oRS.Fields("sName").value = GIS.items.Item(i).Name
                    oRS.Fields("sAlias").value = GIS.items.Item(i).caption
                End If

                If GisUtils.IsInherited(GIS.items.Item(i), "XGIS_LayerPixel") Then
                    If bAddRaster Then
                        'lstLayers.AddItem GIS.items.Item(i).caption
                        ListView1.ListItems.Add , , GIS.items.Item(i).caption
                        oRS.AddNew
                        oRS.Fields("sName").value = GIS.items.Item(i).Name
                        oRS.Fields("sAlias").value = GIS.items.Item(i).caption
                    End If
                End If
            End If
        
        End If
        
    Next

End Sub

Public Function GetItem() As String
        
    GetItem = sLayerName

End Function

Public Function ClearItem() As String
        
    sLayerName = ""

End Function

