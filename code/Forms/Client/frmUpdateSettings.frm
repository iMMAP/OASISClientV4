VERSION 5.00
Begin VB.Form frmUpdateSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Synch/Update Settings"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3405
   Icon            =   "frmUpdateSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Comms"
      Height          =   255
      Left            =   3960
      TabIndex        =   34
      Top             =   240
      Width           =   1365
   End
   Begin VB.Frame FraGeneralSettings 
      Caption         =   "General Settings:"
      Height          =   555
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   3405
      Begin VB.CheckBox chkForceZero 
         Caption         =   "Force Zero"
         Height          =   255
         Left            =   150
         TabIndex        =   29
         Top             =   210
         Width           =   1095
      End
      Begin VB.CheckBox chkManualSynchronisation 
         Caption         =   "Manual Synch"
         Height          =   225
         Left            =   150
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdSynchronizeAll 
      Caption         =   "Synch All"
      Height          =   435
      Left            =   3600
      TabIndex        =   24
      Top             =   6210
      Width           =   795
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   930
      TabIndex        =   14
      Top             =   5070
      Width           =   795
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   315
      Left            =   1800
      TabIndex        =   13
      Top             =   5070
      Width           =   795
   End
   Begin VB.Frame FraSynchUpdate 
      Caption         =   "Synch/Update Methods"
      Height          =   1155
      Left            =   3450
      TabIndex        =   8
      Top             =   810
      Width           =   3405
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   0
         Left            =   2250
         TabIndex        =   15
         Top             =   810
         Width           =   1035
      End
      Begin VB.OptionButton OptMethod 
         Caption         =   "Offline"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1395
      End
      Begin VB.OptionButton OptMethod 
         Caption         =   "Internet Single"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   510
         Width           =   1395
      End
      Begin VB.OptionButton OptMethod 
         Caption         =   "Internet Batch"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.Frame FraExcludedUpdates 
      Caption         =   "Excluded Updates/Synchs"
      Height          =   4425
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3405
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   1
         Left            =   2250
         TabIndex        =   33
         Top             =   300
         Width           =   1035
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   2
         Left            =   2250
         TabIndex        =   32
         Top             =   630
         Width           =   1035
      End
      Begin VB.CheckBox chkDynamicData 
         Caption         =   "Dynamic Data Defs"
         Height          =   345
         Left            =   180
         TabIndex        =   31
         Top             =   3900
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkFeedsDynamic 
         Caption         =   "Feeds/Dynamic Data"
         Height          =   405
         Left            =   180
         TabIndex        =   30
         Top             =   3480
         Value           =   1  'Checked
         Width           =   1635
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   11
         Left            =   2250
         TabIndex        =   26
         Top             =   3930
         Width           =   1035
      End
      Begin VB.CheckBox chkThematics 
         Caption         =   "Thematics"
         Height          =   255
         Left            =   180
         TabIndex        =   25
         Top             =   3180
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   10
         Left            =   2250
         TabIndex        =   23
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   9
         Left            =   2250
         TabIndex        =   22
         Top             =   3240
         Width           =   1035
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   8
         Left            =   2250
         TabIndex        =   21
         Top             =   2880
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   7
         Left            =   2250
         TabIndex        =   20
         Top             =   2490
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   6
         Left            =   2250
         TabIndex        =   19
         Top             =   2130
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   5
         Left            =   2250
         TabIndex        =   18
         Top             =   1740
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   4
         Left            =   2250
         TabIndex        =   17
         Top             =   1350
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdSynchronize 
         Caption         =   "Synch Now"
         Height          =   255
         Index           =   3
         Left            =   2250
         TabIndex        =   16
         Top             =   990
         Width           =   1035
      End
      Begin VB.CheckBox chkCharts 
         Caption         =   "Charts"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   2850
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkAutoUpdate 
         Caption         =   "Auto Update"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   2478
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkMapProducts 
         Caption         =   "Map Products"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   2110
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkPrintTemplates 
         Caption         =   "Print Templates"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1742
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkGeoMarks 
         Caption         =   "Geo Marks"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   1374
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkSynchronisationLayers 
         Caption         =   "Synchronisation Layers"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   1006
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkGISAttribute 
         Caption         =   "GIS Attribute Grid"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   638
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkApplicationSettings 
         Caption         =   "Application Settings"
         Height          =   255
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Value           =   1  'Checked
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmUpdateSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event DoSynch(Index As Integer)
Public Event DoCancel()
Public Event DoApply()
Public Event DoSynchAll()

Private Sub cmdApply_Click()
    With g_udtSynchUpdateOptions
    
        .ApplicationSettings = IIf(chkApplicationSettings.Value = vbChecked, True, False)
        .AutoUpdate = IIf(chkAutoUpdate.Value = vbChecked, True, False)
        .Charts = IIf(chkCharts.Value = vbChecked, True, False)
        .GeoMarks = IIf(chkGeoMarks.Value = vbChecked, True, False)
        .GISAttributeSettings = IIf(chkGISAttribute.Value = vbChecked, True, False)
        
        If OptMethod(0).Value = True Then
            .lMethod = 0
        Else
            .lMethod = IIf(OptMethod(1).Value, 1, 2)
        End If
        
        .ManualSynchronisation = IIf(chkManualSynchronisation.Value = vbChecked, True, False)
        .MapProducts = IIf(chkMapProducts.Value = vbChecked, True, False)
        .PrintTemplates = IIf(chkPrintTemplates.Value = vbChecked, True, False)
        .SynchLayersSettings = IIf(chkSynchronisationLayers.Value = vbChecked, True, False)
        .Thematics = IIf(chkThematics.Value = vbChecked, True, False)
        .Feeds = IIf(chkFeedsDynamic.Value = vbChecked, True, False)
        .DynamDataDefs = IIf(chkDynamicData.Value = vbChecked, True, False)
    End With
    
    RaiseEvent DoApply
    
    Me.Hide
    
End Sub

Private Sub cmdCancel_Click()
    RaiseEvent DoCancel
    Me.Hide
End Sub


Private Sub cmdSynchronize_Click(Index As Integer)
 RaiseEvent DoSynch(Index)
End Sub

Private Sub cmdSynchronizeAll_Click()
    RaiseEvent DoSynchAll
End Sub

Public Sub Init()

    With g_udtSynchUpdateOptions
        chkApplicationSettings.Value = IIf(.ApplicationSettings, vbChecked, vbUnchecked)
        chkAutoUpdate.Value = IIf(.AutoUpdate, vbChecked, vbUnchecked)
        chkCharts.Value = IIf(.Charts, vbChecked, vbUnchecked)
        chkForceZero.Value = IIf(.ForceZero, vbChecked, vbUnchecked)
        chkGeoMarks.Value = IIf(.GeoMarks, vbChecked, vbUnchecked)
        chkGISAttribute.Value = IIf(.GISAttributeSettings, vbChecked, vbUnchecked)
        chkManualSynchronisation.Value = IIf(.ManualSynchronisation, vbChecked, vbUnchecked)
        chkMapProducts.Value = IIf(.MapProducts, vbChecked, vbUnchecked)
        chkPrintTemplates.Value = IIf(.PrintTemplates, vbChecked, vbUnchecked)
        chkSynchronisationLayers.Value = IIf(.SynchLayersSettings, vbChecked, vbUnchecked)
        chkThematics.Value = IIf(.Thematics, vbChecked, vbUnchecked)
        OptMethod(.lMethod).Value = True
        chkDynamicData.Value = IIf(.DynamDataDefs, vbChecked, vbUnchecked)
        chkFeedsDynamic.Value = IIf(.Feeds, vbChecked, vbUnchecked)
    End With

End Sub

Private Sub cmdTest_Click()
        '<EhHeader>
        On Error GoTo cmdTest_Click_Err
        '</EhHeader>
        '100     MsgBox "Do you speak English", vbQuestion
        Dim oRS As ADODB.Recordset
        Dim sString As String

101     'sString = g_sAppServerPath & "/oasis4.asp?ID=" & CheckEncrypt("SELECT * FROM AppSettings")
102     'Set oRS = OpenSilentHttpCommsRS(sString, True)
        Set oRS = OpenServerRSCompressed(g_sAppServerPath & "/oasis4.asp", "id", "SELECT * FROM AppSettings")
108     MsgBox oRS.GetString
    
110     Set oRS = Nothing

        '<EhFooter>
        Exit Sub

cmdTest_Click_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmUpdateSettings.cmdTest_Click " & "at line " & Erl
       
        '</EhFooter>
End Sub

Private Sub Form_Load()
    Init
End Sub

