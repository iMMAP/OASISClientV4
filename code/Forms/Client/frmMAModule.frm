VERSION 5.00
Begin VB.Form frmMAModule 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mine Action Module"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2520
   Icon            =   "frmMAModule.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   2520
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox MineActionAddon1 
      BackColor       =   &H000000FF&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   945
      TabIndex        =   0
      Top             =   0
      Width           =   1000
   End
End
Attribute VB_Name = "frmMAModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event LoadMALyr(oLyr As TatukGIS_XDK9.XGIS_LayerVector)
Public Event UnloadMALyr(sLyrName As String)

Public Sub Init(RSSettings As adodb.Recordset)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
100     'MineActionAddon1.Init RSSettings
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMAModule.init " & _
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

Private Sub MineActionAddon1_LoadMALyr(oLyr As TatukGIS_XDK9.XGIS_LayerVector)
        '<EhHeader>
        On Error GoTo MineActionAddon1_LoadMALyr_Err
        '</EhHeader>
100      RaiseEvent LoadMALyr(oLyr)
         'm_frmMain.AddLayer oLyr
        '<EhFooter>
        Exit Sub

MineActionAddon1_LoadMALyr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMAModule.MineActionAddon1_LoadMALyr " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub MineActionAddon1_UnloadMALyr(sLyrName As String)
        '<EhHeader>
        On Error GoTo MineActionAddon1_UnloadMALyr_Err
        '</EhHeader>
100     RaiseEvent UnloadMALyr(sLyrName)
        '<EhFooter>
        Exit Sub

MineActionAddon1_UnloadMALyr_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMAModule.MineActionAddon1_UnloadMALyr " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


