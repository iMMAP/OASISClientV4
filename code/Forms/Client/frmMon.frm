VERSION 5.00
Begin VB.Form frmMon 
   Caption         =   "OASISCommsMon"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmr_Sequence 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   210
      Top             =   180
   End
End
Attribute VB_Name = "frmMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oSync As SynchWorker

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Set oSync = New SynchWorker
        If oSync Is Nothing Then GoTo Form_Load_Err:
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISCommsMon.frmMon.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub StartNow()
        '<EhHeader>
        On Error GoTo StartNow_Err
        '</EhHeader>
100     With oSync
            
102         If Not .IsRunning Then
        
104             .Paddy = False
106             .RType = sArgs(0)
108             .LocalConnectionString = sArgs(1)
110             .RemoteTablePrefix = sArgs(2)
112             .WebsiteURL = sArgs(3)
114             .HasEncrypt = CBool(sArgs(4))
116             .EncryptKey = sArgs(5)
                .EnableGeoMarkSynch = sArgs(8)
                .EnableGeoMarkSynch = sArgs(9)
                .InitComms
118             .Start

            End If

        End With
        '<EhFooter>
        Exit Sub

StartNow_Err:

        Resume Next
        '</EhFooter>
End Sub

Private Sub tmr_Sequence_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    With oSync
            
        If Not .IsRunning Then
        
            .Paddy = False
            .RType = sArgs(0)
            .LocalConnectionString = sArgs(1)
            .RemoteTablePrefix = sArgs(2)
            .WebsiteURL = sArgs(3)
            .HasEncrypt = CBool(sArgs(4))
            .EncryptKey = sArgs(5)
            .EnableGeoMarkSynch = sArgs(8)
                .EnableGeoMarkSynch = sArgs(9)
            .Start

        End If

    End With
End Sub
