VERSION 5.00
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Client Synch Simulation"
   ClientHeight    =   2070
   ClientLeft      =   5730
   ClientTop       =   3135
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClient_sync.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdIncidentSynch 
      Caption         =   "Incident Synch test..."
      Height          =   465
      Left            =   1680
      TabIndex        =   6
      Top             =   1080
      Width           =   1485
   End
   Begin VB.CommandButton cmdRunInternet 
      Caption         =   "Run Internet Test..."
      Height          =   435
      Left            =   90
      TabIndex        =   5
      Top             =   1080
      Width           =   1515
   End
   Begin VB.TextBox txtIntervall 
      Height          =   315
      Left            =   1230
      TabIndex        =   4
      Text            =   "30"
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdNonAsync 
      Caption         =   "UN-sync..."
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Sto&p"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3180
      TabIndex        =   1
      Top             =   30
      Width           =   1515
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start..."
      Height          =   315
      Left            =   1650
      TabIndex        =   0
      Top             =   30
      Width           =   1515
   End
   Begin VB.Label lblInfo 
      Caption         =   "Test Intervall:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   1125
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_bCancel As Boolean
Private WithEvents m_oSyncWorker As SynchWorker
Attribute m_oSyncWorker.VB_VarHelpID = -1

Private Sub cmdIncidentSynch_Click()
    
    With m_oSyncWorker
    
        If .IsRunning Then Exit Sub
        
        m_bCancel = False
        .Paddy = False
        .RType = IncidentSynch
        .LocalConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\OASIS\Client\data\db\Oasisclient.mdb;"
        .RemoteTablePrefix = "iMMAP"
        .WebsiteURL = "http://afghanistan.humanitariansecurity.org"
        .HasEncrypt = False
        .Start
    
    End With
 
End Sub

Private Sub cmdRunInternet_Click()
   
   If m_oSyncWorker.IsRunning Then Exit Sub
   
   m_bCancel = False
   m_oSyncWorker.Paddy = False
   m_oSyncWorker.RType = InternetConnectionCheck
   m_oSyncWorker.Start
End Sub

Private Sub cmdStart_Click()
   If m_oSyncWorker.IsRunning Then Exit Sub
   
   m_bCancel = False
   m_oSyncWorker.Interval = CInt(txtIntervall.Text)
   m_oSyncWorker.Paddy = True
   m_oSyncWorker.Start
   cmdStop.Enabled = True
   cmdStart.Enabled = False
   cmdNonAsync.Enabled = False
End Sub

Private Sub cmdStop_Click()
   cmdStart.Enabled = False
   cmdStop.Enabled = False
   cmdNonAsync.Enabled = False
   m_bCancel = True
End Sub

Private Sub cmdNonAsync_Click()
   m_oSyncWorker.Interval = CInt(txtIntervall.Text)
   m_oSyncWorker.Paddy = True
   m_oSyncWorker.StartNonAsync
End Sub

Private Sub Form_Load()
   Set m_oSyncWorker = New SynchWorker
   m_oSyncWorker.Interval = 20
End Sub

Private Sub m_oSyncWorker_Cancelled()
   MsgBox "Cancel", vbInformation
   cmdStart.Enabled = True
   cmdNonAsync.Enabled = True
   cmdStop.Enabled = False
End Sub

Private Sub m_oSyncWorker_Complete()
   MsgBox "Complete", vbInformation
   cmdStart.Enabled = True
   cmdNonAsync.Enabled = True
   cmdStop.Enabled = False
End Sub

Private Sub m_oSyncWorker_CompleteEX(enmRunnerType As OASIS_Synch.RunnerType)
    Select Case enmRunnerType
    
        Case OASIS_Synch.enmPaddy
    
        Case OASIS_Synch.GeoMarks
    
        Case OASIS_Synch.IncidentSynch
            MsgBox "INCIDENT SYNC DONE..."
        Case OASIS_Synch.InternetConnectionCheck
            MsgBox "Connected:" & m_oSyncWorker.Connected & " Type:" & m_oSyncWorker.InternetConnectionType
        Case OASIS_Synch.SQLLyrSynch
    
    End Select
End Sub

Private Sub m_oSyncWorker_Status(ByVal i As Long, Cancel As Boolean)
   Debug.Print i
   Cancel = m_bCancel
End Sub
