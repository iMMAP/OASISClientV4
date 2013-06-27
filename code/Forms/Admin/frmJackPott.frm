VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2010
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   2490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2310
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1725
      Top             =   2730
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1065
      Top             =   2790
   End
   Begin VB.CommandButton Spin 
      Caption         =   "Spin IT!"
      Height          =   270
      Left            =   180
      TabIndex        =   0
      Top             =   1665
      Width           =   2100
   End
   Begin MSComctlLib.ImageList IList1 
      Left            =   5490
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   46
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   45
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":0000
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":28D8
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":5041
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":77AA
            Key             =   ""
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":9F13
            Key             =   ""
            Object.Tag             =   "5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":C708
            Key             =   ""
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":EF56
            Key             =   ""
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":116E6
            Key             =   ""
            Object.Tag             =   "8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":13F34
            Key             =   ""
            Object.Tag             =   "9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":16729
            Key             =   ""
            Object.Tag             =   "10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":18F1E
            Key             =   ""
            Object.Tag             =   "11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":1B713
            Key             =   ""
            Object.Tag             =   "12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":1DFEB
            Key             =   ""
            Object.Tag             =   "13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":2077B
            Key             =   ""
            Object.Tag             =   "14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":22F0B
            Key             =   ""
            Object.Tag             =   "15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":2569B
            Key             =   ""
            Object.Tag             =   "16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":27F73
            Key             =   ""
            Object.Tag             =   "17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":2A7C9
            Key             =   ""
            Object.Tag             =   "18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":2D0F1
            Key             =   ""
            Object.Tag             =   "19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":2F85A
            Key             =   ""
            Object.Tag             =   "20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":32182
            Key             =   ""
            Object.Tag             =   "21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":34865
            Key             =   ""
            Object.Tag             =   "22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":36F48
            Key             =   ""
            Object.Tag             =   "23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":3962B
            Key             =   ""
            Object.Tag             =   "24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":3BD94
            Key             =   ""
            Object.Tag             =   "25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":3E5EA
            Key             =   ""
            Object.Tag             =   "26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":40E40
            Key             =   ""
            Object.Tag             =   "27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":43696
            Key             =   ""
            Object.Tag             =   "28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":45E26
            Key             =   ""
            Object.Tag             =   "29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":48674
            Key             =   ""
            Object.Tag             =   "30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":4AE69
            Key             =   ""
            Object.Tag             =   "31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":4D49E
            Key             =   ""
            Object.Tag             =   "32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":4FAD3
            Key             =   ""
            Object.Tag             =   "33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":52108
            Key             =   ""
            Object.Tag             =   "34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":54898
            Key             =   ""
            Object.Tag             =   "35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":57028
            Key             =   ""
            Object.Tag             =   "36"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":5987E
            Key             =   ""
            Object.Tag             =   "37"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":5BF61
            Key             =   ""
            Object.Tag             =   "38"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":5E889
            Key             =   ""
            Object.Tag             =   "39"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":611B1
            Key             =   ""
            Object.Tag             =   "40"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":63AD9
            Key             =   ""
            Object.Tag             =   "41"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":66242
            Key             =   ""
            Object.Tag             =   "42"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":68A37
            Key             =   ""
            Object.Tag             =   "43"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":6B06C
            Key             =   ""
            Object.Tag             =   "44"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJackPott.frx":6D6A1
            Key             =   ""
            Object.Tag             =   "45"
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   195
      Top             =   150
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   4
      X1              =   120
      X2              =   2235
      Y1              =   105
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   3
      X1              =   210
      X2              =   2220
      Y1              =   1575
      Y2              =   150
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   2
      X1              =   180
      X2              =   2235
      Y1              =   1305
      Y2              =   1305
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   180
      X2              =   2220
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   180
      X2              =   2235
      Y1              =   825
      Y2              =   825
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   1545
      Picture         =   "frmJackPott.frx":6FCD6
      Top             =   1095
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   870
      Picture         =   "frmJackPott.frx":7259E
      Top             =   1095
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   180
      Picture         =   "frmJackPott.frx":74E66
      Top             =   1095
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   1545
      Picture         =   "frmJackPott.frx":7772E
      Top             =   615
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   870
      Picture         =   "frmJackPott.frx":79FF6
      Top             =   615
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   180
      Picture         =   "frmJackPott.frx":7C8BE
      Top             =   615
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   1545
      Picture         =   "frmJackPott.frx":7F186
      Top             =   150
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   870
      Picture         =   "frmJackPott.frx":81A4E
      Top             =   150
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   180
      Picture         =   "frmJackPott.frx":84316
      Top             =   150
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Randomize
102     r1 = Int((Rnd * 43) + 1)

104     r2 = Int((Rnd * 43) + 1)

106     r3 = Int((Rnd * 43) + 1)

108     Image1(0).Picture = IList1.ListImages(r1).Picture
110     Image1(1).Picture = IList1.ListImages(r2).Picture
112     Image1(2).Picture = IList1.ListImages(r3).Picture

114     Image1(3).Picture = IList1.ListImages(r1 + 1).Picture
116     Image1(4).Picture = IList1.ListImages(r2 + 1).Picture
118     Image1(5).Picture = IList1.ListImages(r3 + 1).Picture

120     Image1(6).Picture = IList1.ListImages(r1 + 2).Picture
122     Image1(7).Picture = IList1.ListImages(r2 + 2).Picture
124     Image1(8).Picture = IList1.ListImages(r3 + 2).Picture

126     cindex1 = r1
128     cindex2 = r2
130     cindex3 = r3
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.Form1.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Spin_Click()
        '<EhHeader>
        On Error GoTo Spin_Click_Err
        '</EhHeader>
100     Randomize
102     r1 = Int((Rnd * 150) + 101)
104     Randomize
106     r2 = Int((Rnd * 100) + r1)
108     Randomize
110     r3 = Int((Rnd * 50) - r2)
112     rounds1 = 100
114     rounds2 = 100
116     rounds3 = 100
118     Timer1.Enabled = True
120     Timer2.Enabled = True
122     Timer3.Enabled = True
        '<EhFooter>
        Exit Sub

Spin_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.Form1.Spin_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Timer1_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim scount As Integer
    Dim xcount As Integer
    
    
    rounds1 = rounds1 + 1
    cindex1 = cindex1 + 1
    counter1 = cindex1

    If counter1 = 46 Then
        cindex1 = 1
        counter1 = 1
    End If

    Image1(0).Picture = IList1.ListImages(counter1).Picture

    If counter1 = 1 Then
        scount = 45
    Else
        scount = counter1 - 1
    End If

    Image1(3).Picture = IList1.ListImages(scount).Picture

    If scount = 1 Then
        xcount = 45
    Else
        xcount = scount - 1
    End If

    Image1(6).Picture = IList1.ListImages(xcount).Picture

    If rounds1 = r1 Then
        Timer1.Enabled = False
    End If

End Sub

Private Sub Timer2_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
        
    Dim scount1 As Integer
    Dim xcount1 As Integer
    
    rounds2 = rounds2 + 1
    cindex2 = cindex2 + 1
    counter2 = cindex2

    If counter2 = 46 Then
        cindex2 = 1
        counter2 = 1
    End If

    Image1(1).Picture = IList1.ListImages(counter2).Picture

    If counter2 = 1 Then
        scount1 = 45
    Else
        scount1 = counter2 - 1
    End If

    Image1(4).Picture = IList1.ListImages(scount1).Picture

    If scount1 = 1 Then
        xcount1 = 45
    Else
        xcount1 = scount1 - 1
    End If

    Image1(7).Picture = IList1.ListImages(xcount1).Picture

    If rounds2 = r1 Then
        Timer2.Enabled = False
    End If

End Sub

Private Sub Timer3_Timer()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
    Dim scount2 As Integer
    Dim xcount2 As Integer
    
    rounds3 = rounds3 + 1
    cindex3 = cindex3 + 1
    counter3 = cindex3

    If counter3 = 46 Then
        cindex3 = 1
        counter3 = 1
    End If

    Image1(2).Picture = IList1.ListImages(counter3).Picture

    If counter3 = 1 Then
        scount2 = 45
    Else
        scount2 = counter3 - 1
    End If

    Image1(5).Picture = IList1.ListImages(scount2).Picture

    If scount2 = 1 Then
        xcount2 = 45
    Else
        xcount2 = scount2 - 1
    End If

    Image1(8).Picture = IList1.ListImages(xcount2).Picture

    If rounds3 = r2 Then
        Timer3.Enabled = False
    End If

End Sub
