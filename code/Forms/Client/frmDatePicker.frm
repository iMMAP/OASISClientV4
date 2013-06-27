VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDatePicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS Incidents Date Picker"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5835
   Icon            =   "frmDatePicker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4410
      TabIndex        =   4
      Top             =   2925
      Width           =   1320
   End
   Begin VB.Frame FraDateTo 
      Caption         =   "Date To:"
      Height          =   2895
      Left            =   2925
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      Begin MSComCtl2.MonthView dtTo 
         CausesValidation=   0   'False
         Height          =   2370
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   254935042
         CurrentDate     =   39299
      End
   End
   Begin VB.Frame FraDateFrom 
      Caption         =   "Date From:"
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
      Begin MSComCtl2.MonthView dtFrom 
         Height          =   2370
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   254935042
         CurrentDate     =   39299
      End
   End
End
Attribute VB_Name = "frmDatePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub dtFrom_DateClick(ByVal DateClicked As Date)
    cmdOK.SetFocus
End Sub

Private Sub dtTo_DateClick(ByVal DateClicked As Date)
    cmdOK.SetFocus
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

106     dtFrom.Day = Day(Now)
108     dtFrom.Month = Month(Now)
110     dtFrom.Year = Year(Now)
    
112     dtTo.Day = Day(Now)
114     dtTo.Month = Month(Now)
116     dtTo.Year = Year(Now)
    

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmDatePicker.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

