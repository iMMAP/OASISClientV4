VERSION 5.00
Object = "{8D02DC4E-BFE1-4A08-9F2A-F268CB42CDFB}#3.0#0"; "Actbar3.ocx"
Begin VB.Form frmAddons 
   Caption         =   "Addons"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin ActiveBar3LibraryCtl.ActiveBar3 AB 
      Height          =   5340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      _LayoutVersion  =   2
      _ExtentX        =   8043
      _ExtentY        =   9419
      _DataPath       =   ""
      Bands           =   "frmAddons.frx":0000
   End
End
Attribute VB_Name = "frmAddons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Event MenuPressed(i As Integer)

Private Sub AB_ChildBandChange(ByVal Band As ActiveBar3LibraryCtl.Band)
        '<EhHeader>
        On Error GoTo AB_ChildBandChange_Err
        '</EhHeader>
    
100     Select Case Band.Name

            Case "cbShortcuts"
102             RaiseEvent MenuPressed(4)

104         Case "cbOperations"
106             RaiseEvent MenuPressed(1)

108         Case "cbProfile"
110             RaiseEvent MenuPressed(0)
            
112         Case "cbContent"
114             RaiseEvent MenuPressed(2)

116         Case "cbJournal"
            
        End Select

118     AB.RecalcLayout
        '<EhFooter>
        Exit Sub

AB_ChildBandChange_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddons.AB_ChildBandChange " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
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

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmAddons.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
