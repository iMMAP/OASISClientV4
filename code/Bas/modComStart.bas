Attribute VB_Name = "modComStart"
Option Explicit
Public sArgs() As String
Public m_frmMon As frmMon

Public Sub Main()
        '<EhHeader>
        On Error GoTo Main_Err
        '</EhHeader>
        Dim i As Integer
    
100     sArgs = Split(Command$, "^")
102     Set m_frmMon = New frmMon
104     Load m_frmMon
106     m_frmMon.Visible = False
108     m_frmMon.StartNow
110     m_frmMon.tmr_Sequence.Interval = CLng(sArgs(UBound(sArgs) - 2))
112     m_frmMon.tmr_Sequence.Enabled = True
        
        '1^connectionString^RemoteTableprefix^ServerPath^HasEncrypt^sKey^HWND^interval
        
        '<EhFooter>
        Exit Sub

Main_Err:

        Resume Next
        '</EhFooter>
End Sub
