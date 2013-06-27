VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExplorer"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_frmSynchReader As frmSynchReader

Public Sub ShowExplorer(wHWND As Long)
        '<EhHeader>
        On Error GoTo ShowExplorer_Err
        '</EhHeader>
    
        'If m_frmSynchReader Is Nothing Then
        '    Set m_frmSynchReader = New frmSynchReader
        'End If
    
100     m_frmSynchReader.ClearAll
102     m_frmSynchReader.Show 'vbModeless, wHWND
        '<EhFooter>
        Exit Sub

ShowExplorer_Err:
        Err.Raise vbObjectError + 100, _
                  "Synch_Explorer.clsExplorer.ShowExplorer", _
                  "clsExplorer component failure"
        '</EhFooter>
End Sub


Private Sub Class_Initialize()
        '<EhHeader>
        On Error GoTo Class_Initialize_Err
        '</EhHeader>
100     Set m_frmSynchReader = New frmSynchReader
        '<EhFooter>
        Exit Sub

Class_Initialize_Err:
        Err.Raise vbObjectError + 100, _
                  "Synch_Explorer.clsExplorer.Class_Initialize", _
                  "clsExplorer component failure"
        '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>
On Error Resume Next
    Unload m_frmSynchReader
End Sub
