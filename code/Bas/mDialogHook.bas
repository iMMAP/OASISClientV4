Attribute VB_Name = "mDialogHook"
Option Explicit

' ==========================================================================
' Provides functions which can be called via AddressOf for common
' dialog hook support.
' ==========================================================================
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (lpvDest As Any, _
                                       lpvSource As Any, _
                                       ByVal cbCopy As Long)

Private m_cHookedDialog As Long

Property Let HookedDialog(ByRef cThis As GCommonDialog)
    'Set cHookedDialog = cThis
    m_cHookedDialog = ObjPtr(cThis)
End Property
Property Get HookedDialog() As GCommonDialog
    Dim oT As GCommonDialog

    If (m_cHookedDialog <> 0) Then
        ' Turn the pointer into an illegal, uncounted interface
        CopyMemory oT, m_cHookedDialog, 4
        ' Do NOT hit the End button here! You will crash!
        ' Assign to legal reference
        Set HookedDialog = oT
        ' Still do NOT hit the End button here! You will still crash!
        ' Destroy the illegal reference
        CopyMemory oT, 0&, 4
    End If

End Property

Public Sub ClearHookedDialog()
    m_cHookedDialog = 0
End Sub

Public Function DialogHookFunction(ByVal hDlg As Long, _
                                   ByVal msg As Long, _
                                   ByVal wParam As Long, _
                                   ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        DialogHookFunction = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function PrintHookProc(ByVal hDlg As Long, _
                              ByVal msg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        PrintHookProc = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function PrintSetupHookProc(ByVal hDlg As Long, _
                                   ByVal msg As Long, _
                                   ByVal wParam As Long, _
                                   ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        PrintSetupHookProc = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function PageSetupHook(ByVal hDlg As Long, _
                              ByVal msg As Long, _
                              ByVal wParam As Long, _
                              ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        PageSetupHook = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function CCHookProc(ByVal hDlg As Long, _
                           ByVal msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        CCHookProc = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

Public Function CFHookProc(ByVal hDlg As Long, _
                           ByVal msg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
    Dim CD As GCommonDialog
    Set CD = HookedDialog

    If Not (CD Is Nothing) Then
        CFHookProc = CD.DialogHook(hDlg, msg, wParam, lParam)
    End If

End Function

