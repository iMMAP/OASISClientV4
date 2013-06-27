VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const IDHOT_SNAPWINDOW = -1          '/* SHIFT-PRINTSCRN  */
Private Const IDHOT_SNAPDESKTOP = -2         '/* PRINTSCRN        */
Public Enum EHKModifiers
    MOD_ALT = &H1&
    MOD_CONTROL = &H2&
    MOD_SHIFT = &H4&
    MOD_WIN = &H8&
End Enum
Private Const WM_DESTROY = &H2&
Private Const WM_HOTKEY = &H312&
Private Declare Function RegisterHotKey _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal ID As Long, _
                              ByVal fsModifiers As Long, _
                              ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal ID As Long) As Long
Private Declare Function GlobalAddAtom _
                Lib "kernel32" _
                Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Private Declare Function GlobalDeleteAtom _
                Lib "kernel32" (ByVal nAtom As Integer) As Integer
' To Report API errors:
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage _
                Lib "kernel32" _
                Alias "FormatMessageA" (ByVal dwFlags As Long, _
                                        lpSource As Any, _
                                        ByVal dwMessageId As Long, _
                                        ByVal dwLanguageId As Long, _
                                        ByVal lpBuffer As String, _
                                        ByVal nSize As Long, _
                                        Arguments As Long) As Long
Private Declare Function GetTickCount _
                Lib "kernel32" () As Long
Private Declare Function SendMessageByLong _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_RESTORE = &HF120&
Private Declare Function IsIconic _
                Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible _
                Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetForegroundWindow _
                Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ShowWindow _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal nCmdShow As Long) As Long
Private Const SW_SHOW = 5

' Implementation
Private Type tHotKeyInfo
    sName As String
    sAtomName As String
    lID As Long
    eKey As KeyCodeConstants
    eModifiers As EHKModifiers
End Type
Private m_tAtoms() As tHotKeyInfo
Private m_iAtomCount As Long
Private m_hWnd As Long

Public Event HotKeyPress(ByVal sName As String, ByVal eModifiers As EHKModifiers, ByVal eKey As KeyCodeConstants)

Implements ISubclass

Public Sub Attach(ByVal hWndA As Long)
    Clear

    If (hWndA <> 0) Then
        m_hWnd = hWndA
        AttachMessage Me, m_hWnd, WM_HOTKEY
        AttachMessage Me, m_hWnd, WM_DESTROY
    End If

End Sub

Public Sub Clear()
    Dim i As Long

    ' Remove all hot keys and atoms:
    For i = m_iAtomCount To 1 Step -1
        UnregisterKey m_tAtoms(i).sName
    Next i

    ' Stop subclassing:
    If (m_hWnd <> 0) Then
        DetachMessage Me, m_hWnd, WM_HOTKEY
        DetachMessage Me, m_hWnd, WM_DESTROY
        m_hWnd = 0
    End If

End Sub

Public Sub RegisterKey(ByVal sName As String, _
                       ByVal eKey As KeyCodeConstants, _
                       ByVal eModifiers As EHKModifiers)
    Dim lID As Long
    Dim lErr As Long
    Dim lR As Long
    Dim sError As String
    Dim sMsg As String
    Dim i As Long
    Dim sAtomName As String

    ' Check for valid user name:
    If Len(sName) > 32 Then
        Err.Raise vbObjectError + 1048 + 3, App.EXEName & ".cRegHotKey", "Key Name too long (max 32 characters)."
        Exit Sub
    Else

        For i = 1 To m_iAtomCount

            If (m_tAtoms(i).sName = sName) Then
                Err.Raise vbObjectError + 1048 + 4, App.EXEName & ".cRegHotKey", "The Key Name '" & sName & "' is already registered."
                Exit Sub
            End If

        Next i

    End If

    ' Modify the user supplied name to get a more random system name:
    sAtomName = sName & "_" & App.EXEName & "_" & GetTickCount()

    If (Len(sAtomName) > 254) Then
        sAtomName = left$(sAtomName, 254)
    End If

    ' Create a new atom:
    lID = GlobalAddAtom(sAtomName)

    If (lID = 0) Then
        lErr = Err.LastDllError
        sError = WinError(lErr)
        sMsg = "Failed to add GlobalAtom"

        If (sError <> "") Then
            sMsg = sMsg & " [" & sError & "]"
        End If

        Err.Raise vbObjectError + 1048 + 2, App.EXEName & ".cRegHotKey", sMsg
    Else
        ' We have added the atom, now try to Register the
        ' key:
        lR = RegisterHotKey(m_hWnd, lID, eModifiers, eKey)

        If (lR = 0) Then
            lErr = Err.LastDllError
            ' Remove the atom:
            GlobalDeleteAtom lID
            ' Raise the error:
            WinError lErr
            sError = WinError(lErr)
            sMsg = "Failed to Register Hot Key"

            If (sError <> "") Then
                sMsg = sMsg & " [" & sError & "]"
            End If

            'Err.Raise vbObjectError + 1048 + 3, App.EXEName & ".cRegHotKey", sMsg
        Else
            ' Succeeded in adding the hot key:
            m_iAtomCount = m_iAtomCount + 1
            ReDim Preserve m_tAtoms(1 To m_iAtomCount) As tHotKeyInfo

            With m_tAtoms(m_iAtomCount)
                .sName = sName
                .sAtomName = sAtomName
                .lID = lID
                .eModifiers = eModifiers
                .eKey = eKey
            End With

        End If
         
    End If

End Sub

Public Sub UnregisterKey(ByVal sName As String)
    Dim lIndex As Long
    Dim i As Long
    lIndex = AtomIndex(sName)

    If (lIndex > 0) Then
        ' Unregister the key:
        UnregisterHotKey m_hWnd, m_tAtoms(lIndex).lID
        ' Unregister the atom:
        GlobalDeleteAtom m_tAtoms(lIndex).lID

        ' Remove from internal array:
        If (m_iAtomCount > 1) Then

            For i = lIndex To m_iAtomCount - 1
                LSet m_tAtoms(lIndex) = m_tAtoms(lIndex + 1)
            Next i

            m_iAtomCount = m_iAtomCount - 1
            ReDim Preserve m_tAtoms(1 To m_iAtomCount) As tHotKeyInfo
        Else
            m_iAtomCount = 0
            Erase m_tAtoms
        End If
    End If

End Sub
Private Property Get AtomIndex(ByVal sName As String) As Long
    Dim i As Long

    For i = 1 To m_iAtomCount

        If (m_tAtoms(i).sName = sName) Then
            AtomIndex = i
            Exit Property
        End If

    Next i

    'Err.Raise vbObjectError + 1048 + 1, App.EXEName & ".cRegHotKey", "No hot key registered under the name '" & sName & "'"
End Property

Private Function WinError(ByVal lLastDLLError As Long) As String
    Dim sBuff As String
    Dim lCount As Long
    
    ' Return the error message associated with LastDLLError:
    sBuff = String$(256, 0)
    lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)

    If lCount Then
        WinError = left$(sBuff, lCount)
    End If
    
End Function

Public Sub RestoreAndActivate(ByVal hWnd As Long)

    If (IsWindowVisible(hWnd) = 0) Then
        ShowWindow hWnd, SW_SHOW
    End If

    If (IsIconic(hWnd) <> 0) Then
        SendMessageByLong hWnd, WM_SYSCOMMAND, SC_RESTORE, 0
    End If

    SetForegroundWindow hWnd
End Sub

Private Sub Class_Terminate()
    Clear
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    ' ...
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, _
                                      ByVal iMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long
    Dim i As Long
    Dim lIndex As Long

    Select Case iMsg

        Case WM_HOTKEY

            ' Interpret the hotkey.  wParam is the id, the
            ' loword of lParam is the key modifier and the
            ' hiword of lParam is the key code:
            Select Case wParam

                Case IDHOT_SNAPWINDOW
                    RaiseEvent HotKeyPress("Window Snapshot", MOD_SHIFT, vbKeySnapshot)

                Case IDHOT_SNAPDESKTOP
                    RaiseEvent HotKeyPress("Desktop Snapshot", 0, vbKeySnapshot)

                Case Else

                    ' Try to find id:
                    For i = 1 To m_iAtomCount

                        If (m_tAtoms(i).lID = wParam) Then
                            lIndex = i
                            Exit For
                        End If

                    Next i

                    If (lIndex <> 0) Then
                        RaiseEvent HotKeyPress(m_tAtoms(lIndex).sName, m_tAtoms(lIndex).eModifiers, m_tAtoms(lIndex).eKey)
                    Else
                        ' What has happened?
                        RaiseEvent HotKeyPress("Unknown HotKey", (lParam And &HFFFF&), (lParam \ &H10000))
                    End If

            End Select

        Case WM_DESTROY
            Clear
    End Select

End Function
