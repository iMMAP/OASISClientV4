Attribute VB_Name = "modEncryptUtils"
Option Explicit


Public Const BITS_TO_A_BYTE = 8
Public Const BYTES_TO_A_WORD = 4
Public Const BITS_TO_A_WORD = 32

' MD5 stuff
Public m_lOnBits(30)
Public m_l2Power(30)

Public m_InCo(3)
Public m_byt2Power(7)
Public m_bytOnBits(7)

Public m_fbsub(255)
Public m_rbsub(255)
Public m_ptab(255)
Public m_ltab(255)
Public m_ftable(255)
Public m_rtable(255)
Public m_rco(29)

Public m_Nk
Public m_Nb
Public m_Nr
Public m_fi(23)
Public m_ri(23)
Public m_fkey(119)
Public m_rkey(119)

Public Function LShift(lValue, _
                       iShiftBits)
        '<EhHeader>
        On Error GoTo LShift_Err
        '</EhHeader>

100     If iShiftBits = 0 Then
102         LShift = lValue
            Exit Function
104     ElseIf iShiftBits = 31 Then

106         If lValue And 1 Then
108             LShift = &H80000000
            Else
110             LShift = 0
            End If

            Exit Function
112     ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
114         Err.Raise 6
        End If

116     If (lValue And m_l2Power(31 - iShiftBits)) Then
118         LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
        Else
120         LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
        End If

        '<EhFooter>
        Exit Function

LShift_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modEncryptUtils.LShift " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function RShift(lValue, _
                       iShiftBits)
        '<EhHeader>
        On Error GoTo RShift_Err
        '</EhHeader>

100     If iShiftBits = 0 Then
102         RShift = lValue
            Exit Function
104     ElseIf iShiftBits = 31 Then

106         If lValue And &H80000000 Then
108             RShift = 1
            Else
110             RShift = 0
            End If

            Exit Function
112     ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
114         Err.Raise 6
        End If

116     RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)

118     If (lValue And &H80000000) Then
120         RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
        End If

        '<EhFooter>
        Exit Function

RShift_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modEncryptUtils.RShift " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function LShiftByte(bytValue, _
                           bytShiftBits)
        '<EhHeader>
        On Error GoTo LShiftByte_Err
        '</EhHeader>

100     If bytShiftBits = 0 Then
102         LShiftByte = bytValue
            Exit Function
104     ElseIf bytShiftBits = 7 Then

106         If bytValue And 1 Then
108             LShiftByte = &H80
            Else
110             LShiftByte = 0
            End If

            Exit Function
112     ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
114         Err.Raise 6
        End If

116     LShiftByte = ((bytValue And m_bytOnBits(7 - bytShiftBits)) * m_byt2Power(bytShiftBits))
        '<EhFooter>
        Exit Function

LShiftByte_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modEncryptUtils.LShiftByte " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function RShiftByte(bytValue, _
                           bytShiftBits)
        '<EhHeader>
        On Error GoTo RShiftByte_Err
        '</EhHeader>

100     If bytShiftBits = 0 Then
102         RShiftByte = bytValue
            Exit Function
104     ElseIf bytShiftBits = 7 Then

106         If bytValue And &H80 Then
108             RShiftByte = 1
            Else
110             RShiftByte = 0
            End If

            Exit Function
112     ElseIf bytShiftBits < 0 Or bytShiftBits > 7 Then
114         Err.Raise 6
        End If

116     RShiftByte = bytValue \ m_byt2Power(bytShiftBits)
        '<EhFooter>
        Exit Function

RShiftByte_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modEncryptUtils.RShiftByte " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function RotateLeft(lValue, _
                           iShiftBits)
        '<EhHeader>
        On Error GoTo RotateLeft_Err
        '</EhHeader>
100     RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
        '<EhFooter>
        Exit Function

RotateLeft_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modEncryptUtils.RotateLeft " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function RotateLeftByte(bytValue, _
                               bytShiftBits)
        '<EhHeader>
        On Error GoTo RotateLeftByte_Err
        '</EhHeader>
100     RotateLeftByte = LShiftByte(bytValue, bytShiftBits) Or RShiftByte(bytValue, (8 - bytShiftBits))
        '<EhFooter>
        Exit Function

RotateLeftByte_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modEncryptUtils.RotateLeftByte " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

