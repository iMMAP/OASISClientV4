VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBlowfish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic Blowfish Implementation
'
' Standard Blowfish implementation with file support, Base64 conversion,
' and overall optimisations for Visual Basic. Blowfish is considered one
' of the strongest encryption algorithms on the market and is much faster
' than the IDEA cipher. It supports variable length keys up to 448-bits.
' would recommend this cipher for high security risk related solutions since
' it is unpatented and free for use.
'
' Information on the Blowfish algorithm can be found at:
' http://www.counterpane.com/blowfish.html

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Event Progress(Percent As Long)

Private Const Rounds = 16

Private m_pBox(0 To Rounds + 1) As Long
Private m_sBox(0 To 3, 0 To 255) As Long
Private m_KeyValue As String
Private m_RunningCompiled As Boolean
Private m_bytIndex(0 To 63) As Byte
Private m_bytReverseIndex(0 To 255) As Byte

Private Const k_bytEqualSign As Byte = 61
Private Const k_bytMask1 As Byte = 3
Private Const k_bytMask2 As Byte = 15
Private Const k_bytMask3 As Byte = 63
Private Const k_bytMask4 As Byte = 192
Private Const k_bytMask5 As Byte = 240
Private Const k_bytMask6 As Byte = 252
Private Const k_bytShift2 As Byte = 4
Private Const k_bytShift4 As Byte = 16
Private Const k_bytShift6 As Byte = 64
Private Const k_lMaxBytesPerLine As Long = 152
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)

Private Sub Initialize64()
        '<EhHeader>
        On Error GoTo Initialize64_Err
        '</EhHeader>
100     m_bytIndex(0) = 65 'Asc("A")
102     m_bytIndex(1) = 66 'Asc("B")
104     m_bytIndex(2) = 67 'Asc("C")
106     m_bytIndex(3) = 68 'Asc("D")
108     m_bytIndex(4) = 69 'Asc("E")
110     m_bytIndex(5) = 70 'Asc("F")
112     m_bytIndex(6) = 71 'Asc("G")
114     m_bytIndex(7) = 72 'Asc("H")
116     m_bytIndex(8) = 73 'Asc("I")
118     m_bytIndex(9) = 74 'Asc("J")
120     m_bytIndex(10) = 75 'Asc("K")
122     m_bytIndex(11) = 76 'Asc("L")
124     m_bytIndex(12) = 77 'Asc("M")
126     m_bytIndex(13) = 78 'Asc("N")
128     m_bytIndex(14) = 79 'Asc("O")
130     m_bytIndex(15) = 80 'Asc("P")
132     m_bytIndex(16) = 81 'Asc("Q")
134     m_bytIndex(17) = 82 'Asc("R")
136     m_bytIndex(18) = 83 'Asc("S")
138     m_bytIndex(19) = 84 'Asc("T")
140     m_bytIndex(20) = 85 'Asc("U")
142     m_bytIndex(21) = 86 'Asc("V")
144     m_bytIndex(22) = 87 'Asc("W")
146     m_bytIndex(23) = 88 'Asc("X")
148     m_bytIndex(24) = 89 'Asc("Y")
150     m_bytIndex(25) = 90 'Asc("Z")
152     m_bytIndex(26) = 97 'Asc("a")
154     m_bytIndex(27) = 98 'Asc("b")
156     m_bytIndex(28) = 99 'Asc("c")
158     m_bytIndex(29) = 100 'Asc("d")
160     m_bytIndex(30) = 101 'Asc("e")
162     m_bytIndex(31) = 102 'Asc("f")
164     m_bytIndex(32) = 103 'Asc("g")
166     m_bytIndex(33) = 104 'Asc("h")
168     m_bytIndex(34) = 105 'Asc("i")
170     m_bytIndex(35) = 106 'Asc("j")
172     m_bytIndex(36) = 107 'Asc("k")
174     m_bytIndex(37) = 108 'Asc("l")
176     m_bytIndex(38) = 109 'Asc("m")
178     m_bytIndex(39) = 110 'Asc("n")
180     m_bytIndex(40) = 111 'Asc("o")
182     m_bytIndex(41) = 112 'Asc("p")
184     m_bytIndex(42) = 113 'Asc("q")
186     m_bytIndex(43) = 114 'Asc("r")
188     m_bytIndex(44) = 115 'Asc("s")
190     m_bytIndex(45) = 116 'Asc("t")
192     m_bytIndex(46) = 117 'Asc("u")
194     m_bytIndex(47) = 118 'Asc("v")
196     m_bytIndex(48) = 119 'Asc("w")
198     m_bytIndex(49) = 120 'Asc("x")
200     m_bytIndex(50) = 121 'Asc("y")
202     m_bytIndex(51) = 122 'Asc("z")
204     m_bytIndex(52) = 48 'Asc("0")
206     m_bytIndex(53) = 49 'Asc("1")
208     m_bytIndex(54) = 50 'Asc("2")
210     m_bytIndex(55) = 51 'Asc("3")
212     m_bytIndex(56) = 52 'Asc("4")
214     m_bytIndex(57) = 53 'Asc("5")
216     m_bytIndex(58) = 54 'Asc("6")
218     m_bytIndex(59) = 55 'Asc("7")
220     m_bytIndex(60) = 56 'Asc("8")
222     m_bytIndex(61) = 57 'Asc("9")
224     m_bytIndex(62) = 43 'Asc("+")
226     m_bytIndex(63) = 47 'Asc("/")
228     m_bytReverseIndex(65) = 0 'Asc("A")
230     m_bytReverseIndex(66) = 1 'Asc("B")
232     m_bytReverseIndex(67) = 2 'Asc("C")
234     m_bytReverseIndex(68) = 3 'Asc("D")
236     m_bytReverseIndex(69) = 4 'Asc("E")
238     m_bytReverseIndex(70) = 5 'Asc("F")
240     m_bytReverseIndex(71) = 6 'Asc("G")
242     m_bytReverseIndex(72) = 7 'Asc("H")
244     m_bytReverseIndex(73) = 8 'Asc("I")
246     m_bytReverseIndex(74) = 9 'Asc("J")
248     m_bytReverseIndex(75) = 10 'Asc("K")
250     m_bytReverseIndex(76) = 11 'Asc("L")
252     m_bytReverseIndex(77) = 12 'Asc("M")
254     m_bytReverseIndex(78) = 13 'Asc("N")
256     m_bytReverseIndex(79) = 14 'Asc("O")
258     m_bytReverseIndex(80) = 15 'Asc("P")
260     m_bytReverseIndex(81) = 16 'Asc("Q")
262     m_bytReverseIndex(82) = 17 'Asc("R")
264     m_bytReverseIndex(83) = 18 'Asc("S")
266     m_bytReverseIndex(84) = 19 'Asc("T")
268     m_bytReverseIndex(85) = 20 'Asc("U")
270     m_bytReverseIndex(86) = 21 'Asc("V")
272     m_bytReverseIndex(87) = 22 'Asc("W")
274     m_bytReverseIndex(88) = 23 'Asc("X")
276     m_bytReverseIndex(89) = 24 'Asc("Y")
278     m_bytReverseIndex(90) = 25 'Asc("Z")
280     m_bytReverseIndex(97) = 26 'Asc("a")
282     m_bytReverseIndex(98) = 27 'Asc("b")
284     m_bytReverseIndex(99) = 28 'Asc("c")
286     m_bytReverseIndex(100) = 29 'Asc("d")
288     m_bytReverseIndex(101) = 30 'Asc("e")
290     m_bytReverseIndex(102) = 31 'Asc("f")
292     m_bytReverseIndex(103) = 32 'Asc("g")
294     m_bytReverseIndex(104) = 33 'Asc("h")
296     m_bytReverseIndex(105) = 34 'Asc("i")
298     m_bytReverseIndex(106) = 35 'Asc("j")
300     m_bytReverseIndex(107) = 36 'Asc("k")
302     m_bytReverseIndex(108) = 37 'Asc("l")
304     m_bytReverseIndex(109) = 38 'Asc("m")
306     m_bytReverseIndex(110) = 39 'Asc("n")
308     m_bytReverseIndex(111) = 40 'Asc("o")
310     m_bytReverseIndex(112) = 41 'Asc("p")
312     m_bytReverseIndex(113) = 42 'Asc("q")
314     m_bytReverseIndex(114) = 43 'Asc("r")
316     m_bytReverseIndex(115) = 44 'Asc("s")
318     m_bytReverseIndex(116) = 45 'Asc("t")
320     m_bytReverseIndex(117) = 46 'Asc("u")
322     m_bytReverseIndex(118) = 47 'Asc("v")
324     m_bytReverseIndex(119) = 48 'Asc("w")
326     m_bytReverseIndex(120) = 49 'Asc("x")
328     m_bytReverseIndex(121) = 50 'Asc("y")
330     m_bytReverseIndex(122) = 51 'Asc("z")
332     m_bytReverseIndex(48) = 52 'Asc("0")
334     m_bytReverseIndex(49) = 53 'Asc("1")
336     m_bytReverseIndex(50) = 54 'Asc("2")
338     m_bytReverseIndex(51) = 55 'Asc("3")
340     m_bytReverseIndex(52) = 56 'Asc("4")
342     m_bytReverseIndex(53) = 57 'Asc("5")
344     m_bytReverseIndex(54) = 58 'Asc("6")
346     m_bytReverseIndex(55) = 59 'Asc("7")
348     m_bytReverseIndex(56) = 60 'Asc("8")
350     m_bytReverseIndex(57) = 61 'Asc("9")
352     m_bytReverseIndex(43) = 62 'Asc("+")
354     m_bytReverseIndex(47) = 63 'Asc("/")
        '<EhFooter>
        Exit Sub

Initialize64_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.Initialize64", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Sub

Public Function Decode64(sInput As String) As String
        '<EhHeader>
        On Error GoTo Decode64_Err
        '</EhHeader>
100     If sInput = "" Then Exit Function
102     Decode64 = StrConv(DecodeArray64(sInput), vbUnicode)
        '<EhFooter>
        Exit Function

Decode64_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.Decode64", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function

Public Function DecodeArray64(sInput As String) As Byte()
        '<EhHeader>
        On Error GoTo DecodeArray64_Err
        '</EhHeader>
100     If m_bytReverseIndex(47) <> 63 Then Initialize64
        Dim bytInput() As Byte
        Dim bytWorkspace() As Byte
        Dim bytResult() As Byte
        Dim lInputCounter As Long
        Dim lWorkspaceCounter As Long
    
102     bytInput = Replace(Replace(sInput, vbCrLf, ""), "=", "")
104     ReDim bytWorkspace(LBound(bytInput) To (UBound(bytInput) * 2)) As Byte
106     lWorkspaceCounter = LBound(bytWorkspace)
108     For lInputCounter = LBound(bytInput) To UBound(bytInput)
110         bytInput(lInputCounter) = m_bytReverseIndex(bytInput(lInputCounter))
112     Next lInputCounter
    
114     For lInputCounter = LBound(bytInput) To (UBound(bytInput) - ((UBound(bytInput) Mod 8) + 8)) Step 8
116         bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
118         bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And k_bytMask2) * k_bytShift4) + (bytInput(lInputCounter + 4) \ k_bytShift2)
120         bytWorkspace(lWorkspaceCounter + 2) = ((bytInput(lInputCounter + 4) And k_bytMask1) * k_bytShift6) + bytInput(lInputCounter + 6)
122         lWorkspaceCounter = lWorkspaceCounter + 3
124     Next lInputCounter
    
        Select Case (UBound(bytInput) Mod 8):
            Case 3:
126             bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
            Case 5:
128             bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
130             bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And k_bytMask2) * k_bytShift4) + (bytInput(lInputCounter + 4) \ k_bytShift2)
132             lWorkspaceCounter = lWorkspaceCounter + 1
            Case 7:
134             bytWorkspace(lWorkspaceCounter) = (bytInput(lInputCounter) * k_bytShift2) + (bytInput(lInputCounter + 2) \ k_bytShift4)
136             bytWorkspace(lWorkspaceCounter + 1) = ((bytInput(lInputCounter + 2) And k_bytMask2) * k_bytShift4) + (bytInput(lInputCounter + 4) \ k_bytShift2)
138             bytWorkspace(lWorkspaceCounter + 2) = ((bytInput(lInputCounter + 4) And k_bytMask1) * k_bytShift6) + bytInput(lInputCounter + 6)
140             lWorkspaceCounter = lWorkspaceCounter + 2
        End Select
    
142     ReDim bytResult(LBound(bytWorkspace) To lWorkspaceCounter) As Byte
144     If LBound(bytWorkspace) = 0 Then lWorkspaceCounter = lWorkspaceCounter + 1
146     CopyMemory VarPtr(bytResult(LBound(bytResult))), VarPtr(bytWorkspace(LBound(bytWorkspace))), lWorkspaceCounter
148     DecodeArray64 = bytResult
        '<EhFooter>
        Exit Function

DecodeArray64_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.DecodeArray64", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function

Public Function Encode64(ByRef sInput As String) As String
        '<EhHeader>
        On Error GoTo Encode64_Err
        '</EhHeader>
100     If sInput = "" Then Exit Function
        Dim bytTemp() As Byte
102     bytTemp = StrConv(sInput, vbFromUnicode)
104     Encode64 = EncodeArray64(bytTemp)
        '<EhFooter>
        Exit Function

Encode64_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.Encode64", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function

Public Function EncodeArray64(ByRef bytInput() As Byte) As String
        '<EhHeader>
        On Error GoTo EncodeArray64_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
    
100     If m_bytReverseIndex(47) <> 63 Then Initialize64
        Dim bytWorkspace() As Byte, bytResult() As Byte
        Dim bytCrLf(0 To 3) As Byte, lCounter As Long
        Dim lWorkspaceCounter As Long, lLineCounter As Long
        Dim lCompleteLines As Long, lBytesRemaining As Long
        Dim lpWorkSpace As Long, lpResult As Long
        Dim lpCrLf As Long

102     If UBound(bytInput) < 1024 Then
104         ReDim bytWorkspace(LBound(bytInput) To (LBound(bytInput) + 4096)) As Byte
        Else
106         ReDim bytWorkspace(LBound(bytInput) To (UBound(bytInput) * 4)) As Byte
        End If

108     lWorkspaceCounter = LBound(bytWorkspace)

110     For lCounter = LBound(bytInput) To (UBound(bytInput) - ((UBound(bytInput) Mod 3) + 3)) Step 3
112         bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
114         bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex(((bytInput(lCounter) And k_bytMask1) * k_bytShift4) + ((bytInput(lCounter + 1)) \ k_bytShift4))
116         bytWorkspace(lWorkspaceCounter + 4) = m_bytIndex(((bytInput(lCounter + 1) And k_bytMask2) * k_bytShift2) + (bytInput(lCounter + 2) \ k_bytShift6))
118         bytWorkspace(lWorkspaceCounter + 6) = m_bytIndex(bytInput(lCounter + 2) And k_bytMask3)
120         lWorkspaceCounter = lWorkspaceCounter + 8
122     Next lCounter

        Select Case (UBound(bytInput) Mod 3):
            Case 0:
124             bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
126             bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex((bytInput(lCounter) And k_bytMask1) * k_bytShift4)
128             bytWorkspace(lWorkspaceCounter + 4) = k_bytEqualSign
130             bytWorkspace(lWorkspaceCounter + 6) = k_bytEqualSign
            Case 1:
132             bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
134             bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex(((bytInput(lCounter) And k_bytMask1) * k_bytShift4) + ((bytInput(lCounter + 1)) \ k_bytShift4))
136             bytWorkspace(lWorkspaceCounter + 4) = m_bytIndex((bytInput(lCounter + 1) And k_bytMask2) * k_bytShift2)
138             bytWorkspace(lWorkspaceCounter + 6) = k_bytEqualSign
            Case 2:
140             bytWorkspace(lWorkspaceCounter) = m_bytIndex((bytInput(lCounter) \ k_bytShift2))
142             bytWorkspace(lWorkspaceCounter + 2) = m_bytIndex(((bytInput(lCounter) And k_bytMask1) * k_bytShift4) + ((bytInput(lCounter + 1)) \ k_bytShift4))
144             bytWorkspace(lWorkspaceCounter + 4) = m_bytIndex(((bytInput(lCounter + 1) And k_bytMask2) * k_bytShift2) + ((bytInput(lCounter + 2)) \ k_bytShift6))
146             bytWorkspace(lWorkspaceCounter + 6) = m_bytIndex(bytInput(lCounter + 2) And k_bytMask3)
        End Select

148     lWorkspaceCounter = lWorkspaceCounter + 8

150     If lWorkspaceCounter <= k_lMaxBytesPerLine Then
152         EncodeArray64 = Left$(bytWorkspace, InStr(1, bytWorkspace, Chr$(0)) - 1)
        Else
154         bytCrLf(0) = 13
156         bytCrLf(1) = 0
158         bytCrLf(2) = 10
160         bytCrLf(3) = 0
162         ReDim bytResult(LBound(bytWorkspace) To UBound(bytWorkspace))
164         lpWorkSpace = VarPtr(bytWorkspace(LBound(bytWorkspace)))
166         lpResult = VarPtr(bytResult(LBound(bytResult)))
168         lpCrLf = VarPtr(bytCrLf(LBound(bytCrLf)))
170         lCompleteLines = Fix(lWorkspaceCounter / k_lMaxBytesPerLine)
        
172         For lLineCounter = 0 To lCompleteLines
174             CopyMemory lpResult, lpWorkSpace, k_lMaxBytesPerLine
176             lpWorkSpace = lpWorkSpace + k_lMaxBytesPerLine
178             lpResult = lpResult + k_lMaxBytesPerLine
180             CopyMemory lpResult, lpCrLf, 4&
182             lpResult = lpResult + 4&
184         Next lLineCounter
        
186         lBytesRemaining = lWorkspaceCounter - (lCompleteLines * k_lMaxBytesPerLine)
188         If lBytesRemaining > 0 Then CopyMemory lpResult, lpWorkSpace, lBytesRemaining
190         EncodeArray64 = Left$(bytResult, InStr(1, bytResult, Chr$(0)) - 1)
        End If
        Exit Function

ErrorHandler:
192     Erase bytResult
194     EncodeArray64 = bytResult
        '<EhFooter>
        Exit Function

EncodeArray64_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.EncodeArray64", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function

Private Static Sub DecryptBlock(Xl As Long, Xr As Long)
        '<EhHeader>
        On Error GoTo DecryptBlock_Err
        '</EhHeader>
        Dim i As Long, j As Long, K As Long
100     K = Xr
102     Xr = Xl Xor m_pBox(Rounds + 1)
104     Xl = K Xor m_pBox(Rounds)
106     j = Rounds - 2
108     For i = 0 To (Rounds \ 2 - 1)
110         Xl = Xl Xor f(Xr)
112         Xr = Xr Xor m_pBox(j + 1)
114         Xr = Xr Xor f(Xl)
116         Xl = Xl Xor m_pBox(j)
118         j = j - 2
        Next
        '<EhFooter>
        Exit Sub

DecryptBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.DecryptBlock", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Sub
Private Static Sub EncryptBlock(Xl As Long, Xr As Long)
        '<EhHeader>
        On Error GoTo EncryptBlock_Err
        '</EhHeader>
        Dim i As Long, j As Long, Temp As Long
100     j = 0
102     For i = 0 To (Rounds \ 2 - 1)
104         Xl = Xl Xor m_pBox(j)
106         Xr = Xr Xor f(Xl)
108         Xr = Xr Xor m_pBox(j + 1)
110         Xl = Xl Xor f(Xr)
112         j = j + 2
        Next
114     Temp = Xr
116     Xr = Xl Xor m_pBox(Rounds)
118     Xl = Temp Xor m_pBox(Rounds + 1)
        '<EhFooter>
        Exit Sub

EncryptBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.EncryptBlock", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Sub
Public Sub EncryptByte(byteArray() As Byte, Optional Key As String)
        '<EhHeader>
        On Error GoTo EncryptByte_Err
        '</EhHeader>
        Dim Offset As Long, OrigLen As Long, LeftWord As Long, RightWord As Long, CipherLen As Long, CipherLeft As Long, CipherRight As Long, CurrPercent As Long, NextPercent As Long
100     If (Len(Key) > 0) Then Me.Key = Key
102     OrigLen = UBound(byteArray) + 1
104     CipherLen = OrigLen + 12
106     If (CipherLen Mod 8 <> 0) Then CipherLen = CipherLen + 8 - (CipherLen Mod 8)
108     ReDim Preserve byteArray(CipherLen - 1)
110     Call CopyMem(byteArray(12), byteArray(0), OrigLen)
112     Call CopyMem(byteArray(8), OrigLen, 4)
114     Call Randomize
116     Call CopyMem(byteArray(0), CLng(2147483647 * Rnd), 4)
118     Call CopyMem(byteArray(4), CLng(2147483647 * Rnd), 4)
120     For Offset = 0 To (CipherLen - 1) Step 8
122         Call GetWord(LeftWord, byteArray(), Offset)
124         Call GetWord(RightWord, byteArray(), Offset + 4)
126         LeftWord = LeftWord Xor CipherLeft
128         RightWord = RightWord Xor CipherRight
130         Call EncryptBlock(LeftWord, RightWord)
132         Call PutWord(LeftWord, byteArray(), Offset)
134         Call PutWord(RightWord, byteArray(), Offset + 4)
136         CipherLeft = LeftWord
138         CipherRight = RightWord
140         If (Offset >= NextPercent) Then
142             CurrPercent = Int((Offset / CipherLen) * 100)
144             NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
146             RaiseEvent Progress(CurrPercent)
            End If
        Next
148     If (CurrPercent <> 100) Then RaiseEvent Progress(100)
        '<EhFooter>
        Exit Sub

EncryptByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.EncryptByte", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Sub
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputIn64 As Boolean) As String
        '<EhHeader>
        On Error GoTo EncryptString_Err
        '</EhHeader>
        Dim byteArray() As Byte
100     byteArray() = StrConv(Text, vbFromUnicode)
102     Call EncryptByte(byteArray(), Key)
104     EncryptString = StrConv(byteArray(), vbUnicode)
106     If OutputIn64 = True Then EncryptString = Encode64(EncryptString)
108     Erase byteArray(): Key = "": Text = ""
        '<EhFooter>
        Exit Function

EncryptString_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.EncryptString", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextIn64 As Boolean) As String
        '<EhHeader>
        On Error GoTo DecryptString_Err
        '</EhHeader>
        Dim byteArray() As Byte
100     If IsTextIn64 = True Then Text = Decode64(Text)
102     byteArray() = StrConv(Text, vbFromUnicode)
104     Call DecryptByte(byteArray(), Key)
106     DecryptString = StrConv(byteArray(), vbUnicode)
108     Erase byteArray(): Key = "": Text = ""
        '<EhFooter>
        Exit Function

DecryptString_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.DecryptString", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function
Public Sub DecryptByte(byteArray() As Byte, Optional Key As String)
        '<EhHeader>
        On Error GoTo DecryptByte_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
        Dim Offset As Long, OrigLen As Long, LeftWord As Long, RightWord As Long, CipherLen As Long, CipherLeft As Long, CipherRight As Long, CurrPercent As Long, NextPercent As Long
100     If (Len(Key) > 0) Then Me.Key = Key
102     CipherLen = UBound(byteArray) + 1
104     For Offset = 0 To (CipherLen - 1) Step 8
106         Call GetWord(LeftWord, byteArray(), Offset)
108         Call GetWord(RightWord, byteArray(), Offset + 4)
110         Call DecryptBlock(LeftWord, RightWord)
112         LeftWord = LeftWord Xor CipherLeft
114         RightWord = RightWord Xor CipherRight
116         Call GetWord(CipherLeft, byteArray(), Offset)
118         Call GetWord(CipherRight, byteArray(), Offset + 4)
120         Call PutWord(LeftWord, byteArray(), Offset)
122         Call PutWord(RightWord, byteArray(), Offset + 4)
124         If Offset >= NextPercent Then
126             CurrPercent = Int((Offset / CipherLen) * 100)
128             NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
130             RaiseEvent Progress(CurrPercent)
            End If
        Next
132     Call CopyMem(OrigLen, byteArray(8), 4)
134     If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Blowfish decryption")
136     Call CopyMem(byteArray(0), byteArray(12), OrigLen)
138     ReDim Preserve byteArray(OrigLen - 1)
140     If CurrPercent <> 100 Then RaiseEvent Progress(100)

ErrorHandler:
        '<EhFooter>
        Exit Sub

DecryptByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.DecryptByte", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Sub
Private Static Function f(ByVal X As Long) As Long
        '<EhHeader>
        On Error GoTo f_Err
        '</EhHeader>
        Dim xb(0 To 3) As Byte
100     Call CopyMem(xb(0), X, 4)
102     If (m_RunningCompiled) Then f = (((m_sBox(0, xb(3)) + m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))) + m_sBox(3, xb(0))) Else f = UnsignedAdd((UnsignedAdd(m_sBox(0, xb(3)), m_sBox(1, xb(2))) Xor m_sBox(2, xb(1))), m_sBox(3, xb(0)))
        '<EhFooter>
        Exit Function

f_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.f", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function
Private Static Sub GetWord(LongValue As Long, CryptBuffer() As Byte, Offset As Long)
        '<EhHeader>
        On Error GoTo GetWord_Err
        '</EhHeader>
        Dim bb(0 To 3) As Byte
100     bb(3) = CryptBuffer(Offset)
102     bb(2) = CryptBuffer(Offset + 1)
104     bb(1) = CryptBuffer(Offset + 2)
106     bb(0) = CryptBuffer(Offset + 3)
108     Call CopyMem(LongValue, bb(0), 4)
        '<EhFooter>
        Exit Sub

GetWord_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.GetWord", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Sub
Private Static Sub PutWord(LongValue As Long, CryptBuffer() As Byte, Offset As Long)
        '<EhHeader>
        On Error GoTo PutWord_Err
        '</EhHeader>
        Dim bb(0 To 3) As Byte
100     Call CopyMem(bb(0), LongValue, 4)
102     CryptBuffer(Offset) = bb(3)
104     CryptBuffer(Offset + 1) = bb(2)
106     CryptBuffer(Offset + 2) = bb(1)
108     CryptBuffer(Offset + 3) = bb(0)
        '<EhFooter>
        Exit Sub

PutWord_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.PutWord", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Sub
Private Static Function UnsignedAdd(ByVal Data1 As Long, Data2 As Long) As Long
        '<EhHeader>
        On Error GoTo UnsignedAdd_Err
        '</EhHeader>
        Dim x1(0 To 3) As Byte, x2(0 To 3) As Byte, xx(0 To 3) As Byte, Rest As Long, Value As Long, a As Long
100     Call CopyMem(x1(0), Data1, 4)
102     Call CopyMem(x2(0), Data2, 4)
104     Rest = 0
106     For a = 0 To 3
108         Value = CLng(x1(a)) + CLng(x2(a)) + Rest
110         xx(a) = Value And 255
112         Rest = Value \ 256
        Next
114     Call CopyMem(UnsignedAdd, xx(0), 4)
        '<EhFooter>
        Exit Function

UnsignedAdd_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.UnsignedAdd", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function
Private Function UnsignedDel(Data1 As Long, Data2 As Long) As Long
        '<EhHeader>
        On Error GoTo UnsignedDel_Err
        '</EhHeader>
        Dim x1(0 To 3) As Byte, x2(0 To 3) As Byte, xx(0 To 3) As Byte, Rest As Long, Value As Long, a As Long
100     Call CopyMem(x1(0), Data1, 4)
102     Call CopyMem(x2(0), Data2, 4)
104     Call CopyMem(xx(0), UnsignedDel, 4)
106     For a = 0 To 3
108         Value = CLng(x1(a)) - CLng(x2(a)) - Rest
110         If (Value < 0) Then
112             Value = Value + 256
114             Rest = 1
            Else
116             Rest = 0
            End If
118         xx(a) = Value
        Next
120     Call CopyMem(UnsignedDel, xx(0), 4)
        '<EhFooter>
        Exit Function

UnsignedDel_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.UnsignedDel", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function
Public Property Let Key(New_Value As String)
        '<EhHeader>
        On Error GoTo Key_Err
        '</EhHeader>
        Dim i As Long, j As Long, K As Long, dataX As Long, datal As Long, datar As Long, Key() As Byte, KeyLength As Long
100     Class_Initialize
102     If (m_KeyValue = New_Value) Then Exit Property
104     m_KeyValue = New_Value
106     KeyLength = Len(New_Value)
108     Key() = StrConv(New_Value, vbFromUnicode)
110     j = 0
112     For i = 0 To (Rounds + 1)
114         dataX = 0
116         For K = 0 To 3
118             Call CopyMem(ByVal VarPtr(dataX) + 1, dataX, 3)
120             dataX = (dataX Or Key(j))
122             j = j + 1
124             If (j >= KeyLength) Then j = 0
            Next
126         m_pBox(i) = m_pBox(i) Xor dataX
        Next
    
128     datal = 0: datar = 0
130     For i = 0 To (Rounds + 1) Step 2
132         Call EncryptBlock(datal, datar)
134         m_pBox(i) = datal
136         m_pBox(i + 1) = datar
        Next
138     For i = 0 To 3
140         For j = 0 To 255 Step 2
142             Call EncryptBlock(datal, datar)
144             m_sBox(i, j) = datal
146             m_sBox(i, j + 1) = datar
            Next
        Next
        '<EhFooter>
        Exit Property

Key_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.Key", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Property
Private Sub Class_Initialize()
        '<EhHeader>
        On Error GoTo Class_Initialize_Err
        '</EhHeader>
100 On Local Error Resume Next
102   m_RunningCompiled = ((2147483647 + 1) < 0)
104   m_pBox(0) = &H243F6A88
106   m_pBox(1) = &H85A308D3
108   m_pBox(2) = &H13198A2E
110   m_pBox(3) = &H3707344
112   m_pBox(4) = &HA4093822
114   m_pBox(5) = &H299F31D0
116   m_pBox(6) = &H82EFA98
118   m_pBox(7) = &HEC4E6C89
120   m_pBox(8) = &H452821E6
122   m_pBox(9) = &H38D01377
124   m_pBox(10) = &HBE5466CF
126   m_pBox(11) = &H34E90C6C
128   m_pBox(12) = &HC0AC29B7
130   m_pBox(13) = &HC97C50DD
132   m_pBox(14) = &H3F84D5B5
134   m_pBox(15) = &HB5470917
136   m_pBox(16) = &H9216D5D9
138   m_pBox(17) = &H8979FB1B
140   m_sBox(0, 0) = &HD1310BA6
142   m_sBox(1, 0) = &H98DFB5AC
144   m_sBox(2, 0) = &H2FFD72DB
146   m_sBox(3, 0) = &HD01ADFB7
148   m_sBox(0, 1) = &HB8E1AFED
150   m_sBox(1, 1) = &H6A267E96
152   m_sBox(2, 1) = &HBA7C9045
154   m_sBox(3, 1) = &HF12C7F99
156   m_sBox(0, 2) = &H24A19947
158   m_sBox(1, 2) = &HB3916CF7
160   m_sBox(2, 2) = &H801F2E2
162   m_sBox(3, 2) = &H858EFC16
164   m_sBox(0, 3) = &H636920D8
166   m_sBox(1, 3) = &H71574E69
168   m_sBox(2, 3) = &HA458FEA3
170   m_sBox(3, 3) = &HF4933D7E
172   m_sBox(0, 4) = &HD95748F
174   m_sBox(1, 4) = &H728EB658
176   m_sBox(2, 4) = &H718BCD58
178   m_sBox(3, 4) = &H82154AEE
180   m_sBox(0, 5) = &H7B54A41D
182   m_sBox(1, 5) = &HC25A59B5
184   m_sBox(2, 5) = &H9C30D539
186   m_sBox(3, 5) = &H2AF26013
188   m_sBox(0, 6) = &HC5D1B023
190   m_sBox(1, 6) = &H286085F0
192   m_sBox(2, 6) = &HCA417918
194   m_sBox(3, 6) = &HB8DB38EF
196   m_sBox(0, 7) = &H8E79DCB0
198   m_sBox(1, 7) = &H603A180E
200   m_sBox(2, 7) = &H6C9E0E8B
202   m_sBox(3, 7) = &HB01E8A3E
204   m_sBox(0, 8) = &HD71577C1
206   m_sBox(1, 8) = &HBD314B27
208   m_sBox(2, 8) = &H78AF2FDA
210   m_sBox(3, 8) = &H55605C60
212   m_sBox(0, 9) = &HE65525F3
214   m_sBox(1, 9) = &HAA55AB94
216   m_sBox(2, 9) = &H57489862
218   m_sBox(3, 9) = &H63E81440
220   m_sBox(0, 10) = &H55CA396A
222   m_sBox(1, 10) = &H2AAB10B6
224   m_sBox(2, 10) = &HB4CC5C34
226   m_sBox(3, 10) = &H1141E8CE
228   m_sBox(0, 11) = &HA15486AF
230   m_sBox(1, 11) = &H7C72E993
232   m_sBox(2, 11) = &HB3EE1411
234   m_sBox(3, 11) = &H636FBC2A
236   m_sBox(0, 12) = &H2BA9C55D
238   m_sBox(1, 12) = &H741831F6
240   m_sBox(2, 12) = &HCE5C3E16
242   m_sBox(3, 12) = &H9B87931E
244   m_sBox(0, 13) = &HAFD6BA33
246   m_sBox(1, 13) = &H6C24CF5C
248   m_sBox(2, 13) = &H7A325381
250   m_sBox(3, 13) = &H28958677
252   m_sBox(0, 14) = &H3B8F4898
254   m_sBox(1, 14) = &H6B4BB9AF
256   m_sBox(2, 14) = &HC4BFE81B
258   m_sBox(3, 14) = &H66282193
260   m_sBox(0, 15) = &H61D809CC
262   m_sBox(1, 15) = &HFB21A991
264   m_sBox(2, 15) = &H487CAC60
266   m_sBox(3, 15) = &H5DEC8032
268   m_sBox(0, 16) = &HEF845D5D
270   m_sBox(1, 16) = &HE98575B1
272   m_sBox(2, 16) = &HDC262302
274   m_sBox(3, 16) = &HEB651B88
276   m_sBox(0, 17) = &H23893E81
278   m_sBox(1, 17) = &HD396ACC5
280   m_sBox(2, 17) = &HF6D6FF3
282   m_sBox(3, 17) = &H83F44239
284   m_sBox(0, 18) = &H2E0B4482
286   m_sBox(1, 18) = &HA4842004
288   m_sBox(2, 18) = &H69C8F04A
290   m_sBox(3, 18) = &H9E1F9B5E
292   m_sBox(0, 19) = &H21C66842
294   m_sBox(1, 19) = &HF6E96C9A
296   m_sBox(2, 19) = &H670C9C61
298   m_sBox(3, 19) = &HABD388F0
300   m_sBox(0, 20) = &H6A51A0D2
302   m_sBox(1, 20) = &HD8542F68
304   m_sBox(2, 20) = &H960FA728
306   m_sBox(3, 20) = &HAB5133A3
308   m_sBox(0, 21) = &H6EEF0B6C
310   m_sBox(1, 21) = &H137A3BE4
312   m_sBox(2, 21) = &HBA3BF050
314   m_sBox(3, 21) = &H7EFB2A98
316   m_sBox(0, 22) = &HA1F1651D
318   m_sBox(1, 22) = &H39AF0176
320   m_sBox(2, 22) = &H66CA593E
322   m_sBox(3, 22) = &H82430E88
324   m_sBox(0, 23) = &H8CEE8619
326   m_sBox(1, 23) = &H456F9FB4
328   m_sBox(2, 23) = &H7D84A5C3
330   m_sBox(3, 23) = &H3B8B5EBE
332   m_sBox(0, 24) = &HE06F75D8
334   m_sBox(1, 24) = &H85C12073
336   m_sBox(2, 24) = &H401A449F
338   m_sBox(3, 24) = &H56C16AA6
340   m_sBox(0, 25) = &H4ED3AA62
342   m_sBox(1, 25) = &H363F7706
344   m_sBox(2, 25) = &H1BFEDF72
346   m_sBox(3, 25) = &H429B023D
348   m_sBox(0, 26) = &H37D0D724
350   m_sBox(1, 26) = &HD00A1248
352   m_sBox(2, 26) = &HDB0FEAD3
354   m_sBox(3, 26) = &H49F1C09B
356   m_sBox(0, 27) = &H75372C9
358   m_sBox(1, 27) = &H80991B7B
360   m_sBox(2, 27) = &H25D479D8
362   m_sBox(3, 27) = &HF6E8DEF7
364   m_sBox(0, 28) = &HE3FE501A
366   m_sBox(1, 28) = &HB6794C3B
368   m_sBox(2, 28) = &H976CE0BD
370   m_sBox(3, 28) = &H4C006BA
372   m_sBox(0, 29) = &HC1A94FB6
374   m_sBox(1, 29) = &H409F60C4
376   m_sBox(2, 29) = &H5E5C9EC2
378   m_sBox(3, 29) = &H196A2463
380   m_sBox(0, 30) = &H68FB6FAF
382   m_sBox(1, 30) = &H3E6C53B5
384   m_sBox(2, 30) = &H1339B2EB
386   m_sBox(3, 30) = &H3B52EC6F
388   m_sBox(0, 31) = &H6DFC511F
390   m_sBox(1, 31) = &H9B30952C
392   m_sBox(2, 31) = &HCC814544
394   m_sBox(3, 31) = &HAF5EBD09
396   m_sBox(0, 32) = &HBEE3D004
398   m_sBox(1, 32) = &HDE334AFD
400   m_sBox(2, 32) = &H660F2807
402   m_sBox(3, 32) = &H192E4BB3
404   m_sBox(0, 33) = &HC0CBA857
406   m_sBox(1, 33) = &H45C8740F
408   m_sBox(2, 33) = &HD20B5F39
410   m_sBox(3, 33) = &HB9D3FBDB
412   m_sBox(0, 34) = &H5579C0BD
414   m_sBox(1, 34) = &H1A60320A
416   m_sBox(2, 34) = &HD6A100C6
418   m_sBox(3, 34) = &H402C7279
420   m_sBox(0, 35) = &H679F25FE
422   m_sBox(1, 35) = &HFB1FA3CC
424   m_sBox(2, 35) = &H8EA5E9F8
426   m_sBox(3, 35) = &HDB3222F8
428   m_sBox(0, 36) = &H3C7516DF
430   m_sBox(1, 36) = &HFD616B15
432   m_sBox(2, 36) = &H2F501EC8
434   m_sBox(3, 36) = &HAD0552AB
436   m_sBox(0, 37) = &H323DB5FA
438   m_sBox(1, 37) = &HFD238760
440   m_sBox(2, 37) = &H53317B48
442   m_sBox(3, 37) = &H3E00DF82
444   m_sBox(0, 38) = &H9E5C57BB
446   m_sBox(1, 38) = &HCA6F8CA0
448   m_sBox(2, 38) = &H1A87562E
450   m_sBox(3, 38) = &HDF1769DB
452   m_sBox(0, 39) = &HD542A8F6
454   m_sBox(1, 39) = &H287EFFC3
456   m_sBox(2, 39) = &HAC6732C6
458   m_sBox(3, 39) = &H8C4F5573
460   m_sBox(0, 40) = &H695B27B0
462   m_sBox(1, 40) = &HBBCA58C8
464   m_sBox(2, 40) = &HE1FFA35D
466   m_sBox(3, 40) = &HB8F011A0
468   m_sBox(0, 41) = &H10FA3D98
470   m_sBox(1, 41) = &HFD2183B8
472   m_sBox(2, 41) = &H4AFCB56C
474   m_sBox(3, 41) = &H2DD1D35B
476   m_sBox(0, 42) = &H9A53E479
478   m_sBox(1, 42) = &HB6F84565
480   m_sBox(2, 42) = &HD28E49BC
482   m_sBox(3, 42) = &H4BFB9790
484   m_sBox(0, 43) = &HE1DDF2DA
486   m_sBox(1, 43) = &HA4CB7E33
488   m_sBox(2, 43) = &H62FB1341
490   m_sBox(3, 43) = &HCEE4C6E8
492   m_sBox(0, 44) = &HEF20CADA
494   m_sBox(1, 44) = &H36774C01
496   m_sBox(2, 44) = &HD07E9EFE
498   m_sBox(3, 44) = &H2BF11FB4
500   m_sBox(0, 45) = &H95DBDA4D
502   m_sBox(1, 45) = &HAE909198
504   m_sBox(2, 45) = &HEAAD8E71
506   m_sBox(3, 45) = &H6B93D5A0
508   m_sBox(0, 46) = &HD08ED1D0
510   m_sBox(1, 46) = &HAFC725E0
512   m_sBox(2, 46) = &H8E3C5B2F
514   m_sBox(3, 46) = &H8E7594B7
516   m_sBox(0, 47) = &H8FF6E2FB
518   m_sBox(1, 47) = &HF2122B64
520   m_sBox(2, 47) = &H8888B812
522   m_sBox(3, 47) = &H900DF01C
524   m_sBox(0, 48) = &H4FAD5EA0
526   m_sBox(1, 48) = &H688FC31C
528   m_sBox(2, 48) = &HD1CFF191
530   m_sBox(3, 48) = &HB3A8C1AD
532   m_sBox(0, 49) = &H2F2F2218
534   m_sBox(1, 49) = &HBE0E1777
536   m_sBox(2, 49) = &HEA752DFE
538   m_sBox(3, 49) = &H8B021FA1
540   m_sBox(0, 50) = &HE5A0CC0F
542   m_sBox(1, 50) = &HB56F74E8
544   m_sBox(2, 50) = &H18ACF3D6
546   m_sBox(3, 50) = &HCE89E299
548   m_sBox(0, 51) = &HB4A84FE0
550   m_sBox(1, 51) = &HFD13E0B7
552   m_sBox(2, 51) = &H7CC43B81
554   m_sBox(3, 51) = &HD2ADA8D9
556   m_sBox(0, 52) = &H165FA266
558   m_sBox(1, 52) = &H80957705
560   m_sBox(2, 52) = &H93CC7314
562   m_sBox(3, 52) = &H211A1477
564   m_sBox(0, 53) = &HE6AD2065
566   m_sBox(1, 53) = &H77B5FA86
568   m_sBox(2, 53) = &HC75442F5
570   m_sBox(3, 53) = &HFB9D35CF
572   m_sBox(0, 54) = &HEBCDAF0C
574   m_sBox(1, 54) = &H7B3E89A0
576   m_sBox(2, 54) = &HD6411BD3
578   m_sBox(3, 54) = &HAE1E7E49
580   m_sBox(0, 55) = &H250E2D
582   m_sBox(1, 55) = &H2071B35E
584   m_sBox(2, 55) = &H226800BB
586   m_sBox(3, 55) = &H57B8E0AF
588   m_sBox(0, 56) = &H2464369B
590   m_sBox(1, 56) = &HF009B91E
592   m_sBox(2, 56) = &H5563911D
594   m_sBox(3, 56) = &H59DFA6AA
596   m_sBox(0, 57) = &H78C14389
598   m_sBox(1, 57) = &HD95A537F
600   m_sBox(2, 57) = &H207D5BA2
602   m_sBox(3, 57) = &H2E5B9C5
604   m_sBox(0, 58) = &H83260376
606   m_sBox(1, 58) = &H6295CFA9
608   m_sBox(2, 58) = &H11C81968
610   m_sBox(3, 58) = &H4E734A41
612   m_sBox(0, 59) = &HB3472DCA
614   m_sBox(1, 59) = &H7B14A94A
616   m_sBox(2, 59) = &H1B510052
618   m_sBox(3, 59) = &H9A532915
620   m_sBox(0, 60) = &HD60F573F
622   m_sBox(1, 60) = &HBC9BC6E4
624   m_sBox(2, 60) = &H2B60A476
626   m_sBox(3, 60) = &H81E67400
628   m_sBox(0, 61) = &H8BA6FB5
630   m_sBox(1, 61) = &H571BE91F
632   m_sBox(2, 61) = &HF296EC6B
634   m_sBox(3, 61) = &H2A0DD915
636   m_sBox(0, 62) = &HB6636521
638   m_sBox(1, 62) = &HE7B9F9B6
640   m_sBox(2, 62) = &HFF34052E
642   m_sBox(3, 62) = &HC5855664
644   m_sBox(0, 63) = &H53B02D5D
646   m_sBox(1, 63) = &HA99F8FA1
648   m_sBox(2, 63) = &H8BA4799
650   m_sBox(3, 63) = &H6E85076A
652   m_sBox(0, 64) = &H4B7A70E9
654   m_sBox(1, 64) = &HB5B32944
656   m_sBox(2, 64) = &HDB75092E
658   m_sBox(3, 64) = &HC4192623
660   m_sBox(0, 65) = &HAD6EA6B0
662   m_sBox(1, 65) = &H49A7DF7D
664   m_sBox(2, 65) = &H9CEE60B8
666   m_sBox(3, 65) = &H8FEDB266
668   m_sBox(0, 66) = &HECAA8C71
670   m_sBox(1, 66) = &H699A17FF
672   m_sBox(2, 66) = &H5664526C
674   m_sBox(3, 66) = &HC2B19EE1
676   m_sBox(0, 67) = &H193602A5
678   m_sBox(1, 67) = &H75094C29
680   m_sBox(2, 67) = &HA0591340
682   m_sBox(3, 67) = &HE4183A3E
684   m_sBox(0, 68) = &H3F54989A
686   m_sBox(1, 68) = &H5B429D65
688   m_sBox(2, 68) = &H6B8FE4D6
690   m_sBox(3, 68) = &H99F73FD6
692   m_sBox(0, 69) = &HA1D29C07
694   m_sBox(1, 69) = &HEFE830F5
696   m_sBox(2, 69) = &H4D2D38E6
698   m_sBox(3, 69) = &HF0255DC1
700   m_sBox(0, 70) = &H4CDD2086
702   m_sBox(1, 70) = &H8470EB26
704   m_sBox(2, 70) = &H6382E9C6
706   m_sBox(3, 70) = &H21ECC5E
708   m_sBox(0, 71) = &H9686B3F
710   m_sBox(1, 71) = &H3EBAEFC9
712   m_sBox(2, 71) = &H3C971814
714   m_sBox(3, 71) = &H6B6A70A1
716   m_sBox(0, 72) = &H687F3584
718   m_sBox(1, 72) = &H52A0E286
720   m_sBox(2, 72) = &HB79C5305
722   m_sBox(3, 72) = &HAA500737
724   m_sBox(0, 73) = &H3E07841C
726   m_sBox(1, 73) = &H7FDEAE5C
728   m_sBox(2, 73) = &H8E7D44EC
730   m_sBox(3, 73) = &H5716F2B8
732   m_sBox(0, 74) = &HB03ADA37
734   m_sBox(1, 74) = &HF0500C0D
736   m_sBox(2, 74) = &HF01C1F04
738   m_sBox(3, 74) = &H200B3FF
740   m_sBox(0, 75) = &HAE0CF51A
742   m_sBox(1, 75) = &H3CB574B2
744   m_sBox(2, 75) = &H25837A58
746   m_sBox(3, 75) = &HDC0921BD
748   m_sBox(0, 76) = &HD19113F9
750   m_sBox(1, 76) = &H7CA92FF6
752   m_sBox(2, 76) = &H94324773
754   m_sBox(3, 76) = &H22F54701
756   m_sBox(0, 77) = &H3AE5E581
758   m_sBox(1, 77) = &H37C2DADC
760   m_sBox(2, 77) = &HC8B57634
762   m_sBox(3, 77) = &H9AF3DDA7
764   m_sBox(0, 78) = &HA9446146
766   m_sBox(1, 78) = &HFD0030E
768   m_sBox(2, 78) = &HECC8C73E
770   m_sBox(3, 78) = &HA4751E41
772   m_sBox(0, 79) = &HE238CD99
774   m_sBox(1, 79) = &H3BEA0E2F
776   m_sBox(2, 79) = &H3280BBA1
778   m_sBox(3, 79) = &H183EB331
780   m_sBox(0, 80) = &H4E548B38
782   m_sBox(1, 80) = &H4F6DB908
784   m_sBox(2, 80) = &H6F420D03
786   m_sBox(3, 80) = &HF60A04BF
788   m_sBox(0, 81) = &H2CB81290
790   m_sBox(1, 81) = &H24977C79
792   m_sBox(2, 81) = &H5679B072
794   m_sBox(3, 81) = &HBCAF89AF
796   m_sBox(0, 82) = &HDE9A771F
798   m_sBox(1, 82) = &HD9930810
800   m_sBox(2, 82) = &HB38BAE12
802   m_sBox(3, 82) = &HDCCF3F2E
804   m_sBox(0, 83) = &H5512721F
806   m_sBox(1, 83) = &H2E6B7124
808   m_sBox(2, 83) = &H501ADDE6
810   m_sBox(3, 83) = &H9F84CD87
812   m_sBox(0, 84) = &H7A584718
814   m_sBox(1, 84) = &H7408DA17
816   m_sBox(2, 84) = &HBC9F9ABC
818   m_sBox(3, 84) = &HE94B7D8C
820   m_sBox(0, 85) = &HEC7AEC3A
822   m_sBox(1, 85) = &HDB851DFA
824   m_sBox(2, 85) = &H63094366
826   m_sBox(3, 85) = &HC464C3D2
828   m_sBox(0, 86) = &HEF1C1847
830   m_sBox(1, 86) = &H3215D908
832   m_sBox(2, 86) = &HDD433B37
834   m_sBox(3, 86) = &H24C2BA16
836   m_sBox(0, 87) = &H12A14D43
838   m_sBox(1, 87) = &H2A65C451
840   m_sBox(2, 87) = &H50940002
842   m_sBox(3, 87) = &H133AE4DD
844   m_sBox(0, 88) = &H71DFF89E
846   m_sBox(1, 88) = &H10314E55
848   m_sBox(2, 88) = &H81AC77D6
850   m_sBox(3, 88) = &H5F11199B
852   m_sBox(0, 89) = &H43556F1
854   m_sBox(1, 89) = &HD7A3C76B
856   m_sBox(2, 89) = &H3C11183B
858   m_sBox(3, 89) = &H5924A509
860   m_sBox(0, 90) = &HF28FE6ED
862   m_sBox(1, 90) = &H97F1FBFA
864   m_sBox(2, 90) = &H9EBABF2C
866   m_sBox(3, 90) = &H1E153C6E
868   m_sBox(0, 91) = &H86E34570
870   m_sBox(1, 91) = &HEAE96FB1
872   m_sBox(2, 91) = &H860E5E0A
874   m_sBox(3, 91) = &H5A3E2AB3
876   m_sBox(0, 92) = &H771FE71C
878   m_sBox(1, 92) = &H4E3D06FA
880   m_sBox(2, 92) = &H2965DCB9
882   m_sBox(3, 92) = &H99E71D0F
884   m_sBox(0, 93) = &H803E89D6
886   m_sBox(1, 93) = &H5266C825
888   m_sBox(2, 93) = &H2E4CC978
890   m_sBox(3, 93) = &H9C10B36A
892   m_sBox(0, 94) = &HC6150EBA
894   m_sBox(1, 94) = &H94E2EA78
896   m_sBox(2, 94) = &HA5FC3C53
898   m_sBox(3, 94) = &H1E0A2DF4
900   m_sBox(0, 95) = &HF2F74EA7
902   m_sBox(1, 95) = &H361D2B3D
904   m_sBox(2, 95) = &H1939260F
906   m_sBox(3, 95) = &H19C27960
908   m_sBox(0, 96) = &H5223A708
910   m_sBox(1, 96) = &HF71312B6
912   m_sBox(2, 96) = &HEBADFE6E
914   m_sBox(3, 96) = &HEAC31F66
916   m_sBox(0, 97) = &HE3BC4595
918   m_sBox(1, 97) = &HA67BC883
920   m_sBox(2, 97) = &HB17F37D1
922   m_sBox(3, 97) = &H18CFF28
924   m_sBox(0, 98) = &HC332DDEF
926   m_sBox(1, 98) = &HBE6C5AA5
928   m_sBox(2, 98) = &H65582185
930   m_sBox(3, 98) = &H68AB9802
932   m_sBox(0, 99) = &HEECEA50F
934   m_sBox(1, 99) = &HDB2F953B
936   m_sBox(2, 99) = &H2AEF7DAD
938   m_sBox(3, 99) = &H5B6E2F84
940   m_sBox(0, 100) = &H1521B628
942   m_sBox(1, 100) = &H29076170
944   m_sBox(2, 100) = &HECDD4775
946   m_sBox(3, 100) = &H619F1510
948   m_sBox(0, 101) = &H13CCA830
950   m_sBox(1, 101) = &HEB61BD96
952   m_sBox(2, 101) = &H334FE1E
954   m_sBox(3, 101) = &HAA0363CF
956   m_sBox(0, 102) = &HB5735C90
958   m_sBox(1, 102) = &H4C70A239
960   m_sBox(2, 102) = &HD59E9E0B
962   m_sBox(3, 102) = &HCBAADE14
964   m_sBox(0, 103) = &HEECC86BC
966   m_sBox(1, 103) = &H60622CA7
968   m_sBox(2, 103) = &H9CAB5CAB
970   m_sBox(3, 103) = &HB2F3846E
972   m_sBox(0, 104) = &H648B1EAF
974   m_sBox(1, 104) = &H19BDF0CA
976   m_sBox(2, 104) = &HA02369B9
978   m_sBox(3, 104) = &H655ABB50
980   m_sBox(0, 105) = &H40685A32
982   m_sBox(1, 105) = &H3C2AB4B3
984   m_sBox(2, 105) = &H319EE9D5
986   m_sBox(3, 105) = &HC021B8F7
988   m_sBox(0, 106) = &H9B540B19
990   m_sBox(1, 106) = &H875FA099
992   m_sBox(2, 106) = &H95F7997E
994   m_sBox(3, 106) = &H623D7DA8
996   m_sBox(0, 107) = &HF837889A
998   m_sBox(1, 107) = &H97E32D77
1000   m_sBox(2, 107) = &H11ED935F
1002   m_sBox(3, 107) = &H16681281
1004   m_sBox(0, 108) = &HE358829
1006   m_sBox(1, 108) = &HC7E61FD6
1008   m_sBox(2, 108) = &H96DEDFA1
1010   m_sBox(3, 108) = &H7858BA99
1012   m_sBox(0, 109) = &H57F584A5
1014   m_sBox(1, 109) = &H1B227263
1016   m_sBox(2, 109) = &H9B83C3FF
1018   m_sBox(3, 109) = &H1AC24696
1020   m_sBox(0, 110) = &HCDB30AEB
1022   m_sBox(1, 110) = &H532E3054
1024   m_sBox(2, 110) = &H8FD948E4
1026   m_sBox(3, 110) = &H6DBC3128
1028   m_sBox(0, 111) = &H58EBF2EF
1030   m_sBox(1, 111) = &H34C6FFEA
1032   m_sBox(2, 111) = &HFE28ED61
1034   m_sBox(3, 111) = &HEE7C3C73
1036   m_sBox(0, 112) = &H5D4A14D9
1038   m_sBox(1, 112) = &HE864B7E3
1040   m_sBox(2, 112) = &H42105D14
1042   m_sBox(3, 112) = &H203E13E0
1044   m_sBox(0, 113) = &H45EEE2B6
1046   m_sBox(1, 113) = &HA3AAABEA
1048   m_sBox(2, 113) = &HDB6C4F15
1050   m_sBox(3, 113) = &HFACB4FD0
1052   m_sBox(0, 114) = &HC742F442
1054   m_sBox(1, 114) = &HEF6ABBB5
1056   m_sBox(2, 114) = &H654F3B1D
1058   m_sBox(3, 114) = &H41CD2105
1060   m_sBox(0, 115) = &HD81E799E
1062   m_sBox(1, 115) = &H86854DC7
1064   m_sBox(2, 115) = &HE44B476A
1066   m_sBox(3, 115) = &H3D816250
1068   m_sBox(0, 116) = &HCF62A1F2
1070   m_sBox(1, 116) = &H5B8D2646
1072   m_sBox(2, 116) = &HFC8883A0
1074   m_sBox(3, 116) = &HC1C7B6A3
1076   m_sBox(0, 117) = &H7F1524C3
1078   m_sBox(1, 117) = &H69CB7492
1080   m_sBox(2, 117) = &H47848A0B
1082   m_sBox(3, 117) = &H5692B285
1084   m_sBox(0, 118) = &H95BBF00
1086   m_sBox(1, 118) = &HAD19489D
1088   m_sBox(2, 118) = &H1462B174
1090   m_sBox(3, 118) = &H23820E00
1092   m_sBox(0, 119) = &H58428D2A
1094   m_sBox(1, 119) = &HC55F5EA
1096   m_sBox(2, 119) = &H1DADF43E
1098   m_sBox(3, 119) = &H233F7061
1100   m_sBox(0, 120) = &H3372F092
1102   m_sBox(1, 120) = &H8D937E41
1104   m_sBox(2, 120) = &HD65FECF1
1106   m_sBox(3, 120) = &H6C223BDB
1108   m_sBox(0, 121) = &H7CDE3759
1110   m_sBox(1, 121) = &HCBEE7460
1112   m_sBox(2, 121) = &H4085F2A7
1114   m_sBox(3, 121) = &HCE77326E
1116   m_sBox(0, 122) = &HA6078084
1118   m_sBox(1, 122) = &H19F8509E
1120   m_sBox(2, 122) = &HE8EFD855
1122   m_sBox(3, 122) = &H61D99735
1124   m_sBox(0, 123) = &HA969A7AA
1126   m_sBox(1, 123) = &HC50C06C2
1128   m_sBox(2, 123) = &H5A04ABFC
1130   m_sBox(3, 123) = &H800BCADC
1132   m_sBox(0, 124) = &H9E447A2E
1134   m_sBox(1, 124) = &HC3453484
1136   m_sBox(2, 124) = &HFDD56705
1138   m_sBox(3, 124) = &HE1E9EC9
1140   m_sBox(0, 125) = &HDB73DBD3
1142   m_sBox(1, 125) = &H105588CD
1144   m_sBox(2, 125) = &H675FDA79
1146   m_sBox(3, 125) = &HE3674340
1148   m_sBox(0, 126) = &HC5C43465
1150   m_sBox(1, 126) = &H713E38D8
1152   m_sBox(2, 126) = &H3D28F89E
1154   m_sBox(3, 126) = &HF16DFF20
1156   m_sBox(0, 127) = &H153E21E7
1158   m_sBox(1, 127) = &H8FB03D4A
1160   m_sBox(2, 127) = &HE6E39F2B
1162   m_sBox(3, 127) = &HDB83ADF7
1164   m_sBox(0, 128) = &HE93D5A68
1166   m_sBox(1, 128) = &H948140F7
1168   m_sBox(2, 128) = &HF64C261C
1170   m_sBox(3, 128) = &H94692934
1172   m_sBox(0, 129) = &H411520F7
1174   m_sBox(1, 129) = &H7602D4F7
1176   m_sBox(2, 129) = &HBCF46B2E
1178   m_sBox(3, 129) = &HD4A20068
1180   m_sBox(0, 130) = &HD4082471
1182   m_sBox(1, 130) = &H3320F46A
1184   m_sBox(2, 130) = &H43B7D4B7
1186   m_sBox(3, 130) = &H500061AF
1188   m_sBox(0, 131) = &H1E39F62E
1190   m_sBox(1, 131) = &H97244546
1192   m_sBox(2, 131) = &H14214F74
1194   m_sBox(3, 131) = &HBF8B8840
1196   m_sBox(0, 132) = &H4D95FC1D
1198   m_sBox(1, 132) = &H96B591AF
1200   m_sBox(2, 132) = &H70F4DDD3
1202   m_sBox(3, 132) = &H66A02F45
1204   m_sBox(0, 133) = &HBFBC09EC
1206   m_sBox(1, 133) = &H3BD9785
1208   m_sBox(2, 133) = &H7FAC6DD0
1210   m_sBox(3, 133) = &H31CB8504
1212   m_sBox(0, 134) = &H96EB27B3
1214   m_sBox(1, 134) = &H55FD3941
1216   m_sBox(2, 134) = &HDA2547E6
1218   m_sBox(3, 134) = &HABCA0A9A
1220   m_sBox(0, 135) = &H28507825
1222   m_sBox(1, 135) = &H530429F4
1224   m_sBox(2, 135) = &HA2C86DA
1226   m_sBox(3, 135) = &HE9B66DFB
1228   m_sBox(0, 136) = &H68DC1462
1230   m_sBox(1, 136) = &HD7486900
1232   m_sBox(2, 136) = &H680EC0A4
1234   m_sBox(3, 136) = &H27A18DEE
1236   m_sBox(0, 137) = &H4F3FFEA2
1238   m_sBox(1, 137) = &HE887AD8C
1240   m_sBox(2, 137) = &HB58CE006
1242   m_sBox(3, 137) = &H7AF4D6B6
1244   m_sBox(0, 138) = &HAACE1E7C
1246   m_sBox(1, 138) = &HD3375FEC
1248   m_sBox(2, 138) = &HCE78A399
1250   m_sBox(3, 138) = &H406B2A42
1252   m_sBox(0, 139) = &H20FE9E35
1254   m_sBox(1, 139) = &HD9F385B9
1256   m_sBox(2, 139) = &HEE39D7AB
1258   m_sBox(3, 139) = &H3B124E8B
1260   m_sBox(0, 140) = &H1DC9FAF7
1262   m_sBox(1, 140) = &H4B6D1856
1264   m_sBox(2, 140) = &H26A36631
1266   m_sBox(3, 140) = &HEAE397B2
1268   m_sBox(0, 141) = &H3A6EFA74
1270   m_sBox(1, 141) = &HDD5B4332
1272   m_sBox(2, 141) = &H6841E7F7
1274   m_sBox(3, 141) = &HCA7820FB
1276   m_sBox(0, 142) = &HFB0AF54E
1278   m_sBox(1, 142) = &HD8FEB397
1280   m_sBox(2, 142) = &H454056AC
1282   m_sBox(3, 142) = &HBA489527
1284   m_sBox(0, 143) = &H55533A3A
1286   m_sBox(1, 143) = &H20838D87
1288   m_sBox(2, 143) = &HFE6BA9B7
1290   m_sBox(3, 143) = &HD096954B
1292   m_sBox(0, 144) = &H55A867BC
1294   m_sBox(1, 144) = &HA1159A58
1296   m_sBox(2, 144) = &HCCA92963
1298   m_sBox(3, 144) = &H99E1DB33
1300   m_sBox(0, 145) = &HA62A4A56
1302   m_sBox(1, 145) = &H3F3125F9
1304   m_sBox(2, 145) = &H5EF47E1C
1306   m_sBox(3, 145) = &H9029317C
1308   m_sBox(0, 146) = &HFDF8E802
1310   m_sBox(1, 146) = &H4272F70
1312   m_sBox(2, 146) = &H80BB155C
1314   m_sBox(3, 146) = &H5282CE3
1316   m_sBox(0, 147) = &H95C11548
1318   m_sBox(1, 147) = &HE4C66D22
1320   m_sBox(2, 147) = &H48C1133F
1322   m_sBox(3, 147) = &HC70F86DC
1324   m_sBox(0, 148) = &H7F9C9EE
1326   m_sBox(1, 148) = &H41041F0F
1328   m_sBox(2, 148) = &H404779A4
1330   m_sBox(3, 148) = &H5D886E17
1332   m_sBox(0, 149) = &H325F51EB
1334   m_sBox(1, 149) = &HD59BC0D1
1336   m_sBox(2, 149) = &HF2BCC18F
1338   m_sBox(3, 149) = &H41113564
1340   m_sBox(0, 150) = &H257B7834
1342   m_sBox(1, 150) = &H602A9C60
1344   m_sBox(2, 150) = &HDFF8E8A3
1346   m_sBox(3, 150) = &H1F636C1B
1348   m_sBox(0, 151) = &HE12B4C2
1350   m_sBox(1, 151) = &H2E1329E
1352   m_sBox(2, 151) = &HAF664FD1
1354   m_sBox(3, 151) = &HCAD18115
1356   m_sBox(0, 152) = &H6B2395E0
1358   m_sBox(1, 152) = &H333E92E1
1360   m_sBox(2, 152) = &H3B240B62
1362   m_sBox(3, 152) = &HEEBEB922
1364   m_sBox(0, 153) = &H85B2A20E
1366   m_sBox(1, 153) = &HE6BA0D99
1368   m_sBox(2, 153) = &HDE720C8C
1370   m_sBox(3, 153) = &H2DA2F728
1372   m_sBox(0, 154) = &HD0127845
1374   m_sBox(1, 154) = &H95B794FD
1376   m_sBox(2, 154) = &H647D0862
1378   m_sBox(3, 154) = &HE7CCF5F0
1380   m_sBox(0, 155) = &H5449A36F
1382   m_sBox(1, 155) = &H877D48FA
1384   m_sBox(2, 155) = &HC39DFD27
1386   m_sBox(3, 155) = &HF33E8D1E
1388   m_sBox(0, 156) = &HA476341
1390   m_sBox(1, 156) = &H992EFF74
1392   m_sBox(2, 156) = &H3A6F6EAB
1394   m_sBox(3, 156) = &HF4F8FD37
1396   m_sBox(0, 157) = &HA812DC60
1398   m_sBox(1, 157) = &HA1EBDDF8
1400   m_sBox(2, 157) = &H991BE14C
1402   m_sBox(3, 157) = &HDB6E6B0D
1404   m_sBox(0, 158) = &HC67B5510
1406   m_sBox(1, 158) = &H6D672C37
1408   m_sBox(2, 158) = &H2765D43B
1410   m_sBox(3, 158) = &HDCD0E804
1412   m_sBox(0, 159) = &HF1290DC7
1414   m_sBox(1, 159) = &HCC00FFA3
1416   m_sBox(2, 159) = &HB5390F92
1418   m_sBox(3, 159) = &H690FED0B
1420   m_sBox(0, 160) = &H667B9FFB
1422   m_sBox(1, 160) = &HCEDB7D9C
1424   m_sBox(2, 160) = &HA091CF0B
1426   m_sBox(3, 160) = &HD9155EA3
1428   m_sBox(0, 161) = &HBB132F88
1430   m_sBox(1, 161) = &H515BAD24
1432   m_sBox(2, 161) = &H7B9479BF
1434   m_sBox(3, 161) = &H763BD6EB
1436   m_sBox(0, 162) = &H37392EB3
1438   m_sBox(1, 162) = &HCC115979
1440   m_sBox(2, 162) = &H8026E297
1442   m_sBox(3, 162) = &HF42E312D
1444   m_sBox(0, 163) = &H6842ADA7
1446   m_sBox(1, 163) = &HC66A2B3B
1448   m_sBox(2, 163) = &H12754CCC
1450   m_sBox(3, 163) = &H782EF11C
1452   m_sBox(0, 164) = &H6A124237
1454   m_sBox(1, 164) = &HB79251E7
1456   m_sBox(2, 164) = &H6A1BBE6
1458   m_sBox(3, 164) = &H4BFB6350
1460   m_sBox(0, 165) = &H1A6B1018
1462   m_sBox(1, 165) = &H11CAEDFA
1464   m_sBox(2, 165) = &H3D25BDD8
1466   m_sBox(3, 165) = &HE2E1C3C9
1468   m_sBox(0, 166) = &H44421659
1470   m_sBox(1, 166) = &HA121386
1472   m_sBox(2, 166) = &HD90CEC6E
1474   m_sBox(3, 166) = &HD5ABEA2A
1476   m_sBox(0, 167) = &H64AF674E
1478   m_sBox(1, 167) = &HDA86A85F
1480   m_sBox(2, 167) = &HBEBFE988
1482   m_sBox(3, 167) = &H64E4C3FE
1484   m_sBox(0, 168) = &H9DBC8057
1486   m_sBox(1, 168) = &HF0F7C086
1488   m_sBox(2, 168) = &H60787BF8
1490   m_sBox(3, 168) = &H6003604D
1492   m_sBox(0, 169) = &HD1FD8346
1494   m_sBox(1, 169) = &HF6381FB0
1496   m_sBox(2, 169) = &H7745AE04
1498   m_sBox(3, 169) = &HD736FCCC
1500   m_sBox(0, 170) = &H83426B33
1502   m_sBox(1, 170) = &HF01EAB71
1504   m_sBox(2, 170) = &HB0804187
1506   m_sBox(3, 170) = &H3C005E5F
1508   m_sBox(0, 171) = &H77A057BE
1510   m_sBox(1, 171) = &HBDE8AE24
1512   m_sBox(2, 171) = &H55464299
1514   m_sBox(3, 171) = &HBF582E61
1516   m_sBox(0, 172) = &H4E58F48F
1518   m_sBox(1, 172) = &HF2DDFDA2
1520   m_sBox(2, 172) = &HF474EF38
1522   m_sBox(3, 172) = &H8789BDC2
1524   m_sBox(0, 173) = &H5366F9C3
1526   m_sBox(1, 173) = &HC8B38E74
1528   m_sBox(2, 173) = &HB475F255
1530   m_sBox(3, 173) = &H46FCD9B9
1532   m_sBox(0, 174) = &H7AEB2661
1534   m_sBox(1, 174) = &H8B1DDF84
1536   m_sBox(2, 174) = &H846A0E79
1538   m_sBox(3, 174) = &H915F95E2
1540   m_sBox(0, 175) = &H466E598E
1542   m_sBox(1, 175) = &H20B45770
1544   m_sBox(2, 175) = &H8CD55591
1546   m_sBox(3, 175) = &HC902DE4C
1548   m_sBox(0, 176) = &HB90BACE1
1550   m_sBox(1, 176) = &HBB8205D0
1552   m_sBox(2, 176) = &H11A86248
1554   m_sBox(3, 176) = &H7574A99E
1556   m_sBox(0, 177) = &HB77F19B6
1558   m_sBox(1, 177) = &HE0A9DC09
1560   m_sBox(2, 177) = &H662D09A1
1562   m_sBox(3, 177) = &HC4324633
1564   m_sBox(0, 178) = &HE85A1F02
1566   m_sBox(1, 178) = &H9F0BE8C
1568   m_sBox(2, 178) = &H4A99A025
1570   m_sBox(3, 178) = &H1D6EFE10
1572   m_sBox(0, 179) = &H1AB93D1D
1574   m_sBox(1, 179) = &HBA5A4DF
1576   m_sBox(2, 179) = &HA186F20F
1578   m_sBox(3, 179) = &H2868F169
1580   m_sBox(0, 180) = &HDCB7DA83
1582   m_sBox(1, 180) = &H573906FE
1584   m_sBox(2, 180) = &HA1E2CE9B
1586   m_sBox(3, 180) = &H4FCD7F52
1588   m_sBox(0, 181) = &H50115E01
1590   m_sBox(1, 181) = &HA70683FA
1592   m_sBox(2, 181) = &HA002B5C4
1594   m_sBox(3, 181) = &HDE6D027
1596   m_sBox(0, 182) = &H9AF88C27
1598   m_sBox(1, 182) = &H773F8641
1600   m_sBox(2, 182) = &HC3604C06
1602   m_sBox(3, 182) = &H61A806B5
1604   m_sBox(0, 183) = &HF0177A28
1606   m_sBox(1, 183) = &HC0F586E0
1608   m_sBox(2, 183) = &H6058AA
1610   m_sBox(3, 183) = &H30DC7D62
1612   m_sBox(0, 184) = &H11E69ED7
1614   m_sBox(1, 184) = &H2338EA63
1616   m_sBox(2, 184) = &H53C2DD94
1618   m_sBox(3, 184) = &HC2C21634
1620   m_sBox(0, 185) = &HBBCBEE56
1622   m_sBox(1, 185) = &H90BCB6DE
1624   m_sBox(2, 185) = &HEBFC7DA1
1626   m_sBox(3, 185) = &HCE591D76
1628   m_sBox(0, 186) = &H6F05E409
1630   m_sBox(1, 186) = &H4B7C0188
1632   m_sBox(2, 186) = &H39720A3D
1634   m_sBox(3, 186) = &H7C927C24
1636   m_sBox(0, 187) = &H86E3725F
1638   m_sBox(1, 187) = &H724D9DB9
1640   m_sBox(2, 187) = &H1AC15BB4
1642   m_sBox(3, 187) = &HD39EB8FC
1644   m_sBox(0, 188) = &HED545578
1646   m_sBox(1, 188) = &H8FCA5B5
1648   m_sBox(2, 188) = &HD83D7CD3
1650   m_sBox(3, 188) = &H4DAD0FC4
1652   m_sBox(0, 189) = &H1E50EF5E
1654   m_sBox(1, 189) = &HB161E6F8
1656   m_sBox(2, 189) = &HA28514D9
1658   m_sBox(3, 189) = &H6C51133C
1660   m_sBox(0, 190) = &H6FD5C7E7
1662   m_sBox(1, 190) = &H56E14EC4
1664   m_sBox(2, 190) = &H362ABFCE
1666   m_sBox(3, 190) = &HDDC6C837
1668   m_sBox(0, 191) = &HD79A3234
1670   m_sBox(1, 191) = &H92638212
1672   m_sBox(2, 191) = &H670EFA8E
1674   m_sBox(3, 191) = &H406000E0
1676   m_sBox(0, 192) = &H3A39CE37
1678   m_sBox(1, 192) = &HD3FAF5CF
1680   m_sBox(2, 192) = &HABC27737
1682   m_sBox(3, 192) = &H5AC52D1B
1684   m_sBox(0, 193) = &H5CB0679E
1686   m_sBox(1, 193) = &H4FA33742
1688   m_sBox(2, 193) = &HD3822740
1690   m_sBox(3, 193) = &H99BC9BBE
1692   m_sBox(0, 194) = &HD5118E9D
1694   m_sBox(1, 194) = &HBF0F7315
1696   m_sBox(2, 194) = &HD62D1C7E
1698   m_sBox(3, 194) = &HC700C47B
1700   m_sBox(0, 195) = &HB78C1B6B
1702   m_sBox(1, 195) = &H21A19045
1704   m_sBox(2, 195) = &HB26EB1BE
1706   m_sBox(3, 195) = &H6A366EB4
1708   m_sBox(0, 196) = &H5748AB2F
1710   m_sBox(1, 196) = &HBC946E79
1712   m_sBox(2, 196) = &HC6A376D2
1714   m_sBox(3, 196) = &H6549C2C8
1716   m_sBox(0, 197) = &H530FF8EE
1718   m_sBox(1, 197) = &H468DDE7D
1720   m_sBox(2, 197) = &HD5730A1D
1722   m_sBox(3, 197) = &H4CD04DC6
1724   m_sBox(0, 198) = &H2939BBDB
1726   m_sBox(1, 198) = &HA9BA4650
1728   m_sBox(2, 198) = &HAC9526E8
1730   m_sBox(3, 198) = &HBE5EE304
1732   m_sBox(0, 199) = &HA1FAD5F0
1734   m_sBox(1, 199) = &H6A2D519A
1736   m_sBox(2, 199) = &H63EF8CE2
1738   m_sBox(3, 199) = &H9A86EE22
1740   m_sBox(0, 200) = &HC089C2B8
1742   m_sBox(1, 200) = &H43242EF6
1744   m_sBox(2, 200) = &HA51E03AA
1746   m_sBox(3, 200) = &H9CF2D0A4
1748   m_sBox(0, 201) = &H83C061BA
1750   m_sBox(1, 201) = &H9BE96A4D
1752   m_sBox(2, 201) = &H8FE51550
1754   m_sBox(3, 201) = &HBA645BD6
1756   m_sBox(0, 202) = &H2826A2F9
1758   m_sBox(1, 202) = &HA73A3AE1
1760   m_sBox(2, 202) = &H4BA99586
1762   m_sBox(3, 202) = &HEF5562E9
1764   m_sBox(0, 203) = &HC72FEFD3
1766   m_sBox(1, 203) = &HF752F7DA
1768   m_sBox(2, 203) = &H3F046F69
1770   m_sBox(3, 203) = &H77FA0A59
1772   m_sBox(0, 204) = &H80E4A915
1774   m_sBox(1, 204) = &H87B08601
1776   m_sBox(2, 204) = &H9B09E6AD
1778   m_sBox(3, 204) = &H3B3EE593
1780   m_sBox(0, 205) = &HE990FD5A
1782   m_sBox(1, 205) = &H9E34D797
1784   m_sBox(2, 205) = &H2CF0B7D9
1786   m_sBox(3, 205) = &H22B8B51
1788   m_sBox(0, 206) = &H96D5AC3A
1790   m_sBox(1, 206) = &H17DA67D
1792   m_sBox(2, 206) = &HD1CF3ED6
1794   m_sBox(3, 206) = &H7C7D2D28
1796   m_sBox(0, 207) = &H1F9F25CF
1798   m_sBox(1, 207) = &HADF2B89B
1800   m_sBox(2, 207) = &H5AD6B472
1802   m_sBox(3, 207) = &H5A88F54C
1804   m_sBox(0, 208) = &HE029AC71
1806   m_sBox(1, 208) = &HE019A5E6
1808   m_sBox(2, 208) = &H47B0ACFD
1810   m_sBox(3, 208) = &HED93FA9B
1812   m_sBox(0, 209) = &HE8D3C48D
1814   m_sBox(1, 209) = &H283B57CC
1816   m_sBox(2, 209) = &HF8D56629
1818   m_sBox(3, 209) = &H79132E28
1820   m_sBox(0, 210) = &H785F0191
1822   m_sBox(1, 210) = &HED756055
1824   m_sBox(2, 210) = &HF7960E44
1826   m_sBox(3, 210) = &HE3D35E8C
1828   m_sBox(0, 211) = &H15056DD4
1830   m_sBox(1, 211) = &H88F46DBA
1832   m_sBox(2, 211) = &H3A16125
1834   m_sBox(3, 211) = &H564F0BD
1836   m_sBox(0, 212) = &HC3EB9E15
1838   m_sBox(1, 212) = &H3C9057A2
1840   m_sBox(2, 212) = &H97271AEC
1842   m_sBox(3, 212) = &HA93A072A
1844   m_sBox(0, 213) = &H1B3F6D9B
1846   m_sBox(1, 213) = &H1E6321F5
1848   m_sBox(2, 213) = &HF59C66FB
1850   m_sBox(3, 213) = &H26DCF319
1852   m_sBox(0, 214) = &H7533D928
1854   m_sBox(1, 214) = &HB155FDF5
1856   m_sBox(2, 214) = &H3563482
1858   m_sBox(3, 214) = &H8ABA3CBB
1860   m_sBox(0, 215) = &H28517711
1862   m_sBox(1, 215) = &HC20AD9F8
1864   m_sBox(2, 215) = &HABCC5167
1866   m_sBox(3, 215) = &HCCAD925F
1868   m_sBox(0, 216) = &H4DE81751
1870   m_sBox(1, 216) = &H3830DC8E
1872   m_sBox(2, 216) = &H379D5862
1874   m_sBox(3, 216) = &H9320F991
1876   m_sBox(0, 217) = &HEA7A90C2
1878   m_sBox(1, 217) = &HFB3E7BCE
1880   m_sBox(2, 217) = &H5121CE64
1882   m_sBox(3, 217) = &H774FBE32
1884   m_sBox(0, 218) = &HA8B6E37E
1886   m_sBox(1, 218) = &HC3293D46
1888   m_sBox(2, 218) = &H48DE5369
1890   m_sBox(3, 218) = &H6413E680
1892   m_sBox(0, 219) = &HA2AE0810
1894   m_sBox(1, 219) = &HDD6DB224
1896   m_sBox(2, 219) = &H69852DFD
1898   m_sBox(3, 219) = &H9072166
1900   m_sBox(0, 220) = &HB39A460A
1902   m_sBox(1, 220) = &H6445C0DD
1904   m_sBox(2, 220) = &H586CDECF
1906   m_sBox(3, 220) = &H1C20C8AE
1908   m_sBox(0, 221) = &H5BBEF7DD
1910   m_sBox(1, 221) = &H1B588D40
1912   m_sBox(2, 221) = &HCCD2017F
1914   m_sBox(3, 221) = &H6BB4E3BB
1916   m_sBox(0, 222) = &HDDA26A7E
1918   m_sBox(1, 222) = &H3A59FF45
1920   m_sBox(2, 222) = &H3E350A44
1922   m_sBox(3, 222) = &HBCB4CDD5
1924   m_sBox(0, 223) = &H72EACEA8
1926   m_sBox(1, 223) = &HFA6484BB
1928   m_sBox(2, 223) = &H8D6612AE
1930   m_sBox(3, 223) = &HBF3C6F47
1932   m_sBox(0, 224) = &HD29BE463
1934   m_sBox(1, 224) = &H542F5D9E
1936   m_sBox(2, 224) = &HAEC2771B
1938   m_sBox(3, 224) = &HF64E6370
1940   m_sBox(0, 225) = &H740E0D8D
1942   m_sBox(1, 225) = &HE75B1357
1944   m_sBox(2, 225) = &HF8721671
1946   m_sBox(3, 225) = &HAF537D5D
1948   m_sBox(0, 226) = &H4040CB08
1950   m_sBox(1, 226) = &H4EB4E2CC
1952   m_sBox(2, 226) = &H34D2466A
1954   m_sBox(3, 226) = &H115AF84
1956   m_sBox(0, 227) = &HE1B00428
1958   m_sBox(1, 227) = &H95983A1D
1960   m_sBox(2, 227) = &H6B89FB4
1962   m_sBox(3, 227) = &HCE6EA048
1964   m_sBox(0, 228) = &H6F3F3B82
1966   m_sBox(1, 228) = &H3520AB82
1968   m_sBox(2, 228) = &H11A1D4B
1970   m_sBox(3, 228) = &H277227F8
1972   m_sBox(0, 229) = &H611560B1
1974   m_sBox(1, 229) = &HE7933FDC
1976   m_sBox(2, 229) = &HBB3A792B
1978   m_sBox(3, 229) = &H344525BD
1980   m_sBox(0, 230) = &HA08839E1
1982   m_sBox(1, 230) = &H51CE794B
1984   m_sBox(2, 230) = &H2F32C9B7
1986   m_sBox(3, 230) = &HA01FBAC9
1988   m_sBox(0, 231) = &HE01CC87E
1990   m_sBox(1, 231) = &HBCC7D1F6
1992   m_sBox(2, 231) = &HCF0111C3
1994   m_sBox(3, 231) = &HA1E8AAC7
1996   m_sBox(0, 232) = &H1A908749
1998   m_sBox(1, 232) = &HD44FBD9A
2000   m_sBox(2, 232) = &HD0DADECB
2002   m_sBox(3, 232) = &HD50ADA38
2004   m_sBox(0, 233) = &H339C32A
2006   m_sBox(1, 233) = &HC6913667
2008   m_sBox(2, 233) = &H8DF9317C
2010   m_sBox(3, 233) = &HE0B12B4F
2012   m_sBox(0, 234) = &HF79E59B7
2014   m_sBox(1, 234) = &H43F5BB3A
2016   m_sBox(2, 234) = &HF2D519FF
2018   m_sBox(3, 234) = &H27D9459C
2020   m_sBox(0, 235) = &HBF97222C
2022   m_sBox(1, 235) = &H15E6FC2A
2024   m_sBox(2, 235) = &HF91FC71
2026   m_sBox(3, 235) = &H9B941525
2028   m_sBox(0, 236) = &HFAE59361
2030   m_sBox(1, 236) = &HCEB69CEB
2032   m_sBox(2, 236) = &HC2A86459
2034   m_sBox(3, 236) = &H12BAA8D1
2036   m_sBox(0, 237) = &HB6C1075E
2038   m_sBox(1, 237) = &HE3056A0C
2040   m_sBox(2, 237) = &H10D25065
2042   m_sBox(3, 237) = &HCB03A442
2044   m_sBox(0, 238) = &HE0EC6E0E
2046   m_sBox(1, 238) = &H1698DB3B
2048   m_sBox(2, 238) = &H4C98A0BE
2050   m_sBox(3, 238) = &H3278E964
2052   m_sBox(0, 239) = &H9F1F9532
2054   m_sBox(1, 239) = &HE0D392DF
2056   m_sBox(2, 239) = &HD3A0342B
2058   m_sBox(3, 239) = &H8971F21E
2060   m_sBox(0, 240) = &H1B0A7441
2062   m_sBox(1, 240) = &H4BA3348C
2064   m_sBox(2, 240) = &HC5BE7120
2066   m_sBox(3, 240) = &HC37632D8
2068   m_sBox(0, 241) = &HDF359F8D
2070   m_sBox(1, 241) = &H9B992F2E
2072   m_sBox(2, 241) = &HE60B6F47
2074   m_sBox(3, 241) = &HFE3F11D
2076   m_sBox(0, 242) = &HE54CDA54
2078   m_sBox(1, 242) = &H1EDAD891
2080   m_sBox(2, 242) = &HCE6279CF
2082   m_sBox(3, 242) = &HCD3E7E6F
2084   m_sBox(0, 243) = &H1618B166
2086   m_sBox(1, 243) = &HFD2C1D05
2088   m_sBox(2, 243) = &H848FD2C5
2090   m_sBox(3, 243) = &HF6FB2299
2092   m_sBox(0, 244) = &HF523F357
2094   m_sBox(1, 244) = &HA6327623
2096   m_sBox(2, 244) = &H93A83531
2098   m_sBox(3, 244) = &H56CCCD02
2100   m_sBox(0, 245) = &HACF08162
2102   m_sBox(1, 245) = &H5A75EBB5
2104   m_sBox(2, 245) = &H6E163697
2106   m_sBox(3, 245) = &H88D273CC
2108   m_sBox(0, 246) = &HDE966292
2110   m_sBox(1, 246) = &H81B949D0
2112   m_sBox(2, 246) = &H4C50901B
2114   m_sBox(3, 246) = &H71C65614
2116   m_sBox(0, 247) = &HE6C6C7BD
2118   m_sBox(1, 247) = &H327A140A
2120   m_sBox(2, 247) = &H45E1D006
2122   m_sBox(3, 247) = &HC3F27B9A
2124   m_sBox(0, 248) = &HC9AA53FD
2126   m_sBox(1, 248) = &H62A80F00
2128   m_sBox(2, 248) = &HBB25BFE2
2130   m_sBox(3, 248) = &H35BDD2F6
2132   m_sBox(0, 249) = &H71126905
2134   m_sBox(1, 249) = &HB2040222
2136   m_sBox(2, 249) = &HB6CBCF7C
2138   m_sBox(3, 249) = &HCD769C2B
2140   m_sBox(0, 250) = &H53113EC0
2142   m_sBox(1, 250) = &H1640E3D3
2144   m_sBox(2, 250) = &H38ABBD60
2146   m_sBox(3, 250) = &H2547ADF0
2148   m_sBox(0, 251) = &HBA38209C
2150   m_sBox(1, 251) = &HF746CE76
2152   m_sBox(2, 251) = &H77AFA1C5
2154   m_sBox(3, 251) = &H20756060
2156   m_sBox(0, 252) = &H85CBFE4E
2158   m_sBox(1, 252) = &H8AE88DD8
2160   m_sBox(2, 252) = &H7AAAF9B0
2162   m_sBox(3, 252) = &H4CF9AA7E
2164   m_sBox(0, 253) = &H1948C25C
2166   m_sBox(1, 253) = &H2FB8A8C
2168   m_sBox(2, 253) = &H1C36AE4
2170   m_sBox(3, 253) = &HD6EBE1F9
2172   m_sBox(0, 254) = &H90D4F869
2174   m_sBox(1, 254) = &HA65CDEA0
2176   m_sBox(2, 254) = &H3F09252D
2178   m_sBox(3, 254) = &HC208E69F
2180   m_sBox(0, 255) = &HB74E6132
2182   m_sBox(1, 255) = &HCE77E25B
2184   m_sBox(2, 255) = &H578FDFE3
2186   m_sBox(3, 255) = &H3AC372E6
         '<EhFooter>
         Exit Sub

Class_Initialize_Err:
         Err.Raise vbObjectError + 100, _
                   "OASISRemoteAdmin.clsBlowfish.Class_Initialize", _
                   "clsBlowfish component failure"
         '</EhFooter>
End Sub



Public Function EncryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String, Optional OutputIn64 As Boolean) As Boolean
        '<EhHeader>
        On Error GoTo EncryptFile_Err
        '</EhHeader>
100     Class_Initialize
102     If FileExist(InFile) = False Then
104         EncryptFile = False
            Exit Function
        End If
106     If FileExist(OutFile) = True And Overwrite = False Then
108         EncryptFile = False
            Exit Function
        End If
        Dim Buffer() As Byte, FileO As Integer
110     FileO = FreeFile
112     Open InFile For Binary As #FileO
114         ReDim Buffer(0 To LOF(FileO) - 1)
116         Get #FileO, , Buffer()
118     Close #FileO
120     Call EncryptByte(Buffer(), Key)
122     If FileExist(OutFile) = True Then Kill OutFile
124     FileO = FreeFile
126     Open OutFile For Binary As #FileO
128         If OutputIn64 = True Then
130             Put #FileO, , EncodeArray64(Buffer())
            Else
132             Put #FileO, , Buffer()
            End If
134     Close #FileO
136     EncryptFile = True
138     Erase Buffer(): Key = "": InFile = "": OutFile = ""
        Exit Function

ErrorHandler:
140     Erase Buffer(): Key = "": InFile = "": OutFile = ""
142     EncryptFile = False
        '<EhFooter>
        Exit Function

EncryptFile_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.EncryptFile", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function
Public Function DecryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String, Optional IsFileIn64 As Boolean) As Boolean
        '<EhHeader>
        On Error GoTo DecryptFile_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
100     If FileExist(InFile) = False Then
102         DecryptFile = False
            Exit Function
        End If
104     If FileExist(OutFile) = True And Overwrite = False Then
106         DecryptFile = False
            Exit Function
        End If
        Dim Buffer() As Byte, FileO As Integer
108     FileO = FreeFile
110     Open InFile For Binary As #FileO
112         ReDim Buffer(0 To LOF(FileO) - 1)
114         Get #FileO, , Buffer()
116     Close #FileO
118     If IsFileIn64 = True Then Buffer() = DecodeArray64(StrConv(Buffer(), vbUnicode))
120     Call DecryptByte(Buffer(), Key)
122     If FileExist(OutFile) = True Then Kill OutFile
124     Open OutFile For Binary As #FileO
126         Put #FileO, , Buffer()
128     Close #FileO
130     DecryptFile = True
132     Erase Buffer(): Key = "": InFile = "": OutFile = ""
        Exit Function

ErrorHandler:
134     Erase Buffer(): Key = "": InFile = "": OutFile = ""
136     DecryptFile = False
        '<EhFooter>
        Exit Function

DecryptFile_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.DecryptFile", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function
Private Function FileExist(FilePath As String) As Boolean
        '<EhHeader>
        On Error GoTo FileExist_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
100     Call FileLen(FilePath)
102     FileExist = True
        Exit Function

ErrorHandler:
104     FileExist = False
        '<EhFooter>
        Exit Function

FileExist_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsBlowfish.FileExist", _
                  "clsBlowfish component failure"
        '</EhFooter>
End Function
