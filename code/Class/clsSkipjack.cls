VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkipjack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Standard Skipjack implementation with file support, Base64 conversion,
' and overall optimisations for Visual Basic. Skipjack is questionable since
' it was developed by the NSA for use in Clipper and Capstone chips ... ciphers
' from the NSA can suggest backdoors or hidden weaknesses but this appears to
' be a farely decent cipher. I would recommend it for low-risk solutions.
'
' Skipjack is property of the NSA but free for public use.
'
' Information can be found at:
' http://csrc.nist.gov/encryption/tkencryption.html

Private Declare Sub CopyMem _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Event Progress(Percent As Long)
Private m_KeyValue As String
Private m_F(0 To 255) As Byte
Private m_Key(0 To 127) As Byte
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
Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (ByVal Destination As Long, _
                                       ByVal Source As Long, _
                                       ByVal Length As Long)

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
                  "OASISRemoteAdmin.clsSkipjack.Initialize64", _
                  "clsSkipjack component failure"
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
                  "OASISRemoteAdmin.clsSkipjack.Decode64", _
                  "clsSkipjack component failure"
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
                  "OASISRemoteAdmin.clsSkipjack.DecodeArray64", _
                  "clsSkipjack component failure"
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
                  "OASISRemoteAdmin.clsSkipjack.Encode64", _
                  "clsSkipjack component failure"
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
                  "OASISRemoteAdmin.clsSkipjack.EncodeArray64", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Function

Private Function FileExist(Filename As String) As Boolean
        '<EhHeader>
        On Error GoTo FileExist_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
100     Call FileLen(Filename)
102     FileExist = True
        Exit Function
    
ErrorHandler:
104     FileExist = False
        '<EhFooter>
        Exit Function

FileExist_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSkipjack.FileExist", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Function

Public Function DecryptByte(byteArray() As Byte, _
                            Optional Key As String) As String
        '<EhHeader>
        On Error GoTo DecryptByte_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
        Dim i As Long, u As Long, K As Long, Temp As Byte, Round As Long, Offset As Long, OrigLen As Long, CipherLen As Long, g(0 To 5) As Byte, NextPercent As Long, CurrPercent As Long, Counter(0 To 32) As Byte, w(0 To 3, 0 To 33) As Integer

100     If (Len(Key) > 0) Then Me.Key = Key
102     CipherLen = UBound(byteArray) + 1

104     For Offset = 0 To (CipherLen - 1) Step 2
106         Temp = byteArray(Offset)
108         byteArray(Offset) = byteArray(Offset + 1)
110         byteArray(Offset + 1) = Temp
        Next

112     For Offset = 0 To (CipherLen - 1) Step 8
114         Call CopyMem(w(0, 32), byteArray(Offset), 8)
116         K = 32
118         u = 31

120         For i = 0 To 32
122             Counter(i) = i + 1
            Next

124         For Round = 1 To 2
126             For i = 1 To 8
128                 Call CopyMem(g(4), w(1, K), 2)
130                 g(3) = m_F(g(5) Xor m_Key(4 * u + 3)) Xor g(4)
132                 g(2) = m_F(g(3) Xor m_Key(4 * u + 2)) Xor g(5)
134                 g(0) = m_F(g(2) Xor m_Key(4 * u + 1)) Xor g(3)
136                 g(1) = m_F(g(0) Xor m_Key(4 * u)) Xor g(2)
138                 Call CopyMem(w(0, K - 1), g(0), 2)
140                 w(1, K - 1) = w(0, K - 1) Xor w(2, K) Xor Counter(K - 1)
142                 w(2, K - 1) = w(3, K)
144                 w(3, K - 1) = w(0, K)
146                 u = u - 1
148                 K = K - 1
                Next

150             For i = 1 To 8
152                 Call CopyMem(g(4), w(1, K), 2)
154                 g(3) = m_F(g(5) Xor m_Key(4 * u + 3)) Xor g(4)
156                 g(2) = m_F(g(3) Xor m_Key(4 * u + 2)) Xor g(5)
158                 g(0) = m_F(g(2) Xor m_Key(4 * u + 1)) Xor g(3)
160                 g(1) = m_F(g(0) Xor m_Key(4 * u)) Xor g(2)
162                 Call CopyMem(w(0, K - 1), g(0), 2)
164                 w(1, K - 1) = w(2, K)
166                 w(2, K - 1) = w(3, K)
168                 w(3, K - 1) = w(0, K) Xor w(1, K) Xor Counter(K - 1)
170                 u = u - 1
172                 K = K - 1
                Next
            Next

174         w(0, 0) = w(0, 0) Xor w(0, 33)
176         w(1, 0) = w(1, 0) Xor w(1, 33)
178         w(2, 0) = w(2, 0) Xor w(2, 33)
180         w(3, 0) = w(3, 0) Xor w(3, 33)
182         Call CopyMem(byteArray(Offset), w(0, 0), 8)

184         If (Offset >= NextPercent) Then
186             CurrPercent = Int((Offset / CipherLen) * 100)
188             NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1

190             RaiseEvent Progress(CurrPercent)
            End If

192         Call CopyMem(w(0, 33), w(0, 32), 8)
        Next
    
194     For Offset = 0 To (CipherLen - 1) Step 2
196         Temp = byteArray(Offset)
198         byteArray(Offset) = byteArray(Offset + 1)
200         byteArray(Offset + 1) = Temp
        Next

202     Call CopyMem(OrigLen, byteArray(8), 4)

204     If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Skipjack decryption")
206     Call CopyMem(byteArray(0), byteArray(12), OrigLen)
208     ReDim Preserve byteArray(OrigLen - 1)

210     If (CurrPercent <> 100) Then RaiseEvent Progress(100)

ErrorHandler:
        '<EhFooter>
        Exit Function

DecryptByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSkipjack.DecryptByte", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Function

Public Function DecryptString(Text As String, _
                              Optional Key As String, _
                              Optional IsTextIn64 As Boolean) As String
        '<EhHeader>
        On Error GoTo DecryptString_Err
        '</EhHeader>
        Dim byteArray() As Byte

100     If IsTextIn64 = True Then Text = Decode64(Text)
102     byteArray() = StrConv(Text, vbFromUnicode)
104     Call DecryptByte(byteArray(), Key)
106     DecryptString = StrConv(byteArray(), vbUnicode)
        '<EhFooter>
        Exit Function

DecryptString_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSkipjack.DecryptString", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Function

Public Sub EncryptByte(byteArray() As Byte, _
                       Optional Key As String)
        '<EhHeader>
        On Error GoTo EncryptByte_Err
        '</EhHeader>
        Dim i As Long, K As Long, Temp As Byte, Round As Long, Offset As Long, OrigLen As Long, Counter As Long, g(0 To 5) As Byte, CipherLen As Long, NextPercent As Long, CurrPercent As Long, w(0 To 3, 0 To 32) As Integer
    
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

120     For Offset = 0 To (CipherLen - 1) Step 2
122         Temp = byteArray(Offset)
124         byteArray(Offset) = byteArray(Offset + 1)
126         byteArray(Offset + 1) = Temp
        Next

128     For Offset = 0 To (CipherLen - 1) Step 8
130         Call CopyMem(w(0, 0), byteArray(Offset), 8)
132         w(0, 0) = w(0, 0) Xor w(0, 32)
134         w(1, 0) = w(1, 0) Xor w(1, 32)
136         w(2, 0) = w(2, 0) Xor w(2, 32)
138         w(3, 0) = w(3, 0) Xor w(3, 32)
140         K = 0
142         Counter = 1

144         For Round = 1 To 2
146             For i = 1 To 8
148                 Call CopyMem(g(0), w(0, K), 2)
150                 g(2) = m_F(g(0) Xor m_Key(4 * K)) Xor g(1)
152                 g(3) = m_F(g(2) Xor m_Key(4 * K + 1)) Xor g(0)
154                 g(5) = m_F(g(3) Xor m_Key(4 * K + 2)) Xor g(2)
156                 g(4) = m_F(g(5) Xor m_Key(4 * K + 3)) Xor g(3)
158                 Call CopyMem(w(1, K + 1), g(4), 2)
160                 w(0, K + 1) = w(1, K + 1) Xor w(3, K) Xor Counter
162                 w(2, K + 1) = w(1, K)
164                 w(3, K + 1) = w(2, K)
166                 Counter = Counter + 1
168                 K = K + 1
                Next

170             For i = 1 To 8
172                 Call CopyMem(g(0), w(0, K), 2)
174                 g(2) = m_F(g(0) Xor m_Key(4 * K)) Xor g(1)
176                 g(3) = m_F(g(2) Xor m_Key(4 * K + 1)) Xor g(0)
178                 g(5) = m_F(g(3) Xor m_Key(4 * K + 2)) Xor g(2)
180                 g(4) = m_F(g(5) Xor m_Key(4 * K + 3)) Xor g(3)
182                 Call CopyMem(w(1, K + 1), g(4), 2)
184                 w(0, K + 1) = w(3, K)
186                 w(2, K + 1) = w(0, K) Xor w(1, K) Xor Counter
188                 w(3, K + 1) = w(2, K)
190                 Counter = Counter + 1
192                 K = K + 1
                Next
            Next

194         Call CopyMem(byteArray(Offset), w(0, 32), 8)

196         If (Offset >= NextPercent) Then
198             CurrPercent = Int((Offset / CipherLen) * 100)
200             NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1

202             RaiseEvent Progress(CurrPercent)
            End If

        Next

204     For Offset = 0 To (CipherLen - 1) Step 2
206         Temp = byteArray(Offset)
208         byteArray(Offset) = byteArray(Offset + 1)
210         byteArray(Offset + 1) = Temp
        Next

212     If (CurrPercent <> 100) Then RaiseEvent Progress(100)
        '<EhFooter>
        Exit Sub

EncryptByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSkipjack.EncryptByte", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Sub

Public Function EncryptFile(InFile As String, _
                            OutFile As String, _
                            Overwrite As Boolean, _
                            Optional Key As String, _
                            Optional OutputIn64 As Boolean) As Boolean
        '<EhHeader>
        On Error GoTo EncryptFile_Err
        '</EhHeader>
        On Error GoTo ErrorHandler

100     If FileExist(InFile) = False Then
102         EncryptFile = False
            Exit Function
        End If

104     If FileExist(OutFile) = True And Overwrite = False Then
106         EncryptFile = False
            Exit Function
        End If

        Dim FileO As Integer, Buffer() As Byte
108     FileO = FreeFile
110     Open InFile For Binary As #FileO
112     ReDim Buffer(0 To LOF(FileO) - 1)
114     Get #FileO, , Buffer()
116     Close #FileO
118     Call EncryptByte(Buffer(), Key)

120     If FileExist(OutFile) = True Then Kill OutFile
122     FileO = FreeFile
124     Open OutFile For Binary As #FileO

126     If OutputIn64 = True Then
128         Put #FileO, , EncodeArray64(Buffer())
        Else
130         Put #FileO, , Buffer()
        End If

132     Close #FileO
134     EncryptFile = True
        Exit Function
    
ErrorHandler:
136     EncryptFile = False
        '<EhFooter>
        Exit Function

EncryptFile_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSkipjack.EncryptFile", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Function

Public Function DecryptFile(InFile As String, _
                            OutFile As String, _
                            Overwrite As Boolean, _
                            Optional Key As String, _
                            Optional IsFileIn64 As Boolean) As Boolean
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

        Dim FileO As Integer, Buffer() As Byte
108     FileO = FreeFile
110     Open InFile For Binary As #FileO
112     ReDim Buffer(0 To LOF(FileO) - 1)
114     Get #FileO, , Buffer()
116     Close #FileO

118     If IsFileIn64 = True Then Buffer() = DecodeArray64(StrConv(Buffer(), vbUnicode))
120     Call DecryptByte(Buffer(), Key)

122     If FileExist(OutFile) = True Then Kill OutFile
124     FileO = FreeFile
126     Open OutFile For Binary As #FileO
128     Put #FileO, , Buffer()
130     Close #FileO
132     DecryptFile = True
        Exit Function

ErrorHandler:
134     DecryptFile = False
        '<EhFooter>
        Exit Function

DecryptFile_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSkipjack.DecryptFile", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Function

Public Function EncryptString(Text As String, _
                              Optional Key As String, _
                              Optional OutputIn64 As Boolean) As String
        '<EhHeader>
        On Error GoTo EncryptString_Err
        '</EhHeader>
        Dim byteArray() As Byte
100     byteArray() = StrConv(Text, vbFromUnicode)
102     Call EncryptByte(byteArray(), Key)
104     EncryptString = StrConv(byteArray(), vbUnicode)

106     If OutputIn64 = True Then EncryptString = Encode64(EncryptString)
        '<EhFooter>
        Exit Function

EncryptString_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSkipjack.EncryptString", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Function

Public Property Let Key(New_Value As String)
        '<EhHeader>
        On Error GoTo Key_Err
        '</EhHeader>
100     Class_Initialize
        Dim i As Long, pass() As Byte, PassLen As Long

102     If (New_Value = m_KeyValue) Then Exit Property
104     If (Len(New_Value) = 0) Then Exit Property
106     PassLen = Len(New_Value)
108     pass() = StrConv(New_Value, vbFromUnicode)

110     For i = 0 To 127
112         m_Key(i) = pass(i Mod PassLen)
        Next

114     m_KeyValue = New_Value
        '<EhFooter>
        Exit Property

Key_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSkipjack.Key", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Property

Private Sub Class_Initialize()
        '<EhHeader>
        On Error GoTo Class_Initialize_Err
        '</EhHeader>
        Dim a As Long, Ftable As Variant
100     Ftable = Array("A3", "D7", "09", "83", "F8", "48", "F6", "F4", "B3", "21", "15", "78", "99", "B1", "AF", "F9", "E7", "2D", "4D", "8A", "CE", "4C", "CA", "2E", "52", "95", "D9", "1E", "4E", "38", "44", "28", "0A", "DF", "02", "A0", "17", "F1", "60", "68", "12", "B7", "7A", "C3", "E9", "FA", "3D", "53", "96", "84", "6B", "BA", "F2", "63", "9A", "19", "7C", "AE", "E5", "F5", "F7", "16", "6A", "A2", "39", "B6", "7B", "0F", "C1", "93", "81", "1B", "EE", "B4", "1A", "EA", "D0", "91", "2F", "B8", "55", "B9", "DA", "85", "3F", "41", "BF", "E0", "5A", "58", "80", "5F", "66", "0B", "D8", "90", "35", "D5", "C0", "A7", "33", "06", "65", "69", "45", "00", "94", "56", "6D", "98", "9B", "76", "97", "FC", "B2", "C2", "B0", "FE", "DB", "20", "E1", "EB", "D6", "E4", "DD", "47", "4A", "1D", "42", "ED", "9E", "6E", "49", "3C", "CD", "43", "27", "D2", "07", "D4", "DE", "C7", "67", "18", "89", "CB", "30", "1F", "8D", "C6", "8F", "AA", "C8", "74", "DC", "C9", "5D", "5C", "31", "A4", _
           "70", "88", "61", "2C", "9F", "0D", "2B", "87", "50", "82", "54", "64", "26", "7D", "03", "40", "34", "4B", "1C", "73", "D1", "C4", "FD", "3B", "CC", "FB", "7F", "AB", "E6", "3E", "5B", "A5", "AD", "04", "23", "9C", "14", "51", "22", "F0", "29", "79", "71", "7E", "FF", "8C", "0E", "E2", "0C", "EF", "BC", "72", "75", "6F", "37", "A1", "EC", "D3", "8E", "62", "8B", "86", "10", "E8", "08", "77", "11", "BE", "92", "4F", "24", "C5", "32", "36", "9D", "CF", "F3", "A6", "BB", "AC", "5E", "6C", "A9", "13", "57", "25", "B5", "E3", "BD", "A8", "3A", "01", "05", "59", "2A", "46")

102     For a = 0 To 255
104         m_F(a) = Val("&H" & Ftable(a))
        Next

        '<EhFooter>
        Exit Sub

Class_Initialize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSkipjack.Class_Initialize", _
                  "clsSkipjack component failure"
        '</EhFooter>
End Sub
