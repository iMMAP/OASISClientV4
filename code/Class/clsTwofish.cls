VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTwofish"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Standard Twofish implementation with file support, Base64 conversion,
' and overall optimisations for Visual Basic. Twofish is simply another
' powerful encryption algorithm from Bruce Schneier that was a finalist
' for AES. Export restrictions have been lifted on it and I would recommend
' this as another candidate for high risk level solutions.
'
' Information can be found at:
' http://www.counterpane.com/twofish.html

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Event Progress(Percent As Long)

Public Enum TWOFISHKEYLENGTH
    TWOFISH_256 = 256
    TWOFISH_196 = 196
    TWOFISH_128 = 128
    TWOFISH_64 = 64
End Enum

Private Const Rounds = 16
Private Const BLOCK_SIZE = 16
Private Const MAX_ROUNDS = 16
Private Const INPUT_WHITEN = 0
Private Const OUTPUT_WHITEN = INPUT_WHITEN + BLOCK_SIZE / 4
Private Const ROUND_SUBKEYS = OUTPUT_WHITEN + BLOCK_SIZE / 4
Private Const GF256_FDBK_2 = &H169 / 2
Private Const GF256_FDBK_4 = &H169 / 4
Private MDS(0 To 3, 0 To 255) As Long
Private p(0 To 1, 0 To 255) As Byte
Private m_RunningCompiled As Boolean
Private sBox(0 To 1023) As Long
Private sKey() As Long
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
                  "OASISRemoteAdmin.clsTwofish.Initialize64", _
                  "clsTwofish component failure"
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
                  "OASISRemoteAdmin.clsTwofish.Decode64", _
                  "clsTwofish component failure"
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
                  "OASISRemoteAdmin.clsTwofish.DecodeArray64", _
                  "clsTwofish component failure"
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
                  "OASISRemoteAdmin.clsTwofish.Encode64", _
                  "clsTwofish component failure"
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
                  "OASISRemoteAdmin.clsTwofish.EncodeArray64", _
                  "clsTwofish component failure"
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
                  "OASISRemoteAdmin.clsTwofish.FileExist", _
                  "clsTwofish component failure"
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
                  "OASISRemoteAdmin.clsTwofish.GetWord", _
                  "clsTwofish component failure"
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
                  "OASISRemoteAdmin.clsTwofish.PutWord", _
                  "clsTwofish component failure"
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
                  "OASISRemoteAdmin.clsTwofish.UnsignedAdd", _
                  "clsTwofish component failure"
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
                  "OASISRemoteAdmin.clsTwofish.UnsignedDel", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Public Function EncryptFile(InFile As String, OutFile As String, Overwrite As Boolean, Optional Key As String, Optional OutputIn64 As Boolean) As Boolean
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
112         ReDim Buffer(0 To LOF(FileO) - 1)
114         Get #FileO, , Buffer()
116     Close #FileO
118     Call EncryptByte(Buffer(), Key)
120     If FileExist(OutFile) = True Then Kill OutFile
122     FileO = FreeFile
124     Open OutFile For Binary As #FileO
126         If OutputIn64 = True Then
128             Put #FileO, , EncodeArray64(Buffer())
            Else
130             Put #FileO, , Buffer()
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
                  "OASISRemoteAdmin.clsTwofish.EncryptFile", _
                  "clsTwofish component failure"
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
        Dim FileO As Integer, Buffer() As Byte
108     FileO = FreeFile
110     Open InFile For Binary As #FileO
112         ReDim Buffer(0 To LOF(FileO) - 1)
114         Get #FileO, , Buffer()
116     Close #FileO
118     If IsFileIn64 = True Then Buffer() = DecodeArray64(StrConv(Buffer(), vbUnicode))
120     Call DecryptByte(Buffer(), Key)
122     If FileExist(OutFile) = True Then Kill OutFile
124     FileO = FreeFile
126     Open OutFile For Binary As #FileO
128         Put #FileO, , Buffer()
130     Close #FileO
132     DecryptFile = True
        Exit Function
    
ErrorHandler:
134     DecryptFile = False
        '<EhFooter>
        Exit Function

DecryptFile_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.DecryptFile", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputIn64 As Boolean) As String
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
                  "OASISRemoteAdmin.clsTwofish.EncryptString", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextIn64 As Boolean) As String
        '<EhHeader>
        On Error GoTo DecryptString_Err
        '</EhHeader>
100     If IsTextIn64 = True Then Text = Decode64(Text)
        Dim byteArray() As Byte
102     byteArray() = StrConv(Text, vbFromUnicode)
104     Call DecryptByte(byteArray(), Key)
106     DecryptString = StrConv(byteArray(), vbUnicode)
        '<EhFooter>
        Exit Function

DecryptString_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.DecryptString", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Static Function LFSR1(ByRef X As Long) As Long
        '<EhHeader>
        On Error GoTo LFSR1_Err
        '</EhHeader>
100     LFSR1 = lBSR(X, 1) Xor ((X And 1) * GF256_FDBK_2)
        '<EhFooter>
        Exit Function

LFSR1_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.LFSR1", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Static Function LFSR2(ByRef X As Long) As Long
        '<EhHeader>
        On Error GoTo LFSR2_Err
        '</EhHeader>
100     LFSR2 = lBSR(X, 2) Xor ((X And &H2) / &H2 * GF256_FDBK_2) Xor ((X And &H1) * GF256_FDBK_4)
        '<EhFooter>
        Exit Function

LFSR2_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.LFSR2", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Static Function RS_Rem(X As Long) As Long
        '<EhHeader>
        On Error GoTo RS_Rem_Err
        '</EhHeader>
        Dim b As Long, g2 As Long, g3 As Long
100     b = (lBSRU(X, 24) And &HFF)
102     g2 = ((lBSL(b, 1) Xor (b And &H80) / &H80 * &H14D) And &HFF)
104     g3 = (lBSRU(b, 1) Xor ((b And &H1) * lBSRU(&H14D, 1)) Xor g2)
106     RS_Rem = lBSL(X, 8) Xor lBSL(g3, 24) Xor lBSL(g2, 16) Xor lBSL(g3, 8) Xor b
        '<EhFooter>
        Exit Function

RS_Rem_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.RS_Rem", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Static Function F32(k64Cnt As Long, X As Long, k32() As Long) As Long
        '<EhHeader>
        On Error GoTo F32_Err
        '</EhHeader>
        Dim xb(0 To 3) As Byte, Key(0 To 3, 0 To 3) As Byte
100     Call CopyMem(xb(0), X, 4)
102     Call CopyMem(Key(0, 0), k32(0), 16)
104     If ((k64Cnt And 3) = 1) Then
106         F32 = MDS(0, p(0, xb(0)) Xor Key(0, 0)) Xor MDS(1, p(0, xb(1)) Xor Key(1, 0)) Xor MDS(2, p(1, xb(2)) Xor Key(2, 0)) Xor MDS(3, p(1, xb(3)) Xor Key(3, 0))
        Else
108         If ((k64Cnt And 3) = 0) Then
110               xb(0) = p(1, xb(0)) Xor Key(0, 3)
112               xb(1) = p(0, xb(1)) Xor Key(1, 3)
114               xb(2) = p(0, xb(2)) Xor Key(2, 3)
116               xb(3) = p(1, xb(3)) Xor Key(3, 3)
            End If
118         If ((k64Cnt And 3) = 3) Or ((k64Cnt And 3) = 0) Then
120               xb(0) = p(1, xb(0)) Xor Key(0, 2)
122               xb(1) = p(1, xb(1)) Xor Key(1, 2)
124               xb(2) = p(0, xb(2)) Xor Key(2, 2)
126               xb(3) = p(0, xb(3)) Xor Key(3, 2)
            End If
128         F32 = MDS(0, p(0, p(0, xb(0)) Xor Key(0, 1)) Xor Key(0, 0)) Xor MDS(1, p(0, p(1, xb(1)) Xor Key(1, 1)) Xor Key(1, 0)) Xor MDS(2, p(1, p(0, xb(2)) Xor Key(2, 1)) Xor Key(2, 0)) Xor MDS(3, p(1, p(1, xb(3)) Xor Key(3, 1)) Xor Key(3, 0))
        End If
        '<EhFooter>
        Exit Function

F32_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.F32", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Static Function Fe32(X As Long, r As Long) As Long
        '<EhHeader>
        On Error GoTo Fe32_Err
        '</EhHeader>
        Dim xb(0 To 3) As Byte
100     Call CopyMem(xb(0), X, 4)
102     Fe32 = sBox(2 * xb(r Mod 4)) Xor sBox(2 * xb((r + 1) Mod 4) + 1) Xor sBox(&H200 + 2 * xb((r + 2) Mod 4)) Xor sBox(&H200 + 2 * xb((r + 3) Mod 4) + 1)
        '<EhFooter>
        Exit Function

Fe32_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.Fe32", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Static Sub KeyCreate(K() As Byte, KeyLength As Long)
        '<EhHeader>
        On Error GoTo KeyCreate_Err
        '</EhHeader>
        Dim i As Long, lA As Long, lB As Long, b(3) As Byte, k64Cnt As Long, k32e(3) As Long, k32o(3) As Long, subkeyCnt As Long, sBoxKey(3) As Long, Key(0 To 3, 0 To 3) As Byte
100     k64Cnt = KeyLength \ 8
102     subkeyCnt = ROUND_SUBKEYS + 2 * Rounds
104     For i = 0 To IIf(KeyLength < 32, KeyLength \ 8 - 1, 3)
106         Call CopyMem(k32e(i), K(i * 8), 4)
108         Call CopyMem(k32o(i), K(i * 8 + 4), 4)
110         sBoxKey(KeyLength \ 8 - 1 - i) = RS_Rem(RS_Rem(RS_Rem(RS_Rem(RS_Rem(RS_Rem(RS_Rem(RS_Rem(k32o(i))))) Xor k32e(i)))))
        Next
112     ReDim sKey(subkeyCnt)
114     For i = 0 To ((subkeyCnt / 2) - 1)
116         lA = F32(k64Cnt, i * &H2020202, k32e)
118         lB = F32(k64Cnt, i * &H2020202 + &H1010101, k32o)
120         lB = lBSL(lB, 8) Or lBSRU(lB, 24)
122         If (m_RunningCompiled) Then lA = lA + lB Else lA = UnsignedAdd(lA, lB)
124         sKey(2 * i) = lA
126         If (m_RunningCompiled) Then lA = lA + lB Else lA = UnsignedAdd(lA, lB)
128         sKey(2 * i + 1) = lBSL(lA, 9) Or lBSRU(lA, 32 - 9)
        Next
130     Call CopyMem(Key(0, 0), sBoxKey(0), 16)
132     For i = 0 To 255
134         If ((k64Cnt And 3) = 1) Then
136             sBox(2 * i) = MDS(0, p(0, i) Xor Key(0, 0))
138             sBox(2 * i + 1) = MDS(1, p(0, i) Xor Key(1, 0))
140             sBox(&H200 + 2 * i) = MDS(2, p(1, i) Xor Key(2, 0))
142             sBox(&H200 + 2 * i + 1) = MDS(3, p(1, i) Xor Key(3, 0))
            Else
144             b(0) = i
146             b(1) = i
148             b(2) = i
150             b(3) = i
152             If ((k64Cnt And 3) = 0) Then
154                   b(0) = p(1, b(0)) Xor Key(0, 3)
156                   b(1) = p(0, b(1)) Xor Key(1, 3)
158                   b(2) = p(0, b(2)) Xor Key(2, 3)
160                   b(3) = p(1, b(3)) Xor Key(3, 3)
                End If
162             If ((k64Cnt And 3) = 3) Or ((k64Cnt And 3) = 0) Then
164                 b(0) = p(1, b(0)) Xor Key(0, 2)
166                 b(1) = p(1, b(1)) Xor Key(1, 2)
168                 b(2) = p(0, b(2)) Xor Key(2, 2)
170                 b(3) = p(0, b(3)) Xor Key(3, 2)
                End If
172             sBox(2 * i) = MDS(0, p(0, p(0, b(0)) Xor Key(0, 1)) Xor Key(0, 0))
174             sBox(2 * i + 1) = MDS(1, p(0, p(1, b(1)) Xor Key(1, 1)) Xor Key(1, 0))
176             sBox(&H200 + 2 * i) = MDS(2, p(1, p(0, b(2)) Xor Key(2, 1)) Xor Key(2, 0))
178             sBox(&H200 + 2 * i + 1) = MDS(3, p(1, p(1, b(3)) Xor Key(3, 1)) Xor Key(3, 0))
            End If
        Next
        '<EhFooter>
        Exit Sub

KeyCreate_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.KeyCreate", _
                  "clsTwofish component failure"
        '</EhFooter>
End Sub
Private Function lBSL(ByRef lInput As Long, ByRef bShiftBits As Byte) As Long
        '<EhHeader>
        On Error GoTo lBSL_Err
        '</EhHeader>
100     lBSL = (lInput And (2 ^ (31 - bShiftBits) - 1)) * 2 ^ bShiftBits
102     If (lInput And 2 ^ (31 - bShiftBits)) = 2 ^ (31 - bShiftBits) Then lBSL = (lBSL Or &H80000000)
        '<EhFooter>
        Exit Function

lBSL_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.lBSL", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Function lBSR(ByRef lInput As Long, ByRef bShiftBits As Byte) As Long
        '<EhHeader>
        On Error GoTo lBSR_Err
        '</EhHeader>
100     If (bShiftBits = 31) Then
102         If (lInput < 0) Then lBSR = &HFFFFFFFF Else lBSR = 0
        Else
104         lBSR = (lInput And Not (2 ^ bShiftBits - 1)) \ 2 ^ bShiftBits
        End If
        '<EhFooter>
        Exit Function

lBSR_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.lBSR", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Function lBSRU(lInput As Long, bShiftBits As Byte) As Long
        '<EhHeader>
        On Error GoTo lBSRU_Err
        '</EhHeader>
100     If (bShiftBits = 31) Then
102         lBSRU = -(lInput < 0)
        Else
104         lBSRU = (((lInput And Not (2 ^ bShiftBits - 1)) \ 2 ^ bShiftBits) And Not (&H80000000 + (2 ^ bShiftBits - 2) * 2 ^ (31 - bShiftBits)))
        End If
        '<EhFooter>
        Exit Function

lBSRU_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.lBSRU", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Static Sub EncryptBlock(DWord() As Long)
        '<EhHeader>
        On Error GoTo EncryptBlock_Err
        '</EhHeader>
        Dim t0 As Long, t1 As Long, K As Long, r As Long
100     DWord(0) = DWord(0) Xor sKey(INPUT_WHITEN)
102     DWord(1) = DWord(1) Xor sKey(INPUT_WHITEN + 1)
104     DWord(2) = DWord(2) Xor sKey(INPUT_WHITEN + 2)
106     DWord(3) = DWord(3) Xor sKey(INPUT_WHITEN + 3)
108     K = ROUND_SUBKEYS
110     For r = 0 To (Rounds - 1) Step 2
112         If (m_RunningCompiled) Then
114             t0 = Fe32(DWord(0), 0)
116             t1 = Fe32(DWord(1), 3)
118             t0 = t0 + t1
120             DWord(2) = Rot1(DWord(2) Xor (t0 + sKey(K)))
122             K = K + 1
124             DWord(3) = Rot31(DWord(3)) Xor (t0 + t1 + sKey(K))
126             K = K + 1
128             t0 = Fe32(DWord(2), 0)
130             t1 = Fe32(DWord(3), 3)
132             t0 = t0 + t1
134             DWord(0) = Rot1(DWord(0) Xor (t0 + sKey(K)))
136             K = K + 1
138             DWord(1) = Rot31(DWord(1)) Xor (t0 + t1 + sKey(K))
140             K = K + 1
            Else
142             t0 = Fe32(DWord(0), 0)
144             t1 = Fe32(DWord(1), 3)
146             t0 = UnsignedAdd(t0, t1)
148             DWord(2) = Rot1(DWord(2) Xor (UnsignedAdd(t0, sKey(K))))
150             K = K + 1
152             DWord(3) = Rot31(DWord(3)) Xor (UnsignedAdd(UnsignedAdd(t0, t1), sKey(K)))
154             K = K + 1
156             t0 = Fe32(DWord(2), 0)
158             t1 = Fe32(DWord(3), 3)
160             t0 = UnsignedAdd(t0, t1)
162             DWord(0) = Rot1(DWord(0) Xor (UnsignedAdd(t0, sKey(K))))
164             K = K + 1
166             DWord(1) = Rot31(DWord(1)) Xor (UnsignedAdd(UnsignedAdd(t0, t1), sKey(K)))
168             K = K + 1
            End If
        Next
170     DWord(2) = DWord(2) Xor sKey(OUTPUT_WHITEN)
172     DWord(3) = DWord(3) Xor sKey(OUTPUT_WHITEN + 1)
174     DWord(4) = DWord(0) Xor sKey(OUTPUT_WHITEN + 2)
176     DWord(5) = DWord(1) Xor sKey(OUTPUT_WHITEN + 3)
178     Call CopyMem(DWord(0), DWord(2), 16)
        '<EhFooter>
        Exit Sub

EncryptBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.EncryptBlock", _
                  "clsTwofish component failure"
        '</EhFooter>
End Sub
Private Sub DecryptBlock(DWord() As Long)
        '<EhHeader>
        On Error GoTo DecryptBlock_Err
        '</EhHeader>
        Dim K As Long, r As Long, t0 As Long, t1 As Long
100     DWord(2) = DWord(2) Xor sKey(OUTPUT_WHITEN)
102     DWord(3) = DWord(3) Xor sKey(OUTPUT_WHITEN + 1)
104     DWord(0) = DWord(4) Xor sKey(OUTPUT_WHITEN + 2)
106     DWord(1) = DWord(5) Xor sKey(OUTPUT_WHITEN + 3)
108     K = ROUND_SUBKEYS + 2 * Rounds - 1
110     For r = 0 To Rounds - 1 Step 2
112         If (m_RunningCompiled) Then
114             t0 = Fe32(DWord(2), 0)
116             t1 = Fe32(DWord(3), 3)
118             t0 = t0 + t1
120             DWord(1) = Rot1(DWord(1) Xor (t0 + t1 + sKey(K)))
122             K = K - 1
124             DWord(0) = Rot31(DWord(0)) Xor (t0 + sKey(K))
126             K = K - 1
128             t0 = Fe32(DWord(0), 0)
130             t1 = Fe32(DWord(1), 3)
132             t0 = t0 + t1
134             DWord(3) = Rot1(DWord(3) Xor (t0 + t1 + sKey(K)))
136             K = K - 1
138             DWord(2) = Rot31(DWord(2)) Xor (t0 + sKey(K))
140             K = K - 1
            Else
142             t0 = Fe32(DWord(2), 0)
144             t1 = Fe32(DWord(3), 3)
146             t0 = UnsignedAdd(t0, t1)
148             DWord(1) = Rot1(DWord(1) Xor (UnsignedAdd(UnsignedAdd(t0, t1), sKey(K))))
150             K = K - 1
152             DWord(0) = Rot31(DWord(0)) Xor (UnsignedAdd(t0, sKey(K)))
154             K = K - 1
156             t0 = Fe32(DWord(0), 0)
158             t1 = Fe32(DWord(1), 3)
160             t0 = UnsignedAdd(t0, t1)
162             DWord(3) = Rot1(DWord(3) Xor (UnsignedAdd(UnsignedAdd(t0, t1), sKey(K))))
164             K = K - 1
166             DWord(2) = Rot31(DWord(2)) Xor (UnsignedAdd(t0, sKey(K)))
168             K = K - 1
            End If
        Next
170     DWord(0) = DWord(0) Xor sKey(INPUT_WHITEN)
172     DWord(1) = DWord(1) Xor sKey(INPUT_WHITEN + 1)
174     DWord(2) = DWord(2) Xor sKey(INPUT_WHITEN + 2)
176     DWord(3) = DWord(3) Xor sKey(INPUT_WHITEN + 3)
        '<EhFooter>
        Exit Sub

DecryptBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.DecryptBlock", _
                  "clsTwofish component failure"
        '</EhFooter>
End Sub
Private Static Function Rot1(Value As Long) As Long
        '<EhHeader>
        On Error GoTo Rot1_Err
        '</EhHeader>
        Dim Temp As Byte, X(0 To 3) As Byte
100     Call CopyMem(X(0), Value, 4)
102     Temp = X(0)
104     X(0) = (X(0) \ 2) Or ((X(1) And 1) * 128)
106     X(1) = (X(1) \ 2) Or ((X(2) And 1) * 128)
108     X(2) = (X(2) \ 2) Or ((X(3) And 1) * 128)
110     X(3) = (X(3) \ 2) Or ((Temp And 1) * 128)
112     Call CopyMem(Rot1, X(0), 4)
        '<EhFooter>
        Exit Function

Rot1_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.Rot1", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Static Function Rot31(Value As Long) As Long
        '<EhHeader>
        On Error GoTo Rot31_Err
        '</EhHeader>
        Dim Temp As Byte, X(0 To 3) As Byte
100     Call CopyMem(X(0), Value, 4)
102     Temp = X(3)
104     X(3) = ((X(3) And 127) * 2) Or -CBool(X(2) And 128)
106     X(2) = ((X(2) And 127) * 2) Or -CBool(X(1) And 128)
108     X(1) = ((X(1) And 127) * 2) Or -CBool(X(0) And 128)
110     X(0) = ((X(0) And 127) * 2) Or -CBool(Temp And 128)
112     Call CopyMem(Rot31, X(0), 4)
        '<EhFooter>
        Exit Function

Rot31_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.Rot31", _
                  "clsTwofish component failure"
        '</EhFooter>
End Function
Private Sub Class_Initialize()
        '<EhHeader>
        On Error GoTo Class_Initialize_Err
        '</EhHeader>
    Dim i As Long, j As Long, m1(0 To 1) As Long, mX(0 To 1) As Long, mY(0 To 1) As Long
  
100 On Local Error Resume Next
102 m_RunningCompiled = ((2147483647 + 1) < 0)

104   p(0, 0) = &HA9
106   p(0, 1) = &H67
108   p(0, 2) = &HB3
110   p(0, 3) = &HE8
112   p(0, 4) = &H4
114   p(0, 5) = &HFD
116   p(0, 6) = &HA3
118   p(0, 7) = &H76
120   p(0, 8) = &H9A
122   p(0, 9) = &H92
124   p(0, 10) = &H80
126   p(0, 11) = &H78
128   p(0, 12) = &HE4
130   p(0, 13) = &HDD
132   p(0, 14) = &HD1
134   p(0, 15) = &H38
136   p(0, 16) = &HD
138   p(0, 17) = &HC6
140   p(0, 18) = &H35
142   p(0, 19) = &H98
144   p(0, 20) = &H18
146   p(0, 21) = &HF7
148   p(0, 22) = &HEC
150   p(0, 23) = &H6C
152   p(0, 24) = &H43
154   p(0, 25) = &H75
156   p(0, 26) = &H37
158   p(0, 27) = &H26
160   p(0, 28) = &HFA
162   p(0, 29) = &H13
164   p(0, 30) = &H94
166   p(0, 31) = &H48
168   p(0, 32) = &HF2
170   p(0, 33) = &HD0
172   p(0, 34) = &H8B
174   p(0, 35) = &H30
176   p(0, 36) = &H84
178   p(0, 37) = &H54
180   p(0, 38) = &HDF
182   p(0, 39) = &H23
184   p(0, 40) = &H19
186   p(0, 41) = &H5B
188   p(0, 42) = &H3D
190   p(0, 43) = &H59
192   p(0, 44) = &HF3
194   p(0, 45) = &HAE
196   p(0, 46) = &HA2
198   p(0, 47) = &H82
200   p(0, 48) = &H63
202   p(0, 49) = &H1
204   p(0, 50) = &H83
206   p(0, 51) = &H2E
208   p(0, 52) = &HD9
210   p(0, 53) = &H51
212   p(0, 54) = &H9B
214   p(0, 55) = &H7C
216   p(0, 56) = &HA6
218   p(0, 57) = &HEB
220   p(0, 58) = &HA5
222   p(0, 59) = &HBE
224   p(0, 60) = &H16
226   p(0, 61) = &HC
228   p(0, 62) = &HE3
230   p(0, 63) = &H61
232   p(0, 64) = &HC0
234   p(0, 65) = &H8C
236   p(0, 66) = &H3A
238   p(0, 67) = &HF5
240   p(0, 68) = &H73
242   p(0, 69) = &H2C
244   p(0, 70) = &H25
246   p(0, 71) = &HB
248   p(0, 72) = &HBB
250   p(0, 73) = &H4E
252   p(0, 74) = &H89
254   p(0, 75) = &H6B
256   p(0, 76) = &H53
258   p(0, 77) = &H6A
260   p(0, 78) = &HB4
262   p(0, 79) = &HF1
264   p(0, 80) = &HE1
266   p(0, 81) = &HE6
268   p(0, 82) = &HBD
270   p(0, 83) = &H45
272   p(0, 84) = &HE2
274   p(0, 85) = &HF4
276   p(0, 86) = &HB6
278   p(0, 87) = &H66
280   p(0, 88) = &HCC
282   p(0, 89) = &H95
284   p(0, 90) = &H3
286   p(0, 91) = &H56
288   p(0, 92) = &HD4
290   p(0, 93) = &H1C
292   p(0, 94) = &H1E
294   p(0, 95) = &HD7
296   p(0, 96) = &HFB
298   p(0, 97) = &HC3
300   p(0, 98) = &H8E
302   p(0, 99) = &HB5
304   p(0, 100) = &HE9
306   p(0, 101) = &HCF
308   p(0, 102) = &HBF
310   p(0, 103) = &HBA
312   p(0, 104) = &HEA
314   p(0, 105) = &H77
316   p(0, 106) = &H39
318   p(0, 107) = &HAF
320   p(0, 108) = &H33
322   p(0, 109) = &HC9
324   p(0, 110) = &H62
326   p(0, 111) = &H71
328   p(0, 112) = &H81
330   p(0, 113) = &H79
332   p(0, 114) = &H9
334   p(0, 115) = &HAD
336   p(0, 116) = &H24
338   p(0, 117) = &HCD
340   p(0, 118) = &HF9
342   p(0, 119) = &HD8
344   p(0, 120) = &HE5
346   p(0, 121) = &HC5
348   p(0, 122) = &HB9
350   p(0, 123) = &H4D
352   p(0, 124) = &H44
354   p(0, 125) = &H8
356   p(0, 126) = &H86
358   p(0, 127) = &HE7
360   p(0, 128) = &HA1
362   p(0, 129) = &H1D
364   p(0, 130) = &HAA
366   p(0, 131) = &HED
368   p(0, 132) = &H6
370   p(0, 133) = &H70
372   p(0, 134) = &HB2
374   p(0, 135) = &HD2
376   p(0, 136) = &H41
378   p(0, 137) = &H7B
380   p(0, 138) = &HA0
382   p(0, 139) = &H11
384   p(0, 140) = &H31
386   p(0, 141) = &HC2
388   p(0, 142) = &H27
390   p(0, 143) = &H90
392   p(0, 144) = &H20
394   p(0, 145) = &HF6
396   p(0, 146) = &H60
398   p(0, 147) = &HFF
400   p(0, 148) = &H96
402   p(0, 149) = &H5C
404   p(0, 150) = &HB1
406   p(0, 151) = &HAB
408   p(0, 152) = &H9E
410   p(0, 153) = &H9C
412   p(0, 154) = &H52
414   p(0, 155) = &H1B
416   p(0, 156) = &H5F
418   p(0, 157) = &H93
420   p(0, 158) = &HA
422   p(0, 159) = &HEF
424   p(0, 160) = &H91
426   p(0, 161) = &H85
428   p(0, 162) = &H49
430   p(0, 163) = &HEE
432   p(0, 164) = &H2D
434   p(0, 165) = &H4F
436   p(0, 166) = &H8F
438   p(0, 167) = &H3B
440   p(0, 168) = &H47
442   p(0, 169) = &H87
444   p(0, 170) = &H6D
446   p(0, 171) = &H46
448   p(0, 172) = &HD6
450   p(0, 173) = &H3E
452   p(0, 174) = &H69
454   p(0, 175) = &H64
456   p(0, 176) = &H2A
458   p(0, 177) = &HCE
460   p(0, 178) = &HCB
462   p(0, 179) = &H2F
464   p(0, 180) = &HFC
466   p(0, 181) = &H97
468   p(0, 182) = &H5
470   p(0, 183) = &H7A
472   p(0, 184) = &HAC
474   p(0, 185) = &H7F
476   p(0, 186) = &HD5
478   p(0, 187) = &H1A
480   p(0, 188) = &H4B
482   p(0, 189) = &HE
484   p(0, 190) = &HA7
486   p(0, 191) = &H5A
488   p(0, 192) = &H28
490   p(0, 193) = &H14
492   p(0, 194) = &H3F
494   p(0, 195) = &H29
496   p(0, 196) = &H88
498   p(0, 197) = &H3C
500   p(0, 198) = &H4C
502   p(0, 199) = &H2
504   p(0, 200) = &HB8
506   p(0, 201) = &HDA
508   p(0, 202) = &HB0
510   p(0, 203) = &H17
512   p(0, 204) = &H55
514   p(0, 205) = &H1F
516   p(0, 206) = &H8A
518   p(0, 207) = &H7D
520   p(0, 208) = &H57
522   p(0, 209) = &HC7
524   p(0, 210) = &H8D
526   p(0, 211) = &H74
528   p(0, 212) = &HB7
530   p(0, 213) = &HC4
532   p(0, 214) = &H9F
534   p(0, 215) = &H72
536   p(0, 216) = &H7E
538   p(0, 217) = &H15
540   p(0, 218) = &H22
542   p(0, 219) = &H12
544   p(0, 220) = &H58
546   p(0, 221) = &H7
548   p(0, 222) = &H99
550   p(0, 223) = &H34
552   p(0, 224) = &H6E
554   p(0, 225) = &H50
556   p(0, 226) = &HDE
558   p(0, 227) = &H68
560   p(0, 228) = &H65
562   p(0, 229) = &HBC
564   p(0, 230) = &HDB
566   p(0, 231) = &HF8
568   p(0, 232) = &HC8
570   p(0, 233) = &HA8
572   p(0, 234) = &H2B
574   p(0, 235) = &H40
576   p(0, 236) = &HDC
578   p(0, 237) = &HFE
580   p(0, 238) = &H32
582   p(0, 239) = &HA4
584   p(0, 240) = &HCA
586   p(0, 241) = &H10
588   p(0, 242) = &H21
590   p(0, 243) = &HF0
592   p(0, 244) = &HD3
594   p(0, 245) = &H5D
596   p(0, 246) = &HF
598   p(0, 247) = &H0
600   p(0, 248) = &H6F
602   p(0, 249) = &H9D
604   p(0, 250) = &H36
606   p(0, 251) = &H42
608   p(0, 252) = &H4A
610   p(0, 253) = &H5E
612   p(0, 254) = &HC1
614   p(0, 255) = &HE0
616   p(1, 0) = &H75
618   p(1, 1) = &HF3
620   p(1, 2) = &HC6
622   p(1, 3) = &HF4
624   p(1, 4) = &HDB
626   p(1, 5) = &H7B
628   p(1, 6) = &HFB
630   p(1, 7) = &HC8
632   p(1, 8) = &H4A
634   p(1, 9) = &HD3
636   p(1, 10) = &HE6
638   p(1, 11) = &H6B
640   p(1, 12) = &H45
642   p(1, 13) = &H7D
644   p(1, 14) = &HE8
646   p(1, 15) = &H4B
648   p(1, 16) = &HD6
650   p(1, 17) = &H32
652   p(1, 18) = &HD8
654   p(1, 19) = &HFD
656   p(1, 20) = &H37
658   p(1, 21) = &H71
660   p(1, 22) = &HF1
662   p(1, 23) = &HE1
664   p(1, 24) = &H30
666   p(1, 25) = &HF
668   p(1, 26) = &HF8
670   p(1, 27) = &H1B
672   p(1, 28) = &H87
674   p(1, 29) = &HFA
676   p(1, 30) = &H6
678   p(1, 31) = &H3F
680   p(1, 32) = &H5E
682   p(1, 33) = &HBA
684   p(1, 34) = &HAE
686   p(1, 35) = &H5B
688   p(1, 36) = &H8A
690   p(1, 37) = &H0
692   p(1, 38) = &HBC
694   p(1, 39) = &H9D
696   p(1, 40) = &H6D
698   p(1, 41) = &HC1
700   p(1, 42) = &HB1
702   p(1, 43) = &HE
704   p(1, 44) = &H80
706   p(1, 45) = &H5D
708   p(1, 46) = &HD2
710   p(1, 47) = &HD5
712   p(1, 48) = &HA0
714   p(1, 49) = &H84
716   p(1, 50) = &H7
718   p(1, 51) = &H14
720   p(1, 52) = &HB5
722   p(1, 53) = &H90
724   p(1, 54) = &H2C
726   p(1, 55) = &HA3
728   p(1, 56) = &HB2
730   p(1, 57) = &H73
732   p(1, 58) = &H4C
734   p(1, 59) = &H54
736   p(1, 60) = &H92
738   p(1, 61) = &H74
740   p(1, 62) = &H36
742   p(1, 63) = &H51
744   p(1, 64) = &H38
746   p(1, 65) = &HB0
748   p(1, 66) = &HBD
750   p(1, 67) = &H5A
752   p(1, 68) = &HFC
754   p(1, 69) = &H60
756   p(1, 70) = &H62
758   p(1, 71) = &H96
760   p(1, 72) = &H6C
762   p(1, 73) = &H42
764   p(1, 74) = &HF7
766   p(1, 75) = &H10
768   p(1, 76) = &H7C
770   p(1, 77) = &H28
772   p(1, 78) = &H27
774   p(1, 79) = &H8C
776   p(1, 80) = &H13
778   p(1, 81) = &H95
780   p(1, 82) = &H9C
782   p(1, 83) = &HC7
784   p(1, 84) = &H24
786   p(1, 85) = &H46
788   p(1, 86) = &H3B
790   p(1, 87) = &H70
792   p(1, 88) = &HCA
794   p(1, 89) = &HE3
796   p(1, 90) = &H85
798   p(1, 91) = &HCB
800   p(1, 92) = &H11
802   p(1, 93) = &HD0
804   p(1, 94) = &H93
806   p(1, 95) = &HB8
808   p(1, 96) = &HA6
810   p(1, 97) = &H83
812   p(1, 98) = &H20
814   p(1, 99) = &HFF
816   p(1, 100) = &H9F
818   p(1, 101) = &H77
820   p(1, 102) = &HC3
822   p(1, 103) = &HCC
824   p(1, 104) = &H3
826   p(1, 105) = &H6F
828   p(1, 106) = &H8
830   p(1, 107) = &HBF
832   p(1, 108) = &H40
834   p(1, 109) = &HE7
836   p(1, 110) = &H2B
838   p(1, 111) = &HE2
840   p(1, 112) = &H79
842   p(1, 113) = &HC
844   p(1, 114) = &HAA
846   p(1, 115) = &H82
848   p(1, 116) = &H41
850   p(1, 117) = &H3A
852   p(1, 118) = &HEA
854   p(1, 119) = &HB9
856   p(1, 120) = &HE4
858   p(1, 121) = &H9A
860   p(1, 122) = &HA4
862   p(1, 123) = &H97
864   p(1, 124) = &H7E
866   p(1, 125) = &HDA
868   p(1, 126) = &H7A
870   p(1, 127) = &H17
872   p(1, 128) = &H66
874   p(1, 129) = &H94
876   p(1, 130) = &HA1
878   p(1, 131) = &H1D
880   p(1, 132) = &H3D
882   p(1, 133) = &HF0
884   p(1, 134) = &HDE
886   p(1, 135) = &HB3
888   p(1, 136) = &HB
890   p(1, 137) = &H72
892   p(1, 138) = &HA7
894   p(1, 139) = &H1C
896   p(1, 140) = &HEF
898   p(1, 141) = &HD1
900   p(1, 142) = &H53
902   p(1, 143) = &H3E
904   p(1, 144) = &H8F
906   p(1, 145) = &H33
908   p(1, 146) = &H26
910   p(1, 147) = &H5F
912   p(1, 148) = &HEC
914   p(1, 149) = &H76
916   p(1, 150) = &H2A
918   p(1, 151) = &H49
920   p(1, 152) = &H81
922   p(1, 153) = &H88
924   p(1, 154) = &HEE
926   p(1, 155) = &H21
928   p(1, 156) = &HC4
930   p(1, 157) = &H1A
932   p(1, 158) = &HEB
934   p(1, 159) = &HD9
936   p(1, 160) = &HC5
938   p(1, 161) = &H39
940   p(1, 162) = &H99
942   p(1, 163) = &HCD
944   p(1, 164) = &HAD
946   p(1, 165) = &H31
948   p(1, 166) = &H8B
950   p(1, 167) = &H1
952   p(1, 168) = &H18
954   p(1, 169) = &H23
956   p(1, 170) = &HDD
958   p(1, 171) = &H1F
960   p(1, 172) = &H4E
962   p(1, 173) = &H2D
964   p(1, 174) = &HF9
966   p(1, 175) = &H48
968   p(1, 176) = &H4F
970   p(1, 177) = &HF2
972   p(1, 178) = &H65
974   p(1, 179) = &H8E
976   p(1, 180) = &H78
978   p(1, 181) = &H5C
980   p(1, 182) = &H58
982   p(1, 183) = &H19
984   p(1, 184) = &H8D
986   p(1, 185) = &HE5
988   p(1, 186) = &H98
990   p(1, 187) = &H57
992   p(1, 188) = &H67
994   p(1, 189) = &H7F
996   p(1, 190) = &H5
998   p(1, 191) = &H64
1000   p(1, 192) = &HAF
1002   p(1, 193) = &H63
1004   p(1, 194) = &HB6
1006   p(1, 195) = &HFE
1008   p(1, 196) = &HF5
1010   p(1, 197) = &HB7
1012   p(1, 198) = &H3C
1014   p(1, 199) = &HA5
1016   p(1, 200) = &HCE
1018   p(1, 201) = &HE9
1020   p(1, 202) = &H68
1022   p(1, 203) = &H44
1024   p(1, 204) = &HE0
1026   p(1, 205) = &H4D
1028   p(1, 206) = &H43
1030   p(1, 207) = &H69
1032   p(1, 208) = &H29
1034   p(1, 209) = &H2E
1036   p(1, 210) = &HAC
1038   p(1, 211) = &H15
1040   p(1, 212) = &H59
1042   p(1, 213) = &HA8
1044   p(1, 214) = &HA
1046   p(1, 215) = &H9E
1048   p(1, 216) = &H6E
1050   p(1, 217) = &H47
1052   p(1, 218) = &HDF
1054   p(1, 219) = &H34
1056   p(1, 220) = &H35
1058   p(1, 221) = &H6A
1060   p(1, 222) = &HCF
1062   p(1, 223) = &HDC
1064   p(1, 224) = &H22
1066   p(1, 225) = &HC9
1068   p(1, 226) = &HC0
1070   p(1, 227) = &H9B
1072   p(1, 228) = &H89
1074   p(1, 229) = &HD4
1076   p(1, 230) = &HED
1078   p(1, 231) = &HAB
1080   p(1, 232) = &H12
1082   p(1, 233) = &HA2
1084   p(1, 234) = &HD
1086   p(1, 235) = &H52
1088   p(1, 236) = &HBB
1090   p(1, 237) = &H2
1092   p(1, 238) = &H2F
1094   p(1, 239) = &HA9
1096   p(1, 240) = &HD7
1098   p(1, 241) = &H61
1100   p(1, 242) = &H1E
1102   p(1, 243) = &HB4
1104   p(1, 244) = &H50
1106   p(1, 245) = &H4
1108   p(1, 246) = &HF6
1110   p(1, 247) = &HC2
1112   p(1, 248) = &H16
1114   p(1, 249) = &H25
1116   p(1, 250) = &H86
1118   p(1, 251) = &H56
1120   p(1, 252) = &H55
1122   p(1, 253) = &H9
1124   p(1, 254) = &HBE
1126   p(1, 255) = &H91

1128   For i = 0 To 255
1130         j = p(0, i)
1132         m1(0) = j
1134         mX(0) = j Xor LFSR2(j)
1136         mY(0) = j Xor LFSR1(j) Xor LFSR2(j)
1138         j = p(1, i)
1140         m1(1) = j
1142         mX(1) = j Xor LFSR2(j)
1144         mY(1) = j Xor LFSR1(j) Xor LFSR2(j)
1146         MDS(0, i) = (m1(1) Or lBSL(mX(1), 8) Or lBSL(mY(1), 16) Or lBSL(mY(1), 24))
1148         MDS(1, i) = (mY(0) Or lBSL(mY(0), 8) Or lBSL(mX(0), 16) Or lBSL(m1(0), 24))
1150         MDS(2, i) = (mX(1) Or lBSL(mY(1), 8) Or lBSL(m1(1), 16) Or lBSL(mY(1), 24))
1152         MDS(3, i) = (mX(0) Or lBSL(m1(0), 8) Or lBSL(mY(0), 16) Or lBSL(mX(0), 24))
       Next
         '<EhFooter>
         Exit Sub

Class_Initialize_Err:
         Err.Raise vbObjectError + 100, _
                   "OASISRemoteAdmin.clsTwofish.Class_Initialize", _
                   "clsTwofish component failure"
         '</EhFooter>
End Sub
Public Property Let Key(Optional ByVal MinKeyLength As TWOFISHKEYLENGTH, New_Value As String)
        '<EhHeader>
        On Error GoTo Key_Err
        '</EhHeader>
100     Class_Initialize
        Dim KeyLength As Long, Key() As Byte
102     KeyLength = Len(New_Value) * 8
104     Key() = StrConv(New_Value, vbFromUnicode)
106     If (KeyLength < MinKeyLength) Then
108         ReDim Preserve Key(MinKeyLength \ 8 - 1)
110         KeyLength = MinKeyLength
        End If
112     If (KeyLength > 192) Then
114         ReDim Preserve Key(31)
116         KeyLength = 256
118     ElseIf (KeyLength > 128) Then
120         ReDim Preserve Key(23)
122         KeyLength = 192
124     ElseIf (KeyLength > 64) Then
126         ReDim Preserve Key(15)
128         KeyLength = 128
130     ElseIf (KeyLength > 32) Then
132         ReDim Preserve Key(7)
134         KeyLength = 64
        Else
136         ReDim Preserve Key(3)
138         KeyLength = 32
        End If
140     Call KeyCreate(Key, KeyLength \ 8)
        '<EhFooter>
        Exit Property

Key_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.Key", _
                  "clsTwofish component failure"
        '</EhFooter>
End Property
Public Sub EncryptByte(byteArray() As Byte, Optional Key As String)
        '<EhHeader>
        On Error GoTo EncryptByte_Err
        '</EhHeader>
        Dim Offset As Long, OrigLen As Long, CipherLen As Long, CurrPercent As Long, NextPercent As Long, DWord(0 To 5) As Long, CipherWord(0 To 3) As Long
100     If (Len(Key) > 0) Then Me.Key = Key
102     OrigLen = UBound(byteArray) + 1
104     CipherLen = OrigLen + 12
106     If (CipherLen Mod 16 <> 0) Then CipherLen = CipherLen + 16 - (CipherLen Mod 16)
108     ReDim Preserve byteArray(CipherLen - 1)
110     Call CopyMem(byteArray(12), byteArray(0), OrigLen)
112     Call CopyMem(byteArray(8), OrigLen, 4)
114     Call Randomize
116     Call CopyMem(byteArray(0), CLng(2147483647 * Rnd), 4)
118     Call CopyMem(byteArray(4), CLng(2147483647 * Rnd), 4)
120     For Offset = 0 To (CipherLen - 1) Step 16
122           Call CopyMem(DWord(0), byteArray(Offset), 16)
124           DWord(0) = DWord(0) Xor CipherWord(0)
126           DWord(1) = DWord(1) Xor CipherWord(1)
128           DWord(2) = DWord(2) Xor CipherWord(2)
130           DWord(3) = DWord(3) Xor CipherWord(3)
132           Call EncryptBlock(DWord())
134           Call CopyMem(byteArray(Offset), DWord(0), 16)
136           Call CopyMem(CipherWord(0), DWord(0), 16)
138           If (Offset >= NextPercent) Then
140               CurrPercent = Int((Offset / CipherLen) * 100)
142               NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
144               RaiseEvent Progress(CurrPercent)
              End If
        Next
146     If (CurrPercent <> 100) Then RaiseEvent Progress(100)
        '<EhFooter>
        Exit Sub

EncryptByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.EncryptByte", _
                  "clsTwofish component failure"
        '</EhFooter>
End Sub
Public Sub DecryptByte(byteArray() As Byte, Optional Key As String)
        '<EhHeader>
        On Error GoTo DecryptByte_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
        Dim Offset As Long, OrigLen As Long, CipherLen As Long, CurrPercent As Long, NextPercent As Long, DWord(0 To 5) As Long, CipherWord(0 To 3) As Long
100     If (Len(Key) > 0) Then Me.Key = Key
102     CipherLen = UBound(byteArray) + 1
104     For Offset = 0 To (CipherLen - 1) Step 16
106         Call CopyMem(DWord(2), byteArray(Offset), 16)
108         Call DecryptBlock(DWord())
110         DWord(0) = DWord(0) Xor CipherWord(0)
112         DWord(1) = DWord(1) Xor CipherWord(1)
114         DWord(2) = DWord(2) Xor CipherWord(2)
116         DWord(3) = DWord(3) Xor CipherWord(3)
118         Call CopyMem(CipherWord(0), byteArray(Offset), 16)
120         Call CopyMem(byteArray(Offset), DWord(0), 16)
122         If (Offset >= NextPercent) Then
124             CurrPercent = Int((Offset / CipherLen) * 100)
126             NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
128             RaiseEvent Progress(CurrPercent)
            End If
        Next
130     Call CopyMem(OrigLen, byteArray(8), 4)
132     If (CipherLen - OrigLen > 27) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Twofish decryption")
134     Call CopyMem(byteArray(0), byteArray(12), OrigLen)
136     ReDim Preserve byteArray(OrigLen - 1)
138     If (CurrPercent <> 100) Then RaiseEvent Progress(100)

ErrorHandler:
        '<EhFooter>
        Exit Sub

DecryptByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTwofish.DecryptByte", _
                  "clsTwofish component failure"
        '</EhFooter>
End Sub
