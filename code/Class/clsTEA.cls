VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTEA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Standard TEA implementation with file support, Base64 conversion,
' and overall optimisations for Visual Basic. TEA is an excellent
' new cipher which appears to be farely strong but is still being
' reviewed and could show weaknesses in the near future. I would
' recommend TEA for use in medium risk solutions.
'
' Information on the algorithm can be found at:
' http://www.cl.cam.ac.uk/Research/Papers/djw-rmn/djw-rmn-tea.html

Private Declare Sub CopyMem _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Event Progress(Percent As Long)

Private m_RunningCompiled As Boolean
Private Tk(3) As Long
Private Const Rounds = 32
Private Const Delta = &H9E3779B9
Private Const DecryptSum = &HC6EF3720
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
                  "OASISRemoteAdmin.clsTEA.Initialize64", _
                  "clsTEA component failure"
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
                  "OASISRemoteAdmin.clsTEA.Decode64", _
                  "clsTEA component failure"
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
                  "OASISRemoteAdmin.clsTEA.DecodeArray64", _
                  "clsTEA component failure"
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
                  "OASISRemoteAdmin.clsTEA.Encode64", _
                  "clsTEA component failure"
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
                  "OASISRemoteAdmin.clsTEA.EncodeArray64", _
                  "clsTEA component failure"
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
                  "OASISRemoteAdmin.clsTEA.FileExist", _
                  "clsTEA component failure"
        '</EhFooter>
End Function

Private Static Function UnsignedAdd(ByVal Data1 As Long, _
                                    Data2 As Long) As Long
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
                  "OASISRemoteAdmin.clsTEA.UnsignedAdd", _
                  "clsTEA component failure"
        '</EhFooter>
End Function

Private Function UnsignedDel(Data1 As Long, _
                             Data2 As Long) As Long
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
                  "OASISRemoteAdmin.clsTEA.UnsignedDel", _
                  "clsTEA component failure"
        '</EhFooter>
End Function

Private Static Sub PutWord(LongValue As Long, _
                           CryptBuffer() As Byte, _
                           Offset As Long)
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
                  "OASISRemoteAdmin.clsTEA.PutWord", _
                  "clsTEA component failure"
        '</EhFooter>
End Sub
Private Static Sub GetWord(LongValue As Long, _
                           CryptBuffer() As Byte, _
                           Offset As Long)
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
                  "OASISRemoteAdmin.clsTEA.GetWord", _
                  "clsTEA component failure"
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
                  "OASISRemoteAdmin.clsTEA.EncryptFile", _
                  "clsTEA component failure"
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
                  "OASISRemoteAdmin.clsTEA.DecryptFile", _
                  "clsTEA component failure"
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
                  "OASISRemoteAdmin.clsTEA.EncryptString", _
                  "clsTEA component failure"
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
                  "OASISRemoteAdmin.clsTEA.DecryptString", _
                  "clsTEA component failure"
        '</EhFooter>
End Function

Public Sub EncryptByte(byteArray() As Byte, _
                       Optional Key As String)
        '<EhHeader>
        On Error GoTo EncryptByte_Err
        '</EhHeader>
        Dim X As Long, Sum As Long, Offset As Long, OrigLen As Long, LeftWord As Long, RightWord As Long, CipherLen As Long, CipherLeft As Long, CipherRight As Long, CurrPercent As Long, NextPercent As Long, Sl As Long, Sr As Long
    
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
130         Sum = 0

132         For X = 1 To Rounds

134             If (m_RunningCompiled) Then
136                 Sum = (Sum + Delta)
138                 Sr = ((RightWord And &HFFFFFFE0) \ 32) And &H7FFFFFF
140                 LeftWord = LeftWord + (((RightWord * 16) + Tk(0)) Xor (RightWord + Sum) Xor (Sr + Tk(1)))
142                 Sl = ((LeftWord And &HFFFFFFE0) \ 32) And &H7FFFFFF
144                 RightWord = RightWord + (((LeftWord * 16) + Tk(2)) Xor (LeftWord + Sum) Xor (Sl + Tk(3)))
                Else
146                 Sum = UnsignedAdd(Sum, Delta)
148                 LeftWord = UnsignedAdd(LeftWord, (UnsignedAdd(LShift4(RightWord), Tk(0)) Xor UnsignedAdd(RightWord, Sum) Xor UnsignedAdd(RShift5(RightWord), Tk(1))))
150                 RightWord = UnsignedAdd(RightWord, (UnsignedAdd(LShift4(LeftWord), Tk(2)) Xor UnsignedAdd(LeftWord, Sum) Xor UnsignedAdd(RShift5(LeftWord), Tk(3))))
                End If

            Next

152         Call PutWord(LeftWord, byteArray(), Offset)
154         Call PutWord(RightWord, byteArray(), Offset + 4)
156         CipherLeft = LeftWord
158         CipherRight = RightWord

160         If (Offset >= NextPercent) Then
162             CurrPercent = Int((Offset / CipherLen) * 100)
164             NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1

166             RaiseEvent Progress(CurrPercent)
            End If

        Next

168     If (CurrPercent <> 100) Then RaiseEvent Progress(100)
        '<EhFooter>
        Exit Sub

EncryptByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTEA.EncryptByte", _
                  "clsTEA component failure"
        '</EhFooter>
End Sub
Private Static Function LShift4(Data1 As Long) As Long
        '<EhHeader>
        On Error GoTo LShift4_Err
        '</EhHeader>
        Dim x1(0 To 3) As Byte, xx(0 To 3) As Byte
100     Call CopyMem(x1(0), Data1, 4)
102     xx(0) = ((x1(0) And 15) * 16)
104     xx(1) = ((x1(1) And 15) * 16) Or ((x1(0) And 240) \ 16)
106     xx(2) = ((x1(2) And 15) * 16) Or ((x1(1) And 240) \ 16)
108     xx(3) = ((x1(3) And 15) * 16) Or ((x1(2) And 240) \ 16)
110     Call CopyMem(LShift4, xx(0), 4)
        '<EhFooter>
        Exit Function

LShift4_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTEA.LShift4", _
                  "clsTEA component failure"
        '</EhFooter>
End Function
Private Static Function RShift5(Data1 As Long) As Long
        '<EhHeader>
        On Error GoTo RShift5_Err
        '</EhHeader>
        Dim x1(0 To 3) As Byte, xx(0 To 3) As Byte
100     Call CopyMem(x1(0), Data1, 4)
102     xx(0) = ((x1(0) And 224) \ 32) Or ((x1(1) And 31) * 8)
104     xx(1) = ((x1(1) And 224) \ 32) Or ((x1(2) And 31) * 8)
106     xx(2) = ((x1(2) And 224) \ 32) Or ((x1(3) And 31) * 8)
108     xx(3) = ((x1(3) And 224) \ 32)
110     Call CopyMem(RShift5, xx(0), 4)
        '<EhFooter>
        Exit Function

RShift5_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTEA.RShift5", _
                  "clsTEA component failure"
        '</EhFooter>
End Function

Public Sub DecryptByte(byteArray() As Byte, _
                       Optional Key As String)
        '<EhHeader>
        On Error GoTo DecryptByte_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
        Dim X As Long, Sum As Long, Offset As Long, OrigLen As Long, LeftWord As Long, RightWord As Long, CipherLen As Long, CipherLeft As Long, CipherRight As Long, CurrPercent As Long, NextPercent As Long, Sr As Long, Sl As Long
    
100     If (Len(Key) > 0) Then Me.Key = Key
102     CipherLen = UBound(byteArray) + 1

104     For Offset = 0 To (CipherLen - 1) Step 8
106         Call GetWord(LeftWord, byteArray(), Offset)
108         Call GetWord(RightWord, byteArray(), Offset + 4)
110         Sum = DecryptSum

112         For X = 1 To Rounds

114             If (m_RunningCompiled) Then
116                 Sl = ((LeftWord And &HFFFFFFE0) \ 32) And &H7FFFFFF
118                 RightWord = RightWord - (((LeftWord * 16) + Tk(2)) Xor (LeftWord + Sum) Xor (Sl + Tk(3)))
120                 Sr = ((RightWord And &HFFFFFFE0) \ 32) And &H7FFFFFF
122                 LeftWord = LeftWord - (((RightWord * 16) + Tk(0)) Xor (RightWord + Sum) Xor (Sr + Tk(1)))
124                 Sum = (Sum - Delta)
                Else
126                 RightWord = UnsignedDel(RightWord, (UnsignedAdd(LShift4(LeftWord), Tk(2)) Xor UnsignedAdd(LeftWord, Sum) Xor UnsignedAdd(RShift5(LeftWord), Tk(3))))
128                 LeftWord = UnsignedDel(LeftWord, (UnsignedAdd(LShift4(RightWord), Tk(0)) Xor UnsignedAdd(RightWord, Sum) Xor UnsignedAdd(RShift5(RightWord), Tk(1))))
130                 Sum = UnsignedDel(Sum, Delta)
                End If

            Next

132         LeftWord = LeftWord Xor CipherLeft
134         RightWord = RightWord Xor CipherRight
136         Call GetWord(CipherLeft, byteArray(), Offset)
138         Call GetWord(CipherRight, byteArray(), Offset + 4)
140         Call PutWord(LeftWord, byteArray(), Offset)
142         Call PutWord(RightWord, byteArray(), Offset + 4)

144         If (Offset >= NextPercent) Then
146             CurrPercent = Int((Offset / CipherLen) * 100)
148             NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1

150             RaiseEvent Progress(CurrPercent)
            End If

        Next

152     Call CopyMem(OrigLen, byteArray(8), 4)

154     If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in TEA decryption")
156     Call CopyMem(byteArray(0), byteArray(12), OrigLen)
158     ReDim Preserve byteArray(OrigLen - 1)

160     If (CurrPercent <> 100) Then RaiseEvent Progress(100)
    
ErrorHandler:
        '<EhFooter>
        Exit Sub

DecryptByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTEA.DecryptByte", _
                  "clsTEA component failure"
        '</EhFooter>
End Sub

Public Property Let Key(New_Value As String)
        '<EhHeader>
        On Error GoTo Key_Err
        '</EhHeader>
100     Class_Initialize
        Dim K() As Byte, w(0 To 3) As Byte
102     K() = StrConv(New_Value, vbFromUnicode)

104     If (Len(New_Value) < 16) Then ReDim Preserve K(15)
106     w(0) = K(3)
108     w(1) = K(2)
110     w(2) = K(1)
112     w(3) = K(0)
114     Call CopyMem(Tk(0), w(0), 4)
116     w(0) = K(7)
118     w(1) = K(6)
120     w(2) = K(5)
122     w(3) = K(4)
124     Call CopyMem(Tk(1), w(0), 4)
126     w(0) = K(11)
128     w(1) = K(10)
130     w(2) = K(9)
132     w(3) = K(8)
134     Call CopyMem(Tk(2), w(0), 4)
136     w(0) = K(15)
138     w(1) = K(14)
140     w(2) = K(13)
142     w(3) = K(12)
144     Call CopyMem(Tk(3), w(0), 4)
        '<EhFooter>
        Exit Property

Key_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTEA.Key", _
                  "clsTEA component failure"
        '</EhFooter>
End Property

Private Sub Class_Initialize()
        '<EhHeader>
        On Error GoTo Class_Initialize_Err
        '</EhHeader>
100     On Local Error Resume Next
102     m_RunningCompiled = ((2147483647 + 1) < 0)
        '<EhFooter>
        Exit Sub

Class_Initialize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsTEA.Class_Initialize", _
                  "clsTEA component failure"
        '</EhFooter>
End Sub
