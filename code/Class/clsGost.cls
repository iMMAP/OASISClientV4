VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Visual Basic GOST Implementation
' Standard Gost implementation with file support, Base64 conversion,
' and overall optimisations for Visual Basic. Gost supports 256-bit keys
' and is a stable algorithm developed by the former Soviet Union but there
' are many notable weaknesses and it remains in question if it really is
' much stronger than DES. Recommended for medium risk level solutions.
'
' Information on the algorithm can be found at:
' http://www.jetico.sci.fi/index.htm#/gost.htm

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub addLongs Lib "myDLL" (ByRef number1 As Long, ByVal number2 As Long)

Event Progress(Percent As Long)

Private m_KeyValue As String
Private K(1 To 8) As Long
Private k87(0 To 255) As Byte
Private k65(0 To 255) As Byte
Private k43(0 To 255) As Byte
Private k21(0 To 255) As Byte
Private sBox(0 To 7, 0 To 255) As Byte
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
                  "OASISRemoteAdmin.clsGost.Initialize64", _
                  "clsGost component failure"
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
                  "OASISRemoteAdmin.clsGost.Decode64", _
                  "clsGost component failure"
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
                  "OASISRemoteAdmin.clsGost.DecodeArray64", _
                  "clsGost component failure"
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
                  "OASISRemoteAdmin.clsGost.Encode64", _
                  "clsGost component failure"
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
                  "OASISRemoteAdmin.clsGost.EncodeArray64", _
                  "clsGost component failure"
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
                  "OASISRemoteAdmin.clsGost.FileExist", _
                  "clsGost component failure"
        '</EhFooter>
End Function
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
                  "OASISRemoteAdmin.clsGost.UnsignedAdd", _
                  "clsGost component failure"
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
                  "OASISRemoteAdmin.clsGost.UnsignedDel", _
                  "clsGost component failure"
        '</EhFooter>
End Function

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
                  "OASISRemoteAdmin.clsGost.PutWord", _
                  "clsGost component failure"
        '</EhFooter>
End Sub
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
                  "OASISRemoteAdmin.clsGost.GetWord", _
                  "clsGost component failure"
        '</EhFooter>
End Sub
Private Static Sub DecryptBlock(LeftWord As Long, RightWord As Long)
        '<EhHeader>
        On Error GoTo DecryptBlock_Err
        '</EhHeader>
        Dim i As Long
100     RightWord = RightWord Xor f(LeftWord, K(1))
102     LeftWord = LeftWord Xor f(RightWord, K(2))
104     RightWord = RightWord Xor f(LeftWord, K(3))
106     LeftWord = LeftWord Xor f(RightWord, K(4))
108     RightWord = RightWord Xor f(LeftWord, K(5))
110     LeftWord = LeftWord Xor f(RightWord, K(6))
112     RightWord = RightWord Xor f(LeftWord, K(7))
114     LeftWord = LeftWord Xor f(RightWord, K(8))
116     For i = 1 To 3
118         RightWord = RightWord Xor f(LeftWord, K(8))
120         LeftWord = LeftWord Xor f(RightWord, K(7))
122         RightWord = RightWord Xor f(LeftWord, K(6))
124         LeftWord = LeftWord Xor f(RightWord, K(5))
126         RightWord = RightWord Xor f(LeftWord, K(4))
128         LeftWord = LeftWord Xor f(RightWord, K(3))
130         RightWord = RightWord Xor f(LeftWord, K(2))
132         LeftWord = LeftWord Xor f(RightWord, K(1))
        Next
        '<EhFooter>
        Exit Sub

DecryptBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsGost.DecryptBlock", _
                  "clsGost component failure"
        '</EhFooter>
End Sub
Private Static Sub EncryptBlock(LeftWord As Long, RightWord As Long)
        '<EhHeader>
        On Error GoTo EncryptBlock_Err
        '</EhHeader>
        Dim i As Long
100     For i = 1 To 3
102         RightWord = RightWord Xor f(LeftWord, K(1))
104         LeftWord = LeftWord Xor f(RightWord, K(2))
106         RightWord = RightWord Xor f(LeftWord, K(3))
108         LeftWord = LeftWord Xor f(RightWord, K(4))
110         RightWord = RightWord Xor f(LeftWord, K(5))
112         LeftWord = LeftWord Xor f(RightWord, K(6))
114         RightWord = RightWord Xor f(LeftWord, K(7))
116         LeftWord = LeftWord Xor f(RightWord, K(8))
        Next
118     RightWord = RightWord Xor f(LeftWord, K(8))
120     LeftWord = LeftWord Xor f(RightWord, K(7))
122     RightWord = RightWord Xor f(LeftWord, K(6))
124     LeftWord = LeftWord Xor f(RightWord, K(5))
126     RightWord = RightWord Xor f(LeftWord, K(4))
128     LeftWord = LeftWord Xor f(RightWord, K(3))
130     RightWord = RightWord Xor f(LeftWord, K(2))
132     LeftWord = LeftWord Xor f(RightWord, K(1))
        '<EhFooter>
        Exit Sub

EncryptBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsGost.EncryptBlock", _
                  "clsGost component failure"
        '</EhFooter>
End Sub
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
                  "OASISRemoteAdmin.clsGost.EncryptFile", _
                  "clsGost component failure"
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
                  "OASISRemoteAdmin.clsGost.DecryptFile", _
                  "clsGost component failure"
        '</EhFooter>
End Function
Private Static Function f(r As Long, K As Long) As Long
        '<EhHeader>
        On Error GoTo f_Err
        '</EhHeader>
        Dim X As Long, xb(0 To 3) As Byte, xx(0 To 3) As Byte, a As Byte, b As Byte, c As Byte, d As Byte
100     If (m_RunningCompiled) Then X = r + K Else X = UnsignedAdd(r, K)
102     d = X And &HFF
104     X = X \ 256
106     c = X And &HFF
108     X = X \ 256
110     b = X And &HFF
112     X = X \ 256
114     a = X And &HFF
116     xb(0) = k21(a)
118     xb(1) = k43(b)
120     xb(2) = k65(c)
122     xb(3) = k87(d)
124     xx(0) = ((xb(3) And 31) * 8) Or ((xb(2) And 224) \ 32)
126     xx(1) = ((xb(0) And 31) * 8) Or ((xb(3) And 224) \ 32)
128     xx(2) = ((xb(1) And 31) * 8) Or ((xb(0) And 224) \ 32)
130     xx(3) = ((xb(2) And 31) * 8) Or ((xb(1) And 224) \ 32)
132     Call CopyMem(f, xx(0), 4)
        '<EhFooter>
        Exit Function

f_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsGost.f", _
                  "clsGost component failure"
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
        '<EhFooter>
        Exit Function

DecryptString_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsGost.DecryptString", _
                  "clsGost component failure"
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
                  "OASISRemoteAdmin.clsGost.EncryptString", _
                  "clsGost component failure"
        '</EhFooter>
End Function
Private Static Function lBSL(ByVal lInput As Long, bShiftBits As Byte) As Long
        '<EhHeader>
        On Error GoTo lBSL_Err
        '</EhHeader>
100     lBSL = (lInput And (2 ^ (31 - bShiftBits) - 1)) * 2 ^ bShiftBits
102     If (lInput And 2 ^ (31 - bShiftBits)) = 2 ^ (31 - bShiftBits) Then lBSL = (lBSL Or &H80000000)
        '<EhFooter>
        Exit Function

lBSL_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsGost.lBSL", _
                  "clsGost component failure"
        '</EhFooter>
End Function
Private Static Function lBSR(ByVal lInput As Long, bShiftBits As Byte) As Long
        '<EhHeader>
        On Error GoTo lBSR_Err
        '</EhHeader>
100     If bShiftBits = 31 Then
102       If lInput < 0 Then lBSR = &HFFFFFFFF Else lBSR = 0
        Else
104       lBSR = (lInput And Not (2 ^ bShiftBits - 1)) \ 2 ^ bShiftBits
        End If
        '<EhFooter>
        Exit Function

lBSR_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsGost.lBSR", _
                  "clsGost component failure"
        '</EhFooter>
End Function
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
                  "OASISRemoteAdmin.clsGost.EncryptByte", _
                  "clsGost component failure"
        '</EhFooter>
End Sub
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
110         Call DecryptBlock(RightWord, LeftWord)
112         LeftWord = LeftWord Xor CipherLeft
114         RightWord = RightWord Xor CipherRight
116         Call GetWord(CipherLeft, byteArray(), Offset)
118         Call GetWord(CipherRight, byteArray(), Offset + 4)
120         Call PutWord(LeftWord, byteArray(), Offset)
122         Call PutWord(RightWord, byteArray(), Offset + 4)
124         If (Offset >= NextPercent) Then
126             CurrPercent = Int((Offset / CipherLen) * 100)
128             NextPercent = (CipherLen * ((CurrPercent + 1) / 100)) + 1
130             RaiseEvent Progress(CurrPercent)
            End If
        Next
132     Call CopyMem(OrigLen, byteArray(8), 4)
134     If (CipherLen - OrigLen > 19) Or (CipherLen - OrigLen < 12) Then Call Err.Raise(vbObjectError, , "Incorrect size descriptor in Gost decryption")
136     Call CopyMem(byteArray(0), byteArray(12), OrigLen)
138     ReDim Preserve byteArray(OrigLen - 1)
140     If (CurrPercent <> 100) Then RaiseEvent Progress(100)

ErrorHandler:
        '<EhFooter>
        Exit Sub

DecryptByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsGost.DecryptByte", _
                  "clsGost component failure"
        '</EhFooter>
End Sub
Public Property Let Key(New_Value As String)
        '<EhHeader>
        On Error GoTo Key_Err
        '</EhHeader>
100     Class_Initialize
        Dim a As Long, Key() As Byte, KeyLen As Long, byteArray() As Byte
102     If (m_KeyValue = New_Value) Then Exit Property
104     KeyLen = Len(New_Value)
106     Key() = StrConv(New_Value, vbFromUnicode)
108     ReDim byteArray(0 To 31)
110     For a = 0 To 31
112         byteArray(a) = Key(a Mod KeyLen)
        Next
114     Call CopyMem(K(1), byteArray(0), 32)
116     m_KeyValue = New_Value
        '<EhFooter>
        Exit Property

Key_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsGost.Key", _
                  "clsGost component failure"
        '</EhFooter>
End Property
Private Sub Class_Initialize()
        '<EhHeader>
        On Error GoTo Class_Initialize_Err
        '</EhHeader>
        Dim a As Long, b As Long, c As Long, LeftWord As Long, S(0 To 7) As Variant
      
100     On Local Error Resume Next
102     m_RunningCompiled = ((2147483647 + 1) < 0)
      
104     S(0) = Array(6, 5, 1, 7, 14, 0, 4, 10, 11, 9, 3, 13, 8, 12, 2, 15)
106     S(1) = Array(14, 13, 9, 0, 8, 10, 12, 4, 7, 15, 6, 11, 3, 1, 5, 2)
108     S(2) = Array(6, 5, 1, 7, 2, 4, 10, 0, 11, 13, 14, 3, 8, 12, 15, 9)
110     S(3) = Array(8, 7, 3, 9, 6, 4, 14, 5, 2, 13, 0, 12, 1, 11, 10, 15)
112     S(4) = Array(10, 9, 6, 11, 5, 1, 8, 4, 0, 13, 7, 2, 14, 3, 15, 12)
114     S(5) = Array(5, 3, 0, 6, 11, 13, 4, 14, 10, 7, 1, 12, 2, 8, 15, 9)
116     S(6) = Array(2, 1, 12, 3, 11, 13, 15, 7, 10, 6, 9, 14, 0, 8, 4, 5)
118     S(7) = Array(6, 5, 1, 7, 8, 9, 4, 2, 15, 3, 13, 12, 10, 14, 11, 0)
    
120     For a = 0 To 15
122         For b = 0 To 7
124             sBox(b, a) = S(b)(a)
            Next
        Next
126     For a = 0 To 255
128         k87(a) = lBSL(CLng(sBox(7, lBSR(a, 4))), 4) Or sBox(6, a And 15)
130         k65(a) = lBSL(CLng(sBox(5, lBSR(a, 4))), 4) Or sBox(4, a And 15)
132         k43(a) = lBSL(CLng(sBox(3, lBSR(a, 4))), 4) Or sBox(2, a And 15)
134         k21(a) = lBSL(CLng(sBox(1, lBSR(a, 4))), 4) Or sBox(0, a And 15)
        Next
        '<EhFooter>
        Exit Sub

Class_Initialize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsGost.Class_Initialize", _
                  "clsGost component failure"
        '</EhFooter>
End Sub
