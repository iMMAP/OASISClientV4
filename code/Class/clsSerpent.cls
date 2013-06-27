VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSerpent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' 256-bit Serpent VB Implementation
' This is an extremely slow but working Visual Basic implementation of
' the 256-bit Serpent algorithm. Serpent was a finalist for AES and
' appears to be a very secure powerful algorithm. The C implementation
' is by far the fastest and most efficient. Since Visual Basic is bogged
' down by memory problems and lacks the power that C provides, high level
' speeds for this algorithm may never be achieved in VB. If you find ways
' to increase the speed please contact David and let him know.
'
' I would only recommend using this implementation for academic purposes
' and not in commercial schemes due to its speed. It's inherent security
' is a definite plus in any high risk situations.

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Event Progress(Percent As Integer)

#Const BLOCK_REVERSE = True

Private l_key(0 To 139) As MemBlock
Private RegX(0 To 3, 0 To 31) As Byte
Private TheByte(0 To 7) As Byte
Private bStr() As Byte
Private byteArray() As Byte
Private hiByte As Long
Private hiBound As Long

Private Type MemBlock
  Data(0 To 3) As Byte
End Type
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
                  "OASISRemoteAdmin.clsSerpent.Initialize64", _
                  "clsSerpent component failure"
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
                  "OASISRemoteAdmin.clsSerpent.Decode64", _
                  "clsSerpent component failure"
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
                  "OASISRemoteAdmin.clsSerpent.DecodeArray64", _
                  "clsSerpent component failure"
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
                  "OASISRemoteAdmin.clsSerpent.Encode64", _
                  "clsSerpent component failure"
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
                  "OASISRemoteAdmin.clsSerpent.EncodeArray64", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Sub PutBits(ByVal Val As Byte)
        '<EhHeader>
        On Error GoTo PutBits_Err
        '</EhHeader>
        Dim Count As Integer
100     For Count = 7 To 1 Step -1
102       TheByte(Count) = Val Mod 2
104       Val = Val \ 2
        Next
106     TheByte(0) = Val Mod 2
        '<EhFooter>
        Exit Sub

PutBits_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.PutBits", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub GetBits(Val As Byte)
        '<EhHeader>
        On Error GoTo GetBits_Err
        '</EhHeader>
        Dim Count As Integer
100     Val = TheByte(0)
102     For Count = 1 To 7
104         Val = Val * 2
106         Val = Val + TheByte(Count)
        Next
        '<EhFooter>
        Exit Sub

GetBits_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.GetBits", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub PutRegByte(ByVal Val As Byte, ByteOrder As Integer, Reg As Integer)
        '<EhHeader>
        On Error GoTo PutRegByte_Err
        '</EhHeader>
100     PutBits Val
102     RegX(Reg, ByteOrder * 8) = TheByte(0)
104     RegX(Reg, ByteOrder * 8 + 1) = TheByte(1)
106     RegX(Reg, ByteOrder * 8 + 2) = TheByte(2)
108     RegX(Reg, ByteOrder * 8 + 3) = TheByte(3)
110     RegX(Reg, ByteOrder * 8 + 4) = TheByte(4)
112     RegX(Reg, ByteOrder * 8 + 5) = TheByte(5)
114     RegX(Reg, ByteOrder * 8 + 6) = TheByte(6)
116     RegX(Reg, ByteOrder * 8 + 7) = TheByte(7)
        '<EhFooter>
        Exit Sub

PutRegByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.PutRegByte", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub GetRegByte(Val As Byte, ByteOrder As Integer, Reg As Integer)
        '<EhHeader>
        On Error GoTo GetRegByte_Err
        '</EhHeader>
100     TheByte(0) = RegX(Reg, ByteOrder * 8)
102     TheByte(1) = RegX(Reg, ByteOrder * 8 + 1)
104     TheByte(2) = RegX(Reg, ByteOrder * 8 + 2)
106     TheByte(3) = RegX(Reg, ByteOrder * 8 + 3)
108     TheByte(4) = RegX(Reg, ByteOrder * 8 + 4)
110     TheByte(5) = RegX(Reg, ByteOrder * 8 + 5)
112     TheByte(6) = RegX(Reg, ByteOrder * 8 + 6)
114     TheByte(7) = RegX(Reg, ByteOrder * 8 + 7)
116     GetBits Val
        '<EhFooter>
        Exit Sub

GetRegByte_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.GetRegByte", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub PutReg(Val As MemBlock, Reg As Integer)
        '<EhHeader>
        On Error GoTo PutReg_Err
        '</EhHeader>
100     PutRegByte Val.Data(0), 0, Reg
102     PutRegByte Val.Data(1), 1, Reg
104     PutRegByte Val.Data(2), 2, Reg
106     PutRegByte Val.Data(3), 3, Reg
        '<EhFooter>
        Exit Sub

PutReg_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.PutReg", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub GetReg(Val As MemBlock, Reg As Integer)
        '<EhHeader>
        On Error GoTo GetReg_Err
        '</EhHeader>
100     GetRegByte Val.Data(0), 0, Reg
102     GetRegByte Val.Data(1), 1, Reg
104     GetRegByte Val.Data(2), 2, Reg
106     GetRegByte Val.Data(3), 3, Reg
        '<EhFooter>
        Exit Sub

GetReg_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.GetReg", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Function Bitwise_OR(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo Bitwise_OR_Err
        '</EhHeader>
        Dim Count As Integer, Val3 As MemBlock
100     PutReg Val1, 0
102     PutReg Val2, 1
104     RegX(2, 0) = (RegX(0, 0) + RegX(1, 0) + 1) \ 2
106     RegX(2, 1) = (RegX(0, 1) + RegX(1, 1) + 1) \ 2
108     RegX(2, 2) = (RegX(0, 2) + RegX(1, 2) + 1) \ 2
110     RegX(2, 3) = (RegX(0, 3) + RegX(1, 3) + 1) \ 2
112     RegX(2, 4) = (RegX(0, 4) + RegX(1, 4) + 1) \ 2
114     RegX(2, 5) = (RegX(0, 5) + RegX(1, 5) + 1) \ 2
116     RegX(2, 6) = (RegX(0, 6) + RegX(1, 6) + 1) \ 2
118     RegX(2, 7) = (RegX(0, 7) + RegX(1, 7) + 1) \ 2
120     RegX(2, 8) = (RegX(0, 8) + RegX(1, 8) + 1) \ 2
122     RegX(2, 9) = (RegX(0, 9) + RegX(1, 9) + 1) \ 2
124     RegX(2, 10) = (RegX(0, 10) + RegX(1, 10) + 1) \ 2
126     RegX(2, 11) = (RegX(0, 11) + RegX(1, 11) + 1) \ 2
128     RegX(2, 12) = (RegX(0, 12) + RegX(1, 12) + 1) \ 2
130     RegX(2, 13) = (RegX(0, 13) + RegX(1, 13) + 1) \ 2
132     RegX(2, 14) = (RegX(0, 14) + RegX(1, 14) + 1) \ 2
134     RegX(2, 15) = (RegX(0, 15) + RegX(1, 15) + 1) \ 2
136     RegX(2, 16) = (RegX(0, 16) + RegX(1, 16) + 1) \ 2
138     RegX(2, 17) = (RegX(0, 17) + RegX(1, 17) + 1) \ 2
140     RegX(2, 18) = (RegX(0, 18) + RegX(1, 18) + 1) \ 2
142     RegX(2, 19) = (RegX(0, 19) + RegX(1, 19) + 1) \ 2
144     RegX(2, 20) = (RegX(0, 20) + RegX(1, 20) + 1) \ 2
146     RegX(2, 21) = (RegX(0, 21) + RegX(1, 21) + 1) \ 2
148     RegX(2, 22) = (RegX(0, 22) + RegX(1, 22) + 1) \ 2
150     RegX(2, 23) = (RegX(0, 23) + RegX(1, 23) + 1) \ 2
152     RegX(2, 24) = (RegX(0, 24) + RegX(1, 24) + 1) \ 2
154     RegX(2, 25) = (RegX(0, 25) + RegX(1, 25) + 1) \ 2
156     RegX(2, 26) = (RegX(0, 26) + RegX(1, 26) + 1) \ 2
158     RegX(2, 27) = (RegX(0, 27) + RegX(1, 27) + 1) \ 2
160     RegX(2, 28) = (RegX(0, 28) + RegX(1, 28) + 1) \ 2
162     RegX(2, 29) = (RegX(0, 29) + RegX(1, 29) + 1) \ 2
164     RegX(2, 30) = (RegX(0, 30) + RegX(1, 30) + 1) \ 2
166     RegX(2, 31) = (RegX(0, 31) + RegX(1, 31) + 1) \ 2
168     GetReg Val3, 2
170     Bitwise_OR = Val3
        '<EhFooter>
        Exit Function

Bitwise_OR_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Bitwise_OR", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function Bitwise_AND(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo Bitwise_AND_Err
        '</EhHeader>
      Dim Count As Integer, Val3 As MemBlock
  
100   PutReg Val1, 0
102   PutReg Val2, 1
  
104   RegX(2, 0) = RegX(0, 0) * RegX(1, 0)
106   RegX(2, 1) = RegX(0, 1) * RegX(1, 1)
108   RegX(2, 2) = RegX(0, 2) * RegX(1, 2)
110   RegX(2, 3) = RegX(0, 3) * RegX(1, 3)
112   RegX(2, 4) = RegX(0, 4) * RegX(1, 4)
114   RegX(2, 5) = RegX(0, 5) * RegX(1, 5)
116   RegX(2, 6) = RegX(0, 6) * RegX(1, 6)
118   RegX(2, 7) = RegX(0, 7) * RegX(1, 7)
120   RegX(2, 8) = RegX(0, 8) * RegX(1, 8)
122   RegX(2, 9) = RegX(0, 9) * RegX(1, 9)
124   RegX(2, 10) = RegX(0, 10) * RegX(1, 10)
126   RegX(2, 11) = RegX(0, 11) * RegX(1, 11)
128   RegX(2, 12) = RegX(0, 12) * RegX(1, 12)
130   RegX(2, 13) = RegX(0, 13) * RegX(1, 13)
132   RegX(2, 14) = RegX(0, 14) * RegX(1, 14)
134   RegX(2, 15) = RegX(0, 15) * RegX(1, 15)
136   RegX(2, 16) = RegX(0, 16) * RegX(1, 16)
138   RegX(2, 17) = RegX(0, 17) * RegX(1, 17)
140   RegX(2, 18) = RegX(0, 18) * RegX(1, 18)
142   RegX(2, 19) = RegX(0, 19) * RegX(1, 19)
144   RegX(2, 20) = RegX(0, 20) * RegX(1, 20)
146   RegX(2, 21) = RegX(0, 21) * RegX(1, 21)
148   RegX(2, 22) = RegX(0, 22) * RegX(1, 22)
150   RegX(2, 23) = RegX(0, 23) * RegX(1, 23)
152   RegX(2, 24) = RegX(0, 24) * RegX(1, 24)
154   RegX(2, 25) = RegX(0, 25) * RegX(1, 25)
156   RegX(2, 26) = RegX(0, 26) * RegX(1, 26)
158   RegX(2, 27) = RegX(0, 27) * RegX(1, 27)
160   RegX(2, 28) = RegX(0, 28) * RegX(1, 28)
162   RegX(2, 29) = RegX(0, 29) * RegX(1, 29)
164   RegX(2, 30) = RegX(0, 30) * RegX(1, 30)
166   RegX(2, 31) = RegX(0, 31) * RegX(1, 31)

168   GetReg Val3, 2
170   Bitwise_AND = Val3
        '<EhFooter>
        Exit Function

Bitwise_AND_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Bitwise_AND", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function Bitwise_XOR(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo Bitwise_XOR_Err
        '</EhHeader>
        Dim Val3 As MemBlock
100     PutReg Val1, 0
102     PutReg Val2, 1
104     RegX(2, 0) = (RegX(0, 0) + RegX(1, 0)) Mod 2
106     RegX(2, 1) = (RegX(0, 1) + RegX(1, 1)) Mod 2
108     RegX(2, 2) = (RegX(0, 2) + RegX(1, 2)) Mod 2
110     RegX(2, 3) = (RegX(0, 3) + RegX(1, 3)) Mod 2
112     RegX(2, 4) = (RegX(0, 4) + RegX(1, 4)) Mod 2
114     RegX(2, 5) = (RegX(0, 5) + RegX(1, 5)) Mod 2
116     RegX(2, 6) = (RegX(0, 6) + RegX(1, 6)) Mod 2
118     RegX(2, 7) = (RegX(0, 7) + RegX(1, 7)) Mod 2
120     RegX(2, 8) = (RegX(0, 8) + RegX(1, 8)) Mod 2
122     RegX(2, 9) = (RegX(0, 9) + RegX(1, 9)) Mod 2
124     RegX(2, 10) = (RegX(0, 10) + RegX(1, 10)) Mod 2
126     RegX(2, 11) = (RegX(0, 11) + RegX(1, 11)) Mod 2
128     RegX(2, 12) = (RegX(0, 12) + RegX(1, 12)) Mod 2
130     RegX(2, 13) = (RegX(0, 13) + RegX(1, 13)) Mod 2
132     RegX(2, 14) = (RegX(0, 14) + RegX(1, 14)) Mod 2
134     RegX(2, 15) = (RegX(0, 15) + RegX(1, 15)) Mod 2
136     RegX(2, 16) = (RegX(0, 16) + RegX(1, 16)) Mod 2
138     RegX(2, 17) = (RegX(0, 17) + RegX(1, 17)) Mod 2
140     RegX(2, 18) = (RegX(0, 18) + RegX(1, 18)) Mod 2
142     RegX(2, 19) = (RegX(0, 19) + RegX(1, 19)) Mod 2
144     RegX(2, 20) = (RegX(0, 20) + RegX(1, 20)) Mod 2
146     RegX(2, 21) = (RegX(0, 21) + RegX(1, 21)) Mod 2
148     RegX(2, 22) = (RegX(0, 22) + RegX(1, 22)) Mod 2
150     RegX(2, 23) = (RegX(0, 23) + RegX(1, 23)) Mod 2
152     RegX(2, 24) = (RegX(0, 24) + RegX(1, 24)) Mod 2
154     RegX(2, 25) = (RegX(0, 25) + RegX(1, 25)) Mod 2
156     RegX(2, 26) = (RegX(0, 26) + RegX(1, 26)) Mod 2
158     RegX(2, 27) = (RegX(0, 27) + RegX(1, 27)) Mod 2
160     RegX(2, 28) = (RegX(0, 28) + RegX(1, 28)) Mod 2
162     RegX(2, 29) = (RegX(0, 29) + RegX(1, 29)) Mod 2
164     RegX(2, 30) = (RegX(0, 30) + RegX(1, 30)) Mod 2
166     RegX(2, 31) = (RegX(0, 31) + RegX(1, 31)) Mod 2
168     GetReg Val3, 2
170     Bitwise_XOR = Val3
        '<EhFooter>
        Exit Function

Bitwise_XOR_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Bitwise_XOR", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function Bitwise_NOT(Val As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo Bitwise_NOT_Err
        '</EhHeader>
        Dim Count As Integer, RetVal As MemBlock
100     PutReg Val, 0
    
102     RegX(0, 0) = (RegX(0, 0) + 1) Mod 2
104     RegX(0, 1) = (RegX(0, 1) + 1) Mod 2
106     RegX(0, 2) = (RegX(0, 2) + 1) Mod 2
108     RegX(0, 3) = (RegX(0, 3) + 1) Mod 2
110     RegX(0, 4) = (RegX(0, 4) + 1) Mod 2
112     RegX(0, 5) = (RegX(0, 5) + 1) Mod 2
114     RegX(0, 6) = (RegX(0, 6) + 1) Mod 2
116     RegX(0, 7) = (RegX(0, 7) + 1) Mod 2
118     RegX(0, 8) = (RegX(0, 8) + 1) Mod 2
120     RegX(0, 9) = (RegX(0, 9) + 1) Mod 2
122     RegX(0, 10) = (RegX(0, 10) + 1) Mod 2
124     RegX(0, 11) = (RegX(0, 11) + 1) Mod 2
126     RegX(0, 12) = (RegX(0, 12) + 1) Mod 2
128     RegX(0, 13) = (RegX(0, 13) + 1) Mod 2
130     RegX(0, 14) = (RegX(0, 14) + 1) Mod 2
132     RegX(0, 15) = (RegX(0, 15) + 1) Mod 2
134     RegX(0, 16) = (RegX(0, 16) + 1) Mod 2
136     RegX(0, 17) = (RegX(0, 17) + 1) Mod 2
138     RegX(0, 18) = (RegX(0, 18) + 1) Mod 2
140     RegX(0, 19) = (RegX(0, 19) + 1) Mod 2
142     RegX(0, 20) = (RegX(0, 20) + 1) Mod 2
144     RegX(0, 21) = (RegX(0, 21) + 1) Mod 2
146     RegX(0, 22) = (RegX(0, 22) + 1) Mod 2
148     RegX(0, 23) = (RegX(0, 23) + 1) Mod 2
150     RegX(0, 24) = (RegX(0, 24) + 1) Mod 2
152     RegX(0, 25) = (RegX(0, 25) + 1) Mod 2
154     RegX(0, 26) = (RegX(0, 26) + 1) Mod 2
156     RegX(0, 27) = (RegX(0, 27) + 1) Mod 2
158     RegX(0, 28) = (RegX(0, 28) + 1) Mod 2
160     RegX(0, 29) = (RegX(0, 29) + 1) Mod 2
162     RegX(0, 30) = (RegX(0, 30) + 1) Mod 2
164     RegX(0, 31) = (RegX(0, 31) + 1) Mod 2
166     GetReg RetVal, 0
168     Bitwise_NOT = RetVal
        '<EhFooter>
        Exit Function

Bitwise_NOT_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Bitwise_NOT", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function Bitwise_LShift(Val As MemBlock, Pos As Integer) As MemBlock
        '<EhHeader>
        On Error GoTo Bitwise_LShift_Err
        '</EhHeader>
        Dim Count As Integer, RetVal As MemBlock
100     PutReg Val, 0
102     If Pos = 0 Then
104         GetReg RetVal, 0
106         Bitwise_LShift = RetVal
            Exit Function
        End If
108     If Pos >= 32 Or Pos < 0 Then
110         SetInt RetVal, 0
112         Bitwise_LShift = RetVal
            Exit Function
        End If
114     For Count = 0 To 31 - Pos
116         RegX(2, Count) = RegX(0, Pos + Count)
        Next
118     For Count = 32 - Pos To 31
120         RegX(2, Count) = 0
122     Next Count
124     GetReg RetVal, 2
126     Bitwise_LShift = RetVal
        '<EhFooter>
        Exit Function

Bitwise_LShift_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Bitwise_LShift", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function ulBitwise_LShift(Val As MemBlock, ulPos As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo ulBitwise_LShift_Err
        '</EhHeader>
        Dim Pos As Integer
100     GetInt ulPos, Pos
102     Pos = Pos Mod 32
104     ulBitwise_LShift = Bitwise_LShift(Val, Pos)
        '<EhFooter>
        Exit Function

ulBitwise_LShift_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ulBitwise_LShift", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function Bitwise_RShift(Val As MemBlock, Pos As Integer) As MemBlock
        '<EhHeader>
        On Error GoTo Bitwise_RShift_Err
        '</EhHeader>
        Dim Count As Integer, RetVal As MemBlock
100     PutReg Val, 0
102     If Pos = 0 Then
104         GetReg RetVal, 0
106         Bitwise_RShift = RetVal
            Exit Function
        End If
108     If Pos >= 32 Or Pos < 0 Then
110         SetInt RetVal, 0
112         Bitwise_RShift = RetVal
            Exit Function
        End If
114     For Count = 0 To Pos - 1
116         RegX(2, Count) = 0
        Next
118     For Count = Pos To 31
120         RegX(2, Count) = RegX(0, Count - Pos)
122     Next Count
124     GetReg RetVal, 2
126     Bitwise_RShift = RetVal
        '<EhFooter>
        Exit Function

Bitwise_RShift_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Bitwise_RShift", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function ulBitwise_RShift(Val As MemBlock, ulPos As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo ulBitwise_RShift_Err
        '</EhHeader>
        Dim Pos As Integer
100     GetInt ulPos, Pos
102     Pos = Pos Mod 32
104     ulBitwise_RShift = Bitwise_RShift(Val, Pos)
        '<EhFooter>
        Exit Function

ulBitwise_RShift_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ulBitwise_RShift", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function rotl(Val As MemBlock, Pos As Integer) As MemBlock
        '<EhHeader>
        On Error GoTo rotl_Err
        '</EhHeader>
100     rotl = Bitwise_OR(Bitwise_LShift(Val, Pos), Bitwise_RShift(Val, 32 - Pos))
        '<EhFooter>
        Exit Function

rotl_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.rotl", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function rotr(Val As MemBlock, Pos As Integer) As MemBlock
        '<EhHeader>
        On Error GoTo rotr_Err
        '</EhHeader>
100     rotr = Bitwise_OR(Bitwise_RShift(Val, Pos), Bitwise_LShift(Val, 32 - Pos))
        '<EhFooter>
        Exit Function

rotr_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.rotr", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function ul_rotl(Val As MemBlock, ulPos As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo ul_rotl_Err
        '</EhHeader>
        Dim Pos As Integer
100     GetInt ulPos, Pos
102     Pos = Pos Mod 32
104     ul_rotl = Bitwise_OR(Bitwise_LShift(Val, Pos), Bitwise_RShift(Val, 32 - Pos))
        '<EhFooter>
        Exit Function

ul_rotl_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ul_rotl", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function ul_rotr(Val As MemBlock, ulPos As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo ul_rotr_Err
        '</EhHeader>
        Dim Pos As Integer
100     GetInt ulPos, Pos
102     Pos = Pos Mod 32
104     ul_rotr = Bitwise_OR(Bitwise_RShift(Val, Pos), Bitwise_LShift(Val, 32 - Pos))
        '<EhFooter>
        Exit Function

ul_rotr_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ul_rotr", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function bswap(Val As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo bswap_Err
        '</EhHeader>
        Dim Temp1 As MemBlock, Temp2 As MemBlock
100     SetData Temp1, 0, 255, 0, 255
102     SetData Temp2, 255, 0, 255, 0
104     bswap = Bitwise_OR(Bitwise_AND(rotl(Val, 8), Temp1), Bitwise_AND(rotr(Val, 8), Temp2))
        '<EhFooter>
        Exit Function

bswap_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.bswap", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function Sum(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo Sum_Err
        '</EhHeader>
        Dim Count As Integer, Val3 As MemBlock
    
100     PutReg Val1, 0
102     PutReg Val2, 1
    
        Dim NextPos As Integer
104     NextPos = 0
106     For Count = 31 To 0 Step -1
108       RegX(2, Count) = (RegX(0, Count) + RegX(1, Count) + NextPos) Mod 2
110       NextPos = (RegX(0, Count) + RegX(1, Count) + NextPos) \ 2
112     Next Count
114     GetReg Val3, 2
116     Sum = Val3
        '<EhFooter>
        Exit Function

Sum_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Sum", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function Prod(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo Prod_Err
        '</EhHeader>
        Dim Count As Integer, SumVal As MemBlock
100     SetInt SumVal, 0
102     PutReg Val2, 3
104     For Count = 31 To 0 Step -1
106       If RegX(3, Count) = 1 Then SumVal = Sum(SumVal, Bitwise_LShift(Val1, 31 - Count))
        Next
108     Prod = SumVal
        '<EhFooter>
        Exit Function

Prod_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Prod", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function ChangeSign(Val As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo ChangeSign_Err
        '</EhHeader>
        Dim Count As Integer, RetVal As MemBlock, ToChange As Boolean
100     PutReg Val, 0
102     ToChange = False
104     For Count = 31 To 0 Step -1
106         If ToChange Then
108             RegX(2, Count) = (RegX(0, Count) + 1) Mod 2
            Else
110             RegX(2, Count) = RegX(0, Count)
112             If RegX(2, Count) = 1 Then ToChange = True
            End If
        Next
114     GetReg RetVal, 2
116     ChangeSign = RetVal
        '<EhFooter>
        Exit Function

ChangeSign_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ChangeSign", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Function Substr(Val1 As MemBlock, Val2 As MemBlock) As MemBlock
        '<EhHeader>
        On Error GoTo Substr_Err
        '</EhHeader>
100     Substr = Sum(Val1, ChangeSign(Val2))
        '<EhFooter>
        Exit Function

Substr_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Substr", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Sub SetInt(Dest As MemBlock, IntValue As Byte)
        '<EhHeader>
        On Error GoTo SetInt_Err
        '</EhHeader>
100     Dest.Data(0) = 0
102     Dest.Data(1) = 0
104     Dest.Data(2) = 0
106     Dest.Data(3) = IntValue
        '<EhFooter>
        Exit Sub

SetInt_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.SetInt", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub SetData(Dest As MemBlock, IVal1 As Byte, IVal2 As Byte, IVal3 As Byte, IVal4 As Byte)
        '<EhHeader>
        On Error GoTo SetData_Err
        '</EhHeader>
100     Dest.Data(0) = IVal1
102     Dest.Data(1) = IVal2
104     Dest.Data(2) = IVal3
106     Dest.Data(3) = IVal4
        '<EhFooter>
        Exit Sub

SetData_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.SetData", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub GetInt(Src As MemBlock, IntValue As Integer)
        '<EhHeader>
        On Error GoTo GetInt_Err
        '</EhHeader>
100     IntValue = Src.Data(3)
        '<EhFooter>
        Exit Sub

GetInt_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.GetInt", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub MCopy(Src As MemBlock, Dest As MemBlock)
        '<EhHeader>
        On Error GoTo MCopy_Err
        '</EhHeader>
100     PutReg Src, 0
102     GetReg Dest, 0
        '<EhFooter>
        Exit Sub

MCopy_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.MCopy", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub Append(ByRef StringData As String, Optional Length As Long)
        '<EhHeader>
        On Error GoTo Append_Err
        '</EhHeader>
        Dim DataLength As Long
100     If Length > 0 Then DataLength = Length Else DataLength = Len(StringData)
102     If DataLength + hiByte > hiBound Then
104         hiBound = hiBound + 1024
106         ReDim Preserve byteArray(hiBound)
        End If
108     CopyMem ByVal VarPtr(byteArray(hiByte)), ByVal StringData, DataLength
110     hiByte = hiByte + DataLength
        '<EhFooter>
        Exit Sub

Append_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Append", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Property Get GData() As String
        '<EhHeader>
        On Error GoTo GData_Err
        '</EhHeader>
        Dim StringData As String
100     StringData = Space(hiByte)
102     CopyMem ByVal StringData, ByVal VarPtr(byteArray(0)), hiByte
104     GData = StringData
        '<EhFooter>
        Exit Property

GData_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.GData", _
                  "clsSerpent component failure"
        '</EhFooter>
End Property
Private Sub Reset()
        '<EhHeader>
        On Error GoTo Reset_Err
        '</EhHeader>
100     hiByte = 0
102     hiBound = 1024
104     ReDim byteArray(hiBound)
        '<EhFooter>
        Exit Sub

Reset_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.Reset", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub SetBuffer(Buffer() As Byte, Str As String, Length As Double)
        '<EhHeader>
        On Error GoTo SetBuffer_Err
        '</EhHeader>
        Dim CF As Double, StrLength As Double
100     StrLength = Len(Str)
102     For CF = 0 To Length - 1
104         If CF < StrLength Then Buffer(CF) = Asc(Mid$(Str, (CF + 1), 1)) Else Buffer(CF) = 0
        Next
        '<EhFooter>
        Exit Sub

SetBuffer_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.SetBuffer", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub SetBlock(Blk() As MemBlock, Buffer() As Byte, Pos As Double, Length As Double)
        '<EhHeader>
        On Error GoTo SetBlock_Err
        '</EhHeader>
100     On Local Error Resume Next
        Dim Count As Double

102     For Count = 0 To Length - 1
104         Blk(Count \ 4).Data(Count Mod 4) = Buffer(Pos + Count)
        Next
        '<EhFooter>
        Exit Sub

SetBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.SetBlock", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub GetBlock(Blk() As MemBlock, Buffer() As Byte, Pos As Double, Length As Double)
        '<EhHeader>
        On Error GoTo GetBlock_Err
        '</EhHeader>
100     On Local Error Resume Next
        Dim Count As Integer
102     For Count = 0 To Length - 1
104         Buffer(Pos + Count) = Blk(Count \ 4).Data(Count Mod 4)
        Next
        '<EhFooter>
        Exit Sub

GetBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.GetBlock", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub GetBuffer(Buffer() As Byte, Str As String, Length As Double)
        '<EhHeader>
        On Error GoTo GetBuffer_Err
        '</EhHeader>
        Dim Count As Integer
100     Str = ""
102     Reset
104     For Count = 0 To Length - 1
106         Append Chr$(Buffer(Count))
108     Next Count
110     Str = GData
112     Reset
        '<EhFooter>
        Exit Sub

GetBuffer_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.GetBuffer", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Public Function EncryptString(Text As String, Optional Key As String, Optional OutputIn64 As Boolean) As String
        '<EhHeader>
        On Error GoTo EncryptString_Err
        '</EhHeader>
100     On Local Error Resume Next
        Dim kblk(0 To 7) As MemBlock, blk1(0 To 3) As MemBlock, blk2(0 To 3) As MemBlock
        Dim StrBuffer() As Byte, KeyBuffer() As Byte, RetVal() As Byte
        Dim Length As Double, Count As Double, PProg As Integer
102     Length = Len(Text)
104     Length = IIf(Length = 0, 16, 16 * ((Length - 1) \ 16 + 1))
106     ReDim StrBuffer(0 To Length - 1)
108     ReDim RetVal(0 To Length - 1)
110     SetKey Key
112     SetBuffer StrBuffer, Text, Length

114     For Count = 0 To Length - 1 Step 16
116         SetBlock blk1, StrBuffer, Count, 16
118         EncryptBlock blk1, blk2
120         GetBlock blk2, RetVal, Count, 16
122         PProg = (Count / Length) * 100
124         If PProg > 100 Or Count = (Length - 16) Then PProg = 100
126         RaiseEvent Progress(PProg)
128         DoEvents
        Next
130     EncryptString = StrConv(RetVal, vbUnicode)
132     If OutputIn64 = True Then EncryptString = Encode64(EncryptString)
        '<EhFooter>
        Exit Function

EncryptString_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.EncryptString", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Public Function DecryptString(Text As String, Optional Key As String, Optional IsTextIn64 As Boolean) As String
        '<EhHeader>
        On Error GoTo DecryptString_Err
        '</EhHeader>
100     On Local Error Resume Next
        Dim blk1(0 To 3) As MemBlock, blk2(0 To 3) As MemBlock, Code() As Byte
        Dim StrBuffer() As Byte, RetVal As String, Count As Double, Length As Double, PProg As Integer
102     If IsTextIn64 = True Then Text = Decode64(Text)
104     Code() = StrConv(Text, vbFromUnicode)
106     Length = UBound(Code) + 1
108     ReDim StrBuffer(0 To Length - 1)
110     SetKey Key
112     For Count = 0 To Length - 1 Step 16
114         SetBlock blk1, Code, Count, 16
116         DecryptBlock blk1, blk2
118         GetBlock blk2, StrBuffer, Count, 16
120         PProg = (Count / Length) * 100
122         If PProg > 100 Then PProg = 100
124         RaiseEvent Progress(PProg)
126         DoEvents
        Next
128     GetBuffer StrBuffer, RetVal, Length
130     DecryptString = RetVal
        '<EhFooter>
        Exit Function

DecryptString_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.DecryptString", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Sub ib5(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo ib5_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_NOT(c): t2 = Bitwise_AND(b, t1): t3 = Bitwise_XOR(d, t2)
102     t4 = Bitwise_AND(a, t3): t5 = Bitwise_XOR(b, t1): h = Bitwise_XOR(t4, t5)
104     t7 = Bitwise_OR(b, h): t8 = Bitwise_AND(a, t7): f = Bitwise_XOR(t3, t8)
106     t10 = Bitwise_OR(a, d): t11 = Bitwise_XOR(t1, t7): e = Bitwise_XOR(t10, t11)
108     t13 = Bitwise_XOR(a, c): t14 = Bitwise_AND(b, t10): t15 = Bitwise_OR(t4, t13)
110     g = Bitwise_XOR(t14, t15)
        '<EhFooter>
        Exit Sub

ib5_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ib5", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub EncryptBlock(in_blk() As MemBlock, out_blk() As MemBlock)
        '<EhHeader>
        On Error GoTo EncryptBlock_Err
        '</EhHeader>
        Dim a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock
    
        #If BLOCK_REVERSE Then
100       a = bswap(in_blk(3))
102       b = bswap(in_blk(2))
104       c = bswap(in_blk(1))
106       d = bswap(in_blk(0))
        #Else
108       MCopy in_blk(0), a
110       MCopy in_blk(1), b
112       MCopy in_blk(2), c
114       MCopy in_blk(3), d
        #End If
116     k_xor 0, a, b, c, d: sb0 a, b, c, d, e, f, g, h: rot e, f, g, h
118     k_xor 1, e, f, g, h: sb1 e, f, g, h, a, b, c, d: rot a, b, c, d
120     k_xor 2, a, b, c, d: sb2 a, b, c, d, e, f, g, h: rot e, f, g, h
122     k_xor 3, e, f, g, h: sb3 e, f, g, h, a, b, c, d: rot a, b, c, d
124     k_xor 4, a, b, c, d: sb4 a, b, c, d, e, f, g, h: rot e, f, g, h
126     k_xor 5, e, f, g, h: sb5 e, f, g, h, a, b, c, d: rot a, b, c, d
128     k_xor 6, a, b, c, d: sb6 a, b, c, d, e, f, g, h: rot e, f, g, h
130     k_xor 7, e, f, g, h: sb7 e, f, g, h, a, b, c, d: rot a, b, c, d
132     k_xor 8, a, b, c, d: sb0 a, b, c, d, e, f, g, h: rot e, f, g, h
134     k_xor 9, e, f, g, h: sb1 e, f, g, h, a, b, c, d: rot a, b, c, d
136     k_xor 10, a, b, c, d: sb2 a, b, c, d, e, f, g, h: rot e, f, g, h
138     k_xor 11, e, f, g, h: sb3 e, f, g, h, a, b, c, d: rot a, b, c, d
140     k_xor 12, a, b, c, d: sb4 a, b, c, d, e, f, g, h: rot e, f, g, h
142     k_xor 13, e, f, g, h: sb5 e, f, g, h, a, b, c, d: rot a, b, c, d
144     k_xor 14, a, b, c, d: sb6 a, b, c, d, e, f, g, h: rot e, f, g, h
146     k_xor 15, e, f, g, h: sb7 e, f, g, h, a, b, c, d: rot a, b, c, d
148     k_xor 16, a, b, c, d: sb0 a, b, c, d, e, f, g, h: rot e, f, g, h
150     k_xor 17, e, f, g, h: sb1 e, f, g, h, a, b, c, d: rot a, b, c, d
152     k_xor 18, a, b, c, d: sb2 a, b, c, d, e, f, g, h: rot e, f, g, h
154     k_xor 19, e, f, g, h: sb3 e, f, g, h, a, b, c, d: rot a, b, c, d
156     k_xor 20, a, b, c, d: sb4 a, b, c, d, e, f, g, h: rot e, f, g, h
158     k_xor 21, e, f, g, h: sb5 e, f, g, h, a, b, c, d: rot a, b, c, d
160     k_xor 22, a, b, c, d: sb6 a, b, c, d, e, f, g, h: rot e, f, g, h
162     k_xor 23, e, f, g, h: sb7 e, f, g, h, a, b, c, d: rot a, b, c, d
164     k_xor 24, a, b, c, d: sb0 a, b, c, d, e, f, g, h: rot e, f, g, h
166     k_xor 25, e, f, g, h: sb1 e, f, g, h, a, b, c, d: rot a, b, c, d
168     k_xor 26, a, b, c, d: sb2 a, b, c, d, e, f, g, h: rot e, f, g, h
170     k_xor 27, e, f, g, h: sb3 e, f, g, h, a, b, c, d: rot a, b, c, d
172     k_xor 28, a, b, c, d: sb4 a, b, c, d, e, f, g, h: rot e, f, g, h
174     k_xor 29, e, f, g, h: sb5 e, f, g, h, a, b, c, d: rot a, b, c, d
176     k_xor 30, a, b, c, d: sb6 a, b, c, d, e, f, g, h: rot e, f, g, h
178     k_xor 31, e, f, g, h: sb7 e, f, g, h, a, b, c, d: k_xor 32, a, b, c, d
        #If BLOCK_REVERSE Then
180       out_blk(3) = bswap(a)
182       out_blk(2) = bswap(b)
184       out_blk(1) = bswap(c)
186       out_blk(0) = bswap(d)
        #Else
188       MCopy a, out_blk(0)
190       MCopy b, out_blk(1)
192       MCopy c, out_blk(2)
194       MCopy d, out_blk(3)
        #End If
        '<EhFooter>
        Exit Sub

EncryptBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.EncryptBlock", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub DecryptBlock(in_blk() As MemBlock, out_blk() As MemBlock)
        '<EhHeader>
        On Error GoTo DecryptBlock_Err
        '</EhHeader>
        Dim a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock
    
        #If BLOCK_REVERSE Then
100       a = bswap(in_blk(3))
102       b = bswap(in_blk(2))
104       c = bswap(in_blk(1))
106       d = bswap(in_blk(0))
        #Else
108       MCopy in_blk(0), a
110       MCopy in_blk(1), b
112       MCopy in_blk(2), c
114       MCopy in_blk(3), d
        #End If
116     k_xor 32, a, b, c, d: ib7 a, b, c, d, e, f, g, h: k_xor 31, e, f, g, h
118     irot e, f, g, h: ib6 e, f, g, h, a, b, c, d: k_xor 30, a, b, c, d
120     irot a, b, c, d: ib5 a, b, c, d, e, f, g, h: k_xor 29, e, f, g, h
122     irot e, f, g, h: ib4 e, f, g, h, a, b, c, d: k_xor 28, a, b, c, d
124     irot a, b, c, d: ib3 a, b, c, d, e, f, g, h: k_xor 27, e, f, g, h
126     irot e, f, g, h: ib2 e, f, g, h, a, b, c, d: k_xor 26, a, b, c, d
128     irot a, b, c, d: ib1 a, b, c, d, e, f, g, h: k_xor 25, e, f, g, h
130     irot e, f, g, h: ib0 e, f, g, h, a, b, c, d: k_xor 24, a, b, c, d
132     irot a, b, c, d: ib7 a, b, c, d, e, f, g, h: k_xor 23, e, f, g, h
134     irot e, f, g, h: ib6 e, f, g, h, a, b, c, d: k_xor 22, a, b, c, d
136     irot a, b, c, d: ib5 a, b, c, d, e, f, g, h: k_xor 21, e, f, g, h
138     irot e, f, g, h: ib4 e, f, g, h, a, b, c, d: k_xor 20, a, b, c, d
140     irot a, b, c, d: ib3 a, b, c, d, e, f, g, h: k_xor 19, e, f, g, h
142     irot e, f, g, h: ib2 e, f, g, h, a, b, c, d: k_xor 18, a, b, c, d
144     irot a, b, c, d: ib1 a, b, c, d, e, f, g, h: k_xor 17, e, f, g, h
146     irot e, f, g, h: ib0 e, f, g, h, a, b, c, d: k_xor 16, a, b, c, d
148     irot a, b, c, d: ib7 a, b, c, d, e, f, g, h: k_xor 15, e, f, g, h
150     irot e, f, g, h: ib6 e, f, g, h, a, b, c, d: k_xor 14, a, b, c, d
152     irot a, b, c, d: ib5 a, b, c, d, e, f, g, h: k_xor 13, e, f, g, h
154     irot e, f, g, h: ib4 e, f, g, h, a, b, c, d: k_xor 12, a, b, c, d
156     irot a, b, c, d: ib3 a, b, c, d, e, f, g, h: k_xor 11, e, f, g, h
158     irot e, f, g, h: ib2 e, f, g, h, a, b, c, d: k_xor 10, a, b, c, d
160     irot a, b, c, d: ib1 a, b, c, d, e, f, g, h: k_xor 9, e, f, g, h
162     irot e, f, g, h: ib0 e, f, g, h, a, b, c, d: k_xor 8, a, b, c, d
164     irot a, b, c, d: ib7 a, b, c, d, e, f, g, h: k_xor 7, e, f, g, h
166     irot e, f, g, h: ib6 e, f, g, h, a, b, c, d: k_xor 6, a, b, c, d
168     irot a, b, c, d: ib5 a, b, c, d, e, f, g, h: k_xor 5, e, f, g, h
170     irot e, f, g, h: ib4 e, f, g, h, a, b, c, d: k_xor 4, a, b, c, d
172     irot a, b, c, d: ib3 a, b, c, d, e, f, g, h: k_xor 3, e, f, g, h
174     irot e, f, g, h: ib2 e, f, g, h, a, b, c, d: k_xor 2, a, b, c, d
176     irot a, b, c, d: ib1 a, b, c, d, e, f, g, h: k_xor 1, e, f, g, h
178     irot e, f, g, h: ib0 e, f, g, h, a, b, c, d: k_xor 0, a, b, c, d
        #If BLOCK_REVERSE Then
180       out_blk(3) = bswap(a)
182       out_blk(2) = bswap(b)
184       out_blk(1) = bswap(c)
186       out_blk(0) = bswap(d)
        #Else
188       MCopy a, out_blk(0)
190       MCopy b, out_blk(1)
192       MCopy c, out_blk(2)
194       MCopy d, out_blk(3)
        #End If
        '<EhFooter>
        Exit Sub

DecryptBlock_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.DecryptBlock", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub k_get(r As Integer, a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock)
        '<EhHeader>
        On Error GoTo k_get_Err
        '</EhHeader>
100    MCopy a, l_key(4 * r + 8)
102    MCopy b, l_key(4 * r + 9)
104    MCopy c, l_key(4 * r + 10)
106    MCopy d, l_key(4 * r + 11)
        '<EhFooter>
        Exit Sub

k_get_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.k_get", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub k_set(r As Integer, a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock)
        '<EhHeader>
        On Error GoTo k_set_Err
        '</EhHeader>
100    MCopy l_key(4 * r + 8), a
102    MCopy l_key(4 * r + 9), b
104    MCopy l_key(4 * r + 10), c
106    MCopy l_key(4 * r + 11), d
        '<EhFooter>
        Exit Sub

k_set_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.k_set", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub irot(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock)
        '<EhHeader>
        On Error GoTo irot_Err
        '</EhHeader>
100     c = rotr(c, 22)
102     a = rotr(a, 5)
104     c = Bitwise_XOR(c, Bitwise_XOR(d, Bitwise_LShift(b, 7)))
106     a = Bitwise_XOR(a, Bitwise_XOR(b, d))
108     d = rotr(d, 7)
110     b = rotr(b, 1)
112     d = Bitwise_XOR(d, Bitwise_XOR(c, Bitwise_LShift(a, 3)))
114     b = Bitwise_XOR(b, Bitwise_XOR(a, c))
116     c = rotr(c, 3)
118     a = rotr(a, 13)
        '<EhFooter>
        Exit Sub

irot_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.irot", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub ib0(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo ib0_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_XOR(a, d): t2 = Bitwise_XOR(c, d): t3 = Bitwise_NOT(t2)
102     t4 = Bitwise_OR(a, b): g = Bitwise_XOR(t3, t4): t6 = Bitwise_XOR(b, t1)
104     t7 = Bitwise_OR(c, t6): t8 = Bitwise_XOR(a, t7): t9 = Bitwise_AND(t2, t8)
106     f = Bitwise_XOR(t6, t9): t11 = Bitwise_NOT(t8): t12 = Bitwise_AND(b, d)
108     t13 = Bitwise_OR(f, t12): h = Bitwise_XOR(t11, t13): t15 = Bitwise_XOR(t2, t12)
110     t16 = Bitwise_OR(f, h): e = Bitwise_XOR(t15, t16)
        '<EhFooter>
        Exit Sub

ib0_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ib0", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub ib1(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo ib1_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_XOR(a, d): t2 = Bitwise_AND(a, b): t3 = Bitwise_XOR(b, c)
102     t4 = Bitwise_XOR(a, t3): t5 = Bitwise_OR(b, d): h = Bitwise_XOR(t4, t5)
104     t7 = Bitwise_OR(c, t1): t8 = Bitwise_XOR(b, t7): t9 = Bitwise_AND(t4, t8)
106     f = Bitwise_XOR(t1, t9): t11 = Bitwise_NOT(t2): t12 = Bitwise_AND(h, f)
108     t13 = Bitwise_XOR(t9, t11): g = Bitwise_XOR(t12, t13): t15 = Bitwise_AND(a, d)
110     t16 = Bitwise_XOR(c, t13): e = Bitwise_XOR(t15, t16)
        '<EhFooter>
        Exit Sub

ib1_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ib1", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub ib2(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo ib2_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_XOR(b, d): t2 = Bitwise_NOT(t1): t3 = Bitwise_XOR(a, c)
102     t4 = Bitwise_XOR(c, t1): t5 = Bitwise_AND(b, t4): e = Bitwise_XOR(t3, t5)
104     t7 = Bitwise_OR(a, t2): t8 = Bitwise_XOR(d, t7): t9 = Bitwise_OR(t3, t8)
106     h = Bitwise_XOR(t1, t9): t11 = Bitwise_NOT(t4): t12 = Bitwise_OR(e, h)
108     f = Bitwise_XOR(t11, t12): t14 = Bitwise_AND(d, t11): t15 = Bitwise_XOR(t3, t12)
110     g = Bitwise_XOR(t14, t15)
        '<EhFooter>
        Exit Sub

ib2_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ib2", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub ib3(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo ib3_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_XOR(b, c): t2 = Bitwise_OR(b, c): t3 = Bitwise_XOR(a, c)
102     t4 = Bitwise_XOR(t2, t3): t5 = Bitwise_OR(d, t4): e = Bitwise_XOR(t1, t5)
104     t7 = Bitwise_XOR(a, d): t8 = Bitwise_OR(t1, t5): t9 = Bitwise_XOR(t2, t7)
106     g = Bitwise_XOR(t8, t9): t11 = Bitwise_AND(a, t4): t12 = Bitwise_OR(e, t9)
108     f = Bitwise_XOR(t11, t12): t14 = Bitwise_AND(a, g): t15 = Bitwise_XOR(t2, t14)
110     t16 = Bitwise_AND(e, t15): h = Bitwise_XOR(t4, t16)
        '<EhFooter>
        Exit Sub

ib3_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ib3", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub ib4(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo ib4_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_XOR(c, d): t2 = Bitwise_OR(c, d): t3 = Bitwise_XOR(b, t2)
102     t4 = Bitwise_AND(a, t3): f = Bitwise_XOR(t1, t4): t6 = Bitwise_XOR(a, d)
104     t7 = Bitwise_OR(b, d): t8 = Bitwise_AND(t6, t7): h = Bitwise_XOR(t3, t8)
106     t10 = Bitwise_NOT(a): t11 = Bitwise_XOR(c, h): t12 = Bitwise_OR(t10, t11)
108     e = Bitwise_XOR(t3, t12): t14 = Bitwise_OR(c, t4): t15 = Bitwise_XOR(t7, t14)
110     t16 = Bitwise_OR(h, t10): g = Bitwise_XOR(t15, t16)
        '<EhFooter>
        Exit Sub

ib4_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ib4", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub rot(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock)
        '<EhHeader>
        On Error GoTo rot_Err
        '</EhHeader>
100     a = rotl(a, 13)
102     c = rotl(c, 3)
104     d = Bitwise_XOR(d, Bitwise_XOR(c, Bitwise_LShift(a, 3)))
106     b = Bitwise_XOR(b, Bitwise_XOR(a, c))
108     d = rotl(d, 7)
110     b = rotl(b, 1)
112     a = Bitwise_XOR(a, Bitwise_XOR(b, d))
114     c = Bitwise_XOR(c, Bitwise_XOR(d, Bitwise_LShift(b, 7)))
116     a = rotl(a, 5)
118     c = rotl(c, 22)
        '<EhFooter>
        Exit Sub

rot_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.rot", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub sb0(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo sb0_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_XOR(b, d): t2 = Bitwise_NOT(t1): t3 = Bitwise_OR(a, d)
102     t4 = Bitwise_XOR(b, c): h = Bitwise_XOR(t3, t4): t6 = Bitwise_XOR(a, b)
104     t7 = Bitwise_OR(a, t4): t8 = Bitwise_AND(c, t6): t9 = Bitwise_OR(t2, t8)
106     e = Bitwise_XOR(t7, t9): t11 = Bitwise_XOR(a, h): t12 = Bitwise_AND(t1, t6)
108     t13 = Bitwise_XOR(e, t11): f = Bitwise_XOR(t12, t13): t15 = Bitwise_OR(e, f)
110     t16 = Bitwise_AND(t3, t15): g = Bitwise_XOR(b, t16)
        '<EhFooter>
        Exit Sub

sb0_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.sb0", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub sb1(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo sb1_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_XOR(a, d): t2 = Bitwise_XOR(b, d): t3 = Bitwise_AND(a, b)
102     t4 = Bitwise_NOT(c): t5 = Bitwise_XOR(t2, t3): g = Bitwise_XOR(t4, t5)
104     t7 = Bitwise_XOR(a, t2): t8 = Bitwise_OR(b, t4): t9 = Bitwise_OR(d, g)
106     t10 = Bitwise_AND(t7, t9): f = Bitwise_XOR(t8, t10): t12 = Bitwise_XOR(c, d)
108     t13 = Bitwise_OR(t1, t2): t14 = Bitwise_XOR(f, t12): h = Bitwise_XOR(t13, t14)
110     t16 = Bitwise_OR(t1, g): t17 = Bitwise_XOR(t8, t14): e = Bitwise_XOR(t16, t17)
        '<EhFooter>
        Exit Sub

sb1_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.sb1", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub sb2(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo sb2_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_NOT(a): t2 = Bitwise_XOR(b, d): t3 = Bitwise_AND(c, t1)
102     e = Bitwise_XOR(t2, t3): t5 = Bitwise_XOR(c, t1): t6 = Bitwise_XOR(c, e)
104     t7 = Bitwise_AND(b, t6): h = Bitwise_XOR(t5, t7): t9 = Bitwise_OR(d, t7)
106     t10 = Bitwise_OR(e, t5): t11 = Bitwise_AND(t9, t10): g = Bitwise_XOR(a, t11)
108     t13 = Bitwise_OR(d, t1): t14 = Bitwise_XOR(t2, h): t15 = Bitwise_XOR(g, t13)
110     f = Bitwise_XOR(t14, t15)
        '<EhFooter>
        Exit Sub

sb2_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.sb2", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub sb3(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo sb3_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_XOR(a, c): t2 = Bitwise_OR(a, d): t3 = Bitwise_AND(a, b)
102     t4 = Bitwise_AND(a, d): t5 = Bitwise_OR(b, t4): t6 = Bitwise_AND(t1, t2)
104     f = Bitwise_XOR(t5, t6): t8 = Bitwise_XOR(b, d): t9 = Bitwise_OR(c, t3)
106     t10 = Bitwise_XOR(t6, t8): h = Bitwise_XOR(t9, t10): t12 = Bitwise_XOR(c, t3)
108     t13 = Bitwise_AND(t2, h): g = Bitwise_XOR(t12, t13): t15 = Bitwise_NOT(g)
110     t16 = Bitwise_XOR(t2, t3): t17 = Bitwise_AND(f, t15): e = Bitwise_XOR(t16, t17)
        '<EhFooter>
        Exit Sub

sb3_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.sb3", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub sb4(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo sb4_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_NOT(a): t2 = Bitwise_XOR(a, d): t3 = Bitwise_XOR(a, b)
102     t4 = Bitwise_XOR(c, t1): t5 = Bitwise_OR(t2, t3): e = Bitwise_XOR(t4, t5)
104     t7 = Bitwise_NOT(e): t8 = Bitwise_OR(b, t7): h = Bitwise_XOR(t2, t8)
106     t10 = Bitwise_AND(a, e): t11 = Bitwise_XOR(b, h): t12 = Bitwise_AND(t8, t11)
108     g = Bitwise_XOR(t10, t12): t14 = Bitwise_OR(a, t7): t15 = Bitwise_XOR(t3, t14)
110     t16 = Bitwise_AND(h, g): f = Bitwise_XOR(t15, t16)
        '<EhFooter>
        Exit Sub

sb4_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.sb4", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub sb5(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo sb5_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_NOT(a): t2 = Bitwise_XOR(a, b): t3 = Bitwise_XOR(a, d)
102     t4 = Bitwise_XOR(c, t1): t5 = Bitwise_OR(t2, t3): e = Bitwise_XOR(t4, t5)
104     t7 = Bitwise_NOT(d): t8 = Bitwise_AND(e, t7): f = Bitwise_XOR(t2, t8)
106     t10 = Bitwise_OR(b, f): t11 = Bitwise_OR(c, e): t12 = Bitwise_XOR(t7, t10)
108     h = Bitwise_XOR(t11, t12): t14 = Bitwise_OR(d, f): t15 = Bitwise_XOR(t1, t14)
110     t16 = Bitwise_OR(e, h): g = Bitwise_XOR(t15, t16)
        '<EhFooter>
        Exit Sub

sb5_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.sb5", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub sb7(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo sb7_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_NOT(c): t2 = Bitwise_XOR(b, c): t3 = Bitwise_OR(b, t1)
102     t4 = Bitwise_XOR(d, t3): t5 = Bitwise_AND(a, t4): h = Bitwise_XOR(t2, t5)
104     t7 = Bitwise_XOR(a, d): t8 = Bitwise_XOR(b, t5): t9 = Bitwise_OR(t2, t8)
106     f = Bitwise_XOR(t7, t9): t11 = Bitwise_AND(d, t3): t12 = Bitwise_XOR(t5, f)
108     t13 = Bitwise_AND(h, t12): g = Bitwise_XOR(t11, t13): t15 = Bitwise_OR(t1, t4)
110     t16 = Bitwise_XOR(t12, g): e = Bitwise_XOR(t15, t16)
        '<EhFooter>
        Exit Sub

sb7_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.sb7", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub ib6(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo ib6_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_NOT(c): t2 = Bitwise_XOR(a, c): t3 = Bitwise_XOR(b, d)
102     t4 = Bitwise_OR(a, t1): f = Bitwise_XOR(t3, t4): t6 = Bitwise_OR(a, b)
104     t7 = Bitwise_AND(b, t2): t8 = Bitwise_XOR(f, t6): t9 = Bitwise_OR(t7, t8)
106     e = Bitwise_XOR(c, t9): t11 = Bitwise_NOT(f): t12 = Bitwise_OR(d, t2)
108     t13 = Bitwise_XOR(t9, t11): h = Bitwise_XOR(t12, t13): t15 = Bitwise_XOR(b, t11)
110     t16 = Bitwise_AND(e, h): g = Bitwise_XOR(t15, t16)
        '<EhFooter>
        Exit Sub

ib6_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ib6", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub sb6(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo sb6_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_XOR(a, c): t2 = Bitwise_OR(b, d): t3 = Bitwise_XOR(b, c)
102     t4 = Bitwise_NOT(t3): t5 = Bitwise_AND(a, d): f = Bitwise_XOR(t4, t5)
104     t7 = Bitwise_OR(b, c): t8 = Bitwise_XOR(d, t1): t9 = Bitwise_AND(t7, t8)
106     h = Bitwise_XOR(t2, t9): t11 = Bitwise_AND(t1, t7): t12 = Bitwise_XOR(t4, t8)
108     t13 = Bitwise_AND(h, t11): e = Bitwise_XOR(t12, t13): t15 = Bitwise_XOR(t3, t11)
110     t16 = Bitwise_OR(h, t15): g = Bitwise_XOR(t12, t16)
        '<EhFooter>
        Exit Sub

sb6_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.sb6", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub ib7(a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock)
        '<EhHeader>
        On Error GoTo ib7_Err
        '</EhHeader>
        Dim t1 As MemBlock, t2 As MemBlock, t3 As MemBlock, t4 As MemBlock, t5 As MemBlock
        Dim t6 As MemBlock, t7 As MemBlock, t8 As MemBlock, t9 As MemBlock, t10 As MemBlock
        Dim t11 As MemBlock, t12 As MemBlock, t13 As MemBlock, t14 As MemBlock, t15 As MemBlock
        Dim t16 As MemBlock, t17 As MemBlock, t18 As MemBlock
100     t1 = Bitwise_AND(a, b): t2 = Bitwise_OR(a, b): t3 = Bitwise_OR(c, t1)
102     t4 = Bitwise_AND(d, t2): h = Bitwise_XOR(t3, t4): t6 = Bitwise_NOT(d)
104     t7 = Bitwise_XOR(b, t4): t8 = Bitwise_XOR(h, t6): t9 = Bitwise_OR(t7, t8)
106     f = Bitwise_XOR(a, t9): t11 = Bitwise_XOR(c, t7): t12 = Bitwise_OR(d, f)
108     e = Bitwise_XOR(t11, t12): t14 = Bitwise_AND(a, h): t15 = Bitwise_XOR(t3, f)
110     t16 = Bitwise_XOR(e, t14): g = Bitwise_XOR(t15, t16)
        '<EhFooter>
        Exit Sub

ib7_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.ib7", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Sub k_xor(r As Integer, a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock)
        '<EhHeader>
        On Error GoTo k_xor_Err
        '</EhHeader>
100    a = Bitwise_XOR(a, l_key(4 * r + 8))
102    b = Bitwise_XOR(b, l_key(4 * r + 9))
104    c = Bitwise_XOR(c, l_key(4 * r + 10))
106    d = Bitwise_XOR(d, l_key(4 * r + 11))
        '<EhFooter>
        Exit Sub

k_xor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.k_xor", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub
Private Function set_key(in_key() As MemBlock, key_len As Integer) As MemBlock()
        '<EhHeader>
        On Error GoTo set_key_Err
        '</EhHeader>
        Dim i As Integer, lk As Integer, a As MemBlock, b As MemBlock, c As MemBlock, d As MemBlock, e As MemBlock, f As MemBlock, g As MemBlock, h As MemBlock, RetVal(0 To 0) As MemBlock
        Dim lk1 As MemBlock, One As MemBlock, Temp1 As MemBlock, Temp2 As MemBlock

100     If key_len < 0 Or key_len > 256 Then
102         SetInt RetVal(0), 0
104         set_key = RetVal
            Exit Function
        End If

106     i = 0
108     lk = (key_len + 31) \ 32
      
110     While i < lk
            #If BLOCK_REVERSE Then
112             l_key(i) = bswap(in_key(lk - i - 1))
            #Else
114             l_key(i) = in_key(i)
            #End If
116         i = i + 1
        Wend
118     If key_len < 256 Then
120         While i < 8
122             SetInt l_key(i), 0
124             i = i + 1
            Wend
126         i = key_len \ 32
128         SetInt One, 1
130         lk1 = Bitwise_LShift(One, key_len Mod 32)
132         l_key(i) = Bitwise_OR(Bitwise_AND(l_key(i), Substr(lk1, One)), lk1)
        End If
134     For i = 0 To 131
136         lk1 = Bitwise_XOR(l_key(i), l_key(i + 3))
138         lk1 = Bitwise_XOR(lk1, l_key(i + 5))
140         lk1 = Bitwise_XOR(lk1, l_key(i + 7))
142         SetData Temp1, 158, 55, 121, 185
144         SetInt Temp2, CByte(i)
146         lk1 = Bitwise_XOR(lk1, Temp1)
148         lk1 = Bitwise_XOR(lk1, Temp2)
150         l_key(i + 8) = Bitwise_OR(Bitwise_LShift(lk1, 11), Bitwise_RShift(lk1, 21))
        Next
    
152     k_set 0, a, b, c, d: sb3 a, b, c, d, e, f, g, h: k_get 0, e, f, g, h
154     k_set 1, a, b, c, d: sb2 a, b, c, d, e, f, g, h: k_get 1, e, f, g, h
156     k_set 2, a, b, c, d: sb1 a, b, c, d, e, f, g, h: k_get 2, e, f, g, h
158     k_set 3, a, b, c, d: sb0 a, b, c, d, e, f, g, h: k_get 3, e, f, g, h
160     k_set 4, a, b, c, d: sb7 a, b, c, d, e, f, g, h: k_get 4, e, f, g, h
162     k_set 5, a, b, c, d: sb6 a, b, c, d, e, f, g, h: k_get 5, e, f, g, h
164     k_set 6, a, b, c, d: sb5 a, b, c, d, e, f, g, h: k_get 6, e, f, g, h
166     k_set 7, a, b, c, d: sb4 a, b, c, d, e, f, g, h: k_get 7, e, f, g, h
168     k_set 8, a, b, c, d: sb3 a, b, c, d, e, f, g, h: k_get 8, e, f, g, h
170     k_set 9, a, b, c, d: sb2 a, b, c, d, e, f, g, h: k_get 9, e, f, g, h
172     k_set 10, a, b, c, d: sb1 a, b, c, d, e, f, g, h: k_get 10, e, f, g, h
174     k_set 11, a, b, c, d: sb0 a, b, c, d, e, f, g, h: k_get 11, e, f, g, h
176     k_set 12, a, b, c, d: sb7 a, b, c, d, e, f, g, h: k_get 12, e, f, g, h
178     k_set 13, a, b, c, d: sb6 a, b, c, d, e, f, g, h: k_get 13, e, f, g, h
180     k_set 14, a, b, c, d: sb5 a, b, c, d, e, f, g, h: k_get 14, e, f, g, h
182     k_set 15, a, b, c, d: sb4 a, b, c, d, e, f, g, h: k_get 15, e, f, g, h
184     k_set 16, a, b, c, d: sb3 a, b, c, d, e, f, g, h: k_get 16, e, f, g, h
186     k_set 17, a, b, c, d: sb2 a, b, c, d, e, f, g, h: k_get 17, e, f, g, h
188     k_set 18, a, b, c, d: sb1 a, b, c, d, e, f, g, h: k_get 18, e, f, g, h
190     k_set 19, a, b, c, d: sb0 a, b, c, d, e, f, g, h: k_get 19, e, f, g, h
192     k_set 20, a, b, c, d: sb7 a, b, c, d, e, f, g, h: k_get 20, e, f, g, h
194     k_set 21, a, b, c, d: sb6 a, b, c, d, e, f, g, h: k_get 21, e, f, g, h
196     k_set 22, a, b, c, d: sb5 a, b, c, d, e, f, g, h: k_get 22, e, f, g, h
198     k_set 23, a, b, c, d: sb4 a, b, c, d, e, f, g, h: k_get 23, e, f, g, h
200     k_set 24, a, b, c, d: sb3 a, b, c, d, e, f, g, h: k_get 24, e, f, g, h
202     k_set 25, a, b, c, d: sb2 a, b, c, d, e, f, g, h: k_get 25, e, f, g, h
204     k_set 26, a, b, c, d: sb1 a, b, c, d, e, f, g, h: k_get 26, e, f, g, h
206     k_set 27, a, b, c, d: sb0 a, b, c, d, e, f, g, h: k_get 27, e, f, g, h
208     k_set 28, a, b, c, d: sb7 a, b, c, d, e, f, g, h: k_get 28, e, f, g, h
210     k_set 29, a, b, c, d: sb6 a, b, c, d, e, f, g, h: k_get 29, e, f, g, h
212     k_set 30, a, b, c, d: sb5 a, b, c, d, e, f, g, h: k_get 30, e, f, g, h
214     k_set 31, a, b, c, d: sb4 a, b, c, d, e, f, g, h: k_get 31, e, f, g, h
216     k_set 32, a, b, c, d: sb3 a, b, c, d, e, f, g, h: k_get 32, e, f, g, h
218     set_key = l_key
        '<EhFooter>
        Exit Function

set_key_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.set_key", _
                  "clsSerpent component failure"
        '</EhFooter>
End Function
Private Sub SetKey(Key As String)
        '<EhHeader>
        On Error GoTo SetKey_Err
        '</EhHeader>
        Dim kblk(0 To 7) As MemBlock, KeyBuffer() As Byte
100     ReDim KeyBuffer(0 To 31) As Byte
102     SetBuffer KeyBuffer, Key, 32
104     SetBlock kblk, KeyBuffer, 0, 32
106     set_key kblk, 32 * 8
        '<EhFooter>
        Exit Sub

SetKey_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSerpent.SetKey", _
                  "clsSerpent component failure"
        '</EhFooter>
End Sub

