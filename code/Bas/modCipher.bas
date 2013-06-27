Attribute VB_Name = "modCipher"
Option Explicit
' Utilize ebCrypt.dll

Public Enum Algorithms
    BLOWFISH
    IDEA
    TRIPLEDES
    DES
    DESE
    CAST5
    SERPENT128
    SERPENT192
    SERPENT256
    RIJNDAEL128
    RIJNDAEL192
    RIJNDAEL256
    RC4
    TWOFISH
End Enum
Public Enum HashAlgorithms
    MD2
    MD5
    RipeMD160
    SHA1
End Enum

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
Private Initialized As Boolean

Public Function Decode64(sInput As String) As String
        '<EhHeader>
        On Error GoTo Decode64_Err
        '</EhHeader>

100     If sInput = "" Then Exit Function
102     Decode64 = StrConv(DecodeArray64(sInput), vbUnicode)
        '<EhFooter>
        Exit Function

Decode64_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.Decode64 " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DecodeArray64(sInput As String) As Byte()
        '<EhHeader>
        On Error GoTo DecodeArray64_Err
        '</EhHeader>

100     If m_bytReverseIndex(47) <> 63 Then Initialize
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.DecodeArray64 " & _
               "at line " & Erl
        Resume Next
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.Encode64 " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function EncodeArray64(ByRef bytInput() As Byte) As String
        '<EhHeader>
        On Error GoTo EncodeArray64_Err
        '</EhHeader>
        On Error GoTo ErrorHandler

100     If m_bytReverseIndex(47) <> 63 Then Initialize
    
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.EncodeArray64 " & _
               "at line " & Erl
        Resume Next
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
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.FileExist " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function EncryptFile(Which As Algorithms, _
                            InFile As String, _
                            OutFile As String, _
                            Overwrite As Boolean, _
                            Optional OutputIn64 As Boolean, _
                            Optional Key As String, _
                            Optional Salt As String) As Boolean
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

118     If FileExist(OutFile) = True Then Kill OutFile
120     FileO = FreeFile
122     Buffer() = EncryptArray(Which, Buffer(), Key, Salt)

124     Open OutFile For Binary As #FileO

126     If OutputIn64 = True Then
128         Put #FileO, , EncodeArray64(Buffer())
        Else
130         Put #FileO, , Buffer()
        End If

132     Close #FileO
134     EncryptFile = True
136     Erase Buffer(): InFile = "": OutFile = "": Key = "": Salt = ""
        Exit Function

ErrorHandler:
138     EncryptFile = False
140     Erase Buffer(): InFile = "": OutFile = "": Key = "": Salt = ""
        '<EhFooter>
        Exit Function

EncryptFile_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.EncryptFile " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DecryptFile(Which As Algorithms, _
                            InFile As String, _
                            OutFile As String, _
                            Overwrite As Boolean, _
                            Optional IsFileIn64 As Boolean, _
                            Optional Key As String, _
                            Optional Salt As String) As Boolean
        '<EhHeader>
        On Error GoTo DecryptFile_Err
        '</EhHeader>
        On Error GoTo ErrorHandler

100     If FileExist(InFile) = False Then
102         DecryptFile = False
            Exit Function
        End If

104     If FileExist(OutFile) = True Then
106         DecryptFile = False
            Exit Function
        End If

        Dim FileO As Integer, Buffer() As Byte
108     FileO = FreeFile
110     Open InFile For Binary As #FileO
112     ReDim Buffer(0 To LOF(FileO) - 1)
114     Get #FileO, , Buffer()
116     Close #FileO

118     If FileExist(OutFile) = True Then Kill OutFile
120     FileO = FreeFile

122     If IsFileIn64 = True Then Buffer() = DecodeArray64(StrConv(Buffer(), vbUnicode))
124     Buffer() = DecryptArray(Which, Buffer(), Key, Salt)
126     Open OutFile For Binary As #FileO
128     Put #FileO, , Buffer()
130     Close #FileO
132     DecryptFile = True
134     Erase Buffer(): InFile = "": OutFile = "": Key = "": Salt = ""
        Exit Function

ErrorHandler:
136     DecryptFile = False
138     Erase Buffer(): InFile = "": OutFile = "": Key = "": Salt = ""
        '<EhFooter>
        Exit Function

DecryptFile_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.DecryptFile " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function Hash(Which As HashAlgorithms, _
                     Message As String) As String
        '<EhHeader>
        On Error GoTo Hash_Err
        '</EhHeader>

100     If Message = "" Then Exit Function
        Dim hsh As ebcryptlib.eb_c_Hash
102     Set hsh = CreateObject("EbCrypt.eb_c_Hash")

104     If Which = MD2 Then Hash = hsh.HashString(EB_CRYPT_HASH_ALGORITHM_MD2, Message)
106     If Which = MD5 Then Hash = hsh.HashString(EB_CRYPT_HASH_ALGORITHM_MD5, Message)
108     If Which = SHA1 Then Hash = hsh.HashString(EB_CRYPT_HASH_ALGORITHM_SHA1, Message)
110     If Which = RipeMD160 Then Hash = hsh.HashString(EB_CRYPT_HASH_ALGORITHM_RIPEMD160, Message)
        '<EhFooter>
        Exit Function

Hash_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.Hash " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function EncryptString(Which As Algorithms, _
                              Text As String, _
                              Optional OutputIn64 As Boolean, _
                              Optional Key As String, _
                              Optional Salt As String) As String
        '<EhHeader>
        On Error GoTo EncryptString_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
        Dim cipher As ebcryptlib.eb_c_Cipher
100     Set cipher = New ebcryptlib.eb_c_Cipher

102     If Len(Salt) < 8 Then Salt = Salt & Space$(8 - Len(Salt))
104     If Which = BLOWFISH Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_BLOWFISH_OFB, Key, Salt, Text), vbUnicode)
106     If Which = CAST5 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_CAST5_OFB, Key, Salt, Text), vbUnicode)
108     If Which = DES Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_DES_OFB, Key, Salt, Text), vbUnicode)
110     If Which = DESE Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_DESE_OFB, Key, Salt, Text), vbUnicode)
112     If Which = TRIPLEDES Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_DES3_OFB, Key, Salt, Text), vbUnicode)
114     If Which = IDEA Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_IDEA_OFB, Key, Salt, Text), vbUnicode)
116     If Which = RC4 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_RC4, Key, Salt, Text), vbUnicode)
118     If Which = RIJNDAEL128 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_128, Key, Salt, Text), vbUnicode)
120     If Which = RIJNDAEL192 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_192, Key, Salt, Text), vbUnicode)
122     If Which = RIJNDAEL256 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, Key, Salt, Text), vbUnicode)
124     If Which = SERPENT128 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_128, Key, Salt, Text), vbUnicode)
126     If Which = SERPENT192 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_192, Key, Salt, Text), vbUnicode)
128     If Which = SERPENT256 Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, Key, Salt, Text), vbUnicode)
130     If Which = TWOFISH Then EncryptString = StrConv(cipher.EncryptString(EB_CRYPT_CIPHER_ALGORITHM_TWOFISH_CBC, Key, Salt, Text), vbUnicode)
132     If OutputIn64 = True Then EncryptString = Encode64(EncryptString)
134     Key = "": Salt = "": Text = ""
        Exit Function

ErrorHandler:
136     MsgBox Err.Description
138     Key = "": Salt = "": Text = ""
        '<EhFooter>
        Exit Function

EncryptString_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.EncryptString " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function EncryptArray(Which As Algorithms, _
                             InputArray() As Byte, _
                             Optional Key As String, _
                             Optional Salt As String) As Variant
        '<EhHeader>
        On Error GoTo EncryptArray_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
        Dim cipher As ebcryptlib.eb_c_Cipher
100     Set cipher = New ebcryptlib.eb_c_Cipher

102     If Len(Salt) < 8 Then Salt = Salt & Space$(8 - Len(Salt))
    
104     If Which = BLOWFISH Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_BLOWFISH_OFB, Key, Salt, InputArray())
106     If Which = CAST5 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_CAST5_OFB, Key, Salt, InputArray())
108     If Which = DES Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DES_OFB, Key, Salt, InputArray())
110     If Which = DESE Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DESE_OFB, Key, Salt, InputArray())
112     If Which = TRIPLEDES Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DES3_OFB, Key, Salt, InputArray())
114     If Which = IDEA Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_IDEA_OFB, Key, Salt, InputArray())
116     If Which = RC4 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RC4, Key, Salt, InputArray())
118     If Which = RIJNDAEL128 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_128, Key, Salt, InputArray())
120     If Which = RIJNDAEL192 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_192, Key, Salt, InputArray())
122     If Which = RIJNDAEL256 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, Key, Salt, InputArray())
124     If Which = SERPENT128 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_128, Key, Salt, InputArray())
126     If Which = SERPENT192 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_192, Key, Salt, InputArray())
128     If Which = SERPENT256 Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, Key, Salt, InputArray())
130     If Which = TWOFISH Then EncryptArray = cipher.EncryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_TWOFISH_CBC, Key, Salt, InputArray())
132     Erase InputArray(): Key = "": Salt = ""
        Exit Function

ErrorHandler:
134     MsgBox Err.Description
136     Erase InputArray(): Key = "": Salt = ""
        '<EhFooter>
        Exit Function

EncryptArray_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.EncryptArray " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DecryptArray(Which As Algorithms, _
                             InputArray() As Byte, _
                             Optional Key As String, _
                             Optional Salt As String) As Variant
        '<EhHeader>
        On Error GoTo DecryptArray_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
        Dim cipher As ebcryptlib.eb_c_Cipher
100     Set cipher = New ebcryptlib.eb_c_Cipher

102     If Len(Salt) < 8 Then Salt = Salt & Space$(8 - Len(Salt))
    
104     If Which = BLOWFISH Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_BLOWFISH_OFB, Key, Salt, InputArray())
106     If Which = CAST5 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_CAST5_OFB, Key, Salt, InputArray())
108     If Which = DES Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DES_OFB, Key, Salt, InputArray())
110     If Which = DESE Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DESE_OFB, Key, Salt, InputArray())
112     If Which = TRIPLEDES Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_DES3_OFB, Key, Salt, InputArray())
114     If Which = IDEA Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_IDEA_OFB, Key, Salt, InputArray())
116     If Which = RC4 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RC4, Key, Salt, InputArray())
118     If Which = RIJNDAEL128 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_128, Key, Salt, InputArray())
120     If Which = RIJNDAEL192 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_192, Key, Salt, InputArray())
122     If Which = RIJNDAEL256 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, Key, Salt, InputArray())
124     If Which = SERPENT128 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_128, Key, Salt, InputArray())
126     If Which = SERPENT192 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_192, Key, Salt, InputArray())
128     If Which = SERPENT256 Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, Key, Salt, InputArray())
130     If Which = TWOFISH Then DecryptArray = cipher.DecryptBLOB(EB_CRYPT_CIPHER_ALGORITHM_TWOFISH_CBC, Key, Salt, InputArray())
132     Erase InputArray(): Key = "": Salt = ""
        Exit Function

ErrorHandler:
134     MsgBox Err.Description
136     Erase InputArray(): Key = "": Salt = ""
        '<EhFooter>
        Exit Function

DecryptArray_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.DecryptArray " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DecryptString(Which As Algorithms, _
                              CipherText As String, _
                              Optional IsTextIn64 As Boolean, _
                              Optional Key As String, _
                              Optional Salt As String) As String
        '<EhHeader>
        On Error GoTo DecryptString_Err
        '</EhHeader>
        On Error GoTo ErrorHandler
        Dim cipher As ebcryptlib.eb_c_Cipher, BArray() As Byte
100     Set cipher = New ebcryptlib.eb_c_Cipher

102     If IsTextIn64 = True Then CipherText = Decode64(CipherText)
104     If Len(Salt) < 8 Then Salt = Salt & Space$(8 - Len(Salt))
106     BArray() = StrConv(CipherText, vbFromUnicode)

108     If Which = BLOWFISH Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_BLOWFISH_OFB, Key, Salt, BArray())
110     If Which = CAST5 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_CAST5_OFB, Key, Salt, BArray())
112     If Which = DES Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_DES_OFB, Key, Salt, BArray())
114     If Which = DESE Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_DESE_OFB, Key, Salt, BArray())
116     If Which = IDEA Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_IDEA_OFB, Key, Salt, BArray())
118     If Which = TRIPLEDES Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_DES3_OFB, Key, Salt, BArray())
120     If Which = TWOFISH Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_TWOFISH_CBC, Key, Salt, BArray())
122     If Which = RC4 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_RC4, Key, Salt, BArray())
124     If Which = RIJNDAEL128 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_128, Key, Salt, BArray())
126     If Which = RIJNDAEL192 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_192, Key, Salt, BArray())
128     If Which = RIJNDAEL256 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_RIJNDAEL_CBC_256, Key, Salt, BArray())
130     If Which = SERPENT128 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_128, Key, Salt, BArray())
132     If Which = SERPENT192 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_192, Key, Salt, BArray())
134     If Which = SERPENT256 Then DecryptString = cipher.DecryptString(EB_CRYPT_CIPHER_ALGORITHM_SERPENT_CBC_256, Key, Salt, BArray())
    
136     Key = "": Salt = "": CipherText = ""
        Exit Function
    
ErrorHandler:
138     MsgBox Err.Description
140     Key = "": Salt = "": CipherText = ""
        '<EhFooter>
        Exit Function

DecryptString_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.DecryptString " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub Initialize()
        '<EhHeader>
        On Error GoTo Initialize_Err
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

Initialize_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modCipher.Initialize " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
