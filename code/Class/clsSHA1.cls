VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSHA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



'Dim EncryptionUtils: Set EncryptionUtils = New Utils_EncryptionUtils

''
' MD5
'   MD5 is one of a series of message digest algorithms.
'   It is a one way encryption, and although not the securest form of
'   encrytion it can be used for simple encrytion to pass unreadable data
'   into the database to store passwords and other second security rated
'   information. This should not be used to store credit card details or
'   other personal information.
'
'   Derived from the RSA Data Security, Inc. MD5 Message-Digest Algorithm,
'   as set out in the memo RFC1321.
'   MD5 algorithm is one of the industry standard methods for generating digital
'   signatures. It is generically known as a digest, digital signature, one-way
'   encryption, hash or checksum algorithm. A common use for MD5 is for password
'   encryption as it is one-way in nature, that does not mean that your passwords
'   are not free from a dictionary attack.
'
'
' RC4
'   RC4 generates a pseudorandom stream of bits (a "keystream") which, for
'   encryption, is combined with the plaintext using XOR as with any
'   Vernam cipher; decryption is performed the same way. To generate the
'   keystream, the cipher makes use of a secret internal state which
'   consists of two parts:
'       A permutation of all 256 possible bytes (denoted "cipher" below).
'       Two 8-bit index-pointers (denoted "i" and "j").
'
'
' SHA1
'   The SHA (Secure Hash Algorithm) family is a set of related cryptographic hash functions.
'   The most commonly used function in the family, SHA-1, is employed in a large variety of
'   popular security applications and protocols, including TLS, SSL, PGP, SSH, S/MIME, and
'   IPSec. SHA-1 was considered to be the successor to MD5, an earlier, widely-used hash function.
'   The encryption is one way, and is a slightly stronger variation from MD5, although
'   it requires a little more processing time, this could be used for small series of calculations
'
'

    ''''''''''''''' Start SHA1 section
Function AndW(ByRef pBytWord1Ary, ByRef pBytWord2Ary)
        '<EhHeader>
        On Error GoTo AndW_Err
        '</EhHeader>
        Dim lBytWordAry(3)
        Dim lLngIndex

100     For lLngIndex = 0 To 3
102         lBytWordAry(lLngIndex) = CByte(pBytWord1Ary(lLngIndex) And pBytWord2Ary(lLngIndex))
        Next

104     AndW = lBytWordAry
        '<EhFooter>
        Exit Function

AndW_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.AndW", _
                  "clsSHA1 component failure"
        '</EhFooter>
End Function

    Function OrW(ByRef pBytWord1Ary, ByRef pBytWord2Ary)
        '<EhHeader>
        On Error GoTo OrW_Err
        '</EhHeader>
            Dim lBytWordAry(3)
            Dim lLngIndex

100         For lLngIndex = 0 To 3
102             lBytWordAry(lLngIndex) = CByte(pBytWord1Ary(lLngIndex) Or pBytWord2Ary(lLngIndex))
            Next
104         OrW = lBytWordAry
        '<EhFooter>
        Exit Function

OrW_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.OrW", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function XorW(ByRef pBytWord1Ary, ByRef pBytWord2Ary)
        '<EhHeader>
        On Error GoTo XorW_Err
        '</EhHeader>
            Dim lBytWordAry(3)
            Dim lLngIndex

100         For lLngIndex = 0 To 3
102             lBytWordAry(lLngIndex) = CByte(pBytWord1Ary(lLngIndex) Xor pBytWord2Ary(lLngIndex))
            Next
104         XorW = lBytWordAry
        '<EhFooter>
        Exit Function

XorW_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.XorW", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function NotW(ByRef pBytWordAry)
        '<EhHeader>
        On Error GoTo NotW_Err
        '</EhHeader>
            Dim lBytWordAry(3)
            Dim lLngIndex

100         For lLngIndex = 0 To 3
102             lBytWordAry(lLngIndex) = Not CByte(pBytWordAry(lLngIndex))
            Next
104         NotW = lBytWordAry
        '<EhFooter>
        Exit Function

NotW_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.NotW", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function AddW(ByRef pBytWord1Ary, ByRef pBytWord2Ary)
        '<EhHeader>
        On Error GoTo AddW_Err
        '</EhHeader>
            Dim lLngIndex
            Dim lIntTotal
            Dim lBytWordAry(3)

100         For lLngIndex = 3 To 0 Step -1
102             If lLngIndex = 3 Then
104                 lIntTotal = CInt(pBytWord1Ary(lLngIndex)) + pBytWord2Ary(lLngIndex)
106                 lBytWordAry(lLngIndex) = lIntTotal Mod 256
                Else
108                 lIntTotal = CInt(pBytWord1Ary(lLngIndex)) + pBytWord2Ary(lLngIndex) + (lIntTotal \ 256)
110                 lBytWordAry(lLngIndex) = lIntTotal Mod 256
                End If
            Next
112         AddW = lBytWordAry
        '<EhFooter>
        Exit Function

AddW_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.AddW", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function CircShiftLeftW(ByRef pBytWordAry, ByRef pLngShift)
        '<EhHeader>
        On Error GoTo CircShiftLeftW_Err
        '</EhHeader>
            Dim lDbl1
            Dim lDbl2

100         lDbl1 = WordToDouble(pBytWordAry)
102         lDbl2 = lDbl1
104         lDbl1 = CDbl(lDbl1 * (2 ^ pLngShift))
106         lDbl2 = CDbl(lDbl2 / (2 ^ (32 - pLngShift)))
108         CircShiftLeftW = OrW(DoubleToWord(lDbl1), DoubleToWord(lDbl2))
        '<EhFooter>
        Exit Function

CircShiftLeftW_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.CircShiftLeftW", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function


    Function WordToHex(ByRef pBytWordAry)
        '<EhHeader>
        On Error GoTo WordToHex_Err
        '</EhHeader>
            Dim lLngIndex

100         For lLngIndex = 0 To 3
102             WordToHex = WordToHex & Right("0" & Hex(pBytWordAry(lLngIndex)), 2)
            Next
        '<EhFooter>
        Exit Function

WordToHex_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.WordToHex", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function HexToWord(ByRef pStrHex)
        '<EhHeader>
        On Error GoTo HexToWord_Err
        '</EhHeader>
100         HexToWord = DoubleToWord(CDbl("&h" & pStrHex)) ' needs "#" at end for VB?
        '<EhFooter>
        Exit Function

HexToWord_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.HexToWord", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function DoubleToWord(ByRef pDblValue)
        '<EhHeader>
        On Error GoTo DoubleToWord_Err
        '</EhHeader>
            Dim lBytWordAry(3)

100         lBytWordAry(0) = Int(DMod(pDblValue, 2 ^ 32) / (2 ^ 24))
102         lBytWordAry(1) = Int(DMod(pDblValue, 2 ^ 24) / (2 ^ 16))
104         lBytWordAry(2) = Int(DMod(pDblValue, 2 ^ 16) / (2 ^ 8))
106         lBytWordAry(3) = Int(DMod(pDblValue, 2 ^ 8))
108         DoubleToWord = lBytWordAry
        '<EhFooter>
        Exit Function

DoubleToWord_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.DoubleToWord", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function WordToDouble(ByRef pBytWordAry)
        '<EhHeader>
        On Error GoTo WordToDouble_Err
        '</EhHeader>
100         WordToDouble = CDbl((pBytWordAry(0) * (2 ^ 24)) + (pBytWordAry(1) * (2 ^ 16)) + (pBytWordAry(2) * (2 ^ 8)) + pBytWordAry(3))
        '<EhFooter>
        Exit Function

WordToDouble_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.WordToDouble", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function DMod(ByRef pDblValue, ByRef pDblDivisor)
        '<EhHeader>
        On Error GoTo DMod_Err
        '</EhHeader>
            Dim lDblMod

100         lDblMod = CDbl(CDbl(pDblValue) - (Int(CDbl(pDblValue) / CDbl(pDblDivisor)) * CDbl(pDblDivisor)))
102         If lDblMod < 0 Then
104             lDblMod = CDbl(lDblMod + pDblDivisor)
            End If
106         DMod = lDblMod
        '<EhFooter>
        Exit Function

DMod_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.DMod", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function f(ByRef lIntT, ByRef pBytWordBAry, ByRef pBytWordCAry, ByRef pBytWordDAry)
        '<EhHeader>
        On Error GoTo f_Err
        '</EhHeader>

100         If lIntT <= 19 Then
102             f = OrW(AndW(pBytWordBAry, pBytWordCAry), AndW((NotW(pBytWordBAry)), pBytWordDAry))
104         ElseIf lIntT <= 39 Then
106             f = XorW(XorW(pBytWordBAry, pBytWordCAry), pBytWordDAry)
108         ElseIf lIntT <= 59 Then
110             f = OrW(OrW(AndW(pBytWordBAry, pBytWordCAry), AndW(pBytWordBAry, pBytWordDAry)), AndW(pBytWordCAry, pBytWordDAry))
            Else
112             f = XorW(XorW(pBytWordBAry, pBytWordCAry), pBytWordDAry)
            End If
        '<EhFooter>
        Exit Function

f_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.f", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

    Function SHA1(pStrMessage)
        '<EhHeader>
        On Error GoTo SHA1_Err
        '</EhHeader>

            Dim lLngLen, lBytLenW, lLngTempWordWAry, lLngNumBlocks, lLngBlock, lIntT, lBytTempAry
            Dim lVarWordWAry(79), lVarWordKAry(3)
            Dim lStrBlockText, lStrWordText, lStrPadMessage
            Dim lBytWordH0Ary, lBytWordH1Ary, lBytWordH2Ary, lBytWordH3Ary, lBytWordH4Ary
            Dim lBytWordAAry, lBytWordBAry, lBytWordCAry, lBytWordDAry, lBytWordEAry, lBytWordFAry

100         lLngLen = Len(CStr(pStrMessage))

102         lBytLenW = DoubleToWord(CDbl(lLngLen) * 8)
104         lStrPadMessage = pStrMessage & Chr(128) & String((128 - (lLngLen Mod 64) - 9) Mod 64, Chr(0)) & String(4, Chr(0)) & Chr(lBytLenW(0)) & Chr(lBytLenW(1)) & Chr(lBytLenW(2)) & Chr(lBytLenW(3))
106         lLngNumBlocks = Len(lStrPadMessage) / 64

108         lVarWordKAry(0) = HexToWord("5A827999")
110         lVarWordKAry(1) = HexToWord("6ED9EBA1")
112         lVarWordKAry(2) = HexToWord("8F1BBCDC")
114         lVarWordKAry(3) = HexToWord("CA62C1D6")
116         lBytWordH0Ary = HexToWord("67452301")
118         lBytWordH1Ary = HexToWord("EFCDAB89")
120         lBytWordH2Ary = HexToWord("98BADCFE")
122         lBytWordH3Ary = HexToWord("10325476")
124         lBytWordH4Ary = HexToWord("C3D2E1F0")

126         For lLngBlock = 0 To lLngNumBlocks - 1

128             lStrBlockText = Mid(lStrPadMessage, (lLngBlock * 64) + 1, 64)
130             For lIntT = 0 To 15
132                 lStrWordText = Mid(lStrBlockText, (lIntT * 4) + 1, 4)
134                 lVarWordWAry(lIntT) = Array(Asc(Mid(lStrWordText, 1, 1)), Asc(Mid(lStrWordText, 2, 1)), Asc(Mid(lStrWordText, 3, 1)), Asc(Mid(lStrWordText, 4, 1)))

                Next

136             For lIntT = 16 To 79
138                 lVarWordWAry(lIntT) = CircShiftLeftW(XorW(XorW(XorW(lVarWordWAry(lIntT - 3), lVarWordWAry(lIntT - 8)), lVarWordWAry(lIntT - 14)), lVarWordWAry(lIntT - 16)), 1)
                Next

140             lBytWordAAry = lBytWordH0Ary
142             lBytWordBAry = lBytWordH1Ary
144             lBytWordCAry = lBytWordH2Ary
146             lBytWordDAry = lBytWordH3Ary
148             lBytWordEAry = lBytWordH4Ary

150             For lIntT = 0 To 79
152                 lBytWordFAry = f(lIntT, lBytWordBAry, lBytWordCAry, lBytWordDAry)
154                 lBytTempAry = AddW(AddW(AddW(AddW(CircShiftLeftW(lBytWordAAry, 5), lBytWordFAry), lBytWordEAry), lVarWordWAry(lIntT)), lVarWordKAry(lIntT \ 20))
156                 lBytWordEAry = lBytWordDAry
158                 lBytWordDAry = lBytWordCAry
160                 lBytWordCAry = CircShiftLeftW(lBytWordBAry, 30)
162                 lBytWordBAry = lBytWordAAry
164                 lBytWordAAry = lBytTempAry
                Next

166             lBytWordH0Ary = AddW(lBytWordH0Ary, lBytWordAAry)
168             lBytWordH1Ary = AddW(lBytWordH1Ary, lBytWordBAry)
170             lBytWordH2Ary = AddW(lBytWordH2Ary, lBytWordCAry)
172             lBytWordH3Ary = AddW(lBytWordH3Ary, lBytWordDAry)
174             lBytWordH4Ary = AddW(lBytWordH4Ary, lBytWordEAry)

            Next

176         SHA1 = WordToHex(lBytWordH0Ary) & WordToHex(lBytWordH1Ary) & WordToHex(lBytWordH2Ary) & WordToHex(lBytWordH3Ary) & WordToHex(lBytWordH4Ary)

        '<EhFooter>
        Exit Function

SHA1_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsSHA1.SHA1", _
                  "clsSHA1 component failure"
        '</EhFooter>
    End Function

