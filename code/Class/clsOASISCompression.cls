VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OASISCompression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Const ENCRYPTPASSWORD As String = "iamnottellingyou"
Dim oRijndael As New CRijndael
Dim objZlib As New clsZlibWrapper

Public Function CompressStringToByteArray(sTextPassed As Variant) As Byte()
        '<EhHeader>
        On Error GoTo CompressStringToByteArray_Err
        '</EhHeader>
        Dim sText As String
        Dim bByteText() As Byte
        Dim bBytePass() As Byte
        
100     sText = CStr(sTextPassed)
102     bByteText = ConvertStringToByteArray(sText)
104     bBytePass = ConvertStringToByteArray(ENCRYPTPASSWORD)

106     objZlib.CompressByteArray bByteText, 9
108     sText = ConvertByteArrayToString(bByteText)
110     sText = "(" & objZlib.OriginalSize & ")" & sText
112     bByteText = ConvertStringToByteArray(sText)
114     bByteText = oRijndael.EncryptData(bByteText, bBytePass)

        objZlib.CompressByteArray bByteText, 1
        sText = ConvertByteArrayToString(bByteText)
        sText = "(" & objZlib.OriginalSize & ")" & sText
        
116     CompressStringToByteArray = ConvertStringToByteArray(sText)

        '<EhFooter>
        Exit Function

CompressStringToByteArray_Err:
        MsgBox "OASISCompression.CompressStringToByteArray_Err (" & Erl & "): " & Err.Description
        '</EhFooter>
End Function

Public Function DecompressStringToString(sTextPassed As Variant) As String
        '<EhHeader>
        On Error GoTo DecompressStringToString_Err
        '</EhHeader>
        Dim sText As String
        Dim lSize As Long
        Dim ByteText() As Byte
        Dim BytePass() As Byte
        
100     If Len(sTextPassed) > 0 Then
        
104         BytePass = ConvertStringToByteArray(ENCRYPTPASSWORD)
            
            If Len(sTextPassed) > 2 And Left$(sTextPassed, 1) = "(" Then
            
                lSize = CLng(Mid(sTextPassed, 2, InStr(1, sTextPassed, ")") - 2))
                sText = Right$(sTextPassed, Len(sTextPassed) - 2 - Len(CStr(lSize)))
                ByteText = ConvertStringToByteArray(sText)
                objZlib.DecompressByteArray ByteText, lSize

106             ByteText = oRijndael.DecryptData(ByteText, BytePass)
108             sText = ConvertByteArrayToString(ByteText)

110             If Left(sText, Len("OASISStringCompression.CRijndael.DecryptData_Err")) = "OASISStringCompression.CRijndael.DecryptData_Err" Then
112                 DecompressStringToString = "Decryption failed (OASISStringCompression.CRijndael.DecryptData_Err) for sTextPassed: '" & sTextPassed & "'"
                ElseIf Len(sText) > 2 And Left$(sText, 1) = "(" Then
116                 lSize = CLng(Mid(sText, 2, InStr(1, sText, ")") - 2))
118                 sText = Right$(sText, Len(sText) - 2 - Len(CStr(lSize)))
120                 ByteText = ConvertStringToByteArray(sText)
122                 objZlib.DecompressByteArray ByteText, lSize
124                 DecompressStringToString = ConvertByteArrayToString(ByteText)
                Else
126                 DecompressStringToString = "Decryption failed"
                End If
            
            Else
                DecompressStringToString = "Decryption failed"
            End If

        Else
128         DecompressStringToString = "Decryption failed (zero length string passed)"
        End If

        '<EhFooter>
        Exit Function

DecompressStringToString_Err:
        MsgBox "OASISCompression.DecompressStringToString_Err (" & Erl & "): " & Err.Description
        MsgBox "Decryption failed for sTextPassed [" & sTextPassed & "]"
        '</EhFooter>
End Function

Public Function ConvertStringToByteArray(sString As Variant) As Byte()
    ConvertStringToByteArray = StrConv(CStr(sString), vbFromUnicode)
End Function

Public Function ConvertByteArrayToString(bytArray() As Byte) As String
    ConvertByteArrayToString = StrConv(bytArray, vbUnicode)
End Function

Private Sub Class_Terminate()
    Set qlz = Nothing
    Set clEncryption = Nothing
End Sub
