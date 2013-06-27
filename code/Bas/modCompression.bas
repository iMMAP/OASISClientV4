Attribute VB_Name = "modCompression"
Private Declare Function qlz_compress _
                Lib "quick32.dll" (ByRef Source As Byte, _
                                   ByRef Destination As Byte, _
                                   ByVal Length As Long) As Long
Private Declare Function qlz_decompress _
                Lib "quick32.dll" (ByRef Source As Byte, _
                                   ByRef Destination As Byte) As Long
Private Declare Function qlz_size_decompressed _
                Lib "quick32.dll" (ByRef Source As Byte) As Long
Private Declare Function qlz_size_source _
                Lib "quick32.dll" (ByRef Source As Byte) As Long

' If the Visual Basic IDE cannot find quick32.dll even though it's in the system32 directory, try adding a path
' to the quick32.dll file name in the declarations. This should never be neccessary though.

Public Function CompressString(sText As String) As String
    
    Dim Source() As Byte
    Source = ConvertStringToByteArray(sText)
    
    Dim dst() As Byte
    Dim r As Long
    ReDim dst(0 To UBound(Source) * 1.2 + 36000)
    r = qlz_compress(Source(0), dst(0), UBound(Source) + 1)
    ReDim Preserve dst(0 To r - 1)
    CompressString = ConvertByteArrayToString(dst)

End Function

Public Function GetSize(Source() As Byte) As Long
    GetSize = qlz_size_decompressed(Source(0))
End Function

Public Function DecompressString(sText As String) As String
    
    Dim Source() As Byte
    Source = ConvertStringToByteArray(sText)
    
    Dim dst() As Byte
    Dim r As Long
    Dim size As Long
    size = GetSize(Source)
    
    If size > 0 Then
        ' If size < 20 * 1000000 Then ' Visual Basic can crash if you allocate too long strings
        ReDim dst(0 To size - 1)
        r = qlz_decompress(Source(0), dst(0))
        ReDim Preserve dst(0 To r - 1)
        DecompressString = ConvertByteArrayToString(dst)
        '  End If
  
    End If
    
End Function

Private Function ConvertStringToByteArray(sString As String) As Byte()

    ConvertStringToByteArray = StrConv(sString, vbFromUnicode)

End Function

Public Function ConvertByteArrayToString(bytArray() As Byte) As String

    ConvertByteArrayToString = StrConv(bytArray, vbUnicode)

End Function

