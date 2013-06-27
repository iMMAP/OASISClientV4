Attribute VB_Name = "ArrayFunctions"
Option Explicit

Public Sub aAdd(ByRef aArray As Variant, _
                xElement As Variant)
    Dim nLen As Long

    nLen = aLen(aArray)

    ReDim Preserve aArray(nLen)
    On Error GoTo MustBeObject
    aArray(nLen) = xElement
    Exit Sub

MustBeObject:
    Set aArray(nLen) = xElement

End Sub

Public Function aLen(ByRef aArray As Variant)

    On Error GoTo NotAnArray
    aLen = UBound(aArray)
    If aLen < 0 Then
        aLen = 0
    Else
        aLen = aLen + 1
    End If

    Exit Function

NotAnArray:
    aLen = 0

End Function

Public Sub aDel(ByRef aArray As Variant, xElement As Variant)

    Dim nLen As Long
    Dim aTemp() As Variant
    Dim i As Long
    
    nLen = aLen(aArray)
    
    For i = 0 To nLen - 1
        If i <> xElement Then
            aAdd aTemp, aArray(i)
        End If
    Next i
    
    aArray = aTemp
    
End Sub

Function aScan(ArrayIn, strSearch) As Long

    Dim i As Long
    Dim intArrayLen As Long
    Dim bIs2Dim As Boolean

    intArrayLen = aLen(ArrayIn)

    aScan = -1

    If intArrayLen > 0 Then

        If aLen(ArrayIn(0)) > 0 Then

            For i = 0 To intArrayLen - 1

                If ArrayIn(i)(0) = strSearch Then
                    aScan = i
                    Exit For
                End If

            Next i

        Else

            For i = 0 To intArrayLen - 1

                If ArrayIn(i) = strSearch Then
                    aScan = i
                    Exit For
                End If

            Next i

        End If

    End If

End Function

Sub aConcat(ByRef aArray1 As Variant, ByRef aArray2 As Variant)

    Dim i As Long
    Dim intCount As Long
    
    intCount = aLen(aArray2) - 1
    
    For i = 0 To intCount
        aAdd aArray1, aArray2(i)
        DoEvents
    Next

End Sub


