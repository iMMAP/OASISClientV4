Attribute VB_Name = "modGUIDGen"

Declare Function CoCreateGuid _
        Lib "ole32.dll" (pguid As GUID) As Long
        
Declare Function StringFromGUID2 _
        Lib "ole32.dll" (rguid As Any, _
                         ByVal lpstrClsId As Long, _
                         ByVal cbMax As Long) As Long
                         
'GUID STRUCT
Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Public Function GUIDGen() As String
        '<EhHeader>
        On Error GoTo GUIDGen_Err
        '</EhHeader>
        Dim uGUID As GUID
        Dim sGUID As String
        Dim bGUID() As Byte
        Dim lLen As Long
        Dim RetVal As Long
    
100     lLen = 40
102     bGUID = String(lLen, 0)
    
104     CoCreateGuid uGUID
    
106     RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), lLen)
    
108     sGUID = bGUID

110     If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1
    
112     GUIDGen = Left$(sGUID, RetVal)

        '<EhFooter>
        Exit Function

GUIDGen_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.GUIDGen " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
