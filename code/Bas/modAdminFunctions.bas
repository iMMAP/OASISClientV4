Attribute VB_Name = "modAdminFunctions"
Public WebSite As String
Public g_sAppPath As String
Public g_bHasEncrypt As Boolean
Public m_oAES As New clsAES
Public m_sKey As String

Public Type ExcludeType
    sTableName As String
    sFieldName As String
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type


Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long ' = &h3F for version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type


Public m_frmOASISProgress As frmOASISProgress
Public m_frmDebug As frmDebug

Public ExcludeArray() As ExcludeType

Private Declare Sub MoveMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Dest As Any, _
                                       ByVal Source As Long, _
                                       ByVal Length As Long)

Public Function SafeMoveFirst(ByRef oRs As ADODB.Recordset) As Boolean
        '<EhHeader>
        On Error GoTo SafeMoveFirst_Err
        '</EhHeader>

100     If Not oRs.State = adStateClosed Then
    
102         If Not oRs.EOF Or Not oRs.Bof Then
104             oRs.MoveFirst
106             SafeMoveFirst = True
            Else
108             SafeMoveFirst = False
            End If

        Else
110         SafeMoveFirst = False
        End If

        '<EhFooter>
        Exit Function

SafeMoveFirst_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modAdminFunctions.SafeMoveFirst " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DeleteRecordFromRSAndSave(oRs As ADODB.Recordset, _
                                          Optional sSettingValue As String, _
                                          Optional sUserGroupName As String) As Boolean
        '<EhHeader>
        On Error GoTo DeleteRecordFromRSAndSave_Err
        '</EhHeader>

        Dim bReturnValue As Boolean
100     DeleteRecordFromRSAndSave = False

102     If Not oRs Is Nothing Then
    
104         If Not oRs.EOF And Not oRs.Bof Then
        
106             If MsgBox("Do you want to delete this record?", vbYesNo, "Confirm Deletion") = vbYes Then
    
108                 oRs.Delete
110                 oRs.Filter = adFilterAffectedRecords 'adFilterPendingRecords
            
112                 bReturnValue = m_frmOASISProgress.SaveHttpCommsRS(oRs, WebSite & "Oasis.asp", True)
            
114                 If bReturnValue Then

116                     If Len(sSettingValue) > 0 Then
118                         IncrementProfileSettingVersion WebSite, sSettingValue, sUserGroupName
                        End If
                
120                     MsgBox "Deletion successful"
122                     DeleteRecordFromRSAndSave = True
                
                    Else
124                     MsgBox "Deletion unsuccessful"
                    End If
            
126                 oRs.Filter = adFilterNone
                End If
            End If
    
        Else
    
128         DeleteRecordFromRSAndSave = False
        
        End If

        '<EhFooter>
        Exit Function

DeleteRecordFromRSAndSave_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modAdminFunctions.DeleteRecordFromRSAndSave " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub IncrementProfileSettingVersion(sWebsitePassed As String, _
                                          sSettingValue As String, _
                                          sUserName As String)
        '<EhHeader>
        On Error GoTo IncrementProfileSettingVersion_Err
        '</EhHeader>
    
        Dim RSLocalAppSetting As ADODB.Recordset
        Dim sSQL As String
        
100     Set RSLocalAppSetting = New ADODB.Recordset
102     sSQL = sWebsitePassed & "Oasis.asp?ID=" & CheckEncrypt("SELECT * FROM " & sUserName & "AppSettings WHERE SettingName ='ProfileSettings'")
104     Set RSLocalAppSetting = m_frmOASISProgress.OpenHttpCommsRS(sSQL, True)

108     If IsNull(RSLocalAppSetting.fields(sSettingValue).Value) Then
110         RSLocalAppSetting.fields(sSettingValue).Value = 1
        Else
112         RSLocalAppSetting.fields(sSettingValue).Value = CInt(0 & RSLocalAppSetting.fields(sSettingValue).Value) + 1
        End If
    
114     If Not RSLocalAppSetting.State = adStateClosed Then
116         RSLocalAppSetting.Filter = adFilterPendingRecords
118         If Not RSLocalAppSetting.EOF And Not RSLocalAppSetting.Bof Then
120             m_frmOASISProgress.SaveHttpCommsRS RSLocalAppSetting, WebSite & "Oasis.asp", True
            End If
        End If

122     Set RSLocalAppSetting = Nothing

        '<EhFooter>
        Exit Sub

IncrementProfileSettingVersion_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modAdminFunctions.IncrementProfileSettingVersion " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub SetStatus(sText As String)
        '<EhHeader>
        On Error GoTo SetStatus_Err
        '</EhHeader>

100     If Len(frmDatabaseConnect.txtAppStatus.Text) > 5000 Then frmDatabaseConnect.txtAppStatus.Text = ""
102     frmDatabaseConnect.txtAppStatus.Text = sText & vbCrLf & frmDatabaseConnect.txtAppStatus.Text
    
        '<EhFooter>
        Exit Sub

SetStatus_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modAdminFunctions.SetStatus " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub LoadLanguage(Optional v1 As Variant, _
                        Optional v2 As Variant, _
                        Optional v3 As Variant)
    'This is a dummy function which is needed for the form frmDynamicDataMenu to work
End Sub

Public Function ReadVersion(FullFileName As String) As String
        '<EhHeader>
        On Error GoTo ReadVersion_Err
        '</EhHeader>
        Dim rc As Long, lDummy As Long, sBuffer() As Byte, lVerPointer As Long
        Dim lBufferLen As Long, udtVerBuffer As VS_FIXEDFILEINFO
        Dim lVerbufferLen As Long

100     LeerVersion = 0
        '*** Get size ****
102     lBufferLen = GetFileVersionInfoSize(FullFileName, lDummy)

104     If lBufferLen < 1 Then
            Exit Function
        End If

        '**** Store info to udtVerBuffer struct ****
106     ReDim sBuffer(lBufferLen)
108     rc = GetFileVersionInfo(FullFileName, 0&, lBufferLen, sBuffer(0))
110     rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
112     MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
        '**** Determine Product Version number ****
114     ReadVersion = Format$(udtVerBuffer.dwProductVersionMSh) & "." & Format$(udtVerBuffer.dwProductVersionMSl) & "." & Format$(udtVerBuffer.dwProductVersionLSh) & "." & Format$(udtVerBuffer.dwProductVersionLSl)

        '<EhFooter>
        Exit Function

ReadVersion_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modAdminFunctions.ReadVersion " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function CheckEncrypt(sValue As String) As String
        '<EhHeader>
        On Error GoTo CheckEncrypt_Err
        '</EhHeader>
            
100     If g_bHasEncrypt Then
102         CheckEncrypt = m_oAES.AESEncyptString(sValue, m_sKey)
        Else
104         CheckEncrypt = sValue
        End If

        '<EhFooter>
        Exit Function

CheckEncrypt_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modAdminFunctions.CheckEncrypt " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

'Public Sub Create_SynchTables(cn As ADODB.Connection)
'    Dim dbCol As ADOX.Column
'    Dim dbIdx As ADOX.Index
'    Dim dbTbl As ADOX.Table
'    Dim dbCat As ADOX.Catalog
'
'    Set dbCat = New ADOX.Catalog
'
'    Set dbCat.ActiveConnection = cn
'
'    '########################################################################################################################
'    'Code to generate Objects for Table: SynchTables
'    Set dbTbl = New ADOX.Table
'
'    With dbTbl
'        Set .ParentCatalog = dbCat
'        .Name = "SynchTables1"
'
'        .Columns.Append "AllowWrite", adBoolean
'        .Columns("AllowWrite").attributes = 1
'        .Columns.Append "AutoUpdate", adBoolean
'        .Columns("AutoUpdate").attributes = 1
'        .Columns.Append "isGeoTable", adBoolean
'        .Columns("isGeoTable").attributes = 1
'        .Columns.Append "OwnerID", adVarWChar
'        .Columns("OwnerID").attributes = 2
'        .Columns.Append "sDescription", adLongVarWChar
'        .Columns("sDescription").attributes = 2
'        .Columns.Append "sGUID", adGUID
'        .Columns("sGUID").attributes = 3
'        .Columns.Append "sName", adVarWChar
'        .Columns("sName").attributes = 2
'        .Columns.Append "sTableName", adVarWChar
'        .Columns("sTableName").attributes = 2
'        .Columns.Append "SynchFrequency", adInteger
'        .Columns("SynchFrequency").attributes = 3
'    End With
'
'    dbCat.Tables.Append dbTbl
'
'    Set dbIdx = New ADOX.Index
'
'    With dbIdx
'        .Name = "PrimaryKey"
'        .PrimaryKey = True
'        .Unique = True
'        .Columns.Append "sGUID", adGUID
'    End With
'
'    dbTbl.Indexes.Append dbIdx
'
'    Set dbIdx = New ADOX.Index
'
'    With dbIdx
'        .Name = "sGUID"
'        .Unique = True
'        .Columns.Append "sGUID", adGUID
'    End With
'
'    dbTbl.Indexes.Append dbIdx
'
'    Set dbIdx = New ADOX.Index
'
'    With dbIdx
'        .Name = "OwnerID"
'        .Columns.Append "OwnerID", adVarWChar
'    End With
'
'    dbTbl.Indexes.Append dbIdx
'
'    Set dbCat = Nothing
'
'End Sub

Public Function KeyGen(sKey As String) As String
        '<EhHeader>
        On Error GoTo KeyGen_Err
        '</EhHeader>
        Dim oMD5 As New clsMD5

100     KeyGen = oMD5.MD5(sKey)

        On Error Resume Next
102     Set oMD5 = Nothing
        '<EhFooter>
        Exit Function

KeyGen_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modAdminFunctions.KeyGen " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function getASPFileVersionAndDate(ByRef sWebsite As String) As String
        '<EhHeader>
        On Error GoTo getASPFileVersionAndDate_Err
        '</EhHeader>
    
        Dim sRetValueSplit() As String
        Dim sRetValue As String
        Dim i As Integer

100     If Not sWebsite = "" Then
    
102         sRetValue = m_frmOASISProgress.OpenHttpCommsResponse(sWebsite & "Oasis.asp?disvers=1", True)

104         If sRetValue = "-1" Then
106             getASPFileVersionAndDate = "Fetching OASIS.ASP version number FAILED"
108             m_frmDebug.DebugPrint getASPFileVersionAndDate
            Else
110             sRetValueSplit = Split(sRetValue, " ")
112             getASPFileVersionAndDate = "OASIS.ASP version number: " & sRetValueSplit(0) & Chr(13) & Chr(13) & "updated " & sRetValueSplit(1) & " " & sRetValueSplit(2) & " " & sRetValueSplit(3)
            End If
        End If
    
        '<EhFooter>
        Exit Function

getASPFileVersionAndDate_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modAdminFunctions.getASPFileVersionAndDate " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function ConvertDateToSerial(dDate As Date) As Long
        '<EhHeader>
        On Error GoTo ConvertDateToSerial_Err
        '</EhHeader>

100     ConvertDateToSerial = CLng(Format(dDate, "yyyymmdd"))

        '<EhFooter>
        Exit Function

ConvertDateToSerial_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modAdminFunctions.ConvertDateToSerial " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function ConvertSerialToDate(lSerialDate As Long) As Date
        '<EhHeader>
        On Error GoTo ConvertSerialToDate_Err
        '</EhHeader>

        Dim sDate As String
        Dim sDateReversed As String
    
100     sDate = CStr(lSerialDate)
102     sDateReversed = Right$(sDate, 2) & "-" & Mid$(sDate, 5, 2) & "-" & Left$(sDate, 4)

104     ConvertSerialToDate = Format(sDateReversed, "dd-mm-yyyy")
106     ConvertSerialToDate = Format(ConvertSerialToDate, "Medium Date")

        '<EhFooter>
        Exit Function

ConvertSerialToDate_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modAdminFunctions.ConvertSerialToDate " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function SaveSilentHttpCommsRS(oRs As ADODB.Recordset, _
                                      sWebsite As String, _
                                      bInit As Boolean) As Boolean
    'Cascading function for frmDynamicDataMenu to make compatible with client app
    SaveSilentHttpCommsRS = m_frmOASISProgress.SaveHttpCommsRS(oRs, sWebsite, bInit)
    
End Function

Public Function OpenSilentHttpCommsRS(sWebsite As String, _
                                      bInit As Boolean) As ADODB.Recordset
    'Cascading function for frmDynamicDataMenu to make compatible with client app
    Set OpenSilentHttpCommsRS = m_frmOASISProgress.OpenHttpCommsRS(sWebsite, bInit)
    
End Function

Public Function GetConnectionString(sString As String)
    
    'Access to encrypted ACCESS DB from admin not support yet
    GetConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & sString

End Function

