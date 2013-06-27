VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clSynchFolderImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
            
Private sClientDbFullPath As String
Private sSynchFolderPath As String
Private sSQLScriptFilename As String

Private sIMPTables() As String
Private oINIReader As New clIniReader

Private Function UpdateFromFiles() As Boolean
        '<EhHeader>
        On Error GoTo UpdateFromFiles_Err
        '</EhHeader>
    
        Dim i As Integer
        Dim ofs As New FileSystemObject
        Dim oFile As TextStream
        Dim sSQL As String
        Dim sAction As String
        Dim sKey As String

100     oINIReader.Section = "Default"
102     oINIReader.Key = "ClientDbUpdate"
    
104     If Len(oINIReader.Value) > 5 Then
106         Set oFile = ofs.OpenTextFile(sSynchFolderPath & "\" & oINIReader.Value)
108         sSQL = oFile.ReadAll

110         PrintToLog "(Drop folder synchroniser) Begin SQL Command execution of: " & sSynchFolderPath & "\" & oINIReader.Value

112         If LoadRsFromXML("", "", sSQL) Then
114             PrintToLog "(Drop folder synchroniser) --- SQL Command executed successfully"
            Else
116             PrintToLog "(Drop folder synchroniser) --- SQL Command execution failed"
            End If

118         Set oFile = Nothing
           
120         ofs.DeleteFile sSynchFolderPath & "\" & oINIReader.Value, True
        End If
    
122     For i = 0 To UBound(sIMPTables)

124         If Len(sIMPTables(i)) > 1 Then

126             oINIReader.Section = sIMPTables(i)
128             oINIReader.Key = "Action"
130             sAction = oINIReader.Value
132             oINIReader.Key = "PKey"
134             sKey = oINIReader.Value

136             oINIReader.Key = "ImportFileName"

138             PrintToLog "(Drop folder synchroniser) Start synch of table : " & oINIReader.Value

140             If LoadRsFromXML(sSynchFolderPath & "\" & oINIReader.Value, sIMPTables(i), , sAction, sKey) Then
142                 PrintToLog "(Drop folder synchroniser) --- Table Synch executed successfully"
                Else
144                 PrintToLog "(Drop folder synchroniser) --- Table Synch execution failed"
                End If
        
            End If

        Next
                        
        On Error Resume Next
146     Kill sSynchFolderPath & "\settings.ini"
148     UpdateFromFiles = True
        '<EhFooter>
        Exit Function

UpdateFromFiles_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.clSynchFolderImporter.UpdateFromFiles", _
                  "clSynchFolderImporter component failure"
        '</EhFooter>
End Function

Private Sub PrintToLog(sMessage As String)
        '<EhHeader>
        On Error GoTo PrintToLog_Err
        '</EhHeader>

100     DebugPrint sMessage

        '<EhFooter>
        Exit Sub

PrintToLog_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.clSynchFolderImporter.PrintToLog", _
                  "clSynchFolderImporter component failure"
        '</EhFooter>
End Sub

Private Function CountOccurances(ByVal sText As String) As Long
        '<EhHeader>
        On Error GoTo CountOccurances_Err
        '</EhHeader>

        Dim i As Integer
    
100     CountOccurances = 0

102     For i = 1 To Len(sText)

104         If Mid$(sText, i, 1) = "," Then CountOccurances = CountOccurances + 1
        
        Next

        '<EhFooter>
        Exit Function

CountOccurances_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.clSynchFolderImporter.CountOccurances", _
                  "clSynchFolderImporter component failure"
        '</EhFooter>
End Function

Private Function ReadINI(sIniFilePath As String) As Boolean
        '<EhHeader>
        On Error GoTo ReadINI_Err
        '</EhHeader>

100     With oINIReader
    
102         .Path = sIniFilePath
104         .Section = "Default"
106         .Key = "ImportTables"
        
108         ReDim sIMPTables(CountOccurances(.Value))
110         sIMPTables = Split(.Value, ",")
        
112         .Section = "Default"
114         .Key = "ClientDbUpdate"
116         sSQLScriptFilename = .Value
       
        End With
        
118     ReadINI = True
        '<EhFooter>
        Exit Function

ReadINI_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.clSynchFolderImporter.ReadINI", _
                  "clSynchFolderImporter component failure"
        '</EhFooter>
End Function

Private Function LoadRsFromXML(sXMLFilePath As String, _
                               sTableName As String, _
                               Optional sSQLCommand As String, _
                               Optional sAction As String, _
                               Optional sKey As String) As Boolean
        '<EhHeader>
        On Error GoTo LoadRsFromXML_Err
        '</EhHeader>

        Dim RS As New ADODB.Recordset
        Dim oRS As New ADODB.Recordset
        Dim CN As New ADODB.Connection
        Dim i As Long
        'On Error Resume Next

100     If Dir(sXMLFilePath) = "" Then
102         If sSQLCommand <> "" Then
104             'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sClientDbFullPath & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
                CN.ConnectionString = GetConnectionString(sClientDbFullPath)
106             CN.Open
108             CN.Execute sSQLCommand
110             CN.Close
112             Set CN = Nothing
            End If
114         LoadRsFromXML = True
            Exit Function
116     ElseIf sSQLCommand <> "" Then
118         'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sClientDbFullPath & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
            CN.ConnectionString = GetConnectionString(sClientDbFullPath)
120         CN.Open
122         CN.Execute sSQLCommand
124         CN.Close
126         Set CN = Nothing
128         LoadRsFromXML = True
            Exit Function
        End If
        
130     oRS.CursorLocation = g_sGlobalCursorLocation
132     oRS.Open sXMLFilePath, "Provider=MSPersist;", adOpenForwardOnly, adLockReadOnly, adCmdFile
        
134     'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sClientDbFullPath & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
        CN.ConnectionString = GetConnectionString(sClientDbFullPath)
136     If sSQLCommand = "" Then
138         CN.Open
            
140         If Not TableExistsLocal(sTableName, CN) Then
142             CloneTableLocal sTableName, oRS, CN
            End If
            
144         If sAction = "Replace" Then
146             CN.Execute "delete from " & sTableName
            End If
            
148         RS.CursorLocation = g_sGlobalCursorLocation
150         RS.Open "SELECT * FROM " & sTableName, CN, adOpenDynamic, adLockOptimistic
        
152         If Not oRS.Bof Then
154             SafeMoveFirst oRS
            End If
        
156         If sAction = "Replace" Or sAction = "Append" Then

158             Do While Not oRS.EOF
    
160                 With RS
162                     .AddNew
    
164                     For i = 0 To .Fields.Count - 1
166                         .Fields.Item(i).Value = oRS.Fields.Item(i).Value
                        Next
    
168                     .UpDate
                    End With
    
170                 oRS.MoveNext
                Loop

172         ElseIf sAction = "Update" Then

174             If Not oRS.Bof Then
176                 SafeMoveFirst oRS
                End If
            
178             Do While Not oRS.EOF
                    
180                 If sKey = "" Then
                        
182                     sKey = oRS.Fields(0).Name
                    End If
                    
                    Dim sPreFix As String
                    
184                 If oRS.Fields(sKey).Type = adChar Then sPreFix = "'"
                    
186                 With RS
188                     SafeMoveFirst RS
190                     .Find sKey & " = " & sPreFix & oRS.Fields.Item(sKey).Value & sPreFix

192                     If Not .EOF Then

194                         For i = 0 To .Fields.Count - 1

196                             If .Fields.Item(i).Name <> sKey Then
198                                 .Fields.Item(i).Value = oRS.Fields.Item(i).Value
                                End If

                            Next
                        
200                         .UpDate
                        End If

                    End With
    
202                 oRS.MoveNext
                Loop
            
            End If
            
        End If
    
204     LoadRsFromXML = True
    
        'Clean Up
        On Error Resume Next
206     oRS.Close
208     Set oRS = Nothing
210     RS.Close
212     Set RS = Nothing
214     Kill sXMLFilePath
216     CN.Close
218     Set CN = Nothing

        '<EhFooter>
        Exit Function

LoadRsFromXML_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.clSynchFolderImporter.LoadRsFromXML", _
                  "clSynchFolderImporter component failure"
        '</EhFooter>
End Function

Private Function TableExistsLocal(sName As String, _
                                  CN As ADODB.Connection) As Boolean
        '<EhHeader>
        On Error GoTo TableExistsLocal_Err
        '</EhHeader>
        
        Dim RS As New ADODB.Recordset
    
100     TableExistsLocal = True
    
102     RS.Open "SELECT * FROM " & sName, CN

        '<EhFooter>
        Exit Function

TableExistsLocal_Err:
        TableExistsLocal = False
        '</EhFooter>
End Function

Private Sub CloneTableLocal(sTableName As String, _
                            oRSTemplateTable As ADODB.Recordset, _
                            connect As ADODB.Connection)
        '<EhHeader>
        On Error GoTo CloneTableLocal_Err
        '</EhHeader>
        
        Dim FileDoc As Object
        Dim SQLString As String
        Dim sConn As String
        Dim cat As ADOX.Catalog
        Dim tblTemplate As ADOX.Table
        Dim col As ADODB.Field
        Dim tbl As ADOX.Table
        Dim CurrentProperty As ADODB.Property

100     Set cat = CreateObject("ADOX.Catalog")
102     Set tblTemplate = CreateObject("ADOX.Table")
104     Set tbl = CreateObject("ADOX.Table")

106     Set cat.ActiveConnection = connect

108     tbl.Name = sTableName
110     Set tbl.ParentCatalog = cat

112     With tbl.Columns

114         For Each col In oRSTemplateTable.Fields
116             .Append col.Name, col.Type, col.DefinedSize
118             .Item(col.Name).Properties("Nullable").Value = True
120             .Item(col.Name).Properties("Jet OLEDB:Allow Zero Length").Value = True
            Next

        End With

122     cat.Tables.Append tbl

124     Set tbl = Nothing
126     Set cat = Nothing
128     Set FileDoc = Nothing
    
        '<EhFooter>
        Exit Sub

CloneTableLocal_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.clSynchFolderImporter.CloneTableLocal", _
                  "clSynchFolderImporter component failure"
        '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    Set oINIReader = Nothing

End Sub

Private Function FileExists(Filename As String) As Integer
        '<EhHeader>
        On Error GoTo FileExists_Err
        '</EhHeader>
        Dim i As Integer
    
100     On Local Error Resume Next
102     i = Len(Dir$(Filename$))
104     If Err Or i = 0 Then
106         FileExists = False
        Else
108         FileExists = True
        End If
110     On Local Error GoTo 0
        '<EhFooter>
        Exit Function

FileExists_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.clSynchFolderImporter.FileExists", _
                  "clSynchFolderImporter component failure"
        '</EhFooter>
End Function

Public Sub ScanAndProcessSynchFolder(sPassedImportFolderPath As String, _
                                     sPassedClientDbFullPath As String)
        '<EhHeader>
        On Error GoTo ScanAndProcessSynchFolder_Err
        '</EhHeader>

        Dim aTemp As Variant
        Dim i As Integer
        Dim j As Integer
        Dim intFileCount As Integer
        Dim bAllFilesAvailable As Boolean
        
100     sClientDbFullPath = sPassedClientDbFullPath
102     sSynchFolderPath = sPassedImportFolderPath
        
104     If FileExists(sSynchFolderPath & "\" & "settings.ini") Then
        
106         PrintToLog Chr(13) & "OASIS background folder synch initiated....." & Chr(13) & "settings.ini file found.  Synch will start in 5 seconds on: " & sSynchFolderPath & "\" & "settings.ini"

108         If ReadINI(sSynchFolderPath & "\" & "settings.ini") = True Then

                ''''' Check if all files are available '''''''
110             bAllFilesAvailable = True
112             j = 0

114             Do Until j = UBound(sIMPTables) + 1

116                 If bAllFilesAvailable Then
118                     bAllFilesAvailable = FileExists(sSynchFolderPath & "\" & sIMPTables(j) & ".xml")
                    End If

120                 j = j + 1
                Loop
                    
122             If Len(sSQLScriptFilename) > 2 Then
                    
124                 If bAllFilesAvailable Then
126                     bAllFilesAvailable = FileExists(sSynchFolderPath & "\" & sSQLScriptFilename)
                    End If
                    
                End If
                    
128             If bAllFilesAvailable Then

130                 If UpdateFromFiles = False Then
132                     PrintToLog "(Drop folder synchroniser) --- Read failure from file(s) specified in INI file"
                    End If
                    
                Else
                
134                 PrintToLog "(Drop folder synchroniser) --- All files specified in the INI file are not available"
                    
                End If

            Else
136             PrintToLog "(Drop folder synchroniser) Read failure from INI file: " & sSynchFolderPath & "\" & "settings.ini"
            End If
        
        End If

        '<EhFooter>
        Exit Sub

ScanAndProcessSynchFolder_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.clSynchFolderImporter.ScanAndProcessSynchFolder", _
                  "clSynchFolderImporter component failure"
        '</EhFooter>
End Sub
