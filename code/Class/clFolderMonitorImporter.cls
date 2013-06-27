VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clFLDMonImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents objFolderMonitor     As FldMonitor
Attribute objFolderMonitor.VB_VarHelpID = -1
Private Declare Function ShellAbout _
                Lib "shell32.dll" _
                Alias "ShellAboutA" (ByVal hWnd As Long, _
                                     ByVal szApp As String, _
                                     ByVal szOtherStuff As String, _
                                     ByVal hIcon As Long) As Long

Private Declare Sub Sleep _
                Lib "kernel32" (ByVal dwMilliseconds As Long)
            
Dim sClientDbFullPath As String
Dim sImportFolderPath As String
Dim sSQLScriptFilename As String

Dim sIMPTables() As String
Private oINIReader As New clIniReader

Public Sub CommenceMonitoring(sPassedImportFolderPath As String, _
                              sPassedClientDbFullPath As String)
        '<EhHeader>
        On Error GoTo CommenceMonitoring_Err
        '</EhHeader>

        Dim objCurrentControl As Object
    
100     sClientDbFullPath = sPassedClientDbFullPath
102     sImportFolderPath = sPassedImportFolderPath
    
104     Set objFolderMonitor = New FldMonitor
106     objFolderMonitor.WaitTime = 100
108     objFolderMonitor.IncludeSubFolders = False
110     objFolderMonitor.Attributes = True
112     objFolderMonitor.AddFolder sImportFolderPath
114     objFolderMonitor.StartMonitoring

        '<EhFooter>
        Exit Sub

CommenceMonitoring_Err:
        Err.Raise vbObjectError + 100, "OASISFileMonitor.clFLDMonImporter.CommenceMonitoring", "clFLDMonImporter component failure"
        '</EhFooter>
End Sub

Public Sub StopMonitoring()
        '<EhHeader>
        On Error GoTo StopMonitoring_Err
        '</EhHeader>

100     If Not objFolderMonitor Is Nothing Then objFolderMonitor.StopMonitoring

        '<EhFooter>
        Exit Sub

StopMonitoring_Err:
        Err.Raise vbObjectError + 100, "OASISFileMonitor.clFLDMonImporter.StopMonitoring", "clFLDMonImporter component failure"
        '</EhFooter>
End Sub

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
106         Set oFile = ofs.OpenTextFile(sImportFolderPath & "\" & oINIReader.Value)
108         sSQL = oFile.ReadAll

110         PrintToLog "(Drop folder synchroniser) Begin SQL Command execution of: " & sImportFolderPath & "\" & oINIReader.Value

112         If LoadRsFromXML("", "", sSQL) Then
114             PrintToLog "(Drop folder synchroniser) --- SQL Command executed successfully"
            Else
116             PrintToLog "(Drop folder synchroniser) --- SQL Command execution failed"
            End If

118         Set oFile = Nothing
           
120         ofs.DeleteFile sImportFolderPath & "\" & oINIReader.Value, True
        End If
    
122     For i = 0 To UBound(sIMPTables)

124         If Len(sIMPTables(i)) > 1 Then

128             oINIReader.Section = sIMPTables(i)
130             oINIReader.Key = "Action"
132             sAction = oINIReader.Value
134             oINIReader.Key = "PKey"
136             sKey = oINIReader.Value

138             oINIReader.Key = "ImportFileName"

140             PrintToLog "(Drop folder synchroniser) Start synch of table : " & oINIReader.Value

142             If LoadRsFromXML(sImportFolderPath & "\" & oINIReader.Value, sIMPTables(i), , sAction, sKey) Then
144                 PrintToLog "(Drop folder synchroniser) --- Table Synch executed successfully"
                Else
146                 PrintToLog "(Drop folder synchroniser) --- Table Synch execution failed"
                End If
        
            End If

        Next
                        
        On Error Resume Next
148     Kill sImportFolderPath & "\settings.ini"
150     UpdateFromFiles = True
        '<EhFooter>
        Exit Function

UpdateFromFiles_Err:
        UpdateFromFiles = False
        '</EhFooter>
End Function

Private Sub PrintToLog(sMessage As String)

    DebugPrint sMessage

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
        Err.Raise vbObjectError + 100, "OASISFileMonitor.clFLDMonImporter.CountOccurances", "clFLDMonImporter component failure"
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
        
        ReadINI = True
        '<EhFooter>
        Exit Function

ReadINI_Err:
        ReadINI = False
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
104             CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sClientDbFullPath & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
106             CN.Open
108             CN.Execute sSQLCommand
110             CN.Close
112             Set CN = Nothing
            End If
            LoadRsFromXML = True
            Exit Function
114     ElseIf sSQLCommand <> "" Then
116         CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sClientDbFullPath & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
118         CN.Open
120         CN.Execute sSQLCommand
122         CN.Close
124         Set CN = Nothing
            LoadRsFromXML = True
            Exit Function
        End If
        
126     oRS.CursorLocation = g_sGlobalCursorLocation
128     oRS.Open sXMLFilePath, "Provider=MSPersist;", adOpenForwardOnly, adLockReadOnly, adCmdFile
130     CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sClientDbFullPath & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False"
    
132     If sSQLCommand = "" Then
134         CN.Open
            
136         If Not TableExistsLocal(sTableName, CN) Then
138             CloneTableLocal sTableName, oRS, CN
            End If
            
140         If sAction = "Replace" Then
142             CN.Execute "delete from " & sTableName
            End If
            
144         RS.CursorLocation = g_sGlobalCursorLocation
146         RS.Open "SELECT * FROM " & sTableName, CN, adOpenDynamic, adLockOptimistic
        
148         If Not oRS.Bof Then
150             SafeMoveFirst oRS
            End If
        
152         If sAction = "Replace" Or sAction = "Append" Then

154             Do While Not oRS.EOF
    
156                 With RS
158                     .AddNew
    
160                     For i = 0 To .Fields.Count - 1
162                         .Fields.Item(i).Value = oRS.Fields.Item(i).Value
                        Next
    
164                     .UpDate
                    End With
    
166                 oRS.MoveNext
                Loop

168         ElseIf sAction = "Update" Then

170             If Not oRS.Bof Then
172                 SafeMoveFirst oRS
                End If
            
174             Do While Not oRS.EOF
                    
176                 If sKey = "" Then
                        
178                     sKey = oRS.Fields(0).Name
                    End If
                    
                    Dim sPreFix As String
                    
180                 If oRS.Fields(sKey).Type = adChar Then sPreFix = "'"
                    
182                 With RS
184                     SafeMoveFirst RS
186                     .Find sKey & " = " & sPreFix & oRS.Fields.Item(sKey).Value & sPreFix

188                     If Not .EOF Then

190                         For i = 0 To .Fields.Count - 1

192                             If .Fields.Item(i).Name <> sKey Then
194                                 .Fields.Item(i).Value = oRS.Fields.Item(i).Value
                                End If

                            Next
                        
196                         .UpDate
                        End If

                    End With
    
198                 oRS.MoveNext
                Loop
            
            End If
            
        End If
    
200     LoadRsFromXML = True
    
        'Clean Up
        On Error Resume Next
202     oRS.Close
204     Set oRS = Nothing
206     RS.Close
208     Set RS = Nothing
210     Kill sXMLFilePath
212     CN.Close
214     Set CN = Nothing

        '<EhFooter>
        Exit Function

LoadRsFromXML_Err:
        LoadRsFromXML = False
        '</EhFooter>
End Function

Private Function TableExistsLocal(sName As String, _
                                  CN As ADODB.Connection) As Boolean
        '<EhHeader>
        On Error GoTo ErrH
        '</EhHeader>
        
        Dim RS As New ADODB.Recordset
    
100     TableExistsLocal = True
    
102     RS.Open "SELECT * FROM " & sName, CN

        '<EhFooter>
        Exit Function
ErrH:
104     TableExistsLocal = False
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
        Err.Raise vbObjectError + 100, "OASISFileMonitor.clFLDMonImporter.CloneTableLocal", "clFLDMonImporter component failure"
        '</EhFooter>
End Sub

Private Sub Class_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    Set objFolderMonitor = Nothing
    Set oINIReader = Nothing

End Sub

Private Function FileExists(Filename As String) As Integer
    Dim i As Integer
    
    On Local Error Resume Next
    i = Len(Dir$(Filename$))
    If Err Or i = 0 Then
        FileExists = False
    Else
        FileExists = True
    End If
    On Local Error GoTo 0
End Function

Private Sub objFolderMonitor_ChangeOccurred()
        '<EhHeader>
        On Error GoTo objFolderMonitor_ChangeOccurred_Err
        '</EhHeader>

        Dim aTemp As Variant
        Dim i As Integer
        Dim j As Integer
        Dim intFileCount As Integer
        Dim bAllFilesAvailable As Boolean
        
100     If FileExists(sImportFolderPath & "\" & "settings.ini") Then
        
102         PrintToLog Chr(13) & "OASIS background folder synch initiated....." & Chr(13) & "settings.ini file found.  Synch will start in 5 seconds on: " & sImportFolderPath & "\" & "settings.ini"

104         If ReadINI(sImportFolderPath & "\" & "settings.ini") = True Then

                ''''' Check if all files are available '''''''
106             bAllFilesAvailable = True
108             j = 0

110             Do Until j = UBound(sIMPTables) + 1

112                 If bAllFilesAvailable Then
114                     bAllFilesAvailable = FileExists(sImportFolderPath & "\" & sIMPTables(j) & ".xml")
                    End If

116                 j = j + 1
                Loop
                    
118             If Len(sSQLScriptFilename) > 2 Then
                    
120                 If bAllFilesAvailable Then
122                     bAllFilesAvailable = FileExists(sImportFolderPath & "\" & sSQLScriptFilename)
                    End If
                    
                End If
                    
124             If bAllFilesAvailable Then

126                 If UpdateFromFiles = False Then
128                     PrintToLog "(Drop folder synchroniser) --- Read failure from file(s) specified in INI file"
                    End If
                    
                Else
                
130                 PrintToLog "(Drop folder synchroniser) --- All files specified in the INI file are not available"
                    
                End If

            Else
132             PrintToLog "(Drop folder synchroniser) Read failure from INI file: " & sImportFolderPath & "\" & "settings.ini"
            End If
        
        End If

        '<EhFooter>
        Exit Sub

objFolderMonitor_ChangeOccurred_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISFileMonitor.clFLDMonImporter.objFolderMonitor_ChangeOccurred", _
                  "clFLDMonImporter component failure"
        '</EhFooter>
End Sub
