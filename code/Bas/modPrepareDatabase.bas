Attribute VB_Name = "modPrepareDatabase"
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function CreateProcess _
                Lib "kernel32" _
                Alias "CreateProcessA" (ByVal lpApplicationName As String, _
                                        ByVal lpCommandLine As String, _
                                        lpProcessAttributes As Any, _
                                        lpThreadAttributes As Any, _
                                        ByVal bInheritHandles As Long, _
                                        ByVal dwCreationFlags As Long, _
                                        lpEnvironment As Any, _
                                        ByVal lpCurrentDriectory As String, _
                                        lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function OpenProcess _
                Lib "kernel32.dll" (ByVal dwAccess As Long, _
                                    ByVal fInherit As Integer, _
                                    ByVal hObject As Long) As Long

Private Declare Function TerminateProcess _
                Lib "kernel32" (ByVal hProcess As Long, _
                                ByVal uExitCode As Long) As Long

Private Declare Function CloseHandle _
                Lib "kernel32" (ByVal hObject As Long) As Long
         
Private Declare Function WaitForSingleObject _
                Lib "kernel32" (ByVal hHandle As Long, _
                                ByVal dwMilliseconds As Long) As Long

Const SYNCHRONIZE = 1048576
Const NORMAL_PRIORITY_CLASS = &H20&
      
Private Sub ShellAndWait(ByVal program_name As String, _
                         ByVal window_style As VbAppWinStyle)
    Dim process_id As Long
    Dim process_handle As Long

    ' Start the program.
    On Error GoTo ShellError
    process_id = Shell(program_name, window_style)
    On Error GoTo 0

    ' Hide.
    'Me.Visible = False

    DoEvents

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)

    If process_handle <> 0 Then
        WaitForSingleObject process_handle, 60000 ' INFINITE
        CloseHandle process_handle
    End If

    ' Reappear.
    '  Me.Visible = True
    Exit Sub

ShellError:
    MsgBox "Error starting task " & txtProgram.Text & vbCrLf & Err.Description, vbOKOnly Or vbExclamation, "Error"
End Sub

Public Function MSSQL_CheckIfInstalled() As Boolean

    Dim objShell
    On Error Resume Next
    Set objShell = CreateObject("WScript.Shell")
    objShell.RegRead ("HKLM\Software\Microsoft\Microsoft SQL Server\Instance Names\SQL\OASISSQL")
    Set objShell = Nothing

    If Err = 0 Then
        MSSQL_CheckIfInstalled = True
    Else
        MSSQL_CheckIfInstalled = False
    End If
  
End Function

Public Sub MSSQL_PrepareNewDatabase()
    '<EhHeader>
    On Error GoTo MSSQL_PrepareNewDatabase_Err
    '</EhHeader>
    
    Dim CN As New adodb.Connection
    Dim sDatabaseCreationScript As String
    Dim bRestore As Boolean
    
    If FileExists(g_sAppPath & "\data\db\" & Replace(frmLogin.ComServer.Text, "/", "-") & ".bak") Then
    
        If MsgBox("Do you want to restore from a backup database?  Otherwise a blank database will be created.", vbYesNo) = vbYes Then
            bRestore = True
        End If
    
    End If
    
    If bRestore Then
        MSSQL_NewFromBackup
    Else
    
        CN.ConnectionString = Replace(GetConnectionString(""), g_sSQLServerDatabaseName, "master") ' "Driver={SQL Server Native Client 10.0};Server=localhost\oasissql;Database=master;Uid=sa;Pwd=!MM@P2O1O"
        CN.Open
        sDatabaseCreationScript = "CREATE DATABASE [" & g_sSQLServerDatabaseName & "] " & "ON " & "( NAME = '" & g_sSQLServerDatabaseName & "_dat', " & "  FILENAME = '" & g_sAppPath & "\data\db\" & g_sSQLServerDatabaseName & ".mdf' " & " ) " & "LOG ON " & "( NAME = '" & g_sSQLServerDatabaseName & "_log', " & "  FILENAME = '" & g_sAppPath & "\data\db\" & g_sSQLServerDatabaseName & ".ldf' " & " ) "
        CN.Execute sDatabaseCreationScript
        CN.Close

        GetConnectionString ""
        CN.ConnectionString = g_sGlobalConnectionString
        CN.Open
        MSSQL_CreateTablesFromDefaultSchema CN
        CN.Close
        MsgBox "The OASIS MSSQL Database has been restored to default schema", vbInformation

    End If

    Set CN = Nothing
    '<EhFooter>
    Exit Sub

MSSQL_PrepareNewDatabase_Err:
    MsgBox Err.Description & vbCrLf & "in Project1.PrepareDatabase.MSSQL_PrepareNewDatabase " & "at line " & Erl
    'Resume Next

    '</EhFooter>
End Sub

Public Sub MSSQL_RestoreFromBackup()
  
  GetConnectionString ""
  
    If FileExists(g_sAppPath & "\data\db\" & Replace(frmLogin.ComServer.Text, "/", "-") & ".bak") Then
        ShellAndWait "net stop MSSQL$OASISSQL", vbMaximizedFocus
        ShellAndWait "net start MSSQL$OASISSQL", vbMaximizedFocus
        Sleep 3000
        ShellAndWait "OSQL -U sa -P !MM@P2O1O -S " & g_sManualSQLServerPath & "\oasissql -Q ""RESTORE DATABASE [" & g_sSQLServerDatabaseName & "] from DISK = '" & g_sAppPath & "\data\db\" & Replace(frmLogin.ComServer.Text, "/", "-") & ".bak' WITH replace""", vbMaximizedFocus
        MsgBox "The OASIS MSSQL Database has been restored from a backup", vbInformation
    Else
        'MsgBox "FAILED.  File does not exist: " & g_sAppPath & "\data\db\" & Replace(frmLogin.ComServer.Text, "/", "^") & ".bak", vbError
    End If

End Sub

Public Sub MSSQL_NewFromBackup()
  
  GetConnectionString ""
  
    If FileExists(g_sAppPath & "\data\db\" & Replace(frmLogin.ComServer.Text, "/", "-") & ".bak") Then
        ShellAndWait "OSQL -U sa -P !MM@P2O1O -S " & g_sManualSQLServerPath & "\oasissql -Q ""RESTORE DATABASE [" & g_sSQLServerDatabaseName & "] from DISK = '" & g_sAppPath & "\data\db\" & Replace(frmLogin.ComServer.Text, "/", "-") & ".bak' WITH MOVE '" & g_sSQLServerDatabaseName & "_dat' to '" & g_sAppPath & "\data\db\" & g_sSQLServerDatabaseName & ".mdf', MOVE '" & g_sSQLServerDatabaseName & "_log' to '" & g_sAppPath & "\data\db\" & g_sSQLServerDatabaseName & ".ldf'""", vbMaximizedFocus
        MsgBox "The OASIS MSSQL Database has been initialised from a backup", vbInformation
    Else
        'MsgBox "FAILED.  File does not exist: " & g_sAppPath & "\data\db\" & Replace(frmLogin.ComServer.Text, "/", "^") & ".bak", vbError
    End If

End Sub

Private Sub DropAllTables(oCn As adodb.Connection)
        '<EhHeader>
        On Error GoTo DropAllTables_Err
        '</EhHeader>

        On Error Resume Next
    
        Dim objCatalog As New ADOX.Catalog
        Dim objTable As New ADOX.Table
        
100     Set objCatalog.ActiveConnection = oCn
        
102     For Each objTable In objCatalog.Tables

104         If Not objTable.Type = "SYSTEM TABLE" And Not objTable.Type = "SYSTEM VIEW" And Not objTable.Type = "VIEW" And Not Left(objTable.Name, 3) = "dd_" Then

106             oCn.Execute "DROP TABLE [" & objTable.Name & "]"

            End If

        Next

108     Set objCatalog = Nothing
110     Set objTable = Nothing
    
        '<EhFooter>
        Exit Sub

DropAllTables_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.modPrepareDatabase.DropAllTables " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub MSSQL_CreateTablesFromDefaultSchema(oCn As adodb.Connection)
        '<EhHeader>
        On Error GoTo MSSQL_CreateTablesFromDefaultSchema_Err
        '</EhHeader>

    'CREATE TABLES
100     oCn.Execute "CREATE TABLE [AppSettings]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[SettingDesc] ntext, " & _
            "[SettingName] nvarchar(255) Unique NOT NULL, " & _
            "[SettingValue1] ntext, " & _
            "[SettingValue2] ntext, " & _
            "[SettingValue3] ntext, " & _
            "[SettingValue4] ntext, " & _
            "[SettingValue5] ntext, " & _
            "[SettingValue6] ntext, " & _
            "[SettingValue7] ntext, " & _
            "[SettingValue8] ntext, " & _
            "[SettingValue9] ntext, " & _
            "[SettingValue10] ntext, " & _
            "PRIMARY KEY ([ID]))"
            
            oCn.Execute "INSERT INTO [AppSettings] (SettingName,SettingValue1,SettingValue2,SettingValue3,SettingValue4,SettingValue5,SettingValue6,SettingValue7,SettingValue8,SettingValue9,SettingValue10) VALUES ('ProfileSettings', '-1','-1','-1','-1','-1','-1','-1','-1','-1','-1')"
            
102     oCn.Execute "CREATE TABLE [Attachments]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[incidentID] nvarchar(255), " & _
            "[FilePath] nvarchar(255), " & _
            "[DateInserted] nvarchar(255), " & _
            "[DateModified] nvarchar(255), " & _
            "[Updated] bit, " & _
            "[AttachmentTable] nvarchar(255), " & _
            "[Description] ntext, " & _
            "[Source] ntext, " & _
            "[Copyright] ntext, " & _
            "[sGUID] nvarchar(255), " & _
            "[oBLOB] image, " & _
            "PRIMARY KEY ([ID]))"
            
104     oCn.Execute "CREATE TABLE [Draw_Layer_FEA]" & _
            "([UID] int, " & _
            "[NAME] nvarchar(255), " & _
            "[Label] nvarchar(255), " & _
            "[Style] nvarchar(255), " & _
            "PRIMARY KEY ([UID]))"
            
106     oCn.Execute "CREATE TABLE [Draw_Layer_GEO]" & _
            "([UID] int, " & _
            "[XMIN] float, " & _
            "[XMAX] float, " & _
            "[YMIN] float, " & _
            "[YMAX] float, " & _
            "[SHAPETYPE] smallint, " & _
            "[GEOMETRY] image, " & _
            "PRIMARY KEY ([UID]))"
        
108     oCn.Execute "CREATE TABLE [DynamicDataDefs]" & _
            "([DDDefName] nvarchar(255), " & _
            "[Description] ntext, " & _
            "[AccessRights] ntext, " & _
            "[ConnectionString] ntext, " & _
            "[Synch] bit, " & _
            "[EnableDataEntry] bit, " & _
            "[EnableReporting] bit, " & _
            "[ExcludedFields] ntext, " & _
            "[LockedFields] ntext)"
            
110     oCn.Execute "CREATE TABLE [Feeds]" & _
            "([FeedID] int IDENTITY(1,1), " & _
            "[GroupID] int, " & _
            "[CustomID] int, " & _
            "[FeedName] nvarchar(255), " & _
            "[FeedDescription] ntext, " & _
            "[FeedURL] ntext, " & _
            "[FeedImageURL] ntext, " & _
            "[CheckInterval] int, " & _
            "[Subscribed] bit, " & _
            "[LastCheck] date, " & _
            "PRIMARY KEY ([FeedID]))"
            
112     oCn.Execute "CREATE TABLE [GeoBookMarks]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[Name] nvarchar(255), " & _
            "[X] float, " & _
            "[Y] float, " & _
            "[Z] float, " & _
            "[Description] nvarchar(255), " & _
            "[UseSymbol] bit, " & _
            "[SymbolChar] nvarchar(255), " & _
            "[SymbolFont] nvarchar(255), " & _
            "[SymbolSize] nvarchar(255), " & _
            "[MapName] nvarchar(255), " & _
            "[BmkrID] int, " & _
            "[sGUID] nvarchar(255), " & _
            "[dTimeStamp] date, " & _
            "[OwnerGUID] nvarchar(255), " & _
            "[Deleted] bit, " & _
            "[bSubmitted] bit, " & _
            "[isURLMark] bit, " & _
            "[sURL] ntext, " & _
            "PRIMARY KEY ([ID]))"
        
114     oCn.Execute "CREATE TABLE [GeoBookMarksCategories]" & _
            "([ID] int, " & _
            "[Name] nvarchar(255), " & _
            "[Description] nvarchar(255), " & _
            "[sGUID] nvarchar(255), " & _
            "PRIMARY KEY ([ID]))"
                   
116     oCn.Execute "CREATE TABLE [GISGridTableSettings]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[Name] nvarchar(255), " & _
            "[Alias] nvarchar(255), " & _
            "[Visible] bit, " & _
            "[DatasetWarning] bit, " & _
            "[WarningLevel] int, " & _
            "[MaxRec] int, " & _
            "[ExcludedFlds] ntext, " & _
            "[IsUrlLayer] bit, " & _
            "[AutoRunUrls] bit, " & _
            "[URLLayerField] ntext, " & _
            "PRIMARY KEY ([ID]))"
        
118     oCn.Execute "CREATE TABLE [Groups]" & _
            "([GroupID] int, " & _
            "[GroupText] nvarchar(255), " & _
            "[CustomGroup] bit, " & _
            "PRIMARY KEY ([GroupID]))"
        
120     oCn.Execute "CREATE TABLE [Incidents_ChartSettings]" & _
            "([GUID1] nvarchar(255), " & _
            "[QueryName] nvarchar(255), " & _
            "[OCTSettings] image, " & _
            "[SQLCommand] ntext, " & _
            "[MSSQLCommand] ntext, " & _
            "[UseChart] bit, " & _
            "[bAutoLoadReport] bit, " & _
            "[FilterSQL] ntext, " & _
            "[FilterMSSQL] ntext, " & _
            "[Group] nvarchar(255), " & _
            "PRIMARY KEY ([GUID1]))"
        
122     oCn.Execute "CREATE TABLE [IncidentSymbology]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[Incident_Type] nvarchar(255), " & _
            "[Ascii] int, " & _
            "[Font_Name] nvarchar(255), " & _
            "[Character] nvarchar(255), " & _
            "PRIMARY KEY ([ID]))"
            
124     oCn.Execute "CREATE TABLE [IncTarget]" & _
            "([IncTargetID] int IDENTITY(1,1), " & _
            "[Name] nvarchar(255), " & _
            "[Incident_Target_Short_Name] nvarchar(255), " & _
            "[Ascii] int, " & _
            "[Character] nvarchar(255), " & _
            "[Font_Name] nvarchar(255), " & _
            "[Description] nvarchar(255), " & _
            "[Scoring] int, " & _
            "[bgColor] nvarchar(255), " & _
            "[color] nvarchar(255), " & _
            "[size] int, " & _
            "PRIMARY KEY ([IncTargetID]))"
        
126     oCn.Execute "CREATE TABLE [IncTargetCategory]" & _
            "([IncTargetID] int IDENTITY(1,1), " & _
            "[Name] nvarchar(255), " & _
            "[Incident_Target_Short_Name] nvarchar(255), " & _
            "[Ascii] int, " & _
            "[Character] nvarchar(255), " & _
            "[Font_Name] nvarchar(255), " & _
            "[Description] nvarchar(255), " & _
            "[Scoring] int, " & _
            "[bgColor] nvarchar(255), " & _
            "[color] nvarchar(255), " & _
            "PRIMARY KEY ([IncTargetID]))"
            
128     oCn.Execute "CREATE TABLE [IncTimeCategory]" & _
            "([Incident_Time_ID] int IDENTITY(1,1), " & _
            "[Incident_Time_Name] nvarchar(255), " & _
            "[Incident_Time_Short_Name] nvarchar(255), " & _
            "[Ascii] int, " & _
            "[Character] nvarchar(255), " & _
            "[Font_Name] nvarchar(255), " & _
            "[Incident_TimeDescription] nvarchar(255), " & _
            "[Scoring] int, " & _
            "[bgColor] nvarchar(255), " & _
            "[color] nvarchar(255), " & _
            "[size] int, " & _
            "PRIMARY KEY ([Incident_Time_ID]))"
            
130     oCn.Execute "CREATE TABLE [IncTypeCategory]" & _
            "([Incident_Type_ID] int IDENTITY(1,1), " & _
            "[Incident_Type_Name] nvarchar(255), " & _
            "[Incident_Type_Short_Name] nvarchar(255), " & _
            "[Ascii] int, " & _
            "[Character] nvarchar(255), " & _
            "[Font_Name] nvarchar(255), " & _
            "[Incident_TimeDescription] nvarchar(255), " & _
            "[Scoring] int, " & _
            "[bgColor] nvarchar(255), " & _
            "[color] nvarchar(255), " & _
            "[size] int, " & _
            "PRIMARY KEY ([Incident_Type_ID]))"
            
132     oCn.Execute "CREATE TABLE [InternalAppSettings]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[SettingDesc] ntext, " & _
            "[SettingName] nvarchar(255) Unique NOT NULL, " & _
            "[SettingValue1] ntext, " & _
            "[SettingValue2] ntext, " & _
            "[SettingValue3] ntext, " & _
            "[SettingValue4] ntext, " & _
            "[SettingValue5] ntext, " & _
            "[SettingValue6] ntext, " & _
            "[SettingValue7] ntext, " & _
            "[SettingValue8] ntext, " & _
            "[SettingValue9] ntext, " & _
            "[SettingValue10] ntext, " & _
            "PRIMARY KEY ([ID]))"
        
134     oCn.Execute "CREATE TABLE [Lang]" & _
            "([sGUID] nvarchar(255), " & _
            "[Name] nvarchar(255), " & _
            "[inx] int, " & _
            "[Type] nvarchar(255), " & _
            "[Container] nvarchar(255), " & _
            "[Desc] ntext, " & _
            "[Default] ntext, " & _
            "[Swedish] ntext, " & _
            "[KiSwahili] ntext, " & _
            "[German] ntext, " & _
            "[Finish] ntext, " & _
            "[Joe] ntext, " & _
            "[French] ntext, " & _
            "PRIMARY KEY ([sGUID]))"
        
136     oCn.Execute "CREATE TABLE [Maps]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[Name] nvarchar(255), " & _
            "[Alias] nvarchar(255), " & _
            "[MapPath] nvarchar(255), " & _
            "[FileName] nvarchar(255), " & _
            "[Image] nvarchar(255), " & _
            "[ThumbNail] nvarchar(255), " & _
            "[CreatedBy] nvarchar(255), " & _
            "[CreatedDate] nvarchar(255), " & _
            "[Description] ntext, " & _
            "[Contact] nvarchar(255), " & _
            "[Restrictions] nvarchar(255), " & _
            "[Copyright] nvarchar(255), " & _
            "[URL] nvarchar(255), " & _
            "[StandardLyrs] nvarchar(255), " & _
            "[Source] nvarchar(255), " & _
            "[AdminLyr1Name] nvarchar(255), " & _
            "[AdminLyr2Name] nvarchar(255), " & _
            "[AdminLyr3Name] nvarchar(255), " & _
            "[AdminLyr4Name] nvarchar(255), " & _
            "[AdminLyr5Name] nvarchar(255), " & _
            "PRIMARY KEY ([ID]))"
        
'138     oCn.Execute "CREATE TABLE [MapStyles]" & _
'            "([ID] int IDENTITY(1,1), " & _
'            "[StyleName] nvarchar(255), " & _
'            "[LineColor] int, [LineInterleaved] bit, [LineStyle] int, [LineSupportsInterleave] bit, [LineWidth] int, [LineWidthUnit] int, " & _
'            "[MaxVectorSymbolCharacter] int, " & _
'            "[MinVectorSymbolCharacter] int, " & _
'            "[RegionBackColor] int, [RegionBorderColor] int, [RegionBorderStyle] int, [RegionBorderWidth] int, [RegionBorderWidthUnit] int, [RegionColor] int, [RegionPattern] int, [RegionTransparent] bit, " & _
'            "[SupportsBitmapSymbols] bit, " & _
'            "[SymbolBitmapColor] int, [SymbolBitmapTransparent] bit, [SymbolBitmapOverrideColor] bit, [SymbolBitmapName] nvarchar(255), [SymbolBitmapSize] int, [SymbolCharacter] int, [SymbolFont] nvarchar(255), [SymbolFontBackColor] int, [SymbolFontColor] int, [SymbolFontHalo] bit, [SymbolFontOpaque] bit, [SymbolFontRotation] int, [SymbolFontSize] int, [SymbolFontShadow] bit, [SymbolType] int, [SymbolVectorColor] int, [SymbolVectorSize] int, " & _
'            "[TextFontBold] bit, [TextFontCharset] int, [TextFontItalic] bit, [TextFontName] nvarchar(255), [TextFontSize] int, [TextFontStrikethrough] bit, [TextFontUnderline] bit, [TextFontWeight] int, [TextFontAllCaps] bit, [TextFontBackColor] int, [TextFontColor] int, [TextFontDblSpace] bit, [TextFontHalo] bit, [TextFontOpaque] bit, [TextFontRotation] int, [TextFontShadow] bit, " & _
'            "PRIMARY KEY ([ID]))"
            
'140     oCn.Execute "CREATE TABLE [MapThemes]" & _
'            "([ID] int IDENTITY(1,1), " & _
'            "[MapsID] nvarchar(255), " & _
'            "[ThemeID] nvarchar(255), " & _
'            "PRIMARY KEY ([ID]))"
            
            If Not DoesTableExist(g_sGlobalConnectionString, "oincidentstrans") Then
142             oCn.Execute "CREATE TABLE [oincidentstrans]" & _
                    "([UID] int Unique NOT NULL, " & _
                    "[ID] nvarchar(255), " & _
                    "[Name] nvarchar(255), " & _
                    "[Type] nvarchar(255), " & _
                    "[Target] nvarchar(255), " & _
                    "[Incident_Date] date, " & _
                    "[Time00] nvarchar(255), " & _
                    "[Town] nvarchar(255), " & _
                    "[District] nvarchar(255), " & _
                    "[Province] nvarchar(255), " & _
                    "[Description] nvarchar(255), " & _
                    "[Scoring] int, " & _
                    "[TimeStamp] nvarchar(255), " & _
                    "[GUID] nvarchar(255), " & _
                    "[ReportID] int, " & _
                    "PRIMARY KEY ([UID]))"
            End If
            
            If Not DoesTableExist(g_sGlobalConnectionString, "oincidents_FEA") Then
144             oCn.Execute "CREATE TABLE [oincidents_FEA]" & _
                    "([UID] int Unique NOT NULL, " & _
                    "[ID] nvarchar(255) Unique NOT NULL, " & _
                    "[Name] nvarchar(255), " & _
                    "[Type] nvarchar(255), " & _
                    "[Target] nvarchar(255), " & _
                    "[Dead] int, " & _
                    "[Affected] int, " & _
                    "[Violent] int, " & _
                    "[Injured] int, " & _
                    "[Incident_Date] date, " & _
                    "[Time00] nvarchar(255), " & _
                    "[LocDesc] ntext, " & _
                    "[Source] nvarchar(255), " & _
                    "[Town] nvarchar(255), " & _
                    "[District] nvarchar(255), " & _
                    "[Province] nvarchar(255), " & _
                    "[Description] ntext, " & _
                    "[Scoring] int, " & _
                    "[Incident_DateSERIAL] int, " & _
                    "PRIMARY KEY ([UID]))"
            End If
            
            If Not DoesTableExist(g_sGlobalConnectionString, "oincidents_GEO") Then
146             oCn.Execute "CREATE TABLE [oincidents_GEO]" & _
                    "([UID] int, " & _
                    "[XMIN] float, " & _
                    "[XMAX] float, " & _
                    "[YMIN] float, " & _
                    "[YMAX] float, " & _
                    "[SHAPETYPE] smallint, " & _
                    "[GEOMETRY] image, " & _
                    "PRIMARY KEY ([UID]))"
            End If
            
148     oCn.Execute "CREATE TABLE [Personnell]" & _
            "([Personnell_ID] nvarchar(255), " & _
            "[FirstName] nvarchar(255), " & _
            "[FamilyName] nvarchar(255), " & _
            "[Position] nvarchar(255), " & _
            "[UserGroup] nvarchar(255), " & _
            "[RegisteredBy_ID] nvarchar(255), " & _
            "[UserName] nvarchar(255), " & _
            "[pwd] nvarchar(255), " & _
            "[OrganisationID] int, " & _
            "[MedicalDetails] ntext, " & _
            "[CallSign] nvarchar(255), " & _
            "[DOB] nvarchar(255), " & _
            "[Photo] image, " & _
            "[IssuedRadio] bit, " & _
            "[RadioSNR] nvarchar(255), " & _
            "[DoneRequiredSecurityTraining] bit, " & _
            "[ActiveTrackingID] int, " & _
            "[CurrentLocationID] nvarchar(255), " & _
            "[DefaultViewName] nvarchar(255), " & _
            "[DefaultViewX] float, " & _
            "[DefaultViewY] float, " & _
            "[DefaultViewZ] float, " & _
            "[LatestViewX] float, [LatestViewY] float, [LatestViewZ] float, [LatestMapName] nvarchar(255), " & _
            "PRIMARY KEY ([Personnell_ID]))"
            
150     oCn.Execute "INSERT INTO [Personnell] (Personnell_ID,UserName,pwd) VALUES (2, 'bart', 'simpson')"
        
152     oCn.Execute "CREATE TABLE [PrintTemplates]" & _
            "([id] int IDENTITY(1,1), " & _
            "[Name] nvarchar(255), " & _
            "[Format] nvarchar(255), " & _
            "[Width] int, " & _
            "[Height] int, " & _
            "[FileName] nvarchar(255), " & _
            "[Description] ntext, " & _
            "[IsLandscape] bit, " & _
            "[IsPortrait] bit, " & _
            "[Note] ntext, " & _
            "[Copyright] ntext, " & _
            "[MapTitle] nvarchar(255), " & _
            "[MapSubTitle] nvarchar(255), " & _
            "[MapIDPrefix] nvarchar(255), " & _
            "PRIMARY KEY ([id]))"
        
154     oCn.Execute "CREATE TABLE [Render]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[RenderName] nvarchar(255), " & _
            "[Item_Using] nvarchar(255), " & _
            "[Expression] nvarchar(255), " & _
            "[Chart] nvarchar(255), " & _
            "[MinVal] nvarchar(255), " & _
            "[MinValEx] nvarchar(255), " & _
            "[MaxVal] nvarchar(255), " & _
            "[MaxValEx] nvarchar(255), " & _
            "[ColorDefault] nvarchar(255), " & _
            "[StartColor] nvarchar(255), " & _
            "[StartColorEx] nvarchar(255), " & _
            "[EndColor] nvarchar(255), " & _
            "[EndColorEx] nvarchar(255), " & _
            "[SizeDefault] nvarchar(255), " & _
            "[StartSize] nvarchar(255), " & _
            "[StartSizeEx] nvarchar(255), " & _
            "[EndSize] nvarchar(255), " & _
            "[EndSizeEx] nvarchar(255), " & _
            "[Round] nvarchar(255), " & _
            "[Zones] nvarchar(255), " & _
            "[ZonesEx] nvarchar(255), " & _
            "PRIMARY KEY ([ID]))"
            
            
158     oCn.Execute "CREATE TABLE [Style]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[StyleName] nvarchar(255), [Item_Using] nvarchar(255), [Color] nvarchar(255), " & _
            "[Symbol] nvarchar(255), [SymbolSize] nvarchar(255), [SymbolGap] nvarchar(255), [SymbolRotate] nvarchar(255), " & _
            "[Bitmap] nvarchar(255), [Pattern] nvarchar(255), " & _
            "[OutlineColor] nvarchar(255), [OutlineWidth] nvarchar(255), [OutlineStyle] nvarchar(255), [OutlineSymbol] nvarchar(255), [OutlineSymbolGap] nvarchar(255), [OutlineSymbolRotation] nvarchar(255), [OutlineBitmap] nvarchar(255), [OutlinePattern] nvarchar(255), " & _
            "[SmartSize] nvarchar(255), [SmartSizeField] nvarchar(255), " & _
            "[Width] nvarchar(255), [Visible] nvarchar(255), [Allocator] nvarchar(255), [Dublicates] nvarchar(255), [Field] nvarchar(255), " & _
            "[Value] nvarchar(255), [Alignment] nvarchar(255), [Position] nvarchar(255), " & _
            "[FontName] nvarchar(255), [FontStyle] nvarchar(255), [FontColor] nvarchar(255), [Height] nvarchar(255), [Style] nvarchar(255), [Size] nvarchar(255), [Values] nvarchar(255), " & _
            "[Red band assignment] nvarchar(255), [Green band assignment] nvarchar(255), [Blue band assignment] nvarchar(255), [Grid band assignment] nvarchar(255), [Grid no-value assignment] nvarchar(255), [Grid shadow assignment] nvarchar(255), [Red brightness] nvarchar(255), [Green brightness] nvarchar(255), [Blue brightness] nvarchar(255), [Inversion] nvarchar(255), [GrayMapZones] nvarchar(255), [TransparentZones] nvarchar(255), [RedMapZones] nvarchar(255), " & _
            "[GrayScale] nvarchar(255), " & _
            "[GreenMapZones] nvarchar(255), " & _
            "[BlueMapZones] nvarchar(255), " & _
            "[Histogram] nvarchar(255), " & _
            "[HistogramPath] nvarchar(255), " & _
            "PRIMARY KEY ([ID]))"
            
           If Not DoesTableExist(g_sGlobalConnectionString, "SynchHistory") Then
160             oCn.Execute "CREATE TABLE [SynchHistory]" & _
                    "([sID] nvarchar(255), " & _
                    "[sGUID] nvarchar(255), " & _
                    "[sTableName] nvarchar(255), " & _
                    "[sWhen] nvarchar(255), " & _
                    "[sStatus] nvarchar(255), " & _
                    "[Sequence] int, " & _
                    "[sBy] nvarchar(255), " & _
                    "[sDelete] nvarchar(255), " & _
                    "[Updates] nvarchar(255), " & _
                    "[NoConflict] nvarchar(255), " & _
                    "PRIMARY KEY ([sID]))"
            
            End If
                      
164     oCn.Execute "CREATE TABLE [ThemeGroups]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[Name] nvarchar(255), " & _
            "[Description] ntext, " & _
            "PRIMARY KEY ([ID]))"
        
166     oCn.Execute "CREATE TABLE [Themes]" & _
            "([ID] int IDENTITY(1,1), " & _
            "[Name] ntext, " & _
            "[ThemeGroup] int, " & _
            "[Description] ntext, " & _
            "[AnalysisField] nvarchar(255), " & _
            "[Maps] ntext, " & _
            "[AnalysisLayer] nvarchar(255), " & _
            "[ThemeConfigName] nvarchar(255), " & _
            "PRIMARY KEY ([ID]))"
            
168     oCn.Execute "CREATE TABLE [ttkGISLayerSQL]" & _
            "([Name] nvarchar(255), " & _
            "[XMIN] float, " & _
            "[XMAX] float, " & _
            "[YMIN] float, " & _
            "[YMAX] float, " & _
            "[SHAPETYPE] int, " & _
            "PRIMARY KEY ([Name]))"
            
170     oCn.Execute "CREATE TABLE [ttkGISLayerSQLInProject]" & _
            "([LayerName] nvarchar(255), " & _
            "[LayerCaption] nvarchar(255), " & _
            "[Transparency] int, " & _
            "[IsVisible] bit, " & _
            "[IsExpanded] bit, " & _
            "[Dialect] nvarchar(255), " & _
            "[ADO] ntext, " & _
            "[Sequence] int, " & _
            "[XMIN] float, " & _
            "[XMAX] float, " & _
            "[YMIN] float, " & _
            "[YMAX] float, " & _
            "[SHAPETYPE] int, " & _
            "[INISettings] ntext, " & _
            "PRIMARY KEY ([LayerCaption]))"
            
172     oCn.Execute "CREATE TABLE [ttkGISProjectDef]" & _
            "([InUse] bit, " & _
            "[MapData] ntext, " & _
            "[sGUID] nvarchar(255), " & _
            "[XMIN] float, " & _
            "[XMAX] float, " & _
            "[YMIN] float, " & _
            "[YMAX] float, " & _
            "PRIMARY KEY ([sGUID]))"
        
        '<EhFooter>
        Exit Sub

MSSQL_CreateTablesFromDefaultSchema_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.modPrepareDatabase.MSSQL_CreateTablesFromDefaultSchema " & _
               "at line " & Erl
        'Resume Next
        '</EhFooter>
End Sub















