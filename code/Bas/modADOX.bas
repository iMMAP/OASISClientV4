Attribute VB_Name = "modADOX"

Declare Function CoCreateGuid _
        Lib "ole32.dll" (pguid As GUID) As Long
Declare Function StringFromGUID2 _
        Lib "ole32.dll" (rguid As Any, _
                         ByVal lpstrClsId As Long, _
                         ByVal cbMax As Long) As Long
Declare Sub ExitProcess _
        Lib "kernel32" (ByVal uExitCode As Long)

Declare Function GetInputState _
        Lib "user32" () As Long

'GUID STRUCT
Private Type GUID
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4(8) As Byte
End Type

Public Declare Function GetFileVersionInfo _
                Lib "Version.dll" _
                Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, _
                                             ByVal dwhandle As Long, _
                                             ByVal dwlen As Long, _
                                             lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize _
                Lib "Version.dll" _
                Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
                                                 lpdwHandle As Long) As Long
Public Declare Function VerQueryValue _
                Lib "Version.dll" _
                Alias "VerQueryValueA" (pBlock As Any, _
                                        ByVal lpSubBlock As String, _
                                        lplpBuffer As Any, _
                                        puLen As Long) As Long
Private Declare Sub MoveMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Dest As Any, _
                                       ByVal Source As Long, _
                                       ByVal Length As Long)
                                       
'Public g_PictureDialogLarge As StdPicture
'Public g_PictureDialogSmall As StdPicture
'Public g_PictureDialogLogo As StdPicture

Public Sub KillALL()
        '<EhHeader>
        On Error GoTo KillALL_Err
        '</EhHeader>
        On Error Resume Next
    
        Dim i As Integer
    
100     For i = 1 To Forms.Count
102         Unload Forms(i)
104         Set Forms(i) = Nothing
        Next
    
106     DoEvents
108     ExitProcess 1
    
110     End
        '<EhFooter>
        Exit Sub

KillALL_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.KillALL " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

'Public Sub Main()
    '        '<EhHeader>
    '        On Error GoTo Main_Err
    '        '</EhHeader>
    '
    '100     frmLogin.Show vbModal
    '
    '102     If frmLogin.LoginSucceeded Then FrmMain.Show
    '
    '        On Error Resume Next
    '
    '104     Unload frmLogin
    '
    '    '    Load frmConnectionString
    '    '    Load frmDSDefinitions
    '    '    Load frmProgress
    '    '    Load frmServerStatus
    '    '    Load frmSQLchecker
    '    '    Load frmTimeDateTest
    '    '
    '    '    Dim frmloaded As Form
    '    '    Dim i As Integer
    '    '    Dim sVal As String
    '    '
    '    '    For Each frmloaded In Forms
    '    '        DebugPrint "<" & frmloaded.Name & ">"
    '    '        On Error Resume Next
    '    '        'Loop through all form controls
    '    '
    '    '        For i = 0 To frmloaded.Controls.Count - 1
    '    '            sVal = ""
    '    '            sVal = frmloaded.Controls(i).Caption
    '    '
    '    '            If Not Len(sVal) < 1 Then
    '    '                sVal = frmloaded.Controls(i).Name & ".Caption =" & sVal
    '    '            Else
    '    '                sVal = frmloaded.Controls(i).Name & ".Text =" & sVal & frmloaded.Controls(i).Text
    '    '            End If
    '    '
    '    '            DebugPrint sVal
    '    '        Next
    '    '    Next
    '
    '        '<EhFooter>
    '        Exit Sub
    '
    'Main_Err:
    '        MsgBox Err.Description & vbCrLf & _
    '               "in OASISRemoteAdmin.modADOX.Main " & _
    '               "at line " & Erl
    '        Resume Next
    '        '</EhFooter>
'End Sub

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
    
112     GUIDGen = left$(sGUID, RetVal)

        '<EhFooter>
        Exit Function

GUIDGen_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.GUIDGen " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function SetSeed(strTable As String, strAutoNum As String, lngID As Long, CurrentProjectConnection As ADODB.Connection) As Boolean
    'Purpose:   Set the Seed of an AutoNumber using ADOX.
    Dim cat As New ADOX.Catalog
    
    Set cat.ActiveConnection = CurrentProjectConnection
    cat.Tables(strTable).Columns(strAutoNum).Properties("Seed") = lngID
    Set cat = Nothing
    SetSeed = True
End Function

Public Function CreateTable(sTableName As String, _
                            rst As ADODB.Recordset, _
                            oConn As ADODB.Connection)
        '<EhHeader>
        On Error GoTo CreateTable_Err
        '</EhHeader>
                            
        Dim colField As ADOX.Column
        Dim prpField As ADOX.Property
        Dim strField As String
        Dim x As Long
        Dim y As Long
        Dim lngType As Long
        Dim lngSize As Long
    
        Dim rstnew  As ADODB.Recordset
        Dim mcatDB As ADOX.Catalog
        Dim mtblNew As ADOX.Table
    
        If DoTableExists(sTableName, oConn) Then Exit Function
        
100     Set mcatDB = New ADOX.Catalog
102     mcatDB.ActiveConnection = oConn
    
        Set mtblNew = New ADOX.Table
104     Set mtblNew.ParentCatalog = mcatDB
108     mtblNew.Name = sTableName
   
110     With rst

112         For x = 0 To rst.Fields.Count - 1
114             strField = .Fields(x).Name
116             strField = Replace(strField, ".", "_")
118             strField = Replace(strField, "!", "_")
120             strField = Replace(strField, "`", "_")
122             strField = Replace(strField, "[", "_")
124             strField = Replace(strField, "]", "_")
126             Set colField = New ADOX.Column
128             colField.Name = strField
130             lngType = .Fields(x).Type
132             lngSize = .Fields(x).DefinedSize

134             Select Case lngType

                    Case adChar
136                     colField.Type = adWChar
138                     colField.DefinedSize = lngSize

140                 Case adVarChar
142                     colField.Type = adVarWChar
144                     colField.DefinedSize = lngSize

146                 Case adLongVarChar
148                     colField.Type = adLongVarWChar
150                     colField.DefinedSize = lngSize

152                 Case adNumeric
154                     colField.Type = adNumeric

156                     If .Fields(x).Precision > 18 Then
158                         colField.Precision = 18
160                         colField.NumericScale = 2
                        Else
162                         colField.Precision = .Fields(x).Precision
164                         colField.NumericScale = .Fields(x).NumericScale
                        End If

166                 Case adDBTimeStamp
168                     colField.Type = adDate

170                 Case Else
172                     colField.Type = lngType
                        
                End Select

174             mtblNew.Columns.Append colField
176             mtblNew.Columns.Item(colField.Name).Properties("Nullable").value = True
178             mtblNew.Columns.Item(colField.Name).Properties("Jet OLEDB:Allow Zero Length").value = True
180         Next x

        End With

182     mcatDB.Tables.Append mtblNew

        'DATA _________________________________________________________________
    
        Set rstnew = New ADODB.Recordset
184     rstnew.Open "SELECT * FROM " & sTableName, mcatDB.ActiveConnection, adOpenDynamic, adLockBatchOptimistic

186     Do Until rst.EOF
188         rstnew.AddNew

190         For x = 0 To rst.Fields.Count - 1
192             rstnew.Fields(x).value = rst.Fields(x).value
194         Next x

196         rstnew.UpdateBatch
198         rst.MoveNext
        Loop

200     rstnew.Close
202     Set rstnew = Nothing
    
        '<EhFooter>
        Exit Function

CreateTable_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.modADOX.CreateTable " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function CompareTables(sSourceConn As String, _
                               sSourceTableName As String, _
                               sTargetConn As String, _
                               sTargetTableName As String) As Boolean
        '<EhHeader>
        On Error GoTo CompareTables_Err
        '</EhHeader>
    
        Dim sourceConn As New ADODB.Connection
        Dim sourceTablesSchema As ADODB.Recordset
        Dim sourceColumnsSchema As ADODB.Recordset
    
        Dim targetConn As New ADODB.Connection
        Dim targetTablesSchema As ADODB.Recordset
        Dim targetColumnsSchema As ADODB.Recordset
    
        Dim bBreakOutFlag As Boolean
100     bBreakOutFlag = False

102     If DoesTableExist(sSourceConn, sSourceTableName) And DoesTableExist(sTargetConn, sTargetTableName) Then
    
104         sourceConn.ConnectionString = sSourceTableName
106         sourceConn.Open
    
108         targetConn.ConnectionString = sTargetTableName
110         targetConn.Open
    
112         Set sourceTablesSchema = sourceConn.OpenSchema(adSchemaTables)
114         Set targetTablesSchema = targetConn.OpenSchema(adSchemaTables)
    
116         sourceTablesSchema.Find "TABLE_NAME = " & sSourceTableName
118         targetTablesSchema.Find "TABLE_NAME = " & sTargetTableName
    
120         Set sourceColumnsSchema = sourceConn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "" & sourceTablesSchema("TABLE_NAME")))
122         Set targetColumnsSchema = targetConn.OpenSchema(adSchemaColumns, Array(Empty, Empty, "" & targetTablesSchema("TABLE_NAME")))

124         Do While Not sourceColumnsSchema.EOF And Not bBreakOutFlag
            
126             targetColumnsSchema.Find "COLUMN_NAME = " & sourceColumnsSchema.Fields("COLUMN_NAME").value
128             If targetColumnsSchema.EOF Then
130                 bBreakOutFlag = True
132                 CompareTables = False
134             ElseIf Not sourceColumnsSchema.Fields("DATA_TYPE").value = targetColumnsSchema.Fields("DATA_TYPE").value Then
136                 bBreakOutFlag = True
138                 CompareTables = False
                Else
140                 bBreakOutFlag = False
142                 CompareTables = True
                End If
            
144             sourceColumnsSchema.MoveNext
            Loop
     
        End If

        '<EhFooter>
        Exit Function

CompareTables_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.modADOX.CompareTables " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DoesTableExist(sConn As String, _
                                sTableName As String) As Boolean
        '<EhHeader>
        On Error GoTo DoesTableExist_Err
        '</EhHeader>
    
        Dim Conn As New ADODB.Connection
        Dim TablesSchema As ADODB.Recordset
        Dim ColumnsSchema As ADODB.Recordset
        Conn.CursorLocation = g_sGlobalCursorLocation
100     Conn.ConnectionString = sConn
102     Conn.Open

104     Set TablesSchema = Conn.OpenSchema(adSchemaTables)

106     TablesSchema.Find "[TABLE_NAME] = '" & sTableName & "'"
    
108     If TablesSchema.EOF Then
110         DoesTableExist = False
        Else
112         DoesTableExist = True
        End If
        
        '<EhFooter>
        Exit Function

DoesTableExist_Err:
    Err.Clear
    DoesTableExist = False
'</EhFooter>
End Function

Function ShowAllTables(CurrentProjectConnection As ADODB.Connection, Optional bShowFieldsToo As Boolean, Optional oComb As ComboBox)
        'Purpose:   List the tables (and optionally their fields) using ADOX.
        '<EhHeader>
        On Error GoTo ShowAllTables_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog 'Root object of ADOX.
        Dim tbl As ADOX.Table       'Each Table in Tables.
        Dim Col As ADOX.Column      'Each Column in the Table.
    
        'Point the catalog to the current project's connection.
100     Set cat.ActiveConnection = CurrentProjectConnection
    
        'Loop through the tables.
102     For Each tbl In cat.Tables

104         If Not oComb Is Nothing Then
106             If tbl.Type = "TABLE" Then
108                 If Not left(tbl.Name, 1) = "~" Then
110                     oComb.AddItem tbl.Name
                    End If
                End If
            
            Else
112             DebugPrint tbl.Name
                DebugPrint tbl.Type
            End If
        
114         If bShowFieldsToo Then

                'Loop through the columns of the table.
116             For Each Col In tbl.Columns
118                 DebugPrint Col.Name
                    DebugPrint Col.Type
                Next

120             DebugPrint "--------------------------------"
                'Stop
            End If

        Next
    
        'Clean up
122     Set Col = Nothing
124     Set tbl = Nothing
126     Set cat = Nothing
        '<EhFooter>
        Exit Function

ShowAllTables_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.modADOX.ShowAllTables " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ShowPropsADOX(CurrentProjectConnection As ADODB.Connection, strTable As String, Optional bShowPropertiesToo As Boolean)
        'Purpose:   Show the columns in a table, and optionally their properties, using ADOX.
        '<EhHeader>
        On Error GoTo ShowPropsADOX_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog 'Root object of ADOX.
        Dim tbl As ADOX.Table       'Each Table in Tables.
        Dim Col As ADOX.Column      'Each Column in the Table.
        Dim prp As ADOX.Property
    
        'Point the catalog to the current project's connection.
100     Set cat.ActiveConnection = CurrentProjectConnection
102     Set tbl = cat.Tables(strTable)
    
104     For Each Col In tbl.Columns
106         DebugPrint Col.Name        ', col.Properties("Fixed length"), col.Type

108         If bShowPropertiesToo Then

110             For Each prp In Col.Properties
112                 DebugPrint prp.Name
                    DebugPrint prp.Type
                    DebugPrint prp.value
                Next

114             DebugPrint "--------------------------------"
                'Stop
            End If

        Next
    
        'Clean up
116     Set prp = Nothing
118     Set Col = Nothing
120     Set tbl = Nothing
122     Set cat = Nothing
        '<EhFooter>
        Exit Function

ShowPropsADOX_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.ShowPropsADOX " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function CreateTableAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Create a table with various field types, using ADOX.
        '<EhHeader>
        On Error GoTo CreateTableAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim tbl As ADOX.Table
    
100     Set cat.ActiveConnection = CurrentProjectConnection
        'Initialize the Contractor table.
102     Set tbl = New ADOX.Table
104     tbl.Name = "tblAdoxContractor"
    
        'Append the columns.
106     With tbl.Columns
108         .Append "ContractorID", adInteger   'Number (Long Integer)
110         .Append "Surname", adVarWChar, 30   'Text (30 max)
112         .Append "FirstName", adVarWChar, 20 'Text (20 max)
114         .Append "Inactive", adBoolean       'Yes/No
116         .Append "HourlyFee", adCurrency     'Currency
118         .Append "PenaltyRate", adDouble     'Number (Double)
120         .Append "BirthDate", adDate         'Date/Time
122         .Append "Notes", adLongVarWChar     'Memo
124         .Append "Web", adLongVarWChar       'Memo (for hyperlink)
        
            'Set the field properties.
            'AutoNumber
126         With !ContractorID
128             Set .ParentCatalog = cat
            
130             .Properties("Autoincrement") = True     'AutoNumber.
132             .Properties("Description") = "Automatically " & "generated unique identifier for this record."
            End With
        
            'Required field.
134         With !Surname
136             Set .ParentCatalog = cat
138             .Properties("Nullable") = False         'Required.
140             .Properties("Jet OLEDB:Allow Zero Length") = False
            End With
        
            'Set a validation rule.
142         With !BirthDate
144             Set .ParentCatalog = cat
146             .Properties("Jet OLEDB:Column Validation Rule") = "Is Null Or <=Date()"
148             .Properties("Jet OLEDB:Column Validation Text") = "Birth date cannot be future."
            End With
        
            'Hyperlink field.
150         With !web
152             Set .ParentCatalog = cat
154             .Properties("Jet OLEDB:Hyperlink") = True 'Hyperlink.
            End With
        End With
    
        'Save the new table by appending to catalog.
156     cat.Tables.Append tbl
158     DebugPrint "tblAdoxContractor created."
160     Set tbl = Nothing
    
        'Initialize the Booking table
162     Set tbl = New ADOX.Table
164     tbl.Name = "tblAdoxBooking"
    
        'Append the columns.
166     With tbl.Columns
168         .Append "BookingID", adInteger
170         .Append "BookingDate", adDate
172         .Append "ContractorID", adInteger
174         .Append "BookingFee", adCurrency
176         .Append "BookingNote", adWChar, 255
        
            'Set the field properties.
178         With !BookingID                             'AutoNumber.
180             .ParentCatalog = cat
182             .Properties("Autoincrement") = True
            End With

184         With !BookingNote                           'Required.
186             .ParentCatalog = cat
188             .Properties("Nullable") = False
190             .Properties("Jet OLEDB:Allow Zero Length") = False
            End With
        End With
    
        'Save the new table by appending to catalog.
192     cat.Tables.Append tbl
194     DebugPrint "tblAdoxBooking created."
    
        'Clean up
196     Set tbl = Nothing
198     Set cat = Nothing
        '<EhFooter>
        Exit Function

CreateTableAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.CreateTableAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function CopyTableAdox(CurrentProjectConnection As ADODB.Connection, sTableName As String, sTemplateTableName As String)
        'Purpose:   Create a table with various field types, using ADOX.
        '<EhHeader>
        On Error GoTo CopyTableAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog 'Root object of ADOX.
        Dim tblTemplate As ADOX.Table       'Each Table in Tables.
        Dim Col As ADOX.Column      'Each Column in the Table.
        Dim tbl As ADOX.Table
        Dim CurrentProperty As ADOX.Property
    
100     Set cat.ActiveConnection = CurrentProjectConnection
        'Initialize the table.
102     Set tbl = New ADOX.Table
104     tbl.Name = sTableName
106     Set tbl.ParentCatalog = cat
108     Set tblTemplate = cat.Tables.Item(sTemplateTableName)
    
110     With tbl.Columns
      
            'Loop through the columns of the table.
112         For Each Col In tblTemplate.Columns
                'Append the columns.
114             .Append Col.Name, Col.Type, Col.DefinedSize
                '.Item(col.Name).Properties("Nullable") = True
                '.Item(col.Name).Properties("Jet OLEDB:Allow Zero Length").Value = True
            
116             For Each CurrentProperty In Col.Properties
118                 .Item(Col.Name).Properties(CurrentProperty.Name) = CurrentProperty.value
                Next
            
            Next
        
        End With
    
        'Save the new table by appending to catalog.
120     cat.Tables.Append tbl
    
122     CurrentProjectConnection.Execute "INSERT INTO " & sTableName & " SELECT * FROM " & sTemplateTableName
    
        'Clean up
124     Set tbl = Nothing
126     Set cat = Nothing
        '<EhFooter>
        Exit Function

CopyTableAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.CopyTableAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function AddFieldToTable(CurrentProjectConnection As ADODB.Connection, sTableName As String, Col As ADOX.Column)
        '<EhHeader>
        On Error GoTo ModifyTableAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim tbl As ADOX.Table
    
        'Initialize
100     cat.ActiveConnection = CurrentProjectConnection
102     Set tbl = cat.Tables(sTableName)
    
114     tbl.Columns.Append Col
'        tbl.Columns.Refresh
        
116     Set Col = Nothing
    
126     Set tbl = Nothing
128     Set cat = Nothing
        '<EhFooter>
        Exit Function

ModifyTableAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.ModifyTableAdox " & "at line " & Erl
        '</EhFooter>

End Function

Function ModifyTableAdox(CurrentProjectConnection As ADODB.Connection, sTableName As String)
        '<EhHeader>
        On Error GoTo ModifyTableAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim tbl As ADOX.Table
        Dim Col As New ADOX.Column
    
        'Initialize
100     cat.ActiveConnection = CurrentProjectConnection
102     Set tbl = cat.Tables(sTableName)
    
        'Add a new column
104     With Col
106         .Name = "MyDecimal"
108         .Type = adNumeric   'Decimal type.
110         .Precision = 28     '28 digits.
112         .NumericScale = 8   '8 decimal places.
        End With

114     tbl.Columns.Append Col
116     Set Col = Nothing
118     DebugPrint "Column added."
    
        'Delete a column.
120     tbl.Columns.Delete "MyDecimal"
122     DebugPrint "Column deleted."
    
        'Clean up
124     Set Col = Nothing
126     Set tbl = Nothing
128     Set cat = Nothing
        '<EhFooter>
        Exit Function

ModifyTableAdox_Err:
        'MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.ModifyTableAdox " & "at line " & Erl
        'Resume Next
        '</EhFooter>
End Function

Function ModifyFieldPropAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Show how to alter field properties, using ADOX.
        'Note:      You cannot alter the DefinedSize of the field like this.
        '<EhHeader>
        On Error GoTo ModifyFieldPropAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim Col As ADOX.Column
        Dim prp As ADOX.Property

100     cat.ActiveConnection = CurrentProjectConnection
102     Set Col = cat.Tables("MyTable").Columns("MyField")
        'col.ParentCatalog = cat
104     Set prp = Col.Properties("Nullable")
        'Read the property
106     DebugPrint prp.Name
        DebugPrint prp.value
        DebugPrint prp.Type = adBoolean
        'Change the property
108     prp.value = Not prp.value
    
        'Clean up
110     Set prp = Nothing
112     Set Col = Nothing
114     Set cat = Nothing
        '<EhFooter>
        Exit Function

ModifyFieldPropAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.ModifyFieldPropAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function DeleteTableAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Delete a table using ADOX.
        '<EhHeader>
        On Error GoTo DeleteTableAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
    
100     cat.ActiveConnection = CurrentProjectConnection
102     cat.Tables.Delete "MyTable"
104     Set cat = Nothing
        '<EhFooter>
        Exit Function

DeleteTableAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.DeleteTableAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function CreateIndexesAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Show how to create indexes using ADOX.
        '<EhHeader>
        On Error GoTo CreateIndexesAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim tbl As ADOX.Table
        Dim ind As ADOX.Index
    
        'Initialize
100     Set cat.ActiveConnection = CurrentProjectConnection
102     Set tbl = cat.Tables("tblAdoxContractor")

        'Create a primary key index
104     Set ind = New ADOX.Index
106     ind.Name = "PrimaryKey"
108     ind.PrimaryKey = True
110     ind.Columns.Append "ContractorID"
112     tbl.Indexes.Append ind
114     Set ind = Nothing
    
        'Create an index on one column.
116     Set ind = New ADOX.Index
118     ind.Name = "Inactive"
120     ind.Columns.Append "Inactive"
122     tbl.Indexes.Append ind
124     Set ind = Nothing
    
        'Multi-field index.
126     Set ind = New ADOX.Index
128     ind.Name = "FullName"

130     With ind.Columns
132         .Append "Surname"
134         .Append "FirstName"
        End With

136     tbl.Indexes.Append ind
    
        'Clean up
138     Set ind = Nothing
140     Set tbl = Nothing
142     Set cat = Nothing
144     DebugPrint "tblAdoxContractor indexes created."
        '<EhFooter>
        Exit Function

CreateIndexesAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.CreateIndexesAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function DeleteIndexAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Show how to delete indexes using ADOX.
        '<EhHeader>
        On Error GoTo DeleteIndexAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
100     cat.ActiveConnection = CurrentProjectConnection
102     cat.Tables("tblAdoxContractor").Indexes.Delete "Inactive"
104     Set cat = Nothing
        '<EhFooter>
        Exit Function

DeleteIndexAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.DeleteIndexAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function CreateKeyAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Show how to create relationships using ADOX.
        '<EhHeader>
        On Error GoTo CreateKeyAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim tbl As ADOX.Table
        Dim ky As New ADOX.Key
    
100     Set cat.ActiveConnection = CurrentProjectConnection
102     Set tbl = cat.Tables("tblAdoxBooking")
    
        'Create as foreign key to tblAdoxContractor.ContractorID
104     With ky
106         .Type = adKeyForeign
108         .Name = "tblAdoxContractortblAdoxBooking"
110         .RelatedTable = "tblAdoxContractor"
112         .Columns.Append "ContractorID"      'Just one field.
114         .Columns("ContractorID").RelatedColumn = "ContractorID"
116         .DeleteRule = adRISetNull   'Cascade to Null on delete.
        End With

118     tbl.Keys.Append ky
    
120     Set ky = Nothing
122     Set tbl = Nothing
124     Set cat = Nothing
126     DebugPrint "Key created."
        '<EhFooter>
        Exit Function

CreateKeyAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.CreateKeyAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ShowKeyAdox(strTableName As String, CurrentProjectConnection As ADODB.Connection)
        'Purpose:   List relationships using ADOX.
        '<EhHeader>
        On Error GoTo ShowKeyAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim tbl As ADOX.Table
        Dim ky As ADOX.Key
        Dim strRIName As String
    
100     Set cat.ActiveConnection = CurrentProjectConnection
102     Set tbl = cat.Tables(strTableName)
    
104     For Each ky In tbl.Keys

106         With ky

108             Select Case .DeleteRule

                    Case adRINone
110                     strRIName = "No delete rule"

112                 Case adRICascade
114                     strRIName = "Cascade delete"

116                 Case adRISetNull
118                     strRIName = "Cascade to null"

120                 Case adRISetDefault
122                     strRIName = "Cascade to default"

124                 Case Else
126                     strRIName = "DeleteRule of " & .DeleteRule & " unknown."
                End Select

128             DebugPrint "Key: " & .Name & ", to table: " & .RelatedTable & ", with: " & strRIName
            End With

        Next

130     Set ky = Nothing
132     Set tbl = Nothing
134     Set cat = Nothing
        '<EhFooter>
        Exit Function

ShowKeyAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.ShowKeyAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function DeleteKeyAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Delete relationships using ADOX.
        '<EhHeader>
        On Error GoTo DeleteKeyAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim tbl As ADOX.Table
    
100     Set cat.ActiveConnection = CurrentProjectConnection
102     cat.Tables("tblAdoxBooking").Keys.Delete "tblAdoxContractortblAdoxBooking"
    
104     Set cat = Nothing
106     DebugPrint "Key deleted."
        '<EhFooter>
        Exit Function

DeleteKeyAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.DeleteKeyAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function CreateViewAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Create a query using ADOX.
        '<EhHeader>
        On Error GoTo CreateViewAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim cmd As New ADODB.Command
        Dim strSql As String
    
        'Initialize.
100     cat.ActiveConnection = CurrentProjectConnection
    
        'Assign the SQL statement to Command object's CommandText property.
102     strSql = "SELECT BookingID, BookingDate FROM tblDaoBooking;"
104     cmd.CommandText = strSql
    
        'Append the Command to the Views collectiion of the catalog.
106     cat.Views.Append "qryAdoxBooking", cmd
    
        'Clean up.
108     Set cmd = Nothing
110     Set cat = Nothing
112     DebugPrint "View created."
        '<EhFooter>
        Exit Function

CreateViewAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.CreateViewAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function CreateProcedureAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Create a parameter query or action query using ADOX.
        '<EhHeader>
        On Error GoTo CreateProcedureAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim cmd As New ADODB.Command
        Dim strSql As String
    
        'Initialize.
100     cat.ActiveConnection = CurrentProjectConnection
    
        ''Assign the SQL statement to the CommandText property.
102     strSql = "PARAMETERS StartDate DateTime, EndDate DateTime; " & "DELETE FROM tblAdoxBooking " & "WHERE BookingDate Between StartDate And EndDate;"
104     cmd.CommandText = strSql
    
        'Append the Command to the Procedures collection of the catalog.
106     cat.Procedures.Append "qryAdoxDeleteBooking", cmd
    
        'Clean up.
108     Set cmd = Nothing
110     Set cat = Nothing
112     DebugPrint "Procedure created."
        '<EhFooter>
        Exit Function

CreateProcedureAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.CreateProcedureAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ShowProx(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   List the parameter/action queries using ADOX.
        '<EhHeader>
        On Error GoTo ShowProx_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim proc As ADOX.Procedure
        Dim vw As ADOX.View
    
100     cat.ActiveConnection = CurrentProjectConnection
    
102     DebugPrint "Procedures: " & cat.Procedures.Count

104     For Each proc In cat.Procedures
106         DebugPrint proc.Name
        Next

108     DebugPrint cat.Procedures.Count & " procedure(s)"
110     DebugPrint ""
    
112     DebugPrint "Views " & cat.Views.Count

114     For Each vw In cat.Views
116         DebugPrint vw.Name
        Next
    
118     Set cat = Nothing
        '<EhFooter>
        Exit Function

ShowProx_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.ShowProx " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ExecuteProcedureAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Execute a parameter query using ADOX.
        '<EhHeader>
        On Error GoTo ExecuteProcedureAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim cmd As ADODB.Command
        Dim lngCount As Long
    
        'Initialize.
100     cat.ActiveConnection = CurrentProjectConnection
102     Set cmd = cat.Procedures("qryAdoxDeleteBooking").Command
    
        'Supply the parameters
104     cmd.Parameters("StartDate") = #1/1/2004#
106     cmd.Parameters("EndDate") = #12/31/2004#
    
        'Execute the procedure
108     cmd.Execute lngCount
110     DebugPrint lngCount & " record(s) deleted."
    
        'Alternative: specify the parameters in a variant array.
        'cmd.Execute , Array(#1/1/2004#, #12/31/2004#)
    
        'Clean up.
112     Set cmd = Nothing
114     Set cat = Nothing
        '<EhFooter>
        Exit Function

ExecuteProcedureAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.ExecuteProcedureAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function DeleteProcedureAdox(CurrentProjectConnection As ADODB.Connection)
        'Purpose:   Delete a parameter/action query using ADOX.
        '<EhHeader>
        On Error GoTo DeleteProcedureAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim cmd As ADODB.Command
        Dim lngCount As Long
    
        'Initialize.
100     cat.ActiveConnection = CurrentProjectConnection
102     cat.Procedures.Delete "qryAdoxDeleteBooking"
104     Set cat = Nothing
        '<EhFooter>
        Exit Function

DeleteProcedureAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.DeleteProcedureAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function CreateDatabaseAdox(sDatabaseNameAndPath As String)
        'Purpose:   Create a database using ADOX.
        '<EhHeader>
        On Error GoTo CreateDatabaseAdox_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog

102     cat.Create "Provider='Microsoft.Jet.OLEDB.4.0';" & "Data Source='" & sDatabaseNameAndPath & "'"

104     Set cat = Nothing
106     DebugPrint sDatabaseNameAndPath & " created."
        '<EhFooter>
        Exit Function

CreateDatabaseAdox_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.CreateDatabaseAdox " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function DeleteAllAndResetAutoNum(CurrentProjectConnection As ADODB.Connection, strTable As String) As Boolean
        'Purpose:   Delete all records from the table, and reset the AutoNumber using ADOX.
        '           Also illustrates how to find the AutoNumber field.
        'Argument:  Name of the table to reset.
        'Return:    True if sucessful.
        '<EhHeader>
        On Error GoTo DeleteAllAndResetAutoNum_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog
        Dim tbl As ADOX.Table
        Dim Col As ADOX.Column
        Dim strSql As String
    
        'Delete all records.
100     strSql = "DELETE FROM [" & strTable & "];"
102     CurrentProjectConnection.Execute strSql
    
        'Find and reset the AutoNum field.
104     cat.ActiveConnection = CurrentProjectConnection
106     Set tbl = cat.Tables(strTable)

108     For Each Col In tbl.Columns

110         If Col.Properties("Autoincrement") Then
112             Col.Properties("Seed") = 1
114             DeleteAllAndResetAutoNum = True
            End If

        Next

        '<EhFooter>
        Exit Function

DeleteAllAndResetAutoNum_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.DeleteAllAndResetAutoNum " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function GetSeedADOX(CurrentProjectConnection As ADODB.Connection, strTable As String, Optional ByRef strCol As String) As Long
        'Purpose:   Read the Seed of the AutoNumber of a table.
        'Arguments: strTable the table to examine.
        '           strCol = the name of the field. If omited, the code finds it.
        'Return:    The seed value.
        '<EhHeader>
        On Error GoTo GetSeedADOX_Err
        '</EhHeader>
        Dim cat As New ADOX.Catalog 'Root object of ADOX.
        Dim tbl As ADOX.Table       'Each Table in Tables.
        Dim Col As ADOX.Column      'Each Column in the Table.
    
        'Point the catalog to the current project's connection.
100     Set cat.ActiveConnection = CurrentProjectConnection
102     Set tbl = cat.Tables(strTable)
    
        'Loop through the columns to find the AutoNumber.
104     For Each Col In tbl.Columns

106         If Col.Properties("Autoincrement") Then
108             strCol = Col.Name
110             GetSeedADOX = Col.Properties("Seed")
                Exit For    'There can be only one AutoNum.
            End If

        Next
    
        'Clean up
112     Set Col = Nothing
114     Set tbl = Nothing
116     Set cat = Nothing
        '<EhFooter>
        Exit Function

GetSeedADOX_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.GetSeedADOX " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function ResetSeed(CurrentProjectConnection As ADODB.Connection, strTable As String) As String
    '    'Purpose:   Reset the Seed of the AutoNumber, using ADOX.
    '    Dim strAutoNum As String    'Name of the autonumber column.
    '    Dim lngSeed As Long         'Current value of the Seed.
    '    Dim lngNext As Long         'Next unused value.
    '    Dim strSql As String
    '    Dim strResult As String
    '
    '    lngSeed = GetSeedADOX(CurrentProjectConnection, strTable, strAutoNum)
    '    If strAutoNum = vbNullString Then
    '        strResult = "AutoNumber not found."
    '    Else
    '        lngNext = Nz(DMax(strAutoNum, strTable), 0) + 1
    '        If lngSeed = lngNext Then
    '            strResult = strAutoNum & " already correctly set to " & lngSeed & "."
    '        Else
    '            DebugPrint lngNext, lngSeed
    '            strSql = "ALTER TABLE [" & strTable & "] ALTER COLUMN [" & strAutoNum & "] COUNTER(" & lngNext & ", 1);"
    '            DebugPrint strSql
    '            CurrentProjectConnection.Execute strSql
    '            strResult = strAutoNum & " reset from " & lngSeed & " to " & lngNext
    '        End If
    '    End If
    '    ResetSeed = strResult
End Function

Public Function DoFieldExists(oRS As ADODB.Recordset, sFieldName As String) As Boolean
    Dim oFld As ADODB.Field
    
    On Error Resume Next
    
    Set oFld = oRS.Fields.Item(sFieldName)
    
    If oFld Is Nothing Then
        DoFieldExists = False
    Else
        DoFieldExists = True
    End If
    
End Function

Public Function FieldIsString(FieldObject As ADODB.Field) As Boolean
        '<EhHeader>
        On Error GoTo FieldIsString_Err
        '</EhHeader>

100     If Not TypeOf FieldObject Is ADODB.Field Then Exit Function

102     Select Case FieldObject.Type

            Case adBSTR, adChar, adVarChar, adWChar, adVarWChar, adLongVarChar, adLongVarWChar
104             FieldIsString = True

106         Case Else
108             FieldIsString = False
        End Select
        
        '<EhFooter>
        Exit Function

FieldIsString_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.FieldIsString " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function FieldIsBoolean(FieldObject As ADODB.Field) As Boolean
        '<EhHeader>
        On Error GoTo FieldIsBoolean_Err
        '</EhHeader>

100     If Not TypeOf FieldObject Is ADODB.Field Then Exit Function

102     Select Case FieldObject.Type

            Case adBoolean
104             FieldIsBoolean = True

106         Case Else
108             FieldIsBoolean = False
        End Select
        
        '<EhFooter>
        Exit Function

FieldIsBoolean_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.FieldIsBoolean " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function FieldIsNumeric(FieldObject As ADODB.Field) As Boolean
        '<EhHeader>
        On Error GoTo FieldIsNumeric_Err
        '</EhHeader>

100     If Not TypeOf FieldObject Is ADODB.Field Then Exit Function

102     Select Case FieldObject.Type

            Case adBigInt, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adVarNumeric
104             FieldIsNumeric = True

106         Case Else
108             FieldIsNumeric = False
        End Select
        
        '<EhFooter>
        Exit Function

FieldIsNumeric_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.FieldIsNumeric " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function FieldIsTimeDate(FieldObject As ADODB.Field) As Boolean
        '<EhHeader>
        On Error GoTo FieldIsTimeDate_Err
        '</EhHeader>
    
100     If Not TypeOf FieldObject Is ADODB.Field Then Exit Function

102     Select Case FieldObject.Type

            Case adDate, adDBDate, adDBTime, adDBTimeStamp
104             FieldIsTimeDate = True

106         Case Else
108             FieldIsTimeDate = False
        End Select
        
        '<EhFooter>
        Exit Function

FieldIsTimeDate_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.FieldIsTimeDate " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function FieldIsBinary(FieldObject As ADODB.Field) As Boolean
        '<EhHeader>
        On Error GoTo FieldIsBinary_Err
        '</EhHeader>

100     If Not TypeOf FieldObject Is ADODB.Field Then Exit Function

102     Select Case FieldObject.Type

            Case adBinary, adLongVarBinary, adVarBinary
104             FieldIsBinary = True

106         Case Else
108             FieldIsBinary = False
        End Select
        
        '<EhFooter>
        Exit Function

FieldIsBinary_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.FieldIsBinary " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function DoTableExists(sTable As String, _
                              CN As ADODB.Connection, _
                              Optional oRSCheck As ADODB.Recordset) As Boolean
        '<EhHeader>
        On Error GoTo DoTableExists_Err
        '</EhHeader>
    
100     If oRSCheck Is Nothing Then
102         Set oRSCheck = New ADODB.Recordset
        End If
    
        On Error Resume Next
    
104     oRSCheck.Open "SELECT * FROM " & sTable, CN, adOpenForwardOnly, adLockReadOnly

106     If Err.number = 0 Then

108         DoTableExists = True
        Else
110         Err.Clear
        End If

        '    oRSCheck.Close
        '    Set oRSCheck = Nothing

        '<EhFooter>
        Exit Function

DoTableExists_Err:
        MsgBox Err.Description & vbCrLf & "in OASISRemoteAdmin.modADOX.DoTableExists " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function



