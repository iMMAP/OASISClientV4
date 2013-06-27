Attribute VB_Name = "modNomadUtils"

Public Const LB_FINDSTRING As Long = &H18F
Public Const LB_FINDSTRINGEXACT As Long = &H1A2
Public Const CB_ERR As Long = (-1)
Public Const LB_ERR As Long = (-1)
Public Const WM_USER As Long = &H400
Public Const CB_FINDSTRING As Long = &H14C
Public Const CB_SHOWDROPDOWN As Long = &H14F
Public Const XMLHeader = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>"

Public Declare Function SendMessageStr _
               Lib "user32" _
               Alias "SendMessageA" (ByVal hWnd As Long, _
                                     ByVal wMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As String) As Long
                                     
Private Type SystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(32) As Integer
    StandardDate As SystemTime
    StandardBias As Long
    DaylightName(32) As Integer
    DaylightDate As SystemTime
    DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation _
                Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
                
Private Declare Function timeGetTime _
                Lib "winmm.dll" () As Long

Public Function FromUnixTime(ByVal sUnixTime As Long) As Date
    Dim NTime As Date, STime As Date
    Dim TZ As TIME_ZONE_INFORMATION
    STime = #1/1/1970#
    NTime = DateAdd("s", sUnixTime, STime)
    GetTimeZoneInformation TZ
    NTime = DateAdd("n", -TZ.Bias, NTime)
    FromUnixTime = NTime
End Function

Public Function ToUnixTime(ByVal STime As Date) As Long
    Dim NTime As Date, sUnix As Date, sUnixTime As Long
    Dim TZ As TIME_ZONE_INFORMATION
    sUnix = #1/1/1970#
    GetTimeZoneInformation TZ
    NTime = DateAdd("n", TZ.Bias, STime)
    sUnixTime = DateDiff("s", sUnix, NTime)
    ToUnixTime = sUnixTime
End Function

Public Function OASIS2NomadTable(sString) As String

    Select Case LCase(sString)
    
        Case LCase("dd_NOMADALLOC_ddAllocationMonths")
            OASIS2NomadTable = "month"

        Case LCase("dd_NOMADALLOC_ddAllocationRefs")
            OASIS2NomadTable = "reference"

        Case LCase("dd_NOMADALLOC_ddAllocationYears")
            OASIS2NomadTable = "year"

        Case LCase("dd_NOMADALLOC_linkAllocationDetail")
            OASIS2NomadTable = "itemallocation"

        Case LCase("dd_NOMADALLOC_mastertable")
            OASIS2NomadTable = "allocation"

        Case LCase("dd_NOMADCORE_ddDonor")
            OASIS2NomadTable = "donor"

        Case LCase("dd_NOMADCORE_ddItemCategory")
            OASIS2NomadTable = "itemcategory"

        Case LCase("dd_NOMADCORE_ddItems")
            OASIS2NomadTable = "item"

        Case LCase("dd_NOMADCORE_ddItemUnits")
            OASIS2NomadTable = "unit"

        Case LCase("dd_NOMADCORE_ddKebele")
            OASIS2NomadTable = "kebele"

        Case LCase("dd_NOMADCORE_ddReceiver")
            OASIS2NomadTable = "receiver"

        Case LCase("dd_NOMADCORE_ddRegion")
            OASIS2NomadTable = "region"

        Case LCase("dd_NOMADCORE_ddSender")
            OASIS2NomadTable = "sender"

        Case LCase("dd_NOMADCORE_ddSite")
            OASIS2NomadTable = "site"

        Case LCase("dd_NOMADCORE_ddTransportCompany")
            OASIS2NomadTable = "transportcompany"

        Case LCase("dd_NOMADCORE_ddTransporter")
            OASIS2NomadTable = "transporter"

        Case LCase("dd_NOMADCORE_ddWereda")
            OASIS2NomadTable = "wereda"

        Case LCase("dd_NOMADCORE_ddZone")
            OASIS2NomadTable = "zone"

        Case LCase("dd_NOMADCORE_linkItemsDisAndRec")
            OASIS2NomadTable = "itemwaybill"

        Case LCase("dd_NOMADCORE_mastertable")
            OASIS2NomadTable = "goodstransfer"
    
        Case Else
            OASIS2NomadTable = ""
    
    End Select
    
    'MsgBox "OASIS2NomadTable: " & sString & " returned: " & OASIS2NomadTable

End Function

Public Function OASIS2NomadService(sString) As String

    Select Case LCase(sString)
    
        Case LCase("dd_NOMADALLOC_ddAllocationMonths")
            OASIS2NomadService = "months"

        Case LCase("dd_NOMADALLOC_ddAllocationRefs")
            OASIS2NomadService = "references"

        Case LCase("dd_NOMADALLOC_ddAllocationYears")
            OASIS2NomadService = "years"

        Case LCase("dd_NOMADALLOC_linkAllocationDetail")
            OASIS2NomadService = "itemallocation"

        Case LCase("dd_NOMADALLOC_mastertable")
            OASIS2NomadService = "allocations"

        Case LCase("dd_NOMADCORE_ddDonor")
            OASIS2NomadService = "donors"

        Case LCase("dd_NOMADCORE_ddItemCategories")
            OASIS2NomadService = "itemcategorys"

        Case LCase("dd_NOMADCORE_ddItems")
            OASIS2NomadService = "items"

        Case LCase("dd_NOMADCORE_ddItemUnits")
            OASIS2NomadService = "units"

        Case LCase("dd_NOMADCORE_ddKebele")
            OASIS2NomadService = "kebeles"

        Case LCase("dd_NOMADCORE_ddReceiver")
            OASIS2NomadService = "receivers"

        Case LCase("dd_NOMADCORE_ddRegion")
            OASIS2NomadService = "regions"

        Case LCase("dd_NOMADCORE_ddSender")
            OASIS2NomadService = "senders"

        Case LCase("dd_NOMADCORE_ddSite")
            OASIS2NomadService = "sites"

        Case LCase("dd_NOMADCORE_ddTransportCompany")
            OASIS2NomadService = "transportcompanys"

        Case LCase("dd_NOMADCORE_ddTransporter")
            OASIS2NomadService = "transporters"

        Case LCase("dd_NOMADCORE_ddWereda")
            OASIS2NomadService = "weredas"

        Case LCase("dd_NOMADCORE_ddZone")
            OASIS2NomadService = "zones"

        Case LCase("dd_NOMADCORE_linkItemsDisAndRec")
            OASIS2NomadService = "itemwaybills"

        Case LCase("dd_NOMADCORE_mastertable")
            OASIS2NomadService = "goodstransfers"
    
        Case Else
            OASIS2NomadService = ""
    
    End Select
    
    OASIS2NomadService = OASIS2NomadTable(sString)
    'MsgBox "OASIS2NomadService: " & sString & " returned: " & OASIS2NomadService

End Function

Public Function OASIS2NomadField(strTable As String, _
                                 strColumn As String) As String
    
    strTable = Trim$(strTable)
    strColumn = Trim$(strColumn)
    
    Select Case LCase(strTable)
   
        Case LCase("dd_NOMADALLOC_ddAllocationMonths")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "month_id"

                Case LCase("option")
                    OASIS2NomadField = "month_name"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADALLOC_ddAllocationRefs")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "reference_id"

                Case LCase("option")
                    OASIS2NomadField = "reference_name"

                Case LCase("RoundNumber")
                    OASIS2NomadField = "reference_round"

                Case LCase("dd_NOMADALLOC_ddAllocationMonths")
                    OASIS2NomadField = "month"

                Case LCase("dd_NOMADALLOC_ddAllocationYears")
                    OASIS2NomadField = "year"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADALLOC_ddAllocationYears")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "year_id"

                Case LCase("option")
                    OASIS2NomadField = "year_name"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADALLOC_linkAllocationDetail")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "allocation"

                Case LCase("GUID2")
                    OASIS2NomadField = "allocationitemallocation_id"

                Case LCase("dd_NOMADCORE_ddItemCategory")
                    OASIS2NomadField = "category"

                Case LCase("dd_NOMADCORE_ddItemUnits")
                    OASIS2NomadField = "unit"

                Case LCase("dd_NOMADCORE_ddRegion")
                    OASIS2NomadField = "region"

                Case LCase("dd_NOMADCORE_ddZone")
                    OASIS2NomadField = "zone"

                Case LCase("dd_NOMADCORE_ddWereda")
                    OASIS2NomadField = "wereda"

                Case LCase("dd_NOMADCORE_ddSite")
                    OASIS2NomadField = "site"

                Case LCase("Beneficiaries")
                    OASIS2NomadField = "itemallocation_beneficiaries"

                Case LCase("RequestQty")
                    OASIS2NomadField = "itemallocation_quantity"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADALLOC_mastertable")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "allocation_id"

                Case LCase("dd_NOMADALLOC_ddAllocationRefs")
                    OASIS2NomadField = "reference"

                Case LCase("TransactionDate")
                    OASIS2NomadField = "allocation_transactiondate"

                Case LCase("dd_NOMADCORE_ddRegion")
                    OASIS2NomadField = "location"

                Case LCase("Remarks")
                    OASIS2NomadField = "allocation_remarks"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddDonor")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "donor_id"

                Case LCase("option")
                    OASIS2NomadField = "donor_name"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddItemCategory")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "itemcategory_id"

                Case LCase("option")
                    OASIS2NomadField = "itemcategory_name"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddItems")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "item_id"

                Case LCase("option")
                    OASIS2NomadField = "item_description"

                Case LCase("item_itemnumber")
                    OASIS2NomadField = "item_itemnumber"

                Case LCase("item_partnumber")
                    OASIS2NomadField = "item_partnumber"

                Case LCase("dd_NOMADCORE_ddItemCategories")
                    OASIS2NomadField = "itemcategory"

                Case LCase("dd_NOMADCORE_ddItemUnits")
                    OASIS2NomadField = "unit"
                    
                Case LCase("item_unitprice")
                    OASIS2NomadField = "item_unitprice"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddItemUnits")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "unit_id"

                Case LCase("option")
                    OASIS2NomadField = "unit_name"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddKebele")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "kebele_id"

                Case LCase("option")
                    OASIS2NomadField = "kebele_name"

                Case LCase("dd_NOMADCORE_ddRegion")
                    OASIS2NomadField = "region"

                Case LCase("dd_NOMADCORE_ddZone")
                    OASIS2NomadField = "zone"

                Case LCase("dd_NOMADCORE_ddWereda")
                    OASIS2NomadField = "wereda"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddReceiver")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "receiver_id"

                Case LCase("option")
                    OASIS2NomadField = "receiver_name"

                Case LCase("receiver_firstname")
                    OASIS2NomadField = "receiver_firstname"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddRegion")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "region_id"

                Case LCase("option")
                    OASIS2NomadField = "region_name"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddSender")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "sender_id"

                Case LCase("option")
                    OASIS2NomadField = "sender_name"

                Case LCase("sender_firstname")
                    OASIS2NomadField = "sender_firstname"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddSite")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "site_id"

                Case LCase("option")
                    OASIS2NomadField = "site_name"

                Case LCase("dd_NOMADCORE_ddRegion")
                    OASIS2NomadField = "region"
                    
                Case LCase("dd_NOMADCORE_ddZone")
                    OASIS2NomadField = "zone"

                Case LCase("dd_NOMADCORE_ddWereda")
                    OASIS2NomadField = "wereda"

                Case LCase("dd_NOMADCORE_ddKebele")
                    OASIS2NomadField = "kebele"

                Case LCase("Longitude")
                    OASIS2NomadField = "site_coordinateslongitude"

                Case LCase("Latitude")
                    OASIS2NomadField = "site_coordinateslatitude"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddTransportCompany")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "transportcompany_id"

                Case LCase("option")
                    OASIS2NomadField = "transportcompany_name"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddTransporter")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "transporter_id"

                Case LCase("option")
                    OASIS2NomadField = "transporter_name"

                Case LCase("transporter_firstname")
                    OASIS2NomadField = "transporter_firstname"

                Case LCase("dd_NOMADCORE_ddTransportCompany")
                    OASIS2NomadField = "company"

                Case LCase("dd_NOMADCORE_ddRegion")
                    OASIS2NomadField = "region"

                Case LCase("dd_NOMADCORE_ddZone")
                    OASIS2NomadField = "zone"

                Case LCase("dd_NOMADCORE_ddWereda")
                    OASIS2NomadField = "wereda"

                Case LCase("dd_NOMADCORE_ddKebele")
                    OASIS2NomadField = "kebele"

                Case LCase("transporter_idnumber")
                    OASIS2NomadField = "transporter_idnumber"

                Case LCase("transporter_licence")
                    OASIS2NomadField = "transporter_licence"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddWereda")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "wereda_id"

                Case LCase("option")
                    OASIS2NomadField = "wereda_name"

                Case LCase("dd_NOMADCORE_ddRegion")
                    OASIS2NomadField = "region"

                Case LCase("dd_NOMADCORE_ddZone")
                    OASIS2NomadField = "zone"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_ddZone")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "zone_id"

                Case LCase("option")
                    OASIS2NomadField = "zone_name"

                Case LCase("dd_NOMADCORE_ddRegion")
                    OASIS2NomadField = "region"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_linkItemsDisAndRec")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "transfer"

                Case LCase("GUID2")
                    OASIS2NomadField = "itemwaybill_id"

                Case LCase("dd_NOMADCORE_ddSite")
                    OASIS2NomadField = "destination"

                Case LCase("dd_NOMADCORE_ddItems")
                    OASIS2NomadField = "item"
                    
                Case LCase("item_totalprice")
                    OASIS2NomadField = "totalprice"

                Case LCase("item_requestqty")
                    OASIS2NomadField = "requestqty"

                Case LCase("item_deliverqty")
                    OASIS2NomadField = "deliverqty"

                Case LCase("item_receivingqty")
                    OASIS2NomadField = "receivingqty"

                Case Else
    
                    OASIS2NomadField = ""

            End Select

        Case LCase("dd_NOMADCORE_mastertable")

            Select Case LCase(strColumn)

                Case LCase("GUID1")
                    OASIS2NomadField = "goodstransfer_id"

                Case LCase("TransactionDate")
                    OASIS2NomadField = "goodstransfer_transactiondate"

                Case LCase("goodstransfer_sentdate")
                    OASIS2NomadField = "goodstransfer_sentdate"

                Case LCase("goodstransfer_purchasereq")
                    OASIS2NomadField = "goodstransfer_purchasereq"

                Case LCase("dd_NOMADCORE_ddSite")
                    OASIS2NomadField = "warehousesender"

                Case LCase("dd_NOMADCORE_ddSite-PAD1")
                    OASIS2NomadField = "destination"

                Case LCase("dd_NOMADAlloc_ddAllocationRefs")
                    OASIS2NomadField = "allocation"

                Case LCase("dd_NOMADCORE_ddDonor")
                    OASIS2NomadField = "donor"

                Case LCase("dd_NOMADCORE_ddSender")
                    OASIS2NomadField = "sender"

                Case LCase("dd_NOMADCORE_ddReceiver")
                    OASIS2NomadField = "receiver"

                Case LCase("dd_NOMADCORE_ddTransportCompany")
                    OASIS2NomadField = "transportercompany"

                Case LCase("dd_NOMADCORE_ddTransporter")
                    OASIS2NomadField = "transporter"

                Case LCase("goodstransfer_platenumber")
                    OASIS2NomadField = "goodstransfer_platenumber"

                Case LCase("goodstransfer_receiveddate")
                    OASIS2NomadField = "goodstransfer_receiveddate"

                Case LCase("Latitude")
                    OASIS2NomadField = "coordinateslatitude"

                Case LCase("Longitude")
                    OASIS2NomadField = "coordinateslongitude"

                Case Else
    
                    OASIS2NomadField = ""

            End Select
    End Select

    'MsgBox "OASIS2NomadField: " & strTable & "." & strColumn & " returned: " & OASIS2NomadField
    
End Function

Public Sub FindIndexStrEx(ctlSource As Control, _
                          ByVal Str As String)
    
    Dim lngIdx As Long

    If TypeName(ctlSource) = "ComboBox" Then
        lngIdx = SendMessageStr(ctlSource.hWnd, CB_FINDSTRING, -1, Str)
    ElseIf TypeName(ctlSource) = "ListBox" Then
        lngIdx = SendMessageStr(ctlSource.hWnd, LB_FINDSTRING, -1, Str)
    Else
        Exit Sub
    End If

    If lngIdx <> -1 Then
        ctlSource.ListIndex = lngIdx
    End If

End Sub

Public Function ConvertOASIS2Nomad(objRS As ADODB.Recordset, _
                              sRootTag As String, _
                              sTablePreFix As String, _
                              strTable As String, _
                              sConn As String, _
                              sBaseURL As String, _
                              sNomadServiceName As String, _
                              sUser As String, _
                              sPwd As String, _
                              Optional bProxy As Boolean, _
                              Optional sProxy As String) As Boolean
        '<EhHeader>
        On Error GoTo ConvertOASIS2Nomad_Err
        'MsgBox "ConvertOASIS2Nomad"
        '</EhHeader>
        Dim oCOnn As New ADODB.Connection
        ConvertOASIS2Nomad = False
106     oCOnn.ConnectionString = sConn
108     oCOnn.Open
        ConvertOASIS2Nomad = ExecuteNomadXML(objRS, sRootTag, sTablePreFix, strTable, oCOnn, sBaseURL, sNomadServiceName, sUser, sPwd, bProxy, sProxy, True)
112     oCOnn.Close
114     Set oCOnn = Nothing
    
        '<EhFooter>
        Exit Function

ConvertOASIS2Nomad_Err:
        'MsgBox Err.Description & vbCrLf & "in NomadTester.frmNomadGet.ConvertOASIS2Nomad " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Function ExecuteNomadXML(objRS As ADODB.Recordset, _
                           strRootTag As String, _
                           strTblPrefix As String, _
                           strTable As String, _
                           conn As ADODB.Connection, _
                           sBaseURL As String, _
                           sNomadServiceName As String, _
                           sUser As String, _
                           sPwd As String, _
                           Optional bProxy As Boolean, _
                           Optional sProxy As String, _
                           Optional bUseAlias As Boolean) As Boolean
        '<EhHeader>
        On Error GoTo ExecuteNomadXML_Err
        '</EhHeader>
                           
        Dim objField As ADODB.Field
        Dim strName As String
        Dim strValue As String
        Dim NomadXML As String
        Dim sFieldDesc As String
        Dim sID As String
        Dim sStrippedTablename As String
        Dim bSuccess As Boolean

        ExecuteNomadXML = False
100     strName = Trim$(strName)
102     'If bDelete Then
           ' sBaseURL = sBaseURL & "/" & Left(sNomadServiceName, Len(sNomadServiceName) - 1)
       ' Else
            sBaseURL = sBaseURL & "/" & sNomadServiceName
       ' End If
     '   MsgBox 1
104     'If bDelete Or (Not objRS.EOF And Not objRS.BOF) Then
'MsgBox 2
106        ' If Not bDelete And Not objRS.BOF Then objRS.MoveFirst
    
'108         'While Not objRS.EOF
'MsgBox 3
110             sStrippedTablename = Right$(strTable, Len(strTable) - InStrRev(strTable, "_"))

                'MsgBox sStrippedTablename

112             If Left$(sStrippedTablename, 4) = "link" Then
114
                    'If bDelete Then
                       ' sBaseURL = sBaseURL & "/" & objRS(1).OriginalValue
                       ' NomadXML = XMLHeader & "<" & strRootTag & " id=""" & objRS(1).OriginalValue & """>"
                   ' Else
                        NomadXML = XMLHeader & "<" & strRootTag & " id=""" & objRS(1).Value & """>"
                   ' End If
                    
                Else
116
                    'If bDelete Then
                       ' sBaseURL = sBaseURL & "/" & objRS(0).OriginalValue
                       ' NomadXML = XMLHeader & "<" & strRootTag & " id=""" & objRS(0).OriginalValue & """>"
                    'Else
                        NomadXML = XMLHeader & "<" & strRootTag & " id=""" & objRS(0).Value & """>"
                    'End If
                End If
                
                'MsgBox sBaseURL
                
118             For Each objField In objRS.Fields
                    
                    'strName = OASIS2NomadField(strTable, objField.Name)
120                 strName = Replace(OASIS2NomadField(strTable, objField.Name), strTblPrefix, "")
                    
122                 If Len(strName) < 1 Then
                        'strName = objField.Name
124                     strName = Replace(objField.Name, strTblPrefix, "")
                    End If
                    
126                 If Not IsNull(objField.Value) Then
128                     strValue = objField.Value
                    Else
130                     strValue = ""
                    End If
                    
132                 If InStr(LCase$(strName), "date") > 0 And Len(strValue) > 0 Then
134                     strValue = ToUnixTime(Format(strValue, "dd-MMM-yy"))
                    End If
    
136                 If LCase(strName) <> "id" And LCase(strName) <> "uid" Then
138                     NomadXML = NomadXML & "<" & strName & ">" & strValue & "</" & strName & ">"
                    End If

                Next
            
140             NomadXML = NomadXML & "</" & strRootTag & ">"
                
142             ExecuteNomadXML = CommitXML2Nomad(sBaseURL, NomadXML, sUser, sPwd, bProxy, sProxy)

144             NomadXML = ""
                
                'OASIS updates 1 a time so this loop is not necessary
146             'objRS.MoveNext
            
           ' Wend

       ' End If
    
148     Set objRS = Nothing
        '<EhFooter>
        Exit Function

ExecuteNomadXML_Err:
        MsgBox Err.Description & vbCrLf & "in NomadTester.modUtils.ExecuteNomadXML " & "at line " & Erl
        
        '</EhFooter>
End Function

Public Function CommitXML2Nomad(sBaseURL As String, _
                            sNomadXML As String, _
                            sUser As String, _
                            sPwd As String, _
                            Optional bProxy As Boolean, _
                            Optional sProxy As String, Optional bDelete As Boolean = False) As Boolean
        '<EhHeader>
        On Error GoTo CommitXML2Nomad_Err
        '</EhHeader>
        Dim oHttp As New WinHttpRequest
        CommitXML2Nomad = False
        
        'MsgBox "CommitXML2Nomad"
        
102     If bProxy Then oHttp.setProxy HTTPREQUEST_PROXYSETTING_PROXY, sProxy, "*.microsoft.com"
106
        If bDelete Then
            'MsgBox sBaseURL
            oHttp.Open "DELETE", sBaseURL, False
        Else
            oHttp.Open "PUT", sBaseURL, False
        End If
        
108     oHttp.setRequestHeader "Content-type", "application/xml"
110     oHttp.SetCredentials sUser, sPwd, HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
        
        sNomadXML = Replace$(sNomadXML, "<coordinateslongitude>", "<coordinates><longitude>")
        sNomadXML = Replace$(sNomadXML, "<coordinateslatitude>", "<latitude>")
        sNomadXML = Replace$(sNomadXML, "</coordinateslongitude>", "</longitude>")
        sNomadXML = Replace$(sNomadXML, "</coordinateslatitude>", "</latitude></coordinates>")
        
        If bDelete Then
            oHttp.send
        Else
112         oHttp.send sNomadXML
        End If
        
        On Error Resume Next
        'MsgBox sNomadXML
        If CStr(oHttp.Status) = "200" Then CommitXML2Nomad = True
        
        '<EhFooter>
        Exit Function

CommitXML2Nomad_Err:
        'MsgBox Err.Description & vbCrLf & "in NomadTester.frmNomadGet.CommitXML2Nomad " & "at line " & Erl
        'Resume Next
        '</EhFooter>
End Function

