VERSION 5.00
Object = "{84E5CF37-E467-4AC2-89C4-C6002FFB5055}#25.1#0"; "ChartViewer.ocx"
Begin VB.Form frmSecChart 
   Caption         =   "OASIS Security Charts"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10635
   Icon            =   "frmSecChart.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin CDChartViewer.ChartViewer ChartViewer1 
      Height          =   3750
      Left            =   240
      Top             =   150
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   6615
   End
End
Attribute VB_Name = "frmSecChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_frmChartSettings As frmChartSettings
Attribute m_frmChartSettings.VB_VarHelpID = -1
' Internal variables to keep track of dragging of the navigation window
Private isDraggingNavigatePad As Boolean
Private mouseDownXCoor As Double
Private mouseDownYCoor As Double
Private m_sSQL() As String
Private m_Labels As Variant
Dim oWordApp As Object '
Private WithEvents m_frmWordBookMarks As frmWordBookMarks
Attribute m_frmWordBookMarks.VB_VarHelpID = -1

Private Sub ChartViewer1_ViewPortChanged(needUpdateChart As Boolean, _
                                         needUpdateImageMap As Boolean)
        '<EhHeader>
        On Error GoTo ChartViewer1_ViewPortChanged_Err
        '</EhHeader>

100     With m_frmChartSettings

102         If Not isDraggingNavigatePad Then
                ' Update the navigator window size and position to reflect the view port
                Dim internalPadWidth As Long, internalPadHeight As Long
104             internalPadWidth = .NavigatePad.Width - ScaleX(2, vbPixels, ScaleMode)
106             internalPadHeight = .NavigatePad.Height - ScaleY(2, vbPixels, ScaleMode)
        
108             .NavigateWindow.Left = CInt(ChartViewer1.ViewportLeft * internalPadWidth)
110             .NavigateWindow.Top = CInt(ChartViewer1.ViewportTop * internalPadHeight)
112             .NavigateWindow.Width = Int(ChartViewer1.ViewportWidth * internalPadWidth)
114             .NavigateWindow.Height = Int(ChartViewer1.ViewportHeight * internalPadHeight)
            End If
    
            ' Synchronize the zoom bar value with the view port width/height
116         .ZoomBar.Value = Int(0.5 + IIf(ChartViewer1.ViewportWidth > ChartViewer1.ViewportHeight, ChartViewer1.ViewportHeight, ChartViewer1.ViewportWidth) * .ZoomBar.Max)
        
            ' Update chart and image map if necessary
118         If needUpdateChart Then
                'DoTrendAnalysis
                'Call DrawChart(ChartViewer1)
            End If

120         If needUpdateImageMap Then
                'Call updateImageMap(ChartViewer1)
            End If
    
        End With
    
122     Clipboard.Clear
124     Clipboard.SetData ChartViewer1.Picture
126     CheckWordExport
        '<EhFooter>
        Exit Sub

ChartViewer1_ViewPortChanged_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.ChartViewer1_ViewPortChanged " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
        Dim CN As New adodb.Connection
        Dim RS As New adodb.Recordset
        Dim var As Variant
        Dim Data As Variant
        Dim labels As Variant
        Dim fld As adodb.Field
        Dim sFields As String
        Dim intNumFields As Integer
        Dim intNumRows As Integer
        Dim intCurFieldIdx As Integer
        Dim intCurRowIdx As Integer
        Dim g1 As New glasslightbar
    
        Dim i As Integer

If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
    
        'frmMain.Show
    
        'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\OASIS\Client\data\db\Oasisclient.mdb;Mode=ReadWrite|Share Deny None;Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
    
100     RS.Open "SELECT * FROM qryIncByProvince", m_Cnn

102     Set m_frmChartSettings = New frmChartSettings
        Set m_frmWordBookMarks = New frmWordBookMarks
    
104     For Each fld In RS.Fields
106         sFields = sFields & fld.Name & ","
        Next
    
        If RS Is Nothing Then Exit Sub
    
        If RS.EOF And RS.Bof Then
            m_frmChartSettings.SetIncidentFields Split(sFields, ",")
            m_frmChartSettings.Show vbModeless, Me
            Exit Sub
        End If
        
        ChartViewer1.ScrollDirection = cvHorizontal
        ChartViewer1.ZoomDirection = cvHorizontal
            
        ' Viewport is always unzoomed as y-axis is auto-scaled
        ChartViewer1.ViewportTop = 0
        ChartViewer1.ViewportHeight = 1
    
        ' Initially choose the pointer mode (drag to scroll mode)
        m_frmChartSettings_ArrowMode
        m_frmChartSettings_ZoomInMode
    
        m_frmChartSettings.SetIncidentFields Split(sFields, ",")
        m_frmChartSettings.Show vbModeless, Me
        Exit Sub
    
108     var = RS.GetRows
    
110     intNumFields = UBound(var)
112     intNumRows = UBound(var, 2)
    
        'Add the data for the fields
114     For intCurRowIdx = 0 To intNumRows
116         For intCurFieldIdx = 0 To intNumFields
118             DebugPrint CStr(var(intCurFieldIdx, intCurRowIdx))
            Next
        Next
    
120     ReDim labels(UBound(var, 2))
122     ReDim Data(UBound(var, 2))
    
124     For i = LBound(labels) To UBound(labels)
126         labels(i) = var(0, i)
128         Data(i) = var(1, i)
        Next
    
130     m_frmChartSettings.SetIncidentFields Split(sFields, ",")
132     m_frmChartSettings.Show vbModeless, Me
    
134     g1.createChartEx ChartViewer1, labels, Data, , , , , m_frmChartSettings.txtCaption.Text, m_frmChartSettings.txtYCaption.Text

        '        ' The data for the bar chart
        '    Data = Array(450, 560, 630, 800, 1100, 1350, 1600, 1950, 2300, 2700)
        '
        '    ' The labels for the bar chart
        '
        '    Labels = Array("1996", "1997", "1998", "1999", "2000", "2001", "2002", "2003", _
        '        "2004", "2005")
        '

        Exit Sub

        Dim CD As New ChartDirector.API

        'The data for the bar chart
        'Dim Data()
136     Data = Array(85, 156, 179.5, 211, 123)

        'The labels for the bar chart
        ' Dim Labels()
138     labels = Array("Mon", "Tue", "Wed", "Thu", "Fri")

        'Create a XYChart object of size 250 x 250 pixels
        Dim c As Object
140     Set c = CD.XYChart(250, 250)

        'Set the plotarea at (30, 20) and of size 200 x 200 pixels
142     Call c.setPlotArea(30, 20, 200, 200)

        'Add a bar chart layer using the given data
144     Call c.addBarLayer(Data)

        'Set the x axis labels using the given labels
146     Call c.xAxis().setLabels(labels)

        'output the chart
148     Set ChartViewer1.Picture = c.makePicture()

        'include tool tip for the chart
150     ChartViewer1.ImageMap = c.getHTMLImageMap("clickable", "", "title='{xLabel}: US${value}K'")
        
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmSecChart.Form_Load " & "at line " & Erl
        
        '</EhFooter>
End Sub


Public Sub InsertIMG2Word(omWordApp As Object, sImage As String, sBookMarkName As String, Optional bUseClipboard As Boolean = False)
    'Find the Bookmark
    omWordApp.Selection.GoTo What:=-1, Name:=sBookMarkName
    
    If bUseClipboard Then
       omWordApp.Selection.Paste
    Else
        omWordApp.Selection.InlineShapes.AddPicture Filename:=sImage, LinkToFile:=False, SaveWithDocument:=True
    End If
    
End Sub

Public Sub Xport2Word(sImage As String, sBookMarkName As String, Optional bUseClipboard As Boolean = False)
        '<EhHeader>
        On Error GoTo Xport2Word_Err
        '</EhHeader>


100     oWordApp.Selection.GoTo What:=-1, Name:=sBookMarkName
    
102     If bUseClipboard Then
104        oWordApp.Selection.Paste
        Else
106         oWordApp.Selection.InlineShapes.AddPicture Filename:=sImage, LinkToFile:=False, SaveWithDocument:=True
        End If



    '    Dim oWordApp As New Word.Application
      
        
      
    '        oWordApp.Documents.Open "C:\WebDL\Weekly Format Iraq.doc"
            'Selection.GoTo What:=-1, Name:="wassit"
            'InsertIMG2Word Nothing, "", sBookMarkName, True
        
    'With oWordApp
        
    '        With .ActiveDocument.Bookmarks
    '            .DefaultSorting = wdSortByName
    '            .ShowHidden = False
    '        End With

            'Selection.InlineShapes.AddPicture Filename:="C:\Users\Petri\Pictures\1.bmp", LinkToFile:=False, SaveWithDocument:=True
    '        .Documents.Save
    '        .Documents.Close
    '        .Quit
    '        End With

    '    Set oWordApp = Nothing
        '<EhFooter>
        Exit Sub

Xport2Word_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.Xport2Word " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
 ' \\ This function checks if the file is allready opened by another process,
 ' \\ and if the specified type of access is not allowed.
 ' \\ If so the Open method will fail and a error occurs!
Function IsFileLocked(sFile As String) As Boolean
    On Error Resume Next
     
     ' \\ Open the file
    Open sFile For Binary Access Read Write Lock Read Write As #1
     ' \\ Close the file
    Close #1
     
     ' \\ If error occurs the document if open!
    If Err.number <> 0 Then
         ' \\ msgbox for demonstration purposes
        'MsgBox Err.Description, vbExclamation, "Warning File is opened"
         
         '\\ Return true and clear error
        IsFileLocked = True
        Err.Clear
    End If
End Function

Private Sub CheckWord()
        '<EhHeader>
        On Error GoTo CheckWord_Err
        '</EhHeader>
        On Error Resume Next
100     Set oWordApp = GetObject(, "Word.Application")
    
102     If Err.number <> 0 Then
104         Set oWordApp = CreateObject("Word.Application")
        End If
    
106     If m_frmChartSettings.c.Filename = "" Then
108        m_frmChartSettings.c.Filter = "Microsoft Word Documents (*.doc)|*.doc|Microsoft Word Documents (*.rtf)|*.rtf"
110        m_frmChartSettings.c.ShowOpen
        End If
dude:
    
    
112     If Not IsFileLocked(m_frmChartSettings.c.Filename) Then
            If Len(oWordApp.ActiveDocument.FullName) > 0 Then
            If m_frmChartSettings.c.Filename <> oWordApp.ActiveDocument.FullName Then

                oWordApp.ActiveDocument.Close True
            End If
            End If
114         oWordApp.Documents.Open Filename:=m_frmChartSettings.c.Filename
        Else

            'MsgBox "FILE IS OPEN AND LOCKED CHOOSE ANOTHER FILE OR CLOSE THE FILE AND TRY AGAIN!", vbInformation, "OASIS Word Export"
            'm_frmChartSettings.CommonDialog1.Filter = "Microsoft Word Documents (*.doc)|*.doc|Microsoft Word Documents (*.rtf)|*.rtf"
            'm_frmChartSettings.CommonDialog1.ShowOpen
            'GoTo DUDE
        End If

        '        oWordApp.Documents.Open "C:\WebDL\Weekly Format Iraq.doc"
        'Selection.GoTo What:=-1, Name:="wassit"
        'InsertIMG2Word Nothing, "", sBookMarkName, True
        
        'With oWordApp
        
        '        With .ActiveDocument.Bookmarks
        '            .DefaultSorting = wdSortByName
        '            .ShowHidden = False
        '        End With

        'Selection.InlineShapes.AddPicture Filename:="C:\Users\Petri\Pictures\1.bmp", LinkToFile:=False, SaveWithDocument:=True
        '        .Documents.Save
        '        .Documents.Close
        '        .Quit
        '        End With

        '    Set oWordApp = Nothing

        '<EhFooter>
        Exit Sub

CheckWord_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.CheckWord " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
            
        On Error Resume Next
                        
100     If Not oWordApp Is Nothing Then
102         oWordApp.Visible = True
104         oWordApp.Dialogs(84).Show ' wdDialogFileSaveAs = 84
            'oWordApp.Documents.Save NoPrompt:=False
106         oWordApp.Documents.Close
108         oWordApp.Quit
110         Set oWordApp = Nothing
        End If
    
112     Unload m_frmWordBookMarks
    
114     Set m_frmWordBookMarks = Nothing
116     Set m_frmChartSettings = Nothing
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Function IsRunning(ByVal myAppl As String) As Boolean
    Dim applRef As Object
    On Error Resume Next

    Set applRef = GetObject(, myAppl)

    If Err.number = 429 Then
        IsRunning = False
    Else
        IsRunning = True
    End If

    'clear object variable
    Set applRef = Nothing
End Function

Private Sub m_frmChartSettings_Apply(sSQL As String)
        '<EhHeader>
        On Error GoTo m_frmChartSettings_Apply_Err
        '</EhHeader>
        Dim CN As New adodb.Connection
        Dim RS As New adodb.Recordset
        Dim var As Variant
        Dim Data As Variant
        Dim labels As Variant
        Dim fld As adodb.Field
        Dim sFields As String
        Dim intNumFields As Integer
        Dim intNumRows As Integer
        Dim intCurFieldIdx As Integer
        Dim intCurRowIdx As Integer
        Dim g1 As New glasslightbar
        Dim i As Integer
        Dim g2 As New sidelabelpie
        Dim g3 As New Multiline

        'ArrayName = objRS.GetRows(, , Array("ColumnName1", "ColumnName2", ...))

        'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\OASIS\Client\data\db\Oasisclient.mdb;Mode=ReadWrite|Share Deny None;Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
    
100     m_Cnn.Execute "delete from oincidentstrans"
    
102     m_Cnn.Execute sSQL
    
104     If m_frmChartSettings.OptAdminLevel(0).Value Then
106         RS.Open "SELECT * FROM qryIncByProvince", m_Cnn
108     ElseIf m_frmChartSettings.OptAdminLevel(1).Value Then
    
110         RS.Open "SELECT * FROM qryIncByDistrict", m_Cnn
        Else
112         RS.Open "SELECT * FROM qryIncByTown", m_Cnn

        End If
    
114     For Each fld In RS.Fields
116         sFields = sFields & fld.Name & ","
        Next
    
117     'RS.Sort = "Incidents DESC"
        '        If m_frmChartSettings.ComNumOfAreasToAnalyse.ListIndex = 0 Then
    
118     var = RS.GetRows
    
120     intNumFields = UBound(var)
122     intNumRows = UBound(var, 2)
    
        'Add the data for the fields
        '124     For intCurRowIdx = 0 To intNumRows
        '126         For intCurFieldIdx = 0 To intNumFields
        '128             DebugPrint var(intCurFieldIdx, intCurRowIdx)
        '            Next
        '        Next
    
        If Not m_frmChartSettings.OptChartType(2).Value Then
    
130         ReDim labels(UBound(var, 2))
        
132         ReDim Data(UBound(var, 2))
    
134         For i = LBound(labels) To UBound(labels)
136             labels(i) = var(0, i)
138             Data(i) = var(m_frmChartSettings.ComFlds.ListIndex + 1, i)
            Next
    
            'm_frmChartSettings.SetIncidentFields Split(sFields, ",")
140         If Not m_frmChartSettings.Visible Then m_frmChartSettings.Show vbModeless, Me
    
142         With m_frmChartSettings

144             If .OptChartType(0).Value Then
146                 g1.createChartEx ChartViewer1, labels, Data, .txtWidth(0).Text, .txtHeight(0).Text, .txtWidth(1).Text, .txtHeight(1).Text, .txtCaption.Text, .txtYCaption.Text
148             ElseIf .OptChartType(1).Value Then
150                 g2.createChartEx ChartViewer1, labels, Data, .txtWidth(0).Text, .txtHeight(0).Text, .txtWidth(1).Text, .txtHeight(1).Text, .txtCaption.Text, .txtYCaption.Text
152             ElseIf .OptChartType(2).Value Then
154                 'g3.createChartex1 ChartViewer1, ""
                End If

            End With
        
        Else
        
            Dim Datas() As Variant
        
            g3.createChartex1 ChartViewer1, ""
        End If
        
156     Clipboard.Clear
158     Clipboard.SetData ChartViewer1.Picture
        
        CheckWordExport

        '<EhFooter>
        Exit Sub

m_frmChartSettings_Apply_Err:
        MsgBox Err.Description & vbCrLf & "in ChartSettings Apply " & "at line " & Erl & " SQL:" & sSQL
        
        '</EhFooter>
End Sub

Private Sub CheckWordExport()
Dim i As Integer
        '<EhHeader>
        On Error GoTo CheckWordExport_Err
        '</EhHeader>
100         If m_frmChartSettings.chkExportTo = vbChecked Then
102             m_frmWordBookMarks.lvBookMarks.ListItems.Clear

104             CheckWord

106             For i = 1 To oWordApp.ActiveDocument.Bookmarks.Count
108                 m_frmWordBookMarks.lvBookMarks.ListItems.Add Text:=oWordApp.ActiveDocument.Bookmarks.Item(i).Name
                Next
            
110             m_frmWordBookMarks.Show vbModal, Me
            
            End If
        '<EhFooter>
        Exit Sub

CheckWordExport_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.CheckWordExport " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmChartSettings_ArrowMode()
    ChartViewer1.MouseUsage = cvDefaultUsage
End Sub

Private Sub m_frmChartSettings_DoDetailedChart()
        '<EhHeader>
        On Error GoTo m_frmChartSettings_DoDetailedChart_Err
        '</EhHeader>
        Dim sb As New stackedbar
        Dim sSQLWhere As String
    
100     With m_frmChartSettings
    
102         If .chkUseDate1.Value = vbChecked Then
104             If .chkUseQuery1.Value = vbChecked Then
106                 sSQLWhere = " WHERE ((Incident_DATE BETWEEN #" & .dxDtFROM1.EditValue & "# AND #" & .dxDtTo1.EditValue & "#)) " & IIf(.txtQry1.Text <> "", "AND ", "") & .txtQry1.Text
                Else
108                 sSQLWhere = " WHERE ((Incident_DATE BETWEEN #" & .dxDtFROM1.EditValue & "# AND #" & .dxDtTo1.EditValue & "#))"
                End If

            Else

110             If .chkUseQuery1.Value = vbChecked Then
112                 sSQLWhere = " " & IIf(.txtQry1.Text <> "", "WHERE ", "") & .txtQry1.Text
                Else
114                 sSQLWhere = ""
                End If
            End If
    
116         sb.createIncidentChart ChartViewer1, "", CInt(.txtDeLGDTop), CInt(.txtDeLGDLeft), CInt(.txtDeBGWidth), CInt(.txtDeBGHeight), CInt(.txtDeCHTop), CInt(.txtDeCHLeft), CInt(.txtDeCHWidth), CInt(.txtDeCHHgt), .txtHeadCaption, .Text1.Text, "SELECT DISTINCT Province FROM oincidents_FEA", "SELECT * FROM oincidents_FEA" & sSQLWhere, IIf(.chkIgnoreZero.Value = vbChecked, " " & sSQLWhere, ""), CInt(.txtXValRotation(1)), CInt(.txtLargeValueStep(1)), CInt(.txtSmallValueStep(1))
            Clipboard.Clear
180         Clipboard.SetData ChartViewer1.Picture
            CheckWordExport
        End With

        '<EhFooter>
        Exit Sub

m_frmChartSettings_DoDetailedChart_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.m_frmChartSettings_DoDetailedChart " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub m_frmChartSettings_DoTrendAnalysis(sSQL() As String, _
                                               labels As Variant)
    m_sSQL = sSQL
    m_Labels = labels
    
    DoTrendAnalysis
End Sub

Private Sub DoTrendAnalysis()
        '<EhHeader>
        On Error GoTo DoTrendAnalysis_Err
        '</EhHeader>
        Dim RS As New adodb.Recordset
        Dim var As Variant
        Dim Data As Variant
        Dim i As Integer
        Dim j As Integer
        Dim g3 As New Multiline
        Dim Datas() As Variant
        Dim ValueLabels() As String
        Dim iFrqVals() As Long
        Dim rsTrend As adodb.Recordset
        Dim rsTrendValues As adodb.Recordset
        Dim rsCurRecs As adodb.Recordset
        Dim FromDate As Date
        Dim ToDate As Date
        Dim z As Integer
        Dim x As Integer
        Dim ObjLabels() As String
        
100     If Not IsArray(m_Labels) Then Exit Sub
    
102     ReDim ValueLabels(0)
104     ReDim Datas(0)
106     ReDim iFrqVals(0)
        ReDim ObjLabels(0)
    
108     FromDate = m_frmChartSettings.dxDtFROM.EditValue
110     ToDate = m_frmChartSettings.dxDtTo.EditValue
    
112     Do While Not FromDate > ToDate
114         ValueLabels(UBound(ValueLabels)) = CStr(FromDate)
116         FromDate = FromDate + 1

118         If Not FromDate > ToDate Then ReDim Preserve ValueLabels(UBound(ValueLabels) + 1)
        Loop
        
120     For j = LBound(m_sSQL) To UBound(m_sSQL)
122         m_Cnn.Execute "delete from oincidentstrans"
            DebugPrint m_sSQL(j)
124         m_Cnn.Execute m_sSQL(j)
    
126         If RS.State = adStateOpen Then
128             RS.Close
130             Set RS = New adodb.Recordset
            End If
    
132         If m_frmChartSettings.OptAdminLevel(0).Value Then
134             RS.Open "SELECT * FROM qryIncByProvince", m_Cnn
136         ElseIf m_frmChartSettings.OptAdminLevel(1).Value Then
    
138             RS.Open "SELECT * FROM qryIncByDistrict", m_Cnn
            Else
140             RS.Open "SELECT * FROM qryIncByTown", m_Cnn
            End If
        
142         If (Not RS.EOF) And (Not RS.Bof) Then
                ObjLabels(UBound(ObjLabels)) = m_Labels(j)
                ReDim Preserve ObjLabels(UBound(ObjLabels) + 1)
                
144             If m_frmChartSettings.OptFreqSetting(0).Value Then
146                 ReDim iFrqVals(UBound(ValueLabels))
                Else
148                 ReDim iFrqVals(CInt(UBound(ValueLabels) / 7))
                End If
            
                x = 0
            
150             For i = 0 To UBound(ValueLabels)
152                 Set rsCurRecs = New adodb.Recordset
                
154                 If m_frmChartSettings.OptFreqSetting(0).Value Then
156                     rsCurRecs.Open "SELECT ID FROM oincidentstrans WHERE Incident_DATE = #" & ValueLabels(i) & "#", m_Cnn
                    Else
158                     z = IIf((i + 7) >= UBound(ValueLabels), UBound(ValueLabels), i + 7)
160                     rsCurRecs.Open "SELECT ID FROM oincidentstrans WHERE ((Incident_DATE BETWEEN #" & ValueLabels(i) & "# AND #" & ValueLabels(z) & "#))", m_Cnn
162                     i = z
                    End If
                
164                 iFrqVals(x) = rsCurRecs.RecordCount
                    x = x + 1
                Next
            
166             Datas(UBound(Datas)) = iFrqVals

170             RS.MoveNext
171             If Not RS.EOF Then ReDim Preserve Datas(UBound(Datas) + 1)
                ReDim Preserve Datas(UBound(Datas) + 1)
            Else
                
            End If

172     Next j
    
        'm_frmChartSettings.SetIncidentFields Split(sFields, ",")
174     If Not m_frmChartSettings.Visible Then m_frmChartSettings.Show vbModeless, Me
176     g3.createChartex1 ChartViewer1, "", ObjLabels, Datas, ValueLabels, CInt(m_frmChartSettings.txtWidth(0).Text), CInt(m_frmChartSettings.txtHeight(0).Text), CInt(m_frmChartSettings.txtWidth(1).Text), CInt(m_frmChartSettings.txtHeight(1).Text), m_frmChartSettings.txtCaption.Text, m_frmChartSettings.txtYCaption.Text, IIf(m_frmChartSettings.chkTrendLine.Value = vbChecked, True, False), m_frmChartSettings.txtXValRotation(0).Text, m_frmChartSettings.txtLargeValueStep(0).Text, m_frmChartSettings.txtSmallValueStep(0).Text
        
178     Clipboard.Clear
180     Clipboard.SetData ChartViewer1.Picture
        CheckWordExport
        '<EhFooter>
        Exit Sub

DoTrendAnalysis_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.DoTrendAnalysis " & _
               "at line " & Erl
        
        '</EhFooter>
End Sub

Private Sub m_frmChartSettings_ZoomBar()
        '<EhHeader>
        On Error GoTo m_frmChartSettings_ZoomBar_Err
        '</EhHeader>
        
100     With m_frmChartSettings
        
            ' Remember the center point
            Dim centerX As Double, centerY As Double
102         centerX = ChartViewer1.ViewportLeft + ChartViewer1.ViewportWidth / 2
104         centerY = ChartViewer1.ViewportTop + ChartViewer1.ViewportHeight / 2

            ' Aspect ratio and zoom factor
            Dim aspectRatio As Double, zoomTo As Double
106         aspectRatio = ChartViewer1.ViewportWidth / ChartViewer1.ViewportHeight
108         zoomTo = CDbl(.ZoomBar.Value) / .ZoomBar.Max

            ' Zoom while preserving aspect ratio
110         ChartViewer1.ViewportWidth = zoomTo * IIf(aspectRatio > 1, aspectRatio, 1)
112         ChartViewer1.ViewportHeight = zoomTo * IIf(aspectRatio > 1, 1, 1 / aspectRatio)
        
            ' Adjust ViewPortLeft and ViewPortTop to keep center point unchanged
114         ChartViewer1.ViewportLeft = centerX - ChartViewer1.ViewportWidth / 2
116         ChartViewer1.ViewportTop = centerY - ChartViewer1.ViewportHeight / 2
        
            ' Update the chart
118         Call ChartViewer1.updateViewPort(True, False)
   
        End With
 
        '<EhFooter>
        Exit Sub

m_frmChartSettings_ZoomBar_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.m_frmChartSettings_ZoomBar " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


'
' User click on the navigate window
'
Private Sub m_frmChartSettings_NavWindowMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        ' Save the mouse coordinates to keep track of how far the navigateWindow has been dragged.
        isDraggingNavigatePad = True
        mouseDownXCoor = x
        mouseDownYCoor = y
    End If
End Sub

'
' User drags the navigate window
'
Private Sub m_frmChartSettings_NavWindowMouseMove(Button As Integer, _
                                                  Shift As Integer, _
                                                  x As Single, _
                                                  y As Single)
        '<EhHeader>
        On Error GoTo m_frmChartSettings_NavWindowMouseMove_Err
        '</EhHeader>

100     With m_frmChartSettings
    
102         If isDraggingNavigatePad Then
    
                ' Is currently dragging - compute where is the navigateWindow being dragged to
                Dim newLabelLeft As Double, newLabelTop As Double
104             newLabelLeft = .NavigateWindow.Left + x - mouseDownXCoor
106             newLabelTop = .NavigateWindow.Top + y - mouseDownYCoor
    
                ' Ensure the navigateWindow is within the navigatePad container
                Dim internalPadWidth As Long, internalPadHeight As Long
108             internalPadWidth = .NavigatePad.Width - ScaleX(2, vbPixels, ScaleMode)
110             internalPadHeight = .NavigatePad.Height - ScaleY(2, vbPixels, ScaleMode)
        
112             If newLabelLeft < 0 Then
114                 newLabelLeft = 0
116             ElseIf newLabelLeft > internalPadWidth - .NavigateWindow.Width Then
118                 newLabelLeft = internalPadWidth - .NavigateWindow.Width
                End If

120             If newLabelTop < 0 Then
122                 newLabelTop = 0
124             ElseIf newLabelTop > internalPadHeight - .NavigateWindow.Height Then
126                 newLabelTop = internalPadHeight - .NavigateWindow.Height
                End If
    
                ' Update the navigateWindow position as it is being dragged
128             .NavigateWindow.Left = newLabelLeft
130             .NavigateWindow.Top = newLabelTop
        
                ' Update the view port to reflect the navigation window
132             ChartViewer1.ViewportLeft = CDbl(.NavigateWindow.Left) / internalPadWidth
134             ChartViewer1.ViewportTop = CDbl(.NavigateWindow.Top) / internalPadHeight
136             Call ChartViewer1.updateViewPort(True, False)
            End If

        End With

        '<EhFooter>
        Exit Sub

m_frmChartSettings_NavWindowMouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.m_frmChartSettings_NavWindowMouseMove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

'
' User release mouse button on the navigate window
'
Private Sub m_frmChartSettings_NavWindowMouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    isDraggingNavigatePad = False
End Sub

'
' Utility to scroll the view port to the given position if necessary
'
Private Sub scrollChartTo(viewer As ChartViewer, vpLeft As Double, vpTop As Double)
        ' Validate the parameters
        '<EhHeader>
        On Error GoTo scrollChartTo_Err
        '</EhHeader>
100     Call viewer.validateViewPort

        ' Update chart only if the view port has changed
102     If vpLeft <> viewer.ViewportLeft Or vpTop <> viewer.ViewportTop Then
104         viewer.ViewportLeft = vpLeft
106         viewer.ViewportTop = vpTop
108         Call viewer.updateViewPort(True, False)
        End If
        '<EhFooter>
        Exit Sub

scrollChartTo_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.scrollChartTo " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub


Private Sub m_frmChartSettings_ZoomInMode()
    ChartViewer1.MouseUsage = cvZoomIn
End Sub

Private Sub m_frmChartSettings_ZoomOutMode()
        ChartViewer1.MouseUsage = cvZoomOut
End Sub

Private Sub m_frmWordBookMarks_ClosingDown(iBkrMark As Integer, bChoosenMark As Variant)
        '<EhHeader>
        On Error GoTo m_frmWordBookMarks_ClosingDown_Err
        '</EhHeader>
100     Xport2Word "", oWordApp.ActiveDocument.Bookmarks.Item(iBkrMark).Name, True
        '<EhFooter>
        Exit Sub

m_frmWordBookMarks_ClosingDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecChart.m_frmWordBookMarks_ClosingDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
