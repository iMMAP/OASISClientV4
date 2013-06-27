VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "stackedbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Number of charts produced in this demo module
Public Function getNoOfCharts() As Integer
    getNoOfCharts = 1
End Function

'Main code for creating chart.
'Note: the argument img is unused because this demo only has 1 chart
Public Sub createChart(viewer As Object, img As String)

    Dim CD As New ChartDirector.API

    ' The data for the bar chart
    Dim data0()
    data0 = Array(100, 125, 245, 147, 67)
    Dim Data1()
    Data1 = Array(85, 156, 179, 211, 123)
    Dim Data2()
    Data2 = Array(97, 87, 56, 267, 157)

    ' The labels for the bar chart
    Dim labels()
    labels = Array("Mon", "Tue", "Wed", "Thu", "Fri")

    ' Create a XYChart object of size 500 x 320 pixels
    Dim c As XYChart
    Set c = CD.XYChart(500, 320)

    ' Set the plotarea at (100, 40) and of size 280 x 240 pixels
    Call c.setPlotArea(100, 40, 280, 240)

    ' Add a legend box at (400, 100)
    Call c.addLegend(400, 100)

    ' Add a title to the chart using 14 points Times Bold Itatic font
    Call c.addTitle("Weekday Network Load", "timesbi.ttf", 14)

    ' Add a title to the y axis. Draw the title upright (font angle = 0)
    Call c.yAxis().SetTitle("Average<*br*>Workload<*br*>(MBytes<*br*>Per Hour)" _
        ).setFontAngle(0)

    ' Set the labels on the x axis
    Call c.xAxis().setLabels(labels)

    ' Add a stacked bar layer and set the layer 3D depth to 8 pixels
    Dim layer As BarLayer
    Set layer = c.addBarLayer2(CD.Stack, 8)

    ' Add the three data sets to the bar layer
    Call layer.addDataSet(data0, &HFF8080, "Server # 1")
    Call layer.addDataSet(Data1, &H80FF80, "Server # 2")
    Call layer.addDataSet(Data2, &H8080FF, "Server # 3")

    ' Enable bar label for the whole bar
    Call layer.setAggregateLabelStyle

    ' Enable bar label for each segment of the stacked bar
    Call layer.setDataLabelStyle

    ' output the chart
    Set viewer.Picture = c.makePicture()

    'include tool tip for the chart
    viewer.ImageMap = c.getHTMLImageMap("clickable", "", _
        "title='{dataSetName} on {xLabel}: {value} MBytes/hour'")

End Sub

'Main code for creating chart.
'Note: the argument img is unused because this demo only has 1 chart
Public Sub createIncidentChart(viewer As Object, _
                               img As String, _
                               ilgdTop As Integer, _
                               ilgdLeft As Integer, _
                               iPlotWidth As Integer, _
                               iPlotHeight As Integer, _
                               iChartTop As Integer, _
                               iChartLeft As Integer, _
                               iChartWidth As Integer, _
                               iChartHeight As Integer, _
                               sTitle As String, _
                               sYTitle As String, _
                               sLocSQL As String, _
                               sIncidentSQL As String, _
                               sFilterSQL As String, _
                               Optional iRotation As Integer = 0, _
                               Optional ilStep As Integer = 1, _
                               Optional isStep As Integer = 0)
        '<EhHeader>
        On Error GoTo createIncidentChart_Err
        '</EhHeader>
        Dim RSInc As ADODB.Recordset
        Dim RSProv As New ADODB.Recordset
        Dim CN As New ADODB.Connection
        Dim sAdmRegions() As String
        Dim VarVals() As Variant
        Dim strVals As String
        Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim sMorningVals As String
        Dim sAfternoonVals As String
        Dim sEveningVals As String
        Dim sNightVals As String
        Dim sEmptyVals As String
        Dim sArMorningVals() As String
        Dim sArAfternoonVals() As String
        Dim sArEveningVals() As String
        Dim sArNightVals() As String
        Dim sArEmptyVals() As String

        Dim c As XYChart
        Dim layer As BarLayer
        Dim CD As New ChartDirector.API

        ' The data for the bar chart
        Dim data0()
100     data0 = Array(100, 125, 245, 147, 67)
        Dim Data1()
102     Data1 = Array(85, 156, 179, 211, 123)
        Dim Data2()
104     Data2 = Array(97, 87, 56, 267, 157)

        ' The labels for the bar chart
        Dim labels()
106     labels = Array("Mon", "Tue", "Wed", "Thu", "Fri")

108     'cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\OASIS\Client\data\db\Oasisclient.mdb;Mode=ReadWrite|Share Deny None;Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;"
110     'cn.Open
    
        '************************************************************
    
        Dim oRS As New ADODB.Recordset
    
        '***********************************************************
112     RSProv.CursorLocation = g_sGlobalCursorLocation
114     RSProv.Open sLocSQL & sFilterSQL, m_Cnn, adOpenDynamic, adLockReadOnly
        'RSProv.Open "SELECT DISTINCT Province FROM oincidents_FEA", cn, adOpenDynamic, adLockReadOnly
    
116     If RSProv.EOF And RSProv.Bof Then Exit Sub
    
118     SafeMoveFirst RSProv
    
120     ReDim VarVals(0)
    
122     ReDim sAdmRegions(RSProv.RecordCount - 1)
    
124     Set oRS = New ADODB.Recordset
126     oRS.CursorLocation = g_sGlobalCursorLocation
    
128     oRS.Open sIncidentSQL, m_Cnn, adOpenDynamic, adLockReadOnly

        '    oRS.Open "SELECT * FROM oincidents_FEA", cn

130     For k = LBound(sAdmRegions) To UBound(sAdmRegions)
        
132         sAdmRegions(k) = RSProv.Fields.Item("Province").Value
        
134         oRS.Filter = "TIME00 = 'Morning (05-12)' AND Province = '" & sAdmRegions(k) & "'"
136         sMorningVals = sMorningVals & IIf(Len(sMorningVals) > 0, ",", "") & oRS.RecordCount
        
138         oRS.Filter = "TIME00 ='Afternoon (12-18)' AND Province = '" & sAdmRegions(k) & "'"
140         sAfternoonVals = sAfternoonVals & IIf(Len(sAfternoonVals) > 0, ",", "") & oRS.RecordCount
        
142         oRS.Filter = "TIME00 ='Evening (18-22)' AND Province = '" & sAdmRegions(k) & "'"
144         sEveningVals = sEveningVals & IIf(Len(sEveningVals) > 0, ",", "") & oRS.RecordCount
        
146         oRS.Filter = "TIME00 ='Night (22-05)' AND Province = '" & sAdmRegions(k) & "'"
148         sNightVals = sNightVals & IIf(Len(sNightVals) > 0, ",", "") & oRS.RecordCount
        
150         oRS.Filter = "TIME00 ='' AND Province = '" & sAdmRegions(k) & "'"
152         sEmptyVals = sEmptyVals & IIf(Len(sEmptyVals) > 0, ",", "") & oRS.RecordCount
        
154         RSProv.MoveNext
        Next

156     sArMorningVals = Split(sMorningVals, ",")
158     sArAfternoonVals = Split(sAfternoonVals, ",")
160     sArEveningVals = Split(sEveningVals, ",")
162     sArNightVals = Split(sNightVals, ",")
164     sArEmptyVals = Split(sEmptyVals, ",")

        '******************************

        ' Create a XYChart object of size 500 x 320 pixels
166     Set c = CD.XYChart(iPlotWidth, iPlotHeight)

        ' Set the plotarea at (100, 40) and of size 280 x 240 pixels
168     Call c.setPlotArea(iChartLeft, iChartTop, iChartWidth, iChartHeight)

        ' Add a legend box at (400, 100)
170     Call c.addLegend(ilgdLeft, ilgdTop)

        ' Add a title to the chart using 14 points Times Bold Itatic font
172     Call c.addTitle(sTitle, "timesbi.ttf", 14)

        ' Add a title to the y axis. Draw the title upright (font angle = 0)
        '"Incident<*br*>Distribution<*br*>Over Day"
174     Call c.yAxis().SetTitle(sYTitle).setFontAngle(90)

        ' Set the labels on the x axis
176     Call c.xAxis().setLabels(sAdmRegions).setFontAngle(iRotation) 'labels)
    
178     Call c.xAxis().setLabelStep(ilStep, isStep)
    
        ' Add a stacked bar layer and set the layer 3D depth to 8 pixels
180     Set layer = c.addBarLayer2(CD.Stack, 8)

        ' Add the three data sets to the bar layer
    
182     Call layer.addDataSet(sArMorningVals, , "Morning")
184     Call layer.addDataSet(sArAfternoonVals, , "Afternoon")
186     Call layer.addDataSet(sArEveningVals, , "Evening")
188     Call layer.addDataSet(sArNightVals, , "Night")
190     Call layer.addDataSet(sArEmptyVals, , "N/A")
    
        'For i = LBound(sAdmRegions) To UBound(sAdmRegions)
        '    Call layer.addDataSet(Split(VarVals(i), ","), , "Line" & i)
        'Next i
    
        '    Call layer.addDataSet(data0, &HFF8080, "Server # 1")
        '    Call layer.addDataSet(data1, &H80FF80, "Server # 2")
        '    Call layer.addDataSet(data2, &H8080FF, "Server # 3")

        ' Enable bar label for the whole bar
192     Call layer.setAggregateLabelStyle

        ' Enable bar label for each segment of the stacked bar
194     Call layer.setDataLabelStyle

        ' output the chart
196     Set viewer.Picture = c.makePicture()

        'include tool tip for the chart
198     viewer.ImageMap = c.getHTMLImageMap("clickable", "", "title='{dataSetName} Incidents In {xLabel}: {value} Incidents Frequency/Day'")

        Exit Sub

        '********************************

200     Do While Not RSProv.EOF
    
            '**********************************
        
202         Set oRS = New ADODB.Recordset
204         oRS.CursorLocation = g_sGlobalCursorLocation
206         oRS.Open "SELECT * FROM oincidents_FEA", CN
208         DebugPrint oRS.RecordCount
210         oRS.Filter = "TIME00 ='Morning (05-12)' AND Province = '" & RSProv.Fields.Item("Province").Value & "'"
212         DebugPrint oRS.RecordCount
        
214         Set oRS = New ADODB.Recordset
216         oRS.Open "SELECT * FROM oincidents_FEA WHERE TIME00 ='Afternoon (12-18)'", CN
        
218         Set oRS = New ADODB.Recordset
220         oRS.Open "SELECT * FROM oincidents_FEA WHERE TIME00 =", CN
        
222         Set oRS = New ADODB.Recordset
224         oRS.Open "SELECT * FROM oincidents_FEA WHERE TIME00 =", CN
        
226         Set oRS = New ADODB.Recordset
228         oRS.Open "SELECT * FROM oincidents_FEA WHERE TIME00 =''", CN
        
            '*********************************

230         If j = 10000 Then

232             DebugPrint RSProv.Fields.Item("Province").Value
234             sAdmRegions(UBound(sAdmRegions)) = RSProv.Fields.Item("Province").Value
236             Set RSInc = New ADODB.Recordset
238             RSInc.CursorLocation = g_sGlobalCursorLocation
240             RSInc.Open "SELECT * FROM oincidents_FEA WHERE Province = '" & RSProv.Fields.Item("Province").Value & "'", CN, adOpenDynamic, adLockReadOnly
        
242             SafeMoveFirst RSInc
        
                'DebugPrint RSInc.RecordCount
        
244             RSInc.Filter = "TIME00 = 'Morning (05-12)'"
246             strVals = RSInc.RecordCount
248             RSInc.Filter = "TIME00 = 'Afternoon (12-18)'"
250             strVals = strVals & "," & RSInc.RecordCount
252             RSInc.Filter = "TIME00 = 'Evening (18-22)'"
254             strVals = strVals & "," & RSInc.RecordCount
256             RSInc.Filter = "TIME00 = 'Night (22-05)'"
258             strVals = strVals & "," & RSInc.RecordCount
        
260             VarVals(UBound(VarVals)) = Split(strVals, ",")
        
                'DebugPrint RSInc.RecordCount
                'RSInc.Filter = "TIME00 = ''"
                'DebugPrint RSInc.RecordCount
        
            End If
        
262         RSProv.MoveNext
        
264         If Not RSProv.EOF Then
266             ReDim Preserve VarVals(UBound(VarVals) + 1)
268             ReDim Preserve sAdmRegions(UBound(sAdmRegions) + 1)
            End If
            
        Loop

        Exit Sub

        '   ReDim Preserve sAdmRegions(UBound(sAdmRegions) - 1)

        ' Create a XYChart object of size 500 x 320 pixels
270     Set c = CD.XYChart(500, 320)

        ' Set the plotarea at (100, 40) and of size 280 x 240 pixels
272     Call c.setPlotArea(100, 40, 280, 240)

        ' Add a legend box at (400, 100)
274     Call c.addLegend(400, 100)

        ' Add a title to the chart using 14 points Times Bold Itatic font
276     Call c.addTitle("Weekday Network Load", "timesbi.ttf", 14)

        ' Add a title to the y axis. Draw the title upright (font angle = 0)
278     Call c.yAxis().SetTitle("Average<*br*>Workload<*br*>(MBytes<*br*>Per Hour)").setFontAngle(0)

        ' Set the labels on the x axis
280     Call c.xAxis().setLabels(sAdmRegions) 'labels)

        ' Add a stacked bar layer and set the layer 3D depth to 8 pixels
282     Set layer = c.addBarLayer2(CD.Stack, 8)

        ' Add the three data sets to the bar layer
    
284     For i = LBound(sAdmRegions) To UBound(sAdmRegions)
286         Call layer.addDataSet(Split(VarVals(i), ","), , "Line" & i)
288     Next i
    
        '    Call layer.addDataSet(data0, &HFF8080, "Server # 1")
        '    Call layer.addDataSet(data1, &H80FF80, "Server # 2")
        '    Call layer.addDataSet(data2, &H8080FF, "Server # 3")

        ' Enable bar label for the whole bar
290     Call layer.setAggregateLabelStyle

        ' Enable bar label for each segment of the stacked bar
292     Call layer.setDataLabelStyle

        ' output the chart
294     Set viewer.Picture = c.makePicture()

        'include tool tip for the chart
296     viewer.ImageMap = c.getHTMLImageMap("clickable", "", "title='{dataSetName} on {xLabel}: {value} MBytes/hour'")

        '<EhFooter>
        Exit Sub

createIncidentChart_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.stackedbar.createIncidentChart", _
                  "stackedbar component failure"
        '</EhFooter>
End Sub

