VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "multiline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' The earliest date and the duration in seconds for horizontal scrolling
Private minDate As Date
Private dateRange As Double
        
' The vertical range of the chart for vertical scrolling
Private maxValue As Double
Private minValue As Double

'Number of charts produced in this demo module
Public Function getNoOfCharts() As Integer
    getNoOfCharts = 1
End Function

'Main code for creating chart.
'Note: the argument img is unused because this demo only has 1 chart
Public Sub CreateChart(viewer As Object, img As String)

    Dim cd As New ChartDirector.API

    ' The data for the line chart
    Dim data0()
    data0 = Array(42, 49, 33, 38, 51, 46, 29, 41, 44, 57, 59, 52, 37, 34, 51, 56, _
        56, 60, 70, 76, 63, 67, 75, 64, 51)
    Dim data1()
    data1 = Array(50, 55, 47, 34, 42, 49, 63, 62, 73, 59, 56, 50, 64, 60, 67, 67, _
        58, 59, 73, 77, 84, 82, 80, 84, 98)
    Dim data2()
    data2 = Array(36, 28, 25, 33, 38, 20, 22, 30, 25, 33, 30, 24, 28, 15, 21, 26, _
        46, 42, 48, 45, 43, 52, 64, 60, 70)

    ' The labels for the line chart
    Dim Labels()
    Labels = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", _
        "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", _
        "24")

    ' Create an XYChart object of size 600 x 300 pixels, with a light blue (EEEEFF)
    ' background, black border, 1 pxiel 3D border effect and rounded corners
    Dim c As XYChart
    Set c = cd.XYChart(600, 300, &HEEEEFF, &H0, 1)
    Call c.setRoundedFrame

    ' Set the plotarea at (55, 58) and of size 520 x 195 pixels, with white
    ' background. Turn on both horizontal and vertical grid lines with light grey
    ' color (0xcccccc)
    Call c.setPlotArea(55, 58, 520, 195, &HFFFFFF, -1, -1, &HCCCCCC, &HCCCCCC)

    ' Add a legend box at (50, 30) (top of the chart) with horizontal layout. Use 9
    ' pts Arial Bold font. Set the background and border color to Transparent.
    Call c.addLegend(50, 30, False, "arialbd.ttf", 9).setBackground(cd.Transparent)

    ' Add a title box to the chart using 15 pts Times Bold Italic font, on a light
    ' blue (CCCCFF) background with glass effect. white (0xffffff) on a dark red
    ' (0x800000) background, with a 1 pixel 3D border.
    Call c.addTitle("Application Server Throughput", "timesbi.ttf", 15 _
        ).setBackground(&HCCCCFF, &H0, cd.glassEffect())

    ' Add a title to the y axis
    Call c.yAxis().setTitle("MBytes per hour")

    ' Set the labels on the x axis.
    Call c.xAxis().setLabels(Labels)

    ' Display 1 out of 3 labels on the x-axis.
    Call c.xAxis().setLabelStep(3)

    ' Add a title to the x axis
    Call c.xAxis().setTitle("Jun 12, 2006")

    ' Add a line layer to the chart
    Dim layer As LineLayer
    Set layer = c.addLineLayer2()

    ' Set the default line width to 2 pixels
    Call layer.setLineWidth(2)

    ' Add the three data sets to the line layer. For demo purpose, we use a dash line
    ' color for the last line
    Call layer.addDataSet(data0, &HFF0000, "Server #1")
    Call layer.addDataSet(data1, &H8800, "Server #2")
    Call layer.addDataSet(data2, c.dashLineColor(&H3333FF, cd.DashLine), _
        "Server #3")

    ' output the chart
    Set viewer.Picture = c.makePicture()

    'include tool tip for the chart
    viewer.ImageMap = c.getHTMLImageMap("clickable", "", _
        "title='[{dataSetName}] Hour {xLabel}: {value} MBytes'")

End Sub

'Main code for creating chart.
'Note: the argument img is unused because this demo only has 1 chart
Public Sub createChartex1(viewer As Object, _
                          img As String, _
                          Optional Labels As Variant, _
                          Optional Data As Variant, _
                          Optional ValueLabels As Variant, _
                          Optional bgWidth As Integer = 600, _
                          Optional bgHeight As Integer = 360, _
                          Optional ChartWidth As Integer = 480, _
                          Optional ChartHeight As Integer = 240, _
                          Optional xCaption As String, _
                          Optional YCaption As String, Optional bUseTrend As Boolean, Optional iRotation As Integer = 0, Optional ilStep As Integer = 1, Optional isStep As Integer = 0)

    Dim cd As New ChartDirector.API

    ' Create an XYChart object of size 600 x 300 pixels, with a light blue (EEEEFF)
    ' background, black border, 1 pxiel 3D border effect and rounded corners
    Dim c As XYChart
    Set c = cd.XYChart(bgWidth, bgHeight, &HEEEEFF, &H0, 1)
    Call c.setRoundedFrame

    ' Set the plotarea at (55, 58) and of size 520 x 195 pixels, with white
    ' background. Turn on both horizontal and vertical grid lines with light grey
    ' color (0xcccccc)
    Call c.setPlotArea(55, 58, ChartWidth, ChartHeight, &HFFFFFF, -1, -1, &HCCCCCC, &HCCCCCC)

    ' Add a legend box at (50, 30) (top of the chart) with horizontal layout. Use 9
    ' pts Arial Bold font. Set the background and border color to Transparent.
    Call c.addLegend(50, 30, False, "arialbd.ttf", 9).setBackground(cd.Transparent)

    ' Add a title box to the chart using 15 pts Times Bold Italic font, on a light
    ' blue (CCCCFF) background with glass effect. white (0xffffff) on a dark red
    ' (0x800000) background, with a 1 pixel 3D border.
    Call c.addTitle(xCaption, "timesbi.ttf", 15).setBackground(&HCCCCFF, &H0, cd.glassEffect())

    ' Add a title to the y axis
    Call c.yAxis().setTitle(YCaption)

    ' Set the labels on the x axis.
    Call c.xAxis().setLabels(ValueLabels).setFontAngle(iRotation)

    ' Display 1 out of 3 labels on the x-axis.
    Call c.xAxis().setLabelStep(ilStep, isStep)

    ' Add a title to the x axis
    Call c.xAxis().setTitle(xCaption)

    ' Add a line layer to the chart
    Dim layer As LineLayer
    Set layer = c.addLineLayer2()

    ' Set the default line width to 2 pixels
    Call layer.setLineWidth(2)

    Dim k As Long
    
    ' Add a trend line layer using the same data with a dark green (0x008000) color.
    ' Set the line width to 2 pixels
    'Dim oLyr As Object
    Dim trendLayer() As trendLayer
    
    ReDim Preserve trendLayer(0)
    
    For k = 0 To UBound(Data) - 1
    
        Call layer.addDataSet(Data(k), -1, Labels(k))
        
        If bUseTrend Then
            Set trendLayer(k) = c.addTrendLayer(Data(k), -1, Labels(k) & " Trend")
            Call trendLayer(k).setLineWidth(3)
                ' Add a 95% confidence band for the line
            Call trendLayer(k).addConfidenceBand(0.95, &H806666FF)
            ' Add a 95% confidence band (prediction band) for the points
            Call trendLayer(k).addPredictionBand(0.95, &H8066FF66)
            If Not k = UBound(Data) - 1 Then ReDim Preserve trendLayer(UBound(trendLayer) + 1)
        End If
    Next


'    '================================================================================
'    ' Step 3 - Set up x-axis scale
'    '================================================================================
'
'    ' Set x-axis date scale to the view port date range. ChartDirector auto-scaling will
'    ' automatically determine the ticks on the axis.
'
'    Dim viewPortStartDate As Date
'    Dim viewPortEndDate As Date
'
'    dateRange = DateDiff("d", ValueLabels(0), ValueLabels(UBound(ValueLabels)))
'
'    viewPortStartDate = DateAdd("d", viewer.ViewportLeft * dateRange, ValueLabels(0))
'    viewPortEndDate = DateAdd("d", viewer.ViewportWidth * dateRange, viewPortStartDate)
'
'        Call c.xAxis().setDateScale(viewPortStartDate, viewPortEndDate)
'    '
'    ' If all ticks are yearly aligned, then we use "yyyy" as the label format.
'    Call c.xAxis().setFormatCondition("align", 31104000)
'    Call c.xAxis().setLabelFormat("{value|yyyy}")
'
'    ' If all ticks are monthly aligned, then we use "mmm yyyy" in bold font as the first label of a
'    ' year, and "mmm" for other labels.
'    Call c.xAxis().setFormatCondition("align", 2592000)
'    Call c.xAxis().setMultiFormat(cd.StartOfYearFilter(), "<*font=bold*>{value|mmm yyyy}", cd.AllPassFilter(), "{value|mmm}")
'
'    ' If all ticks are daily algined, then we use "mmm dd<*br*>yyyy" in bold font as the first label
'    ' of a year, and "mmm dd" in bold font as the first label of a month, and "dd" for other labels.
'    Call c.xAxis().setFormatCondition("align", 86400)
'    Call c.xAxis().setMultiFormat(cd.StartOfYearFilter(), "<*block,halign=left*><*font=bold*>{value|mmm dd<*br*>yyyy}", cd.StartOfMonthFilter(), "<*font=bold*>{value|mmm dd}")
'    Call c.xAxis().setMultiFormat2(cd.AllPassFilter(), "{value|dd}")
'
'    ' For all other cases (sub-daily ticks), use "hh:nn<*br*>mmm dd" for the first label of a day,
'    ' and "hh:nn" for other labels.
'    Call c.xAxis().setFormatCondition("else")
'    Call c.xAxis().setMultiFormat(cd.StartOfDayFilter(), "<*font=bold*>{value|hh:nn<*br*>mmm dd}", cd.AllPassFilter(), "{value|hh:nn}")
'
'
'    '================================================================================
'    ' Step 4 - Set up y-axis scale
'    '================================================================================
'
'    maxValue = 200
'    minValue = 0
'
'    If viewer.ZoomDirection = cvHorizontal Or minValue = maxValue Then
'        ' y-axis is auto-scaled - save the chosen y-axis scaled to support xy-zoom mode
'        Call c.layout
'        minValue = c.yAxis().getMinValue()
'        maxValue = c.yAxis().getMaxValue()
'    Else
'        ' xy-zoom mode - compute the actual axis scale in the view port
'        Dim axisLowerLimit As Double
'        Dim axisUpperLimit As Double
'        axisLowerLimit = maxValue - (maxValue - minValue) * _
'            (viewer.ViewportTop + viewer.ViewportHeight)
'        axisUpperLimit = maxValue - (maxValue - minValue) * viewer.ViewportTop
'        ' *** use the following formula if you are using a log scale axis ***
'        ' axisLowerLimit = maxValue * ((minValue / maxValue) ^ _
'        '     (viewer.ViewPortTop + viewer.ViewPortHeight))
'        ' axisUpperLimit = maxValue * ((minValue / maxValue) ^ viewer.ViewPortTop)
'
'        ' use the zoomed-in scale
'        Call c.yAxis().setLinearScale(axisLowerLimit, axisUpperLimit)
'        Call c.yAxis().setRounding(False, False)
'    End If
'
'
'
'    Set viewer.Chart = c
    Set viewer.Picture = c.makePicture()

    'include tool tip for the chart
    viewer.ImageMap = c.getHTMLImageMap("clickable", "", "title='[{dataSetName}] Frequency {xLabel}: {value} Incidents'")

End Sub

