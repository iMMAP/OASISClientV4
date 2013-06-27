VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "glasslightbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Number of charts produced in this demo module
Public Function getNoOfCharts() As Integer
    getNoOfCharts = 1
End Function

Public Sub createChartEx(viewer As ChartViewer, Labels As Variant, Data As Variant, Optional bgWidth As Integer = 600, Optional bgHeight As Integer = 360, Optional ChartWidth As Integer = 480, Optional ChartHeight As Integer = 240, Optional xCaption As String, Optional YCaption As String, Optional img As String)

    Dim cd As New ChartDirector.API


    ' Create a XYChart object of size 600 x 360 pixels
    Dim c As XYChart
    Set c = cd.XYChart(bgWidth, bgHeight)

    ' Set the plotarea at (60, 40) and of size 480 x 280 pixels. Use a vertical
    ' gradient color from light blue (eeeeff) to deep blue (0000cc) as background.
    ' Set border and grid lines to white (ffffff).
    Call c.setPlotArea(60, 40, ChartWidth, ChartHeight, c.linearGradientColor(60, 40, 60, 280, _
        &HEEEEFF, &HCC), -1, &HFFFFFF, &HFFFFFF)

    ' Add a title to the chart using 18pts Times Bold Italic font
    Call c.addTitle(xCaption, "timesbi.ttf", 18)

    ' Add a multi-color bar chart layer using the supplied data. Use glass lighting
    ' effect with light direction from the left.
    Call c.addBarLayer3(Data).setBorderColor(cd.Transparent, cd.glassEffect( _
        cd.NormalGlare, cd.Left))

    ' Set the x axis labels
    Call c.xAxis().setLabels(Labels)

    ' Show the same scale on the left and right y-axes
    Call c.syncYAxis

    ' Set the left y-axis and right y-axis title using 10pt Arial Bold font
    Call c.yAxis().setTitle(YCaption, "arialbd.ttf", 10)
    Call c.yAxis2().setTitle(YCaption, "arialbd.ttf", 10)

    ' Set all axes to transparent
    Call c.xAxis().setColors(cd.Transparent)
    Call c.yAxis().setColors(cd.Transparent)
    Call c.yAxis2().setColors(cd.Transparent)

    ' Set the label styles of all axes to 8pt Arial Bold font
    Call c.xAxis().setLabelStyle("arialbd.ttf", 8, , 45)
    Call c.yAxis().setLabelStyle("arialbd.ttf", 8)
    Call c.yAxis2().setLabelStyle("arialbd.ttf", 8)

    ' output the chart
    Set viewer.Picture = c.makePicture()

    'include tool tip for the chart
    viewer.ImageMap = c.getHTMLImageMap("clickable", "", _
        "title='Year {xLabel}: US$ {value}M'")

End Sub


'Main code for creating chart.
'Note: the argument img is unused because this demo only has 1 chart
Public Sub createChart(viewer As Object, img As String)

    Dim cd As New ChartDirector.API

    ' The data for the bar chart
    Dim Data()
    Data = Array(450, 560, 630, 800, 1100, 1350, 1600, 1950, 2300, 2700)

    ' The labels for the bar chart
    Dim Labels()
    Labels = Array("1996", "1997", "1998", "1999", "2000", "2001", "2002", "2003", _
        "2004", "2005")

    ' Create a XYChart object of size 600 x 360 pixels
    Dim c As XYChart
    Set c = cd.XYChart(600, 360)

    ' Set the plotarea at (60, 40) and of size 480 x 280 pixels. Use a vertical
    ' gradient color from light blue (eeeeff) to deep blue (0000cc) as background.
    ' Set border and grid lines to white (ffffff).
    Call c.setPlotArea(60, 40, 480, 280, c.linearGradientColor(60, 40, 60, 280, _
        &HEEEEFF, &HCC), -1, &HFFFFFF, &HFFFFFF)

    ' Add a title to the chart using 18pts Times Bold Italic font
    Call c.addTitle("Annual Revenue for Star Tech", "timesbi.ttf", 18)

    ' Add a multi-color bar chart layer using the supplied data. Use glass lighting
    ' effect with light direction from the left.
    Call c.addBarLayer3(Data).setBorderColor(cd.Transparent, cd.glassEffect( _
        cd.NormalGlare, cd.Left))

    ' Set the x axis labels
    Call c.xAxis().setLabels(Labels)

    ' Show the same scale on the left and right y-axes
    Call c.syncYAxis

    ' Set the left y-axis and right y-axis title using 10pt Arial Bold font
    Call c.yAxis().setTitle("USD (millions)", "arialbd.ttf", 10)
    Call c.yAxis2().setTitle("USD (millions)", "arialbd.ttf", 10)

    ' Set all axes to transparent
    Call c.xAxis().setColors(cd.Transparent)
    Call c.yAxis().setColors(cd.Transparent)
    Call c.yAxis2().setColors(cd.Transparent)

    ' Set the label styles of all axes to 8pt Arial Bold font
    Call c.xAxis().setLabelStyle("arialbd.ttf", 8)
    Call c.yAxis().setLabelStyle("arialbd.ttf", 8)
    Call c.yAxis2().setLabelStyle("arialbd.ttf", 8)

    ' output the chart
    Set viewer.Picture = c.makePicture()

    'include tool tip for the chart
    viewer.ImageMap = c.getHTMLImageMap("clickable", "", _
        "title='Year {xLabel}: US$ {value}M'")

End Sub
