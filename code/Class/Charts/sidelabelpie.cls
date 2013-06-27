VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "sidelabelpie"
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

    ' The data for the pie chart
    'Dim Data()
    'Data = Array(25, 18, 15, 12, 8, 30, 35)

    ' The labels for the pie chart
    'Dim Labels()
    'Labels = Array("Labor", "Licenses", "Taxes", "Legal", "Insurance", _
        "Facilities", "Production")

    ' Create a PieChart object of size 500 x 230 pixels
    Dim c As PieChart
    Set c = cd.PieChart(bgWidth, bgHeight)

    ' Set the center of the pie at (250, 120) and the radius to 100 pixels
    Call c.setPieSize(250, 120, ChartWidth)

    ' Add a title box using 15 points Times Bold Italic as font
    Call c.addTitle(xCaption, "timesbi.ttf", 15)

    ' Draw the pie in 3D
    Call c.set3D

    ' Use the side label layout method
    Call c.setLabelLayout(cd.SideLayout)

    ' Set the pie data and the pie labels
    Call c.SetData(Data, Labels)

    ' output the chart
    Set viewer.Picture = c.makePicture()

    'include tool tip for the chart
    viewer.ImageMap = c.getHTMLImageMap("clickable", "", _
        "title='{label}: US${value}K ({percent}%)'")

End Sub

'Main code for creating chart.
'Note: the argument img is unused because this demo only has 1 chart
Public Sub createChart(viewer As Object, img As String)

    Dim cd As New ChartDirector.API

    ' The data for the pie chart
    Dim Data()
    Data = Array(25, 18, 15, 12, 8, 30, 35)

    ' The labels for the pie chart
    Dim Labels()
    Labels = Array("Labor", "Licenses", "Taxes", "Legal", "Insurance", _
        "Facilities", "Production")

    ' Create a PieChart object of size 500 x 230 pixels
    Dim c As PieChart
    Set c = cd.PieChart(500, 230)

    ' Set the center of the pie at (250, 120) and the radius to 100 pixels
    Call c.setPieSize(250, 120, 100)

    ' Add a title box using 15 points Times Bold Italic as font
    Call c.addTitle("Project Cost Breakdown", "timesbi.ttf", 15)

    ' Draw the pie in 3D
    Call c.set3D

    ' Use the side label layout method
    Call c.setLabelLayout(cd.SideLayout)

    ' Set the pie data and the pie labels
    Call c.SetData(Data, Labels)

    ' output the chart
    Set viewer.Picture = c.makePicture()

    'include tool tip for the chart
    viewer.ImageMap = c.getHTMLImageMap("clickable", "", _
        "title='{label}: US${value}K ({percent}%)'")

End Sub
