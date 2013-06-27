VERSION 5.00
Object = "{84E5CF37-E467-4AC2-89C4-C6002FFB5055}#25.1#0"; "ChartViewer.ocx"
Begin VB.Form frmSecurityCharts 
   Caption         =   "Security Charts"
   ClientHeight    =   5790
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   2790
      TabIndex        =   0
      Top             =   2655
      Width           =   1185
   End
   Begin CDChartViewer.ChartViewer ChartViewer1 
      Height          =   1950
      Left            =   405
      Top             =   225
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   3440
   End
End
Attribute VB_Name = "frmSecurityCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Test()
        '<EhHeader>
        On Error GoTo Test_Err
        '</EhHeader>
        Dim c As Object
        Dim CD As New ChartDirector.API
        Dim DBTable As ChartDirector.DBTable
        Dim RS As New ADODB.Recordset
        Dim CN As New ADODB.Connection
        '
        ' Displays the monthly revenue for the selected year. The selected year should be
        ' passed in as a query parameter called "year"
        '
    '    selectedYear = Request("year")

    '    If selectedYear = "" Then selectedYear = 2001

        ' SQL statement to get the monthly revenues for the selected year.
    '    SQL = "Select Software, Hardware, Services From revenue Where Year(TimeStamp) = " & selectedYear & " Order By TimeStamp"

        '
        ' Connect to database and read the query result into arrays
        '

    
100     RS.Open "SELECT * FROM qryIncByProvince", m_Cnn
    
102     Set DBTable = CD.DBTable(RS)
    
104     RS.Close

        Dim Province As Variant
        Dim Incidents As Variant
        Dim NumberAmbush As Variant
        Dim layer As ChartDirector.BarLayer
    
106     Province = DBTable.getCol(0)
108     Incidents = DBTable.getCol(1)
110     NumberAmbush = DBTable.getCol(2)





        '
        ' Now we have read data into arrays, we can draw the chart using ChartDirector
        '

        ' Create a XYChart object of size 600 x 300 pixels, with a light grey (eeeeee)
        ' background, black border, 1 pixel 3D border effect and rounded corners.
112     Set c = CD.XYChart(600, 300, &HEEEEEE, &H0, 1)
114     Call c.setRoundedFrame

        ' Set the plotarea at (60, 60) and of size 520 x 200 pixels. Set background color to
        ' white (ffffff) and border and grid colors to grey (cccccc)
116     Call c.setPlotArea(60, 60, 520, 200, &HFFFFFF, -1, &HCCCCCC, &HCCCCCCC)

        ' Add a title to the chart using 15pts Times Bold Italic font, with a light blue
        ' (ccccff) background and with glass lighting effects.
118     Call c.addTitle("Scoring ", "timesbi.ttf", 15).setBackground(&HCCCCFF, &H0, CD.glassEffect())

        ' Add a legend box at (70, 32) (top of the plotarea) with 9pts Arial Bold font
120     Call c.addLegend(70, 32, False, "arialbd.ttf", 9).setBackground(CD.Transparent)

        ' Add a stacked bar chart layer using the supplied data
122     Set layer = c.addBarLayer2(CD.Stack)
124     Call layer.addDataSet(Province, , "province")
126     Call layer.addDataSet(Incident, , "Incident")
128     Call layer.addDataSet(NumberAmbush, , "NumberAmbush")

        ' Use soft lighting effect with light direction from the left
130     Call layer.setBorderColor(CD.Transparent, CD.softLighting(CD.Left))

        ' Set the x axis labels. In this example, the labels must be Jan - Dec.
        Dim labels()
132     labels = Array("NumberAmbush", "Incidents")
134     Call c.xAxis().setLabels(labels)

        ' Draw the ticks between label positions (instead of at label positions)
136     Call c.xAxis().setTickOffset(0.5)

        ' Set the y axis title
138     Call c.yAxis().SetTitle("Num Incidents")

        ' Set axes width to 2 pixels
140     Call c.xAxis().setWidth(2)
142     Call c.yAxis().setWidth(2)

144     Set ChartViewer1.Picture = c.makePicture()

        'include tool tip for the chart
146     ChartViewer1.ImageMap = c.getHTMLImageMap("clickable", "", _
            "title='{xLabel}: US${value}K'")

        '<EhFooter>
        Exit Sub

Test_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecurityCharts.Test " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub createChart()
        '<EhHeader>
        On Error GoTo createChart_Err
        '</EhHeader>

        Dim CD As New ChartDirector.API
            Dim DBTable As ChartDirector.DBTable
        Dim RS As New ADODB.Recordset
        Dim CN As New ADODB.Connection
    
    
100     RS.Open "SELECT * FROM qryIncByProvince", m_Cnn
    
102     Set DBTable = CD.DBTable(RS)
    
104     RS.Close

        Dim Province As Variant
        Dim Incidents As Variant
        Dim NumberAerialBombardment As Variant
        Dim NumberAirAccident As Variant
        Dim NumberAirIncident As Variant
        Dim NumberAmbush As Variant
        Dim NumberAnti_personnel As Variant
        Dim NumberArmedAttack As Variant
        Dim NumberAssassination As Variant
        Dim NumberAssault As Variant
        Dim NumberBombing As Variant
        Dim NumberIED As Variant
        Dim NumberMine As Variant
        Dim NumberMortar As Variant
        Dim NumberRiot As Variant
        Dim NumberRPG As Variant
        Dim NumberSmall_ArmsFire  As Variant
        Dim NumberSniping  As Variant
    
        Dim layer As ChartDirector.BarLayer
    
106     Province = DBTable.getCol(0)
108     Incidents = DBTable.getCol(1)
110     NumberAerialBombardment = DBTable.getCol(2)
112     NumberAirAccident = DBTable.getCol(3)
114     NumberAirIncident = DBTable.getCol(4)
116     NumberAmbush = DBTable.getCol(5)
118     NumberAnti_personnel = DBTable.getCol(6)
120     NumberArmedAttack = DBTable.getCol(7)
122     NumberAssassination = DBTable.getCol(8)
124     NumberAssault = DBTable.getCol(9)
126     NumberBombing = DBTable.getCol(10)
128     NumberIED = DBTable.getCol(11)
130     NumberMine = DBTable.getCol(12)
132     NumberMortar = DBTable.getCol(13)
134     NumberRiot = DBTable.getCol(14)
136     NumberRPG = DBTable.getCol(15)
138     NumberSmall_ArmsFire = DBTable.getCol(16)
140     NumberSniping = DBTable.getCol(17)

        ' The data for the bar chart
        Dim Data()
142     Data = Incidents

        ' The labels for the bar chart
        Dim labels()
144     labels = Province

        ' The colors for the bar chart
        Dim colors()
146     colors = Array(&HB8BC9C, &HA0BDC4, &H999966, &H333366, &HC3C3E6)

        ' Create a XYChart object of size 300 x 220 pixels. Use golden background color.
        ' Use a 2 pixel 3D border.
        Dim c As XYChart
148     Set c = CD.XYChart(ScaleX(Me.Height, vbTwips, vbPixels), ScaleY(Me.Width, vbTwips, vbPixels), CD.goldColor(), -1, 2)

        ' Add a title box using 10 point Arial Bold font. Set the background color to
        ' metallic blue (9999FF) Use a 1 pixel 3D border.
150     Call c.addTitle("number of incidents", "arialbd.ttf", 10).setBackground( _
            CD.metalColor(&H9999FF), -1, 1)

        ' Set the plotarea at (40, 40) and of 240 x 150 pixels in size
152     Call c.setPlotArea(0, 0, ScaleX(Me.Height, vbTwips, vbPixels) - 60, ScaleY(Me.Width, vbTwips, vbPixels) - 90)

        ' Add a multi-color bar chart layer using the given data and colors. Use a 1
        ' pixel 3D border for the bars.
154     Call c.addBarLayer3(Data, colors).setBorderColor(-1, 1)

        ' Set the labels on the x axis.
156     Call c.xAxis().setLabels(labels)
        'c.yAxis().setLabels (NumberAmbush)

        ' output the chart
158     Set ChartViewer1.Picture = c.makePicture()

        'include tool tip for the chart
160     ChartViewer1.ImageMap = c.getHTMLImageMap("clickable", "", _
            "title='{xLabel}: {value} Num'")

        '<EhFooter>
        Exit Sub

createChart_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecurityCharts.createChart " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Chart2()
        '<EhHeader>
        On Error GoTo Chart2_Err
        '</EhHeader>
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

        ' Create a XYChart object of size 500 x 320 pixels
        Dim c As XYChart
108     Set c = CD.XYChart(500, 320)

        ' Set the plotarea at (100, 40) and of size 280 x 240 pixels
110     Call c.setPlotArea(100, 40, 280, 240)

        ' Add a legend box at (405, 100)
112     Call c.addLegend(405, 100)

        ' Add a title to the chart
114     Call c.addTitle("Weekday Network Load")

        ' Add a title to the y axis. Draw the title upright (font angle = 0)
116     Call c.yAxis().SetTitle("Average<*br*>Workload<*br*>(MBytes<*br*>Per Hour)" _
            ).setFontAngle(0)

        ' Set the labels on the x axis
118     Call c.xAxis().setLabels(labels)

        ' Add three bar layers, each representing one data set. The bars are drawn in
        ' semi-transparent colors.
120     Call c.addBarLayer(data0, &H808080FF, "Server # 1", 5)
122     Call c.addBarLayer(Data1, &H80FF0000, "Server # 2", 5)
124     Call c.addBarLayer(Data2, &H8000FF00, "Server # 3", 5)

        ' output the chart
126     Set ChartViewer1.Picture = c.makePicture()

        'include tool tip for the chart
128     ChartViewer1.ImageMap = c.getHTMLImageMap("clickable", "", _
            "title='{dataSetName} on {xLabel}: {value} MBytes/hour'")

        '<EhFooter>
        Exit Sub

Chart2_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecurityCharts.Chart2 " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdCommand1_Click()
        '<EhHeader>
        On Error GoTo cmdCommand1_Click_Err
        '</EhHeader>
100     createChart
        '<EhFooter>
        Exit Sub

cmdCommand1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecurityCharts.cmdCommand1_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>

        Dim CD As New ChartDirector.API
        Dim DBTable As ChartDirector.DBTable
        
        If Not g_sLanguage = "" Then
        If Not m_Cnn.State = adStateClosed Then
            LoadLanguage Me.Name, g_sLanguage, m_Cnn
        End If
    End If
    
100     DebugPrint m_Cnn.ConnectionString
   
    '
    'rs is the ADO RecordSet object
    'cd is the ChartDirector.API object
    '
    'Set DBTable = cd.DBTable(rs)
    'col0 = DBTable.getCol(0)
    'col1 = DBTable.getCol(1)
102     createChart
        'test

    Exit Sub

        'The data for the bar chart
        Dim Data()
104     Data = Array(85, 156, 179.5, 211, 123)

        'The labels for the bar chart
        Dim labels()
106     labels = Array("Mon", "Tue", "Wed", "Thu", "Fri")

        'Create a XYChart object of size 250 x 250 pixels
        Dim c As Object
108     Set c = CD.XYChart(250, 250)

        'Set the plotarea at (30, 20) and of size 200 x 200 pixels
110     Call c.setPlotArea(30, 20, 200, 200)

        'Add a bar chart layer using the given data
112     Call c.addBarLayer(Data)

        'Set the x axis labels using the given labels
114     Call c.xAxis().setLabels(labels)

        'output the chart
116     Set ChartViewer1.Picture = c.makePicture()

        'include tool tip for the chart
118     ChartViewer1.ImageMap = c.getHTMLImageMap("clickable", "", _
            "title='{xLabel}: US${value}K'")

        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmSecurityCharts.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
