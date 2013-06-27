VERSION 5.00
Object = "{E9DF30CA-4B30-4235-BF0C-7150F6466080}#1.0#0"; "ChartFX.ClientServer.Core.dll"
Begin VB.UserControl OASISChartingVer2 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H0080C0FF&
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   ControlContainer=   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   5580
   Begin Cfx62ClientServerCtl.Chart Chart1 
      Height          =   1995
      Left            =   1320
      TabIndex        =   0
      Top             =   780
      Width           =   2775
      _Data_          =   "OASISChartingV2.ctx":0000
   End
End
Attribute VB_Name = "OASISChartingVer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private cxAnnotation As Cfx62ClientServerAnnotation.AnnotationX

Public Sub MakeVisible()
    Chart1.Visible = True
End Sub

Public Sub MakeInvisible()
    Chart1.Visible = False
End Sub

Public Sub ResizeHeight(iPercentage As Integer)
    Chart1.Height = Chart1.Height * (iPercentage / 100)
End Sub



Public Sub ResetChart()
        '<EhHeader>
        On Error GoTo ResetChart_Err
        '</EhHeader>
100     Chart1.Legend.Clear
        Chart1.ClearData ClearDataFlag_Labels
102     Chart1.LegendBox = False
104     'Chart1.ClearData ClearDataFlag_AllData
106     'Chart1.SerLegBoxObj.Invalidate
108     'Chart1.LegendBoxObj.Invalidate
110     'Chart1.ClearData ClearDataFlag_Strings
112     Chart1.ContextMenus = True
114     RefreshChart
        '<EhFooter>
        Exit Sub

ResetChart_Err:
        MsgBox "OASISChartingVer2.ResetChart_Err (line " & Erl & "): " & Err.Description
        '</EhFooter>
End Sub

Public Sub RefreshChart()
        '<EhHeader>
        On Error GoTo RefreshChart_Err
        '</EhHeader>
100     'C1Elastic1.Refresh
        '<EhFooter>
        Exit Sub

RefreshChart_Err:
        MsgBox "OASISChartingVer2.RefreshChart_Err (line " & Erl & "): " & Err.Description
        '</EhFooter>
End Sub

Public Sub ClearDataAll()
        '<EhHeader>
        On Error GoTo ClearDataAll_Err
        '</EhHeader>
100     Chart1.ClearData ClearDataFlag_AllData
        '<EhFooter>
        Exit Sub

ClearDataAll_Err:
        MsgBox "OASISChartingVer2.ClearDataAll_Err (line " & Erl & "): " & Err.Description
        '</EhFooter>
End Sub

Public Sub ClearDataLabels()
    Chart1.ClearData ClearDataFlag_Labels
End Sub

Public Sub ClearDataStrings()
        '<EhHeader>
        On Error GoTo ClearDataStrings_Err
        '</EhHeader>
100     Chart1.ClearData ClearDataFlag_Strings
        '<EhFooter>
        Exit Sub

ClearDataStrings_Err:
        MsgBox "OASISChartingVer2.ClearDataStrings_Err (line " & Erl & "): " & Err.Description
        '</EhFooter>
End Sub

Public Sub TemplateLoad(sPath As String)
    '<EhHeader>
    On Error GoTo TemplateLoad_Err
    '</EhHeader>

    ResetChart
    
    With Chart1
     
        If Len(sPath) > 0 Then
            .ClearData ClearDataFlag_Strings
            .Import FileFormat_Binary, sPath
        End If

        Chart1.ContextMenus = True
        
    End With

    '<EhFooter>
    Exit Sub

TemplateLoad_Err:
    ResetChart
    MsgBox "OASISChartingVer2.TemplateLoad_Err (line " & Erl & "): " & Err.Description
    '</EhFooter>
End Sub

Public Sub TemplateSave(sPath As String)
    '<EhHeader>
    On Error GoTo TemplateSave_Err
    '</EhHeader>
        
 
    With Chart1
            
        Chart1.ClearData ClearDataFlag_Labels
        .Legend.Clear
        .FileMask = FileMask_All
        .Export FileFormat_Binary, sPath
        
    End With

    '<EhFooter>
    Exit Sub

TemplateSave_Err:
    MsgBox "OASISChartingVer2.TemplateSave_Err (line " & Erl & "): " & Err.Description
    '</EhFooter>
End Sub
Public Sub LoadRS(oRS As ADODB.Recordset)
        '<EhHeader>
        On Error GoTo LoadRS_Err
        '</EhHeader>

        Dim i As Long
        Dim sRet As String
        Dim j As Long
        Dim iCountOfVal As String
    
100     i = 0
102     j = 0
104     iCountOfVal = 0
106     Chart1.MultipleColors = True
    
108     Do Until i = oRS.Fields.Count
    
110         If oRS.Fields(i).Type = 5 Or oRS.Fields(i).Type = 3 Then iCountOfVal = iCountOfVal + 1
112         i = i + 1
        
        Loop
    
114     i = 0

116     Do Until i = oRS.Fields.Count
    
118         If oRS.Fields(i).Type = 5 Or oRS.Fields(i).Type = 3 Then
        
120             Chart1.dataType.Item(i) = DataType_Value
            
122             If iCountOfVal > 1 Then
124                 Chart1.SerLeg(j) = Trim(oRS.Fields(i).Name)
126                 Chart1.MultipleColors = False
128                 j = j + 1
                End If
            
130         ElseIf oRS.Fields(i).Type = 202 Then
        
132             Chart1.dataType.Item(i) = DataType_Label

            Else
       
134             Chart1.dataType.Item(i) = DataType_Default
            
            End If
    
136         i = i + 1
        Loop
    
138     Chart1.DataStyle = DataStyle_ReadXValues
140     Chart1.DataSource = oRS
        'LoadRS = True
    
142     i = 0

144     Do Until i = Chart1.AxisX.Label.Count
    
146         Chart1.AxisX.Label(i) = Trim$(Chart1.AxisX.Label(i))
148         i = i + 1
        Loop
    
150     Chart1.AxisX.AdjustScale
    
        '<EhFooter>
        Exit Sub

LoadRS_Err:
        MsgBox "OASISChartingVer2.LoadRS_Err (line " & Erl & "): " & Err.Description
        '</EhFooter>
End Sub


Public Sub ImageSave(sPath As String)
        '<EhHeader>
        On Error GoTo ImageSave_Err
        '</EhHeader>

100     With Chart1

102         .FileMask = FileMask_All
104         .Export FileFormat_Bitmap, sPath
        
        End With

        '<EhFooter>
        Exit Sub

ImageSave_Err:
        MsgBox "OASISChartingVer2.ImageSave_Err (line " & Erl & "): " & Err.Description
        '</EhFooter>
End Sub

Private Sub UserControl_Initialize()
        '<EhHeader>
        On Error GoTo UserControl_Initialize_Err
        '</EhHeader>

100     If cxAnnotation Is Nothing Then
        
102         Set cxAnnotation = New Cfx62ClientServerAnnotation.AnnotationX
104         Chart1.Extensions.Add cxAnnotation
106         cxAnnotation.Enabled = True
108         cxAnnotation.ToolBar = False
        
        End If
        
        Chart1.Move 0, 0, UserControl.Width, UserControl.Height

        '<EhFooter>
        Exit Sub

UserControl_Initialize_Err:
        MsgBox "OASISChartingVer2.UserControl_Initialize_Err (line " & Erl & "): " & Err.Description
        '</EhFooter>
End Sub

Private Sub UserControl_Resize()
    Chart1.Move 0, 0, UserControl.Width, UserControl.Height
    Chart1.RecalcScale
    Chart1.UpdateSizeNow
    UserControl.Refresh
End Sub
