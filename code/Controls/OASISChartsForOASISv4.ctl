VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{E9DF30CA-4B30-4235-BF0C-7150F6466080}#1.0#0"; "ChartFX.ClientServer.Core.dll"
Begin VB.UserControl OASISChartsForOASISv4 
   ClientHeight    =   4530
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   ScaleHeight     =   4530
   ScaleWidth      =   5580
   ToolboxBitmap   =   "OASISChartsForOASISv4.ctx":0000
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   4530
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5580
      _cx             =   9843
      _cy             =   7990
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   1
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"OASISChartsForOASISv4.ctx":0312
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin Cfx62ClientServerCtl.Chart Chart1 
         Height          =   4530
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5580
         _Data_          =   "OASISChartsForOASISv4.ctx":0346
      End
   End
End
Attribute VB_Name = "OASISChartsForOASISv4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private cxAnnotation As Cfx62ClientServerAnnotation.AnnotationX

Public Sub ResetChart()
    Chart1.Legend.Clear
    Chart1.LegendBox = False
    Chart1.ClearData ClearDataFlag_AllData
    Chart1.SerLegBoxObj.Invalidate
    Chart1.LegendBoxObj.Invalidate
    Chart1.ClearData ClearDataFlag_Strings
    Chart1.ContextMenus = True
    RefreshChart
End Sub

Public Sub SayKukuri()
MsgBox "Kukuri!"
End Sub

Public Sub RefreshChart()
    C1Elastic1.Refresh
End Sub

Public Sub ClearDataAll()
    Chart1.ClearData ClearDataFlag_AllData
End Sub

Public Sub ClearDataStrings()
    Chart1.ClearData ClearDataFlag_Strings
End Sub

Public Sub ClearLegend()
    Chart1.Legend.Clear
End Sub

Public Sub TemplateLoad(sPath As String)
        '<EhHeader>
        On Error GoTo TemplateLoad_Err
        '</EhHeader>

        ' ResetChart
        Set cxAnnotation = New Cfx62ClientServerAnnotation.AnnotationX
100     With Chart1

            .Legend.Clear
     
102         If Len(sPath) > 0 Then
                '.ToolBar = False
                'cxAnnotation.ToolBar = False
104             .ClearData ClearDataFlag_Strings
106             .Import FileFormat_Binary, sPath
            End If

108         Chart1.ContextMenus = True
        
        End With
        
110     If Chart1.Extensions.Count = 0 Then
           ' MsgBox "Adding extension cxAnnotation"
112        ' Set cxAnnotation = New Cfx62ClientServerAnnotation.AnnotationX
114         Chart1.Extensions.Add cxAnnotation
116         cxAnnotation.Enabled = True
118         cxAnnotation.ToolBar = False
        End If

        '<EhFooter>
        Exit Sub

TemplateLoad_Err:
        Err.Raise vbObjectError + 100, "OASISChartingV2.OASISChartingVer2.TemplateLoad", "OASISChartingVer2 component failure"
        '</EhFooter>
End Sub


Public Sub TemplateSave(sPath As String)
        '<EhHeader>
        On Error GoTo TemplateSave_Err
        '</EhHeader>
        
 
100     With Chart1
            '.ToolBar = False
            'cxAnnotation.ToolBar = False
101           .ClearData ClearDataFlag_XValues
102            .Legend.Clear
103      .FileMask = FileMask_All
104         .Export FileFormat_Binary, sPath
        
        End With

        '<EhFooter>
        Exit Sub

TemplateSave_Err:
        MsgBox "TemplateSave_Err (" & Erl & ") " & Err.Description
        
End Sub

Public Sub LoadRS(oRS As ADODB.Recordset)

    Dim i As Long
    Dim sRet As String
    Dim j As Long
    Dim iCountOfVal As String
    
    i = 0
    j = 0
    iCountOfVal = 0
    Chart1.MultipleColors = True
    
    Do Until i = oRS.Fields.Count
    
        If oRS.Fields(i).Type = 5 Or oRS.Fields(i).Type = 3 Then iCountOfVal = iCountOfVal + 1
        i = i + 1
        
    Loop
    
    i = 0

    Do Until i = oRS.Fields.Count
    
        If oRS.Fields(i).Type = 5 Or oRS.Fields(i).Type = 3 Then
        
            Chart1.DataType.Item(i) = DataType_Value
            
            If iCountOfVal > 1 Then
                Chart1.SerLeg(j) = Trim(oRS.Fields(i).Name)
                Chart1.MultipleColors = False
                j = j + 1
            End If
            
        ElseIf oRS.Fields(i).Type = 202 Then
        
            Chart1.DataType.Item(i) = DataType_Label

        Else
       
            Chart1.DataType.Item(i) = DataType_Default
            
        End If
    
        i = i + 1
    Loop
    
    Chart1.DataStyle = DataStyle_ReadXValues
    Chart1.DataSource = oRS
    'LoadRS = True
    
    i = 0

    Do Until i = Chart1.AxisX.Label.Count
    
        Chart1.AxisX.Label(i) = Trim$(Chart1.AxisX.Label(i))
        i = i + 1
    Loop
    
    Chart1.AxisX.AdjustScale
    
End Sub

Public Sub ImageSave(sPath As String)

    With Chart1

        .FileMask = FileMask_All
        .Export FileFormat_Bitmap, sPath
        
    End With

End Sub

Private Sub UserControl_Initialize()

    If cxAnnotation Is Nothing Then
        
        Set cxAnnotation = New Cfx62ClientServerAnnotation.AnnotationX
        Chart1.Extensions.Add cxAnnotation
        cxAnnotation.Enabled = True
        cxAnnotation.ToolBar = False
        
    End If

End Sub

