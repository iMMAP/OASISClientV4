VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{5664FAD6-05FD-11D4-AABA-00105A6F87AB}#1.0#0"; "dXEditrs.dll"
Object = "{EC799235-EE6A-11D2-AE7E-444553540000}#1.0#0"; "dXCtrls.dll"
Begin VB.UserControl OASISThemeRangePicker 
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5925
   ScaleHeight     =   5025
   ScaleWidth      =   5925
   Begin C1SizerLibCtl.C1Elastic C1Elastic2 
      Height          =   5025
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5925
      _cx             =   10451
      _cy             =   8864
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   6
      ChildSpacing    =   4
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
      GridRows        =   2
      GridCols        =   3
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"ThemeRangePicker.ctx":0000
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   1200
         Left            =   4455
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   90
         Width           =   1380
         _cx             =   2434
         _cy             =   2117
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
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   6
         ChildSpacing    =   4
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
         GridRows        =   2
         GridCols        =   2
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"ThemeRangePicker.ctx":005A
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.CommandButton cmdBlend 
            Caption         =   "Blend Colours"
            Height          =   480
            Left            =   105
            TabIndex        =   14
            Top             =   630
            Width           =   1185
         End
         Begin VB.CommandButton cmdResetSizes 
            Caption         =   "Distribute Evenly"
            Height          =   480
            Left            =   105
            TabIndex        =   13
            Top             =   90
            Width           =   1185
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   1200
         Left            =   885
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   90
         Width           =   3510
         _cx             =   6191
         _cy             =   2117
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
         BackColor       =   16777215
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   6
         ChildSpacing    =   4
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
         GridRows        =   0
         GridCols        =   0
         Frame           =   3
         FrameStyle      =   0
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin VB.Label lblThemeIntervals 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Number of ranges"
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   10
            Top             =   180
            Width           =   1260
         End
         Begin VB.Label lblIntervalSelection 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Click below to select a range"
            Height          =   195
            Left            =   30
            TabIndex        =   9
            Top             =   900
            Width           =   3285
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Start colour for this theme"
            Height          =   195
            Left            =   30
            TabIndex        =   8
            Top             =   540
            Width           =   3285
         End
      End
      Begin VB.Frame fraProgressBars 
         BackColor       =   &H8000000E&
         BorderStyle     =   0  'None
         Height          =   3585
         Left            =   90
         TabIndex        =   5
         Top             =   1350
         Width           =   5745
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1200
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   735
         Begin XpressEditorsLibCtl.dxSpinEdit dxThemeIntervals 
            Height          =   315
            Left            =   90
            OleObjectBlob   =   "ThemeRangePicker.ctx":009C
            TabIndex        =   2
            Top             =   120
            Width           =   570
         End
         Begin XpressEditorsLibCtl.dxColorEdit dxStartColour 
            Height          =   315
            Left            =   90
            OleObjectBlob   =   "ThemeRangePicker.ctx":01B0
            TabIndex        =   3
            Top             =   480
            Width           =   585
         End
         Begin XpressEditorsLibCtl.dxColorEdit dxIntervalColor 
            Height          =   315
            Left            =   90
            OleObjectBlob   =   "ThemeRangePicker.ctx":02A3
            TabIndex        =   4
            Top             =   840
            Width           =   585
         End
      End
   End
   Begin VB.TextBox txtRatio 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   0
      Width           =   2220
   End
   Begin CONTROLSLibCtl.dxProgressBar dxProgressBar 
      Height          =   165
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   375
      _Version        =   65536
      _cx             =   661
      _cy             =   291
      ForeColor       =   0
      BackColor       =   15790320
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MinPos          =   0
      MaxPos          =   100
      Pos             =   100
      Step            =   10
      ShowText        =   0   'False
      Orientation     =   0
      StartColor      =   16777215
      EndColor        =   255
      DrawBorderStyle =   0
      ShowTextStyle   =   0
      DrawBarStyle    =   2
      DrawBarBorderStyle=   0
   End
End
Attribute VB_Name = "OASISThemeRangePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim iSelectedInterval As Integer
Dim lRatios(10) As Long
Dim lColours(10) As Long
Dim iThemeIntervals As Integer

Public Sub Init()
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>

        Dim i As Integer
100     i = 1
    
102     lRatios(0) = 0
    
104     Do Until i = 10
    
106         lRatios(i) = lRatios(i - 1) + 256
108         lColours(i) = lRatios(i - 1) + 56
110         i = i + 1
        Loop
        
112     iThemeIntervals = 5
114     dxThemeIntervals = 5

        '<EhFooter>
        Exit Sub

Init_Err:
        Err.Raise vbObjectError + 100, "OASISClient.OASISThemeRangePicker.Init", "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Public Function GetNumberOfIntervals() As Integer
        '<EhHeader>
        On Error GoTo GetNumberOfIntervals_Err
        '</EhHeader>
100     GetNumberOfIntervals = iThemeIntervals
        '<EhFooter>
        Exit Function

GetNumberOfIntervals_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.GetNumberOfIntervals", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Function

Public Function GetThemeStartColor() As Long
        '<EhHeader>
        On Error GoTo GetThemeStartColor_Err
        '</EhHeader>
100     GetThemeStartColor = lColours(0)
        '<EhFooter>
        Exit Function

GetThemeStartColor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.GetThemeStartColor", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Function

Public Function GetIntervalEndColor(iIntervalNumber As Integer) As Long
        '<EhHeader>
        On Error GoTo GetIntervalEndColor_Err
        '</EhHeader>
100     GetIntervalEndColor = lColours(iIntervalNumber)
        '<EhFooter>
        Exit Function

GetIntervalEndColor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.GetIntervalEndColor", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Function

Public Function GetIntervalRatioValue(iIntervalNumber As Integer) As Long
        '<EhHeader>
        On Error GoTo GetIntervalRatioValue_Err
        '</EhHeader>
100     GetIntervalRatioValue = lRatios(iIntervalNumber)
        '<EhFooter>
        Exit Function

GetIntervalRatioValue_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.GetIntervalRatioValue", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Function

Public Sub SetNumberOfIntervals(iNumberOfIntervals As Integer)
        '<EhHeader>
        On Error GoTo SetNumberOfIntervals_Err
        '</EhHeader>
100     iThemeIntervals = iNumberOfIntervals
102     dxThemeIntervals = iNumberOfIntervals
        '<EhFooter>
        Exit Sub

SetNumberOfIntervals_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.SetNumberOfIntervals", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Public Sub Render()
        '<EhHeader>
        On Error GoTo Render_Err
        '</EhHeader>
100     lblIntervalSelection = "Click below to select a range"
102     dxIntervalColor.Enabled = False
104     C1Elastic2.Refresh
106     Call dxThemeIntervals_Change
        '<EhFooter>
        Exit Sub

Render_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.Render", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Public Function SetThemeStartColor(lStartColour As Long)
        '<EhHeader>
        On Error GoTo SetThemeStartColor_Err
        '</EhHeader>
100     lColours(0) = lStartColour
102     dxStartColour = lStartColour
        '<EhFooter>
        Exit Function

SetThemeStartColor_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.SetThemeStartColor", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Function

Public Function SetInterval(iIntervalNumber As Integer, _
                            lEndColour As Long, _
                            lRatio As Long)
        '<EhHeader>
        On Error GoTo SetInterval_Err
        '</EhHeader>
100     lColours(iIntervalNumber) = lEndColour
102     lRatios(iIntervalNumber) = lRatio
        '<EhFooter>
        Exit Function

SetInterval_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.SetInterval", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Function

Private Sub cmdBlend_Click()
        '<EhHeader>
        On Error GoTo cmdBlend_Click_Err
        '</EhHeader>
    
        Dim i As Integer
        
100     i = 1
        
102     Do Until i = iThemeIntervals + 1
            
104         lColours(i) = BlendColors(lColours(0), lColours(iThemeIntervals), 100 * (i / iThemeIntervals))
106         i = i + 1
        Loop
        
108     Call Render
        '<EhFooter>
        Exit Sub

cmdBlend_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.cmdBlend_Click", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub cmdResetSizes_Click()
        '<EhHeader>
        On Error GoTo cmdResetSizes_Click_Err
        '</EhHeader>
    
        Dim i As Integer
        
100     i = 1
102     lRatios(0) = 0
        
104     Do Until i = iThemeIntervals + 1
            
106         lRatios(i) = lRatios(i - 1) + (lRatios(iThemeIntervals) / iThemeIntervals)
108         i = i + 1
        Loop
        
110     Call Render
    
        '<EhFooter>
        Exit Sub

cmdResetSizes_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.cmdResetSizes_Click", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub dxIntervalColor_Change()
        '<EhHeader>
        On Error GoTo dxIntervalColor_Change_Err
        '</EhHeader>

100     If iSelectedInterval > 0 Then
    
102         lColours(iSelectedInterval) = dxIntervalColor
104         dxProgressBar(iSelectedInterval).EndColor = lColours(iSelectedInterval)
        
106         If dxThemeIntervals > (iSelectedInterval) Then
108             dxProgressBar(iSelectedInterval + 1).StartColor = lColours(iSelectedInterval)
            End If
        
        End If

        '<EhFooter>
        Exit Sub

dxIntervalColor_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.dxIntervalColor_Change", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub dxProgressBar_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo dxProgressBar_Click_Err
        '</EhHeader>

100     iSelectedInterval = Index
102     dxIntervalColor = lColours(iSelectedInterval)
104     lblIntervalSelection.caption = "End colour for range " & iSelectedInterval
106     dxIntervalColor.Enabled = True
    
        '<EhFooter>
        Exit Sub

dxProgressBar_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.dxProgressBar_Click", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub dxStartColour_Change()
        '<EhHeader>
        On Error GoTo dxStartColour_Change_Err
        '</EhHeader>
    
100     If Not dxProgressBar.UBound = 0 Then
102         lColours(0) = dxStartColour
104         dxProgressBar(1).StartColor = lColours(0)
        End If
    
        '<EhFooter>
        Exit Sub

dxStartColour_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.dxStartColour_Change", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub dxThemeIntervals_Change()
        '<EhHeader>
        On Error GoTo dxThemeIntervals_Change_Err
        '</EhHeader>

        Dim i As Integer
        Dim lProgressBarWidths(10) As Long

100     iThemeIntervals = dxThemeIntervals
102     fraProgressBars.Refresh
104     i = dxProgressBar.Count - 1

106     Do Until i = 0
108         Unload dxProgressBar(i)
110         Unload txtRatio(i)
112         i = i - 1
        Loop

114     i = 1

116     Do Until i = 10

118         If lRatios(i) <= lRatios(i - 1) Then
120             lRatios(i) = Abs(lRatios(i)) + Abs(lRatios(i - 1))
            End If

122         i = i + 1
        Loop

124     i = 1
126     lProgressBarWidths(0) = 0

128     Do Until i > iThemeIntervals

130         If lRatios(i) <= lRatios(i - 1) Then
132             lRatios(i) = Abs(lRatios(i) - lRatios(i - 1))
            End If


134         lProgressBarWidths(0) = lProgressBarWidths(0) + (lRatios(i) - lRatios(i - 1))
136         i = i + 1
        Loop

138     i = 1

140     Do Until i > iThemeIntervals

142         lProgressBarWidths(i) = Round(fraProgressBars.Width * ((lRatios(i) - lRatios(i - 1)) / lProgressBarWidths(0)), 0)

            'this is an override to force even distribution or ranges on screen
            lProgressBarWidths(i) = (fraProgressBars.Width / iThemeIntervals)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
144         i = i + 1
        Loop

146     lProgressBarWidths(0) = 0
148     i = 1

150     Do Until i > iThemeIntervals

152         If i = 1 Then
154             displayProgressBar 0, 250, lProgressBarWidths(i), fraProgressBars.Height - 250, lColours(0), lColours(i)
156             displayRatioTxt 0, 0, lProgressBarWidths(i), 250, lRatios(i)
            Else
158             displayProgressBar dxProgressBar(i - 1).left + lProgressBarWidths(i - 1) - 5, 250, lProgressBarWidths(i), fraProgressBars.Height - 250, lColours(i - 1), lColours(i)
                'displayProgressBar lProgressBarWidths(i - 1) - 5, 250, lProgressBarWidths(i), fraProgressBars.Height - 250, lColours(0), lColours(i)
160             displayRatioTxt dxProgressBar(i - 1).left + lProgressBarWidths(i - 1) - 5, 0, lProgressBarWidths(i), 250, lRatios(i)
            End If

162         i = i + 1
        Loop

164     iSelectedInterval = 0
166     lblIntervalSelection.caption = "Click below to select a range"
168     dxIntervalColor.Enabled = False
        '<EhFooter>
        Exit Sub

dxThemeIntervals_Change_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.dxThemeIntervals_Change", _
                  "OASISThemeRangePicker component failure"
                  On Error Resume Next
        '</EhFooter>
End Sub

Private Sub displayProgressBar(x As Integer, _
                               y As Integer, _
                               w As Long, _
                               h As Long, _
                               cColourStart As Long, _
                               cColourMax As Long)
        '<EhHeader>
        On Error GoTo displayProgressBar_Err
        '</EhHeader>

100     Load dxProgressBar(dxProgressBar.Count)
102     Set dxProgressBar(dxProgressBar.UBound).Container = fraProgressBars 'cContainingFrame
104     dxProgressBar(dxProgressBar.UBound).Move x, y, w, h
106     dxProgressBar(dxProgressBar.UBound).Visible = True

108     dxProgressBar(dxProgressBar.UBound).StartColor = cColourStart
110     dxProgressBar(dxProgressBar.UBound).EndColor = cColourMax
112     dxProgressBar(dxProgressBar.UBound).toolTipText = "Interval " & dxProgressBar.UBound
        '<EhFooter>
        Exit Sub

displayProgressBar_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.displayProgressBar", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub displayRatioTxt(x As Integer, _
                            y As Integer, _
                            w As Long, _
                            h As Integer, _
                            iNumber As Long)
        '<EhHeader>
        On Error GoTo displayRatioTxt_Err
        '</EhHeader>

100     Load txtRatio(txtRatio.Count)
102     Set txtRatio(txtRatio.UBound).Container = fraProgressBars
104     txtRatio(txtRatio.UBound).Move x, y, w, h
106     txtRatio(txtRatio.UBound).Visible = True
    
108     txtRatio(txtRatio.UBound).toolTipText = "Interval #" & txtRatio.UBound & " ratio"
110     txtRatio(txtRatio.UBound).Text = iNumber
        '<EhFooter>
        Exit Sub

displayRatioTxt_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.displayRatioTxt", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub txtRatio_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo txtRatio_Click_Err
        '</EhHeader>
100     dxProgressBar_Click Index
        '<EhFooter>
        Exit Sub

txtRatio_Click_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.txtRatio_Click", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub txtRatio_KeyDown(Index As Integer, _
                             KeyCode As Integer, _
                             Shift As Integer)
        '<EhHeader>
        On Error GoTo txtRatio_KeyDown_Err
        '</EhHeader>
100     dxProgressBar_Click Index
        '<EhFooter>
        Exit Sub

txtRatio_KeyDown_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.txtRatio_KeyDown", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub txtRatio_KeyUp(Index As Integer, _
                           KeyCode As Integer, _
                           Shift As Integer)
        '<EhHeader>
        On Error GoTo txtRatio_KeyUp_Err
        '</EhHeader>

        On Error GoTo resetvalue

100     If KeyCode = 13 Then
102         If Not CLng(txtRatio(Index).Text) = lRatios(Index) Then
104             ChangedRatio Index
106             txtRatio(Index).SetFocus

            End If
        End If
    
        Exit Sub
     
resetvalue:
108     txtRatio(Index) = lRatios(Index)

        '<EhFooter>
        Exit Sub

txtRatio_KeyUp_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.txtRatio_KeyUp", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub txtRatio_LostFocus(Index As Integer)
        '<EhHeader>
        On Error GoTo txtRatio_LostFocus_Err
        '</EhHeader>
        On Error GoTo resetvalue

100     If Not txtRatio(Index) = lRatios(Index) Then ChangedRatio Index
        Exit Sub
     
resetvalue:
102     txtRatio(Index) = lRatios(Index)
        '<EhFooter>
        Exit Sub

txtRatio_LostFocus_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.txtRatio_LostFocus", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub

Private Sub ChangedRatio(Index As Integer)
        '<EhHeader>
        On Error GoTo ChangedRatio_Err
        '</EhHeader>
    
        Dim bFlag As Boolean
100     bFlag = False
    
102     If Not (CLng(txtRatio(Index).Text) <= lRatios(Index - 1)) Then
    
104         bFlag = True
        
106         If (Index < dxThemeIntervals) Then
        
108             If (CLng(txtRatio(Index)) >= CLng(txtRatio(Index + 1))) Then
                
110                 bFlag = False
                End If
            End If

        End If
    
112     If Not bFlag And Index = iThemeIntervals Then

114         If MsgBox("This value is smaller or equal to the max value of the previous range.  Proceeding will reset the distribution of all ranges.  Do you want to proceed?", vbYesNo, "Reset ranges") = vbYes Then
                If GetNumberOfIntervals > txtRatio(Index) Then txtRatio(Index) = GetNumberOfIntervals
116             lRatios(Index) = txtRatio(Index)
118             Call cmdResetSizes_Click
            Else
                txtRatio(Index) = lRatios(Index)
            End If

120     ElseIf bFlag Then
122         lRatios(Index) = txtRatio(Index)
124         Call dxThemeIntervals_Change
        Else
    
126         MsgBox "Value must be between " & (lRatios(Index - 1) + 1) & " and " & (lRatios(Index + 1) - 1)
128         txtRatio(Index) = lRatios(Index)
        
        End If
    
        '<EhFooter>
        Exit Sub

ChangedRatio_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISThemeRange.OASISThemeRangePicker.ChangedRatio", _
                  "OASISThemeRangePicker component failure"
        '</EhFooter>
End Sub


Private Sub UserControl_Resize()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    If bCanResize Then Call Render
End Sub

Private Sub UserControl_Terminate()
    '<EhHeader>
    On Error Resume Next
    '</EhHeader>

    Dim i As Integer
    i = 9
    
    Do Until i < dxProgressBar.Count - 1
        Unload dxProgressBar(i)
        Unload txtRatio(i)
        i = i - 1
    Loop

End Sub
