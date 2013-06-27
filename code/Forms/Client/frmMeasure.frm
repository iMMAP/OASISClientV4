VERSION 5.00
Begin VB.Form frmMeasure 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "OASIS Measurement Tool"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4170
   FillColor       =   &H000000FF&
   Icon            =   "frmMeasure.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbMeasurement 
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1500
      Width           =   2295
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Total"
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox listSegments 
      Height          =   1035
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblMEasurement 
      Caption         =   "Measurement unit:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblTotal 
      Caption         =   "Total Line Measured:"
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblSegmentsIn 
      Caption         =   "Segments"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmMeasure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event Reset(bResetTool As Boolean)
Private dDistanceTotal As Double
Private dLastX As Double
Private dLastY As Double
Private sLastSetMeasurement As String
Private bCalculateArea As Boolean



Public Sub SetAreaOfPolygon(sArea As String)
    'sArea = left(sArea, Len(sArea) - 2)
        '<EhHeader>
        On Error GoTo SetAreaOfPolygon_Err
        '</EhHeader>
100     txtTotal.Text = ConvertFromKM(CDbl(sArea) * 10000, cmbMeasurement.Text)
102     txtTotal.Text = txtTotal.Text & " " & Chr(178)
        '<EhFooter>
        Exit Sub

SetAreaOfPolygon_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.SetAreaOfPolygon " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Sub Init(bCalcArea As Boolean)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
    
100     bCalculateArea = bCalcArea
    
102     If bCalculateArea Then
104         lblTotal.caption = "Total Area Measured:"
106         txtTotal.Text = " -- "
        Else
108         lblTotal.caption = "Total Line Length:"
110         txtTotal.Text = 0
        End If
    
112     cmdClear_Click
    
        '<EhFooter>
        Exit Sub

Init_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.Init " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Public Function IsMeasuringArea() As Boolean
        '<EhHeader>
        On Error GoTo IsMeasuringArea_Err
        '</EhHeader>
100     IsMeasuringArea = bCalculateArea
        '<EhFooter>
        Exit Function

IsMeasuringArea_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.IsMeasuringArea " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function polygonArea(x() As Double, _
                             y() As Double) As Double
        '<EhHeader>
        On Error GoTo polygonArea_Err
        '</EhHeader>

        'THIS FUNCTION IS NOT USED BUT IT IS LEFT HERE FOR REFERENCE
        'FOR THE MOMENT WE USE TATUKGIS FOR CALCULATING AREAS
    
        'Reference:
        'http://www.mathopenref.com/coordpolygonarea.html
        'http://alienryderflex.com/polygon_area/
    
        'this looks like how TATUKGIS and QGIS calculates it - this Is Not 100% and does not
        'take curvature of the earth into consideration
    
        'this would be a better implementation:
        'http://mathworld.wolfram.com/SphericalPolygon.html
        'http://forum.worldwindcentral.com/showthread.php?p=69704
    
100     If UBound(x) > 1 Then
            Dim area As Double
            Dim i As Long
            Dim j As Long
102         area = 0

104         j = UBound(x) - 1

106         Do Until i = UBound(x)
108             area = area + (x(j) + x(i)) * (y(j) - y(i))
110             j = i
112             i = i + 1
            Loop

114         polygonArea = area * 0.5
        End If

        '<EhFooter>
        Exit Function

polygonArea_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.polygonArea " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub ClearAll()
        '<EhHeader>
        On Error GoTo ClearAll_Err
        '</EhHeader>
100     dDistanceTotal = 0
102     listSegments.Clear

104     If bCalculateArea Then
106         txtTotal.Text = " -- "
        Else
108         txtTotal.Text = ""
        End If

110     dLastX = 666
112     dLastY = 666
114     RaiseEvent Reset(False)
        '<EhFooter>
        Exit Sub

ClearAll_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.ClearAll " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmbMeasurement_Click()
        '<EhHeader>
        On Error GoTo cmbMeasurement_Click_Err
        '</EhHeader>

        Dim i As Long

100     If Not txtTotal.Text = " -- " Then
102         If bCalculateArea Then txtTotal.Text = left(txtTotal.Text, Len(txtTotal.Text) - 2)
        
104         txtTotal.Text = CStr(ConvertToKM(CDbl("0" & txtTotal.Text), sLastSetMeasurement))
106         txtTotal.Text = ConvertFromKM("0" & txtTotal.Text, cmbMeasurement.Text)
        
108         If bCalculateArea Then txtTotal.Text = txtTotal.Text & " " & Chr(178)
        End If

110     Do Until i = listSegments.ListCount

112         listSegments.List(i) = ConvertToKM(listSegments.List(i), sLastSetMeasurement)
114         listSegments.List(i) = ConvertFromKM(listSegments.List(i), cmbMeasurement.Text)

116         i = i + 1
        Loop

118     sLastSetMeasurement = cmbMeasurement.Text
    
        '<EhFooter>
        Exit Sub

cmbMeasurement_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.cmbMeasurement_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub cmdClear_Click()
        '<EhHeader>
        On Error GoTo cmdClear_Click_Err
        '</EhHeader>
100     RaiseEvent Reset(False)
102     ClearAll
        '<EhFooter>
        Exit Sub

cmdClear_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.cmdClear_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     ClearAll
102     cmbMeasurement.Clear
104     cmbMeasurement.AddItem "Kilometers"
106     cmbMeasurement.AddItem "Meters"
108     cmbMeasurement.AddItem "Miles"
110     cmbMeasurement.AddItem "Yards"
112     cmbMeasurement.AddItem "Feet"
114     sLastSetMeasurement = "Kilometers"
116     cmbMeasurement.Text = "Kilometers"
    
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Function ConvertToKM(dOld As Double, _
                             sOldMeasurement As String) As Double
        '<EhHeader>
        On Error GoTo ConvertToKM_Err
        '</EhHeader>

100     Select Case sOldMeasurement
        
            Case "Kilometers"
102             ConvertToKM = Round(dOld, 5)

104         Case "Meters"
106             ConvertToKM = Round(dOld / 1000, 2)

108         Case "Miles"
110             ConvertToKM = Round(dOld / 0.621371192, 5)

112         Case "Yards"
114             ConvertToKM = Round((dOld / 1000) / 1.0936133, 2)

116         Case "Feet"
118             ConvertToKM = Round((dOld / 1000) / 3.2808399, 2)
        
        End Select
        
        '<EhFooter>
        Exit Function

ConvertToKM_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.ConvertToKM " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Function ConvertFromKM(dOld As Double, _
                               sOldMeasurement As String) As Double
        '<EhHeader>
        On Error GoTo ConvertFromKM_Err
        '</EhHeader>

100     Select Case sOldMeasurement
        
            Case "Kilometers"
102             ConvertFromKM = Round(dOld, 5)

104         Case "Meters"
106             ConvertFromKM = Round(dOld * 1000, 2)

108         Case "Miles"
110             ConvertFromKM = Round(dOld * 0.621371192, 5)

112         Case "Yards"
114             ConvertFromKM = Round((dOld * 1000) * 1.0936133, 2)

116         Case "Feet"
118             ConvertFromKM = Round((dOld * 1000) * 3.2808399, 2)
        
        End Select
        
        '<EhFooter>
        Exit Function

ConvertFromKM_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.ConvertFromKM " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function


Public Sub UpdateDistance(x As Double, _
                          y As Double)
        '<EhHeader>
        On Error GoTo UpdateDistance_Err
        '</EhHeader>
    
        Dim dDist As Double

100     If Not dLastX = 666 Then
    
102         dDist = HaversineDistance(x, dLastX, y, dLastY)
104         dDist = ConvertFromKM(dDist, cmbMeasurement.Text)
        
106         If dDist > 0 Then
108             listSegments.AddItem dDist
110             listSegments.ListIndex = listSegments.ListCount - 1
112             If Not bCalculateArea Then
114                 dDistanceTotal = dDistanceTotal + dDist
116                 txtTotal.Text = Round(dDistanceTotal, 2)
                Else
118                 dDistanceTotal = 0
120                 txtTotal.Text = " -- "
                End If
            
            End If
        
        End If
    
122     dLastX = x
124     dLastY = y
    
        '<EhFooter>
        Exit Sub

UpdateDistance_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.UpdateDistance " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Unload(Cancel As Integer)
        '<EhHeader>
        On Error GoTo Form_Unload_Err
        '</EhHeader>
100     RaiseEvent Reset(True)
        '<EhFooter>
        Exit Sub

Form_Unload_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISClient.frmMeasure.Form_Unload " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
