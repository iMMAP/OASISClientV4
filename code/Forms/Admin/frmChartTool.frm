VERSION 5.00
Begin VB.Form frmChartTool 
   Caption         =   "OASIS Chart Preview"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7755
   Icon            =   "frmChartTool.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6165
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmChartTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetParent _
                Lib "user32" (ByVal hWndChild As Long, _
                              ByVal hWndNewParent As Long) As Long
                              
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_BOTTOM = 1
Private Const HWND_TOP = 0
Private ChartHandle As Long
'Private oChartObj As OASISChartObj

Public Sub SetChart(objChart As OASISChartObj)
   Dim ochart As New OASISCharting.ChartProvider
        
    ochart.AllowAutoResize = True
    'ochart.AllowAutoCloseDBLClick = True
    ChartHandle = ochart.InitChart(objChart, False, , Me.hwnd)
    
    SetWindowPos ChartHandle, HWND_TOP, 0, _
    0, ScaleX(Me.Width - 235, vbTwips, vbPixels), ScaleY(Me.Height - 545, vbTwips, vbPixels), 0
 
End Sub

Public Function OpenChartTemplate(sPath As String) As OASISChartObj
   Dim ochart As New OASISCharting.ChartProvider
   Dim oOASISChartObj As OASISChartObj
   
    ochart.AllowAutoResize = True
    oOASISChartObj = ochart.GetOASISObjSettings(sPath)
    ChartHandle = oOASISChartObj.udtGeneric.lParentHwnd
    SetParent ChartHandle, Me.hwnd
    SetWindowPos ChartHandle, HWND_TOP, 0, _
    0, ScaleX(Me.Width - 235, vbTwips, vbPixels), ScaleY(Me.Height - 545, vbTwips, vbPixels), 0

    OpenChartTemplate = oOASISChartObj

End Function

Private Sub Form_Resize()

    If Not ChartHandle = 0 Then
        SetWindowPos ChartHandle, HWND_TOP, 0, 0, ScaleX(Me.Width - 235, vbTwips, vbPixels), ScaleY(Me.Height - 545, vbTwips, vbPixels), 0
    End If

End Sub


