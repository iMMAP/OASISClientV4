Attribute VB_Name = "mHover"
Option Explicit

'Private declarations
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetMenu Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
'
''Private types
'Private Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type
'
'Private Type POINTAPI
'    X As Long
'    Y As Long
'End Type

'Get the workspace of the form
Private Function GetFormWorkspace(hWnd As Long) As RECT
Dim WndRect As RECT, WrkSpc As RECT
Dim X As Long, Y As Long, CaptionBar As Long, MenuBar As Long

'Get metrics
X = (GetSystemMetrics(7) - 1) * 2
Y = (GetSystemMetrics(8) - 1) * 2
CaptionBar = GetSystemMetrics(4) - 1
If GetMenu(hWnd) = 0 Then MenuBar = 0 Else MenuBar = GetSystemMetrics(15) - 1

'Get sizes
GetWindowRect hWnd, WndRect
With WrkSpc
    .Left = WndRect.Left + X
    .Top = WndRect.Top + Y + CaptionBar + MenuBar
    .Right = WndRect.Right + X
    .Bottom = WndRect.Bottom + Y
End With

'Return
GetFormWorkspace = WrkSpc
End Function

'Get the space of an object
Private Function GetObjectSpace(Obj As Object) As RECT
Dim WrkSpc As RECT, ObjRect As RECT

'Get the space
WrkSpc = GetFormWorkspace(Obj.Parent.hWnd)
With ObjRect
    .Left = WrkSpc.Left + (Obj.Left / Screen.TwipsPerPixelX)
    .Top = WrkSpc.Top + (Obj.Top / Screen.TwipsPerPixelY)
    .Right = .Left + (Obj.Width / Screen.TwipsPerPixelX)
    .Bottom = .Top + (Obj.Height / Screen.TwipsPerPixelY)
End With

'Return
GetObjectSpace = ObjRect
End Function

'Check if the cursor is hovering on an object
Public Function Hover(Obj As Object) As Boolean
Dim ObjRect As RECT
Dim Pos As POINTAPI

ObjRect = GetObjectSpace(Obj)
GetCursorPos Pos

If Pos.X > ObjRect.Left And Pos.X < ObjRect.Right Then
    If Pos.Y > ObjRect.Top And Pos.Y < ObjRect.Bottom Then
        Hover = True
            Else
        Hover = False
    End If
End If
End Function

'Check if the cursor is hovering on an object
Public Function GetObjPos(Obj As Object) As RECT
Dim ObjRect As RECT


GetObjPos = GetObjectSpace(Obj)


End Function

