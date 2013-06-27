VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMouseCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ***********************************************************************************
' Purpose:  To change the mouse cursor to any value desired by user and then
'           restore the original cursor when the procedure that instantiated
'           this class ends.  This makes the changing of the mouse cursor easy
'           and almost automatic.
'           Note that the class stores the mouse cursor only the first time the
'           SetCursor is invoked; if you omit any argument when calling the method,
'           the cursor is changed to an hourglass. Here is an example of how
'           this class can be used:
'
'Sub MyRoutine()
'    Dim mouse As New cMouseCursor
'    ' save current cursor, enforce a new one
'    mouse.SetCursor vbArrowQuestion
'    ' ...
'    ' change to an hourglass
'    mouse.SetCursor
'    ' ...
'    ' the original mouse cursor is automatically
'    ' restored when this procedure is exited
'End Sub
'
' ***********************************************************************************

Private lngOldMousePointer As Long

' enforce a new cursor
Public Sub SetCursor(Optional NewCursor As MousePointerConstants = vbHourglass)

    If IsEmpty(lngOldMousePointer) Then
        ' save the original cursor, but only the first
        ' time this method is invoked in this instance
        lngOldMousePointer = Screen.MousePointer
    End If
    Screen.MousePointer = NewCursor

End Sub

Private Sub Class_Terminate()

    ' restore original mouse cursor, if it changed
    If Not IsEmpty(lngOldMousePointer) Then
        Screen.MousePointer = lngOldMousePointer
    End If

End Sub
