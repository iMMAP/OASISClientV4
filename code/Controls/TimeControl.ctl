VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl TimeControl 
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   ScaleHeight     =   300
   ScaleWidth      =   1275
   ToolboxBitmap   =   "TimeControl.ctx":0000
   Begin MSComCtl2.UpDown UpDown1 
      CausesValidation=   0   'False
      Height          =   255
      Left            =   1020
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   15
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   450
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "Text1"
      BuddyDispid     =   196610
      OrigLeft        =   1905
      OrigTop         =   1410
      OrigRight       =   2145
      OrigBottom      =   1650
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "00:00:00 AM"
      Top             =   0
      Width           =   1005
   End
End
Attribute VB_Name = "TimeControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Public Event Change()     ' Change event. Obvious, isnt' it? :p

' private variables

Private mH As Integer       ' current hour
Private mMin As Integer     ' current minute
Private mS As Integer       ' current second
Private mbH As Byte         ' which one selected; 0 = hour, 1 = minute, 2 = second, 3 = AM/PM
Private down As Boolean     ' just something to avoid running out of call stack space and stuff
Private mHour24 As Boolean  ' hour format (12 or 24 -hour)
Private mFormat As String   ' format
Private mWidth As Integer   ' width

'*********************************************
'* Set format and lenght of the time control *
'*********************************************

Private Sub SetFormat()
    mFormat = "hh:mm:ss"    ' set default format
    If Not mHour24 Then mFormat = mFormat & " AMPM"     ' add "AMPM" if hour24 = false
    If Len(mFormat) = 8 Then    ' width control
        mWidth = 990
    ElseIf Len(mFormat) = 13 Then
        mWidth = 1275
    End If
    Text1.Text = Format(mH & ":" & mMin & ":" & mS, mFormat) ' set timebox time
    UserControl.Width = mWidth  ' set width
End Sub

'*********************************************
'* Should be obvious                         *
'*********************************************

Public Property Let Hour24(val As Boolean)
Attribute Hour24.VB_Description = "Display AM/PM if false."
Attribute Hour24.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    mHour24 = val
    SetFormat
End Property

Public Property Get Hour24() As Boolean
    Hour24 = mHour24
End Property

'***********************************************
'* Sets updown control max and value and stuff *
'***********************************************

Private Property Let bH(valu As Byte)
    If valu < 4 And valu >= 0 Then
        mbH = valu                  ' save value
        If mbH = 0 Then             ' Hour selected
            Dim i As Integer
            i = mH                  ' save the current hour, mH might change without the code knowing, don't know why...
            UpDown1.Max = 23        ' set max value of updown control
            UpDown1.Value = i       ' set current hour as the value of updown control
        ElseIf mbH = 3 Then
            down = True             ' prevent ... something, already forgot what it did without this ;p
            If H > 11 Then UpDown1.Value = 1    ' 1 = PM, 0 = AM
            If H < 12 Then UpDown1.Value = 0
            UpDown1.Max = 1
            down = False
        Else
            UpDown1.Max = 59        ' set minute and second values
            If mbH = 1 Then UpDown1.Value = mMin
            If mbH = 2 Then UpDown1.Value = mS
        End If
    End If
End Property

Private Property Get bH() As Byte
    bH = mbH
End Property

'***********************************************
'* Sets hour, minute and second the same time  *
'***********************************************

Public Property Let Value(NewValue As String)
Attribute Value.VB_Description = "Returns/Sets the value of an object."
Attribute Value.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    If IsDate(NewValue) Then
        H = Hour(NewValue)
        Min = Minute(NewValue)
        S = Second(NewValue)
    End If
End Property

'*****************************************************
'* get hour, minute and second in proper time format *
'*****************************************************

Public Property Get Value() As String
    Value = Format(mH & ":" & mMin, mFormat)
End Property

'***********************************************
'* Second value control                        *
'***********************************************

Public Property Let S(valu As Integer)
Attribute S.VB_Description = "Second value."
Attribute S.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    If valu > 59 Then valu = 0 ' check if over 59 or under 0 and make corrections
    If valu < 0 Then valu = 59
    mS = valu                   ' set value
    Text1.Text = Format(mH & ":" & mMin & ":" & mS, mFormat)    ' show it
    Text1.SelStart = 6  ' keep selection
    Text1.SelLength = 2
    If bH = 2 Then
        down = True             ' prevents running out of call stack space
        UpDown1.Value = valu    ' keep updown control value up-to-date
        down = False
    End If
End Property

Public Property Get S() As Integer
    S = mS
End Property

'**************************************************
'* Minute value control (check Second value       *
'* control comments, this is pretty much the same *
'**************************************************

Public Property Let Min(valu As Integer)
Attribute Min.VB_Description = "Minute value."
Attribute Min.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    If valu > 59 Then valu = 0
    If valu < 0 Then valu = 59
    mMin = valu
    Text1.Text = Format(mH & ":" & mMin & ":" & mS, mFormat)
    Text1.SelStart = 3
    Text1.SelLength = 2
    If bH = 1 Then
        down = True
        UpDown1.Value = valu
        down = False
    End If
End Property

Public Property Get Min() As Integer
    Min = mMin
End Property

'***********************************************
'* Hour value control                          *
'***********************************************

Public Property Let H(valu As Integer)
Attribute H.VB_Description = "Hour value."
Attribute H.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
    If valu > 23 Then valu = 0  ' make sure that value is not invalid
    If valu < 0 Then valu = 23
    mH = valu                   ' set value
    Text1.Text = Format(mH & ":" & mMin & ":" & mS, mFormat) ' show it
    If bH <> 3 Then         ' H value changes when changing from AM to PM and vice versa, so...
        Text1.SelStart = 0  ' if AM/PM wasn't changed, keep current selection
    Else
        Text1.SelStart = 9  ' AM/PM was changed
    End If
    Text1.SelLength = 2
    If bH = 0 Then
        down = True          ' prevent running out of call stack space
        UpDown1.Value = valu ' keep updown control up-to-date
        down = False
    End If
End Property

Public Property Get H() As Integer
    H = mH
End Property

Private Sub Text1_Change()
    RaiseEvent Change   ' give change event control to the user
End Sub

'***********************************************
'* Sets correct selection                      *
'***********************************************

Private Sub Text1_DblClick()
    If bH = 0 Then
        Text1.SelStart = 0
    ElseIf bH = 1 Then
        Text1.SelStart = 3
    ElseIf bH = 2 Then
        Text1.SelStart = 6
    ElseIf bH = 3 Then
        Text1.SelStart = 9
    End If
    Text1.SelLength = 2
End Sub

'***********************************************
'* Sets selection to hour                      *
'***********************************************

Private Sub Text1_GotFocus()
    If Text1.SelStart < 3 Then
        Text1.SelStart = 0
        Text1.SelLength = 2
        bH = 0
    End If
End Sub

'*************************************************
'* Text1_KeyDown event, up/down keys changes     *
'* currently selected value and left/right moves *
'* to next/previous value                        *
'*************************************************

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If bH = 0 Then
        If KeyCode = vbKeyRight Then    ' select next value
            bH = 1
            Text1.SelStart = 3
            Text1.SelLength = 2
        ElseIf KeyCode = vbKeyUp Then
            H = H + 1                   ' UP-key: increment hour value
        ElseIf KeyCode = vbKeyDown Then
            H = H - 1                   ' DOWN-key: decrement hour value
        End If
    ElseIf bH = 1 Then
        If KeyCode = vbKeyRight Then    ' select next value
            bH = 2
            Text1.SelStart = 6
            Text1.SelLength = 2
        ElseIf KeyCode = vbKeyLeft Then ' select previous value
            bH = 0
            Text1.SelStart = 0
            Text1.SelLength = 2
        ElseIf KeyCode = vbKeyUp Then
            Min = Min + 1               ' UP-key: increment minute value
        ElseIf KeyCode = vbKeyDown Then
            Min = Min - 1               ' DOWN-key: decrement minute value
        End If
    ElseIf bH = 2 Then
        If KeyCode = vbKeyLeft Then     ' select previous value
            bH = 1
            Text1.SelStart = 3
            Text1.SelLength = 2
        ElseIf KeyCode = vbKeyRight And Not Hour24 Then ' select next value but only if using 12-hour system
            bH = 3
            Text1.SelStart = 9
            Text1.SelLength = 2
        ElseIf KeyCode = vbKeyUp Then
            S = S + 1                   ' UP-key: increment second value
        ElseIf KeyCode = vbKeyDown Then
            S = S - 1                   ' DOWN-key: decrement second value
        End If
    ElseIf bH = 3 And Not Hour24 Then
        If KeyCode = vbKeyLeft Then     ' select previous value
            bH = 2
            Text1.SelStart = 6
            Text1.SelLength = 2
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            If H > 11 Then
                H = H - 12              ' change PM -> AM
            ElseIf H < 12 Then
                H = H + 12              ' change AM -> PM
            End If
        End If
    End If
    KeyCode = 0
End Sub

'***********************************************
'* Manage pressed keys                         *
'***********************************************

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) Then            ' check that the key pressed was numeric
        If bH = 0 Then                          ' hour selected
            If mH = 0 Then
                H = Chr(KeyAscii)               ' if hour = 0 then hour = the value of the pressed key
            Else
                If mH < 3 And Chr(KeyAscii) < 4 Then
                    H = H * 10 + Chr(KeyAscii)  ' the value stays lower than 24 so we can just add the currently pressed key
                ElseIf (mH = 2 And Chr(KeyAscii) > 2) Or (mH > 2) Then
                    H = 23                      ' the value went larger than 23 so we must put it back to 23
                End If
                bH = 1                          ' select minutes
                Text1.SelStart = 3
                Text1.SelLength = 2
            End If
        ElseIf bH = 1 Then                      ' minute selected (about the same comments as above)
            If mMin = 0 Then
                Min = Chr(KeyAscii)
            Else
                If mMin < 6 Then
                    Min = Min * 10 + Chr(KeyAscii) ' difference: the value must be lower than 60
                Else
                    Min = 59                       ' but it was larger
                End If
                bH = 2                          ' select seconds
                Text1.SelStart = 6
                Text1.SelLength = 2
            End If
        ElseIf bH = 2 Then                      ' second selected (see comments above)
            If mS = 0 Then
                mS = Chr(KeyAscii)
            Else
                If mS < 6 Then
                    mS = mS * 10 + Chr(KeyAscii)
                Else
                    mS = 59
                End If
            End If
            If Not Hour24 Then                  ' if using 12-hour system
                bH = 3                          ' select AM/PM
                Text1.SelStart = 9
                Text1.SelLength = 2
            End If
        End If
    ElseIf KeyAscii = vbKeyBack Then            ' backspace control
        If bH = 0 Then                          ' hour selected
            If mH > 9 Then                      ' if hour has two digits
                H = Int(mH / 10)                ' remove one
            Else
                H = 0                           ' it didn't, so it becomes 0
            End If
        ElseIf bH = 1 Then                      ' see above (minute selected)
            If mMin > 9 Then
                Min = Int(mMin / 10)
            ElseIf Min = 0 Then
                bH = 0                          ' the minute value was already 0
                Text1.SelStart = 0              ' select hour
                Text1.SelLength = 2
            Else
                Min = 0
            End If
        ElseIf bH = 2 Then                      ' see above (second selected)
            If mS > 9 Then
                S = Int(mS / 10)
            ElseIf S = 0 Then
                bH = 1
                Text1.SelStart = 3              ' select minute
                Text1.SelLength = 2
            Else
                S = 0
            End If
        End If
    End If
End Sub

'***********************************************
'* control loses focus                         *
'***********************************************

Private Sub Text1_LostFocus()
    If Text1.SelStart > 2 Then
        Text1.SelStart = 0      ' just for fun (and some other important things I've already forgotten) set beginning of selection to 0 (just before first letter)
    End If
End Sub

'***********************************************
'* Mouse management                            *
'***********************************************

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Text1.SelStart < 3 Then      ' hour area was pressed
        Text1.SelStart = 0          ' select it
        bH = 0
    ElseIf Text1.SelStart > 2 And Text1.SelStart < 6 Then   ' minute area
        Text1.SelStart = 3
        bH = 1
    ElseIf Text1.SelStart > 5 And Text1.SelStart < 9 Then   ' seconds area
        Text1.SelStart = 6
        bH = 2
    Else                                                    ' AM/PM area
        Text1.SelStart = 9
        bH = 3
    End If
    Text1.SelLength = 2
End Sub

'***********************************************
'* Prevents from "painting" selection          *
'***********************************************

Private Sub Text1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If bH = 0 Then
        Text1.SelStart = 0
    ElseIf bH = 1 Then
        Text1.SelStart = 3
    ElseIf bH = 2 Then
        Text1.SelStart = 6
    ElseIf bH = 3 Then
        Text1.SelStart = 9
    End If
    Text1.SelLength = 2
End Sub

'************************************************
'* if tab key is pressed then select next value *
'************************************************

Private Sub Text1_Validate(Cancel As Boolean)
    If bH = 0 Then
        Cancel = True
        Text1.SelStart = 3
        Text1.SelLength = 2
        bH = 1
    ElseIf bH = 1 Then
        Cancel = True
        Text1.SelStart = 6
        Text1.SelLength = 2
        bH = 2
    ElseIf bH = 2 And Not Hour24 Then
        Cancel = True
        Text1.SelStart = 9
        Text1.SelLength = 2
        bH = 3
    End If
End Sub

'***********************************************
'* Updown control management                   *
'***********************************************

Private Sub UpDown1_Change()
    If Not down Then
        If bH = 0 Then
            H = UpDown1.Value   ' hour was selected and control pressed, set new hour value
        ElseIf bH = 1 Then
            Min = UpDown1.Value ' see above
        ElseIf bH = 2 Then
            S = UpDown1.Value   ' above
        ElseIf bH = 3 And Not Hour24 Then
            If H > 11 Then
                H = H - 12      ' change PM -> AM
            ElseIf H < 12 Then
                H = H + 12      ' change AM -> PM
            End If
        End If
    End If
End Sub

'***********************************************
'* initialization of values                    *
'***********************************************

Private Sub UserControl_Initialize()
    mbH = 0
    down = False
    mHour24 = True
    mFormat = "hh:mm:ss"
    Text1.Text = Format("0", mFormat)
    mWidth = 990
End Sub

'***********************************************
'* Read saved property values                  *
'***********************************************

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mMin = PropBag.ReadProperty("Min", 0)
    mH = PropBag.ReadProperty("H", 0)
    mHour24 = PropBag.ReadProperty("Hour24", True)
    SetFormat
    mWidth = PropBag.ReadProperty("Width", 990)
    UserControl.Width = mWidth
End Sub

'***********************************************
'* Prevents resizing                           *
'***********************************************

Private Sub UserControl_Resize()
    UserControl.Width = mWidth
    If mWidth <> 0 Then
        Text1.Width = mWidth - UpDown1.Width
        UpDown1.Left = Text1.Width
    End If
    UserControl.Height = 285
End Sub

'***********************************************
'* Save property values                        *
'***********************************************

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "H", mH, 0
    PropBag.WriteProperty "Min", mMin, 0
    PropBag.WriteProperty "Hour24", mHour24, True
    PropBag.WriteProperty "Width", mWidth, 990
End Sub
