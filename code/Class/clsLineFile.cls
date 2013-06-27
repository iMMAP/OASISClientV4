VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const LF_UNLOADED = 0
Private Const LF_READ = 1
Private Const LF_WRITE = 2

Private ReadOrWrite As Integer
Private aFileNumber As Integer

'Temporary read-write variables for private read-only properties.
Private tempEndOfFile As Boolean
Private tempFileName As String
Private tempNumLines As Long
Private tempPosition As Long

Private Sub Class_Initialize()

ReadOrWrite = LF_UNLOADED

End Sub

Public Sub CloseFile()

Close #aFileNumber

End Sub

Private Function CountLines() As Long

Dim tString As String
Dim i As Long
i = 0
'Loop thru and read all lines, counting them.
Do Until EOF(aFileNumber)
  Line Input #aFileNumber, tString
  i = i + 1
Loop
'Populate the NumLines property with the result.
Let NumLines = i
CountLines = i
'Close then reopen the file to reset the pointer.
Close #aFileNumber
aFileNumber = FreeFile
Open tempFileName For Input As aFileNumber
Let Filename = tempFileName
ReadOrWrite = LF_READ
'Update the Position property back to zero.
tempPosition = 0
Let Position = tempPosition

End Function

Public Sub GoToBeg()

'Set the pointer back to the beginning by
'closing, then reopening the same file
Close #aFileNumber
aFileNumber = FreeFile
Open tempFileName For Input As aFileNumber
Let Filename = tempFileName
ReadOrWrite = LF_READ
'Count the number of lines
Call CountLines
'Reset the Position property back to zero.
tempPosition = 0
Let Position = tempPosition

End Sub

Public Sub GoToPos(ByVal aPosition As Long)

Dim str As String
Dim i As Long

'if the position to go to is ahead of the current
'pointer, then move ahead to that position.
If aPosition > tempPosition Then
  For i = tempPosition To aPosition - 1
    Line Input #aFileNumber, str
    tempPosition = tempPosition + 1
    Let Position = tempPosition
  Next
  'if the position to go to is behind the current
  'pointer, then close the file, reopen it, then
  'scan down to the new position.
 ElseIf aPosition <= tempPosition Then
  Close #aFileNumber
  aFileNumber = FreeFile
  Open tempFileName For Input As aFileNumber
  Let Filename = tempFileName
  ReadOrWrite = LF_READ
  Call CountLines
  For i = 0 To aPosition - 1
    Line Input #aFileNumber, str
    tempPosition = tempPosition + 1
    Let Position = tempPosition
  Next
End If
'if the position to go to is the same as the current
'position, do nothing.

End Sub

Public Sub MakeNew(ByVal aFileName As String)

'Use the argument provided to make a new file for
'writing.  If a file by this name exists, overwrite it.
aFileNumber = FreeFile
Open aFileName For Output As aFileNumber
ReadOrWrite = LF_WRITE
'Initialize the read-only properties.
Let Filename = aFileName
Let EndOfFile = True
Let Position = 0
Let NumLines = 0

End Sub

Public Sub OpenExisting(ByVal aFileName As String)

On Error GoTo FileNotFound

'Open an existing file for reading.
aFileNumber = FreeFile
Open aFileName For Input As aFileNumber

'Jason's error check for file not found
If Err.number = 76 Then ' File not found
  Let Filename = ""
  Err.Clear
  Exit Sub
End If
'**************************************

ReadOrWrite = LF_READ
'Initialize the read-only properties.
Let Filename = aFileName
Call CountLines
tempPosition = 0
Let Position = tempPosition

Exit Sub
FileNotFound:
  Let Filename = ""
  Err.Clear

End Sub

Public Function ReadElt() As String

Dim ele As String

'If there is no active file, or the active file is open for
'writing, generate an error and bail out.
Select Case ReadOrWrite
  Case LF_WRITE
    MsgBox "File is currently open for writing only.", vbCritical
    Exit Function
  Case LF_UNLOADED
    MsgBox "You must use OpenFile method before reading.", vbCritical
    Exit Function
End Select

Line Input #aFileNumber, ele
Let Position = Position + 1

If EOF(aFileNumber) Then
  Let EndOfFile = True
End If

ReadElt = ele

End Function

Public Sub WriteElt(ByVal aLine As String)

Select Case ReadOrWrite
  Case LF_READ
    MsgBox "File is currently open for reading only.", vbCritical
    Exit Sub
  Case LF_UNLOADED
    MsgBox "You must use MakeNew method before writing.", vbCritical
    Exit Sub
End Select

Print #aFileNumber, aLine
Let EndOfFile = True
Let Position = Position + 1
Let NumLines = NumLines + 1

End Sub

Public Property Get EndOfFile() As Boolean

EndOfFile = EOF(aFileNumber)

End Property

Private Property Let EndOfFile(ByVal bNewValue As Boolean)

tempEndOfFile = bNewValue

End Property

Public Property Get Filename() As String

Filename = tempFileName

End Property

Private Property Let Filename(ByVal itsName As String)

tempFileName = itsName

End Property

Public Property Get NumLines() As Long

NumLines = tempNumLines

End Property

Private Property Let NumLines(ByVal aLength As Long)

tempNumLines = aLength

End Property

Public Property Get Position() As Long

Position = tempPosition

End Property

Private Property Let Position(ByVal newPosition As Long)

tempPosition = newPosition

End Property


