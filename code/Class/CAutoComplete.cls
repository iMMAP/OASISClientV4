VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mAutoAdd As Boolean
Private WithEvents mComboBox As ComboBox
Attribute mComboBox.VB_VarHelpID = -1

Private Declare Function SendMessage _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hwnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      lParam As Any) As Long

Private Const CB_ERR = (-1)
Private Const CB_FINDSTRING = &H14C
Private Const CB_FINDSTRINGEXACT = &H158
Private Const CB_SHOWDROPDOWN = &H14F

Public Property Let AutoAdd(var As Boolean)
        '<EhHeader>
        On Error GoTo AutoAdd_Err
        '</EhHeader>
100     mAutoAdd = var
        '<EhFooter>
        Exit Property

AutoAdd_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.CAutoComplete.AutoAdd", _
                  "CAutoComplete component failure"
        '</EhFooter>
End Property

Public Property Get AutoAdd() As Boolean
        '<EhHeader>
        On Error GoTo AutoAdd_Err
        '</EhHeader>
100     AutoAdd = mAutoAdd
        '<EhFooter>
        Exit Property

AutoAdd_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.CAutoComplete.AutoAdd", _
                  "CAutoComplete component failure"
        '</EhFooter>
End Property

Private Function FindItem(sFind As String, _
                          Optional IndexStart As Long = -1, _
                          Optional bExact As Boolean = False) As Long
        '<EhHeader>
        On Error GoTo FindItem_Err
        '</EhHeader>
        Dim wMsg As Long
100     wMsg = IIf(bExact, CB_FINDSTRINGEXACT, CB_FINDSTRING)
102     sFind = sFind & Chr(0)
104     FindItem = SendMessage(mComboBox.hwnd, wMsg, IndexStart, ByVal sFind)
        '<EhFooter>
        Exit Function

FindItem_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.CAutoComplete.FindItem", _
                  "CAutoComplete component failure"
        '</EhFooter>
End Function

Public Property Set LinkedComboBox(cbo As ComboBox)
        '<EhHeader>
        On Error GoTo LinkedComboBox_Err
        '</EhHeader>
100     Set mComboBox = cbo
        '<EhFooter>
        Exit Property

LinkedComboBox_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.CAutoComplete.LinkedComboBox", _
                  "CAutoComplete component failure"
        '</EhFooter>
End Property

Public Property Get LinkedComboBox() As ComboBox
        '<EhHeader>
        On Error GoTo LinkedComboBox_Err
        '</EhHeader>
100     Set LinkedComboBox = mComboBox
        '<EhFooter>
        Exit Property

LinkedComboBox_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.CAutoComplete.LinkedComboBox", _
                  "CAutoComplete component failure"
        '</EhFooter>
End Property

Private Sub ShowDropDown(bShow As Boolean)
        '<EhHeader>
        On Error GoTo ShowDropDown_Err
        '</EhHeader>
100     SendMessage mComboBox.hwnd, CB_SHOWDROPDOWN, bShow, 0
        '<EhFooter>
        Exit Sub

ShowDropDown_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.CAutoComplete.ShowDropDown", _
                  "CAutoComplete component failure"
        '</EhFooter>
End Sub

Private Sub mComboBox_KeyPress(KeyAscii As Integer)
        '<EhHeader>
        On Error GoTo mComboBox_KeyPress_Err
        '</EhHeader>
100     If mAutoAdd Then

102         With mComboBox
                Dim iLoc As Long

104             If KeyAscii = 13 Then
106                 iLoc = FindItem(.Text, , True)

108                 If iLoc = -1 Then .AddItem .Text
                End If

            End With

        End If

        '<EhFooter>
        Exit Sub

mComboBox_KeyPress_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.CAutoComplete.mComboBox_KeyPress", _
                  "CAutoComplete component failure"
        '</EhFooter>
End Sub

Private Sub mComboBox_KeyUp(KeyCode As Integer, _
                            Shift As Integer)
        '<EhHeader>
        On Error GoTo mComboBox_KeyUp_Err
        '</EhHeader>
100     With mComboBox
            Dim iLoc As Integer, sText As String

102         If KeyCode <> vbKeyBack And KeyCode > 48 Then
104             sText = .Text
106             iLoc = FindItem(sText, , False)

108             If iLoc <> -1 Then
110                 ShowDropDown True
112                 .Text = .List(iLoc)
114                 .SelStart = Len(sText) - 1
116                 .SelLength = Len(.Text) - .SelStart
                Else
                    Dim iSelStart As Integer, iSelLength As Integer
118                 iSelStart = .SelStart
120                 iSelLength = .SelLength
122                 ShowDropDown False
124                 .Text = sText
126                 .SelStart = iSelStart
128                 .SelLength = iSelLength
                End If
            End If

        End With

        '<EhFooter>
        Exit Sub

mComboBox_KeyUp_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.CAutoComplete.mComboBox_KeyUp", _
                  "CAutoComplete component failure"
        '</EhFooter>
End Sub
