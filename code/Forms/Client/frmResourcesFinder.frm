VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{6A24B331-7634-11D3-A5B0-0050044A7E1A}#1.7#0"; "DXDBGrid.dll"
Begin VB.Form frmResourcesFinder 
   Caption         =   "Resource Finder"
   ClientHeight    =   3915
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   9345
   Icon            =   "frmResourcesFinder.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   9345
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   3915
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9345
      _cx             =   16484
      _cy             =   6906
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
      BackColor       =   -2147483633
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
      GridRows        =   5
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmResourcesFinder.frx":6852
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin DXDBGRIDLibCtl.dxDBGrid dxDBGrid1 
         Height          =   3735
         Left            =   90
         OleObjectBlob   =   "frmResourcesFinder.frx":68E9
         TabIndex        =   1
         Top             =   90
         Width           =   9165
      End
   End
   Begin VB.Menu Export 
      Caption         =   "Export"
   End
End
Attribute VB_Name = "frmResourcesFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private sComms As String
Private sRelief As String
Private sShelter As String
Private sWash As String
Private sTransport As String
Private mRS As adodb.Recordset
Public Event GetResources()
Public Event GetGISImage()
Public Event CloseResources()

Public Sub SetResources(sReourceNames As String, sItems As String)

    Dim sResourceNameArray() As String
    Dim sResourceItemsArray() As String
    Dim sLineItemArray() As String

    Dim sGrouping As String
    Dim i As Long
    Dim j As Long
    Dim bFirstRun As Boolean
    Dim bSecondRun As Boolean
    Dim lKeyField As Long
    
    If Not mRS.EOF Then lKeyField = mRS.Fields(0).Value

    If mRS.RecordCount > 0 Then
        mRS.MoveFirst
        mRS.CancelBatch
    End If
    
    sResourceNameArray = Split(sReourceNames, "&&&")
    sResourceItemsArray = Split(sItems, "&&&")
    i = 0
    
    Do Until i > UBound(sResourceNameArray) Or UBound(sResourceNameArray) = -1
  
        sLineItemArray = Split(sResourceItemsArray(i), "|||")
        sGrouping = sResourceNameArray(i)

        j = 0
        Do Until j = UBound(sLineItemArray) Or UBound(sLineItemArray) = -1
            mRS.AddNew
            mRS.Fields(0).Value = mRS.RecordCount + 1
            mRS.Fields(1).Value = sGrouping
            mRS.Fields(2).Value = sLineItemArray(j)
            j = j + 1

        Loop

        i = i + 1
    Loop
    
    If Not mRS.EOF Or Not mRS.BOF Then
        mRS.MoveFirst
        mRS.Find "KeyField = " & lKeyField
    End If
    
    dxDBGrid1.Columns(2).Sorted = csUp
    dxDBGrid1.M.FullCollapse
    On Error Resume Next
    dxDBGrid1.ex.FocusedNode.Expand True
    
End Sub

Private Sub Export_Click()

    Dim oRS As New adodb.Recordset
    oRS.Fields.Append "Layer", adVarChar, 50
    oRS.Fields.Append "Items", adLongVarChar, 8000
    oRS.Open
    
    With dxDBGrid1.Dataset
    
        If Not .EOF Or Not .BOF Then
        
            .DisableControls
            .First
            mRS.Sort = "[Items]"
        
            Do Until .EOF
        
                oRS.AddNew
                oRS.Fields(0).Value = .FieldValues(oRS.Fields.Item(0).Name)
                oRS.Fields(1).Value = .FieldValues(oRS.Fields.Item(1).Name)
                .Next
            Loop
            
            .EnableControls
        
        End If
    
    End With
    
    Clipboard.Clear
    RaiseEvent GetGISImage
    oRS.Sort = "Items"
    frmReportsFromRS.SetReportRS "Nearby Resources", oRS, "Layer" & ":::KeyField", Clipboard.GetData(vbCFEMetafile), "", "", mRS.Sort
    frmReportsFromRS.ShowReport
    frmReportsFromRS.Show vbModal, Me
    oRS.Close
    Set oRS = Nothing

End Sub

Private Sub Form_Activate()
RaiseEvent GetResources
End Sub

Private Sub Form_Load()

    Set mRS = New adodb.Recordset
    mRS.Fields.Append "KeyField", adBigInt
    mRS.Fields.Append "Layer", adVarChar, 50
    mRS.Fields.Append "Items", adLongVarChar, 8000
    mRS.Open

    Set dxDBGrid1.DataSource = mRS
    dxDBGrid1.KeyField = "KeyField"
    dxDBGrid1.Option = egoCanNavigation
    dxDBGrid1.OptionEnabled = False
    dxDBGrid1.Columns.RetrieveFields
    dxDBGrid1.Columns(1).GroupIndex = 0
    dxDBGrid1.Columns(0).Visible = False
    
    RaiseEvent GetResources
    Me.Width = Screen.Width
    Me.Top = Screen.Height - Me.Height - (Screen.Height * 0.05)
     
End Sub

Private Sub Form_Unload(Cancel As Integer)
RaiseEvent CloseResources
End Sub
