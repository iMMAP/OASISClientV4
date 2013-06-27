VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Begin VB.UserControl ctlMapFilePointer 
   ClientHeight    =   1815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   ScaleHeight     =   1815
   ScaleWidth      =   3495
   ToolboxBitmap   =   "ctlMapProjectPointer.ctx":0000
   Begin C1SizerLibCtl.C1Elastic C1ElasticFrame 
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   3495
      _cx             =   6165
      _cy             =   3201
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   2
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
      FrameWidth      =   0
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"ctlMapProjectPointer.ctx":0312
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Container 
         Height          =   1755
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   3435
         _cx             =   6059
         _cy             =   3096
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
         Align           =   0
         AutoSizeChildren=   8
         BorderWidth     =   1
         ChildSpacing    =   1
         Splitter        =   0   'False
         FloodDirection  =   0
         FloodPercent    =   0
         CaptionPos      =   0
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
         GridRows        =   6
         GridCols        =   4
         Frame           =   4
         FrameStyle      =   1
         FrameWidth      =   0
         FrameColor      =   8421631
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   $"ctlMapProjectPointer.ctx":034A
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin C1SizerLibCtl.C1Elastic C1EMyMap 
            Height          =   225
            Left            =   2565
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   375
            Width           =   855
            _cx             =   1508
            _cy             =   397
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
            BackColor       =   12648384
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "My Map"
            Align           =   0
            AutoSizeChildren=   0
            BorderWidth     =   0
            ChildSpacing    =   0
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   3
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
            FrameWidth      =   0
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   ""
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
         End
         Begin VB.CommandButton cmdInfo 
            Appearance      =   0  'Flat
            Caption         =   "Info"
            Height          =   270
            Left            =   2565
            TabIndex        =   4
            Top             =   615
            Width           =   855
         End
         Begin VB.CommandButton cmdDelete 
            Appearance      =   0  'Flat
            Caption         =   "Delete"
            Height          =   270
            Left            =   2565
            TabIndex        =   8
            Top             =   1185
            Width           =   855
         End
         Begin VB.PictureBox pic1 
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000009&
            Height          =   1365
            Left            =   15
            ScaleHeight     =   1305
            ScaleWidth      =   2475
            TabIndex        =   7
            Top             =   375
            Width           =   2535
         End
         Begin VB.CommandButton cmdPreview 
            Caption         =   "Preview"
            Height          =   510
            Left            =   900
            TabIndex        =   5
            Top             =   375
            Width           =   1650
         End
         Begin VB.CommandButton cmdOpen 
            Appearance      =   0  'Flat
            Caption         =   "Open"
            Height          =   270
            Left            =   2565
            TabIndex        =   3
            Top             =   1470
            Width           =   855
         End
         Begin C1SizerLibCtl.C1Elastic C1EMapTitle 
            Height          =   345
            Left            =   15
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   15
            Width           =   3405
            _cx             =   6006
            _cy             =   609
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
            BackColor       =   12648447
            ForeColor       =   0
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "Map Title"
            Align           =   0
            AutoSizeChildren=   0
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
         End
         Begin VB.CommandButton cmdEdit 
            Appearance      =   0  'Flat
            Caption         =   "Edit"
            Height          =   270
            Left            =   2565
            TabIndex        =   9
            Top             =   900
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "ctlMapFilePointer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event MapLoad(ByVal sMapPath As String, bIsFileBased As Boolean)
Public Event MapInfo(sMapInfo As String, sGUID As String)
Public Event MapEdit(sGUID As String)
Public Event MapDelete(sGUID As String)
Public Event MapPreview(oPicture As StdPicture)

Private sFilePath As String
Private sMapInfo As String
Private mImagePreview As StdPicture
Private mExtent As XGIS_Extent
Private msGUID As String
Private m_lMapIndex As Long

Public Enum MapTypeValue
    FileBased = 0
    UserGroup = 1
    UserCustom = 2
End Enum

Public MapType As MapTypeValue
Public IsFileBased As Boolean

Public Function MapGUID() As String
    MapGUID = msGUID
End Function

Public Sub SetMapType(MapTypePassed As MapTypeValue)

    MapType = MapTypePassed
    
    Select Case MapTypePassed
    
        Case FileBased
            C1EMyMap.caption = "Project file"
            cmdDelete.Enabled = False
            C1EMyMap.BackColor = vbWhite
            cmdEdit.Enabled = False
        
        Case UserCustom
            C1EMyMap.caption = "My map"
            cmdDelete.Enabled = True
            cmdEdit.Enabled = True
            C1EMyMap.BackColor = vbGreen

        Case UserGroup
          
            C1EMyMap.caption = "UG Map"
            cmdDelete.Enabled = False
            cmdEdit.Enabled = False
            C1EMyMap.BackColor = vbYellow
    End Select
    
End Sub

Public Function GetMapType() As MapTypeValue
    GetMapType = MapType
End Function

Public Sub SetImage(oPicture As PictureBox)
    pic1.Picture = oPicture.Image
End Sub

Public Function GetImage() As PictureBox
    GetImage = pic1
End Function

Public Sub oGUID(sGUID As String)
    msGUID = sGUID
End Sub

Public Sub SetExtent(oExtent)
    Set mExtent = oExtent
End Sub

Public Function GetExtent() As XGIS_Extent
    Set GetExtent = mExtent
End Function

Public Sub SetTitle(sTitle As String)
    C1EMapTitle.caption = sTitle
End Sub

Public Sub ClearTitle()
    C1EMapTitle.caption = ""
End Sub

Public Function GetTitle() As String
    GetTitle = Replace(C1EMapTitle.caption, " (active)", "")
End Function

Public Sub SetBackColour(oColour As ColorConstants)
    C1Container.BackColor = oColour
    C1EMapTitle.BackColor = oColour
End Sub

Public Sub SetActive()
    C1Container.ForeColor = vbBlack
    C1EMapTitle.ForeColor = vbBlack ' &HC0&       'vbRed
    C1Container.BackColor = vbWhite
    C1EMapTitle.BackColor = vbWhite
    C1ElasticFrame.BackColor = vbRed
    C1EMapTitle.Font.Bold = False
    SetMapType GetMapType
    If MapType = UserCustom Then cmdEdit.Enabled = True
    C1Container.BorderWidth = 1
    C1Container.FrameWidth = 1
    
End Sub

Public Sub SetInactive()
    C1Container.ForeColor = vbBlack
    C1EMapTitle.ForeColor = vbBlack
    C1EMapTitle.Font.Bold = False
    C1Container.BackColor = UserControl.BackColor
    C1EMapTitle.BackColor = UserControl.BackColor
    C1ElasticFrame.BackColor = UserControl.BackColor
    cmdEdit.Enabled = False
    C1Container.BorderWidth = 0
    C1Container.FrameWidth = 0
    'cmdDelete.Enabled = False
    'If MapType = UserCustom Then cmdDelete = True
End Sub

Public Sub SetCaptionColour(oColour As ColorConstants)
    C1Container.ForeColor = oColour
    C1EMapTitle.ForeColor = oColour
End Sub

Public Sub SetFilePath(sPath As String)
    sFilePath = sPath
End Sub

Public Function GetFilePath() As String
    GetFilePath = sFilePath
End Function

Private Sub cmdDelete_Click()
    RaiseEvent MapDelete(msGUID)
End Sub

Private Sub cmdEdit_Click()
    RaiseEvent MapEdit(msGUID)
End Sub

Private Sub cmdInfo_Click()
    RaiseEvent MapInfo(sMapInfo, msGUID)
End Sub

Private Sub cmdOpen_Click()
    RaiseEvent MapLoad(sFilePath, IsFileBased)
End Sub

Public Sub LoadMap()
    Call cmdOpen_Click
End Sub

Public Sub SetInfo(sInfo As String)
    pic1.toolTipText = sInfo
    sMapInfo = sInfo
End Sub

Private Sub cmdPreview_Click()
    RaiseEvent MapPreview(mImagePreview)
End Sub

Private Sub pic1_DblClick()
   ' RaiseEvent MapLoad(sFilePath)
End Sub
