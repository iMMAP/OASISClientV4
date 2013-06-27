VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.UserControl OASISAttachments 
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   ScaleHeight     =   6240
   ScaleWidth      =   5055
   ToolboxBitmap   =   "ctrOASISAttachments.ctx":0000
   Begin C1SizerLibCtl.C1Elastic elMAin 
      Height          =   6240
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5055
      _cx             =   8916
      _cy             =   11007
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
      GridRows        =   1
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"ctrOASISAttachments.ctx":0312
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin vbalTreeViewLib6.vbalTreeView tvwXml 
         Height          =   6060
         Left            =   90
         TabIndex        =   1
         Top             =   90
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   10689
         NoCustomDraw    =   0   'False
         HotTracking     =   0   'False
         LineColor       =   12632256
         LineStyle       =   0
         OLEDropMode     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbalIml6.vbalImageList imgl 
         Left            =   0
         Top             =   0
         _ExtentX        =   953
         _ExtentY        =   953
         ColourDepth     =   16
         Size            =   117096
         Images          =   "ctrOASISAttachments.ctx":0349
         Version         =   131072
         KeyCount        =   102
         Keys            =   $"ctrOASISAttachments.ctx":1CCD1
      End
   End
End
Attribute VB_Name = "OASISAttachments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public sURL As String
Private m_fntItalic As StdFont

Private Sub AddChild(nodChild As cTreeViewNode, _
                     curr As ChilkatXml)
        '<EhHeader>
        On Error GoTo AddChild_Err
        '</EhHeader>
        Dim nodElement As cTreeViewNode

100     Set nodChild = nodChild.Children.Add(, etvwChild, Now() & Rnd(100) & curr.GetChildContent("itemtitle") & curr.GetChildContent("fileid") & curr.GetChildContent("filesize"), "File: " & curr.GetChildContent("itemtitle"), imgl.ItemIndex(97))

102     Set nodElement = nodChild.Children.Add(, etvwChild, Now() & Rnd(100) & curr.GetChildContent("itemtitle") & curr.GetChildContent("fileid") & curr.GetChildContent("filesize"), "Filetype: " & curr.GetChildContent("filetype"), imgl.ItemIndex(39))
      
104     nodChild.Children.Add , etvwChild, Rnd(100) & curr.GetChildContent("itemtitle") & curr.GetChildContent("fileid") & Now(), "Filesize: " & curr.GetChildContent("filesize"), imgl.ItemIndex(34)
106     Set nodElement = nodChild.Children.Add(, etvwChild, "URL**" & curr.GetChildContent("fileurl"), "Load File", imgl.ItemIndex(101))
108     nodElement.Font = m_fntItalic
110     nodElement.ForeColor = vbBlue
      
112     nodChild.ShowPlusMinus = True

        '<EhFooter>
        Exit Sub

AddChild_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISAttachments.AddChild", _
                  "OASISAttachments component failure"
        '</EhFooter>
End Sub

Private Sub pLoad()
        '<EhHeader>
        On Error GoTo pLoad_Err
        '</EhHeader>
    
        Dim curr As New ChilkatXml
        Dim n As Long
        Dim i As Long
        Dim j As Long
        Dim q As Long
        Dim nodTop As cTreeViewNode
        Dim xml As ChilkatXml
        Dim xmlOrginal As ChilkatXml
        Dim nodChild As cTreeViewNode
        Dim nodElement As cTreeViewNode
        Dim iCount As Long
        Dim sKey As String
    
        Dim nodDocument As cTreeViewNode '38 DOCUMENT
        Dim nodImages As cTreeViewNode '77 PHOTO_PORTRAIT
        Dim nodMultimedia As cTreeViewNode ' 19 FILM
        Dim nodMedical As cTreeViewNode ' 52 FIRST_AID
        Dim nodMaps As cTreeViewNode ' 21 COMPASS
        Dim nodData As cTreeViewNode '38 DISKS
        Dim nodOther As cTreeViewNode ' 72 IMPORT2

100     tvwXml.ImageList = imgl.hIml
102     tvwXml.Nodes.Clear
    
104     Set m_fntItalic = New StdFont
106     m_fntItalic.Name = "Tahoma"
108     m_fntItalic.Size = 8.25
110     m_fntItalic.Underline = True
112     m_fntItalic.Bold = True
       
114     Set xmlOrginal = curr.HttpGet(sURL)

116     If (xmlOrginal Is Nothing) Then
118         MsgBox curr.LastErrorText
            Exit Sub
        End If

        '  First, get the Description
120     q = xmlOrginal.NumChildrenHavingTag("description")
    
122     For j = 0 To q - 1
124         Set xml = xmlOrginal
126         Set curr = xmlOrginal.GetNthChildWithTag("description", j)
    
128         Set nodTop = tvwXml.Nodes.Add(, etvwChild, curr.GetChildContent("parentid"), curr.GetChildContent("desc"), imgl.ItemIndex(28))
130         nodTop.Bold = True
132         nodTop.ShowPlusMinus = True
    
            '**************** Prepare the file Chategories *********************
    
        Next

134     n = xml.NumChildrenHavingTag("attachment")

136     For i = 0 To n - 1
138         Set curr = xml.GetNthChildWithTag("attachment", i)
      
140         Select Case curr.GetChildContent("filecat")
        
                Case "Document"
142                 If Not tvwXml.Nodes.Exists("docs" & curr.GetChildContent("parentid")) Then Set nodDocument = tvwXml.Nodes.Item(curr.GetChildContent("parentid")).Children.Add(, etvwChild, "docs" & curr.GetChildContent("parentid"), "Documents", imgl.ItemIndex(39))    'cTreeViewNode '38 DOCUMENT
144                 Set nodDocument = tvwXml.Nodes.Item("docs" & curr.GetChildContent("parentid"))
146                 AddChild nodDocument, curr
148                 nodDocument.ShowPlusMinus = True

150             Case "Images"
152                 If Not tvwXml.Nodes.Exists("img" & curr.GetChildContent("parentid")) Then Set nodImages = tvwXml.Nodes.Item(curr.GetChildContent("parentid")).Children.Add(, etvwChild, "img" & curr.GetChildContent("parentid"), "Images", imgl.ItemIndex(78)) '77 PHOTO_PORTRAIT
154                 Set nodImages = tvwXml.Nodes.Item("img" & curr.GetChildContent("parentid"))
156                 AddChild nodImages, curr
158                 nodImages.ShowPlusMinus = True

160             Case "Multimedia"
162                 If Not tvwXml.Nodes.Exists("multi" & curr.GetChildContent("parentid")) Then Set nodMultimedia = tvwXml.Nodes.Item(curr.GetChildContent("parentid")).Children.Add(, etvwChild, "multi" & curr.GetChildContent("parentid"), "Multimedia", imgl.ItemIndex(20))  ' 19 FILM
164                 Set nodMultimedia = tvwXml.Nodes.Item("multi" & curr.GetChildContent("parentid"))
166                 AddChild nodMultimedia, curr
168                 nodMultimedia.ShowPlusMinus = True

170             Case "Medical"
172                 If Not tvwXml.Nodes.Exists("medical" & curr.GetChildContent("parentid")) Then Set nodMedical = tvwXml.Nodes.Item(curr.GetChildContent("parentid")).Children.Add(, etvwChild, "medical" & curr.GetChildContent("parentid"), "Medical", imgl.ItemIndex(53)) ' 52 FIRST_AID
174                 Set nodMedical = tvwXml.Nodes.Item("medical" & curr.GetChildContent("parentid"))
176                 AddChild nodMedical, curr
178                 nodMedical.ShowPlusMinus = True

180             Case "Maps"
182                 If Not tvwXml.Nodes.Exists("maps" & curr.GetChildContent("parentid")) Then Set nodMaps = tvwXml.Nodes.Item(curr.GetChildContent("parentid")).Children.Add(, etvwChild, "maps" & curr.GetChildContent("parentid"), "Maps", imgl.ItemIndex(22))         ' 21 COMPASS
184                 Set nodMaps = tvwXml.Nodes.Item("maps" & curr.GetChildContent("parentid"))
186                 AddChild nodMaps, curr
188                 nodMaps.ShowPlusMinus = True

190             Case "Data"
192                 If Not tvwXml.Nodes.Exists("data" & curr.GetChildContent("parentid")) Then Set nodData = tvwXml.Nodes.Item(curr.GetChildContent("parentid")).Children.Add(, etvwChild, "data" & curr.GetChildContent("parentid"), "Data", imgl.ItemIndex(38))         '38 DISKS
194                 Set nodData = tvwXml.Nodes.Item("data" & curr.GetChildContent("parentid"))
196                 AddChild nodData, curr
198                 nodData.ShowPlusMinus = True

200             Case Else
202                 If Not tvwXml.Nodes.Exists("other" & curr.GetChildContent("parentid")) Then Set nodOther = tvwXml.Nodes.Item(curr.GetChildContent("parentid")).Children.Add(, etvwChild, "other" & curr.GetChildContent("parentid"), "Other", imgl.ItemIndex(73)) ' 72 IMPORT2
204                 Set nodOther = tvwXml.Nodes.Item("other" & curr.GetChildContent("parentid"))
206                 AddChild nodOther, curr
208                 nodOther.ShowPlusMinus = True
            End Select
        


        Next
      
210     q = xmlOrginal.NumChildrenHavingTag("description")
    
212     For j = 0 To q - 1
      
214         Set curr = xmlOrginal.GetNthChildWithTag("description", j)
    
216         Set nodTop = tvwXml.Nodes.Item(curr.GetChildContent("parentid"))
218         nodTop.Selected = True
220         nodTop.Expanded = True
        Next
    
        '<EhFooter>
        Exit Sub

pLoad_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISAttachments.pLoad", _
                  "OASISAttachments component failure"
        '</EhFooter>
End Sub

Private Sub tvwXml_SelectedNodeChanged()
        '<EhHeader>
        On Error GoTo tvwXml_SelectedNodeChanged_Err
        '</EhHeader>
100    If InStr(tvwXml.SelectedItem.Key, "URL**") Then
102      ShellExecute UserControl.Parent.hwnd, vbNullString, Mid$(tvwXml.SelectedItem.Key, 6), vbNullString, "C:\", SW_SHOWNORMAL
       End If
   
        '<EhFooter>
        Exit Sub

tvwXml_SelectedNodeChanged_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISAttachments.tvwXml_SelectedNodeChanged", _
                  "OASISAttachments component failure"
        '</EhFooter>
End Sub

Public Sub Init(sSourceURL As String, ocon As ADODB.Connection)
        '<EhHeader>
        On Error GoTo Init_Err
        '</EhHeader>
        sURL = sSourceURL
100     pLoad
        '<EhFooter>
        Exit Sub

Init_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.OASISAttachments.Init", _
                  "OASISAttachments component failure"
        '</EhFooter>
End Sub

