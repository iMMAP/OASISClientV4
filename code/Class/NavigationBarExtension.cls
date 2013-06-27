VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NavigationBarExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Const NavigationPopupName As String = "popNavBar"
Private Const NavigationAddRemovePopupName As String = "popNavBarAddRemove"
Private Const NavigationAddRemoveToolPrefix As String = "navAddRemoveTool"
Private Const NavigationShowMoreToolName As String = "navBarToolShowMore"
Private Const NavigationShowFewerToolName As String = "navBarToolShowFewer"
Private Const NavigationPaneOptionsToolName As String = "navBarToolNavPaneOptions"
Private Const NavigationAddRemoveToolName As String = "navBarToolAddRemove"

Private WithEvents m_ab As ActiveBar3
Attribute m_ab.VB_VarHelpID = -1
Private m_common As AB3Common

Private m_navigationBar As Band
Private m_addRemovePopup As Band
Private m_customPopup As Band

Private m_showMore As Tool
Private m_showFewer As Tool
Private m_navBarOptions As Tool
Private m_addRemoveButtons As Tool

Private Sub ThrowException(ByVal Message As String, Optional ByVal errorType As AB3ErrorEnum = LogicError)
    Call m_common.ThrowException(Message, , errorType)
End Sub


Friend Sub Init(ByVal ab3 As ActiveBar3, ByVal common As AB3Common)
    
    ' set the module variable
    Set m_ab = ab3
    
    ' this is used for building tools and so on...
    Set m_common = common
        
End Sub

Friend Sub Extend()
        '<EhHeader>
        On Error GoTo Extend_Err
        '</EhHeader>
    
        ' check to see if the main popup has been created yet
100     If m_customPopup Is Nothing Then
        
            ' have we already extended it from another extender?
102         Set m_customPopup = m_common.FindBand(NavigationPopupName)
    
            ' if not, build it
104         If m_customPopup Is Nothing Then
            
                ' add it to the bands collection
106             Set m_customPopup = m_ab.Bands.Add(NavigationPopupName)
108             m_customPopup.Type = ddBTPopup
                        
                ' add associating tools that go with it
            
                ' SHOW MORE
110             Set m_showMore = m_customPopup.Tools.Add(m_common.GetUniqueToolID(), NavigationShowMoreToolName)
112             m_showMore.caption = "Show &More Buttons"
            
                ' SHOW FEWER
114             Set m_showFewer = m_customPopup.Tools.Add(m_common.GetUniqueToolID(), NavigationShowFewerToolName)
116             m_showFewer.caption = "Show &Fewer Buttons"
            
                ' PANE OPTIONS
118             Set m_navBarOptions = m_customPopup.Tools.Add(m_common.GetUniqueToolID(), NavigationPaneOptionsToolName)
120             m_navBarOptions.caption = "Na&vigation Pane Options..."
            
                ' ADD REMOVE
122             Set m_addRemoveButtons = m_customPopup.Tools.Add(m_common.GetUniqueToolID(), NavigationAddRemoveToolName)
124             m_addRemoveButtons.caption = "&Add or Remove Buttons"
            
                ' find or create a popup containing all of the child band names
                ' NOTE: Names and Checks get set in the 'SetNavigationBar()' method
126             Set m_addRemovePopup = m_common.FindBand(NavigationAddRemovePopupName)
128             If m_addRemovePopup Is Nothing Then
                
                    ' add the new popup
130                 Set m_addRemovePopup = m_ab.Bands.Add(NavigationAddRemovePopupName)
132                 m_addRemovePopup.Type = ddBTPopup
                
                End If
            
                ' associate the new add/remove popup with the add/remove button
134             m_addRemoveButtons.SubBand = m_addRemovePopup.Name
                        
            End If
        
        End If
    
        '<EhFooter>
        Exit Sub

Extend_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.NavigationBarExtension.Extend", _
                  "NavigationBarExtension component failure Error on Line:" & Erl & Err.Description & " Error Number:" & Err.number
        '</EhFooter>
End Sub

Friend Sub CleanUp()
    
    ' clean up the addremove tool popup menu
    If Not m_addRemovePopup Is Nothing Then
        m_addRemovePopup.Tools.RemoveAll
        Set m_addRemovePopup = Nothing
    End If
    
    ' clean up the navigation popup menu
    If Not m_customPopup Is Nothing Then
        m_customPopup.Tools.RemoveAll
        Set m_customPopup = Nothing
    End If
    
    ' clean up the other objects
    Set m_navigationBar = Nothing
    Set m_ab = Nothing
    Set m_common = Nothing
    
End Sub

Private Sub SetNavigationBar(ByVal navBand As Band)
    
    Dim eachChildBand As Band
    Dim childBandTool As Tool
    Dim toolIndex As Integer
    
    ' can only set bands with child bands
    If navBand.Type <> ddBTNormal Then
        ThrowException "Expecting a band with childbands"
    End If
    
    ' if this band already has the childband style set, get out
    If navBand.ChildBandStyle <> ddCBSNavigationBar Then
        ThrowException "Expecting a navigation bar band"
    End If
    
    ' find the add/remove popup so we can rebuild it
    If m_customPopup Is Nothing Or m_addRemovePopup Is Nothing Then
        ThrowException "Must call Extend() method first"
    End If
    
    ' remove all the old tools
    Set m_navigationBar = navBand
    
    ' by default, ab3 sets largebuttoncount = -1
    If m_navigationBar.LargeButtonCount = -1 Then
        m_navigationBar.LargeButtonCount = m_navigationBar.ChildBands.Count
    End If
    
    ' remove all the old tools
    For toolIndex = m_addRemovePopup.Tools.Count - 1 To 0 Step -1
        m_addRemovePopup.Tools.Remove toolIndex
    Next
    
    ' now add each childband as a tool
    For Each eachChildBand In m_navigationBar.ChildBands
        Set childBandTool = m_addRemovePopup.Tools.Add(m_common.GetUniqueToolID(), NavigationAddRemoveToolPrefix + eachChildBand.Name)
        childBandTool.caption = eachChildBand.caption
        childBandTool.Checked = eachChildBand.Visible
        If Not eachChildBand.Picture Is Nothing Then
            childBandTool.SetPicture ddITNormal, eachChildBand.Picture
        End If
    Next
    
End Sub

Private Function GetVisibleChildBandCount() As Long
    Dim Index As Long
    Dim counter As Long
    For Index = 0 To m_navigationBar.ChildBands.Count - 1
        If m_navigationBar.ChildBands(Index).Visible Then
            counter = counter + 1
        End If
    Next
    GetVisibleChildBandCount = counter
End Function

Private Sub m_ab_BandOpen(ByVal Band As ActiveBar3LibraryCtl.Band, ByVal Cancel As ActiveBar3LibraryCtl.ReturnBool)
    
    Dim eachChildBand As Band
    Dim childBandTool As Tool
    
    ' if it is the 'add/remove' popup, we have to set the checks
    If Band.Name = NavigationAddRemovePopupName Then
        
        ' loop through all the child bands in the navigation bar
        ' and set the tool's check properties
        For Each eachChildBand In m_navigationBar.ChildBands
            
            ' get a reference to the tool that is associated with this childband
            Set childBandTool = m_addRemovePopup.Tools(NavigationAddRemoveToolPrefix + eachChildBand.Name)
            
            ' set the checked property accordingly
            childBandTool.Checked = eachChildBand.Visible
            
        Next
    End If
    
End Sub

Private Sub m_ab_NavBarCustomize(ByVal Band As ActiveBar3LibraryCtl.Band)

    ' if this band is clicked twice, no need to rebuild the add/remove buttons
    If Not Band Is m_navigationBar Then
    
        ' set the navigationbar
        SetNavigationBar Band
        
    End If
    
    ' enable/disable buttons
    m_showFewer.Enabled = CBool(m_navigationBar.LargeButtonCount > 0)
    m_showMore.Enabled = CBool(m_navigationBar.LargeButtonCount < GetVisibleChildBandCount())
    
    ' popup the menu bar
    m_customPopup.PopupMenu
    
End Sub

Private Sub m_ab_ToolClick(ByVal Tool As ActiveBar3LibraryCtl.Tool)

    Select Case Tool.Name
        Case m_showFewer.Name
            
            ' show one less button
            m_navigationBar.LargeButtonCount = m_navigationBar.LargeButtonCount - 1
            m_navigationBar.Refresh
            
        Case m_showMore.Name
            
            ' show one more button
            m_navigationBar.LargeButtonCount = m_navigationBar.LargeButtonCount + 1
            m_navigationBar.Refresh
            
        Case m_navBarOptions.Name
        
            ' user should display a dialog here
            DebugPrint "Display your own dialog here"
            
        Case Else
        
            ' only trap events for the add/remove buttons on the navigation bar
            If InStr(Tool.Name, NavigationAddRemoveToolPrefix) > 0 Then
                
                Dim childBandName As String
                
                ' get the child band name from the tool's name (NOTE: Defined in 'SetNavigationBar()'
                childBandName = Replace$(Tool.Name, NavigationAddRemoveToolPrefix, "")
                
                ' set the visible property of the associating childband
                m_navigationBar.ChildBands(childBandName).Visible = Not Tool.Checked
                
                ' show the changes
                m_ab.RecalcLayout
            End If
            
    End Select
End Sub
