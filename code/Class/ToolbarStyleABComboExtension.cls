VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolbarStyleABCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_ab As ActiveBar3LibraryCtl.ActiveBar3
Attribute m_ab.VB_VarHelpID = -1
Private m_common As AB3Common
Private m_tbsCombo As Tool

Friend Property Get Combo() As Tool
    Set Combo = m_tbsCombo
End Property

Friend Property Set Combo(ByVal settingCombo As Tool)
    
    ' ensure that we have the right tooltype
    If settingCombo.ControlType <> ddTTCombobox Then
        m_common.ThrowException "Expecting a tool with ControlType = ddTTCombobox", "ToolbarStyleABComboExtension"
    End If
    
    ' set the combo object
    Set m_tbsCombo = settingCombo
    
End Property

Friend Sub CleanUp()
    
    ' reset the combo object and clean up memory
    If Not m_tbsCombo Is Nothing Then
        m_tbsCombo.CBClear
        Set m_tbsCombo = Nothing
    End If
    
    ' clean up all resources
    Set m_ab = Nothing
    Set m_common = Nothing
    
End Sub

Friend Sub Extend()

    ' if this tool was not found, add it a new floating band
    If m_tbsCombo Is Nothing Then
        m_common.ThrowException "Combo tool is not set", "ToolbarStyleABCombo"
    End If
    
    ' setup the combo box
    With m_tbsCombo
    
        ' set the defaults
        .ControlType = ddTTCombobox
        .caption = "Look: "
        .Style = ddSIconText
        .CBWidth = 1500
        .Width = 2000
        .CBStyle = ddCBSReadOnly
        
        ' add each toolbar style
        .CBClear
        
        ' OFFICE 2000
        .CBAddItem "Classic OASIS"  '"Office 2000"
        .CBItemData(.CBNewIndex) = ToolbarStyles.ddTSOffice2000
        
        ' OFFICE XP
        .CBAddItem "XP OASIS"  '"Office XP"
        .CBItemData(.CBNewIndex) = ToolbarStyles.ddTSOfficeXP
        
        ' OFFICE 2003
        .CBAddItem "Standard OASIS" '"Office 2003"
        .CBItemData(.CBNewIndex) = ToolbarStyles.ddTSOffice2003
        
        ' VISUAL STUDIO 2005
        .CBAddItem "iMMAP Look" ' "Next Generation OASIS" '"Visual Studio 2005"
        .CBItemData(.CBNewIndex) = ToolbarStyles.ddTSVisualStudio2005
        
        .CBListIndex = m_common.FindCBItemData(m_tbsCombo, m_ab.ToolbarStyle)
    End With
                
    ' refresh the control
    m_ab.RecalcLayout

End Sub

Private Sub m_ab_ComboSelChange(ByVal Tool As ActiveBar3LibraryCtl.Tool)

    ' handle the click
    If Tool.Name = m_tbsCombo.Name Then
        m_ab.ToolbarStyle = Tool.CBItemData(Tool.CBListIndex)
        m_ab.RecalcLayout
    End If
    
End Sub

Friend Sub Init(ByVal AB As ActiveBar3LibraryCtl.ActiveBar3, ByVal common As AB3Common)
    
    ' referenced for events and so on
    Set m_ab = AB
    
    ' used for building tools
    Set m_common = common
    
End Sub
