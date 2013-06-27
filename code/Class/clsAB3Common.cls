VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AB3Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents m_ab As ActiveBar3
Attribute m_ab.VB_VarHelpID = -1
Private m_tbsCombo As Tool

Public Enum AB3ErrorEnum
    UnknownError = 1001
    LogicError = 1002
End Enum

Friend Sub Init(ByVal ab3 As ActiveBar3)
    Set m_ab = ab3
End Sub

Friend Sub ThrowException(ByVal Message As String, Optional ByVal errSource As String = "AB3Common", Optional ByVal exType As AB3ErrorEnum = LogicError)
    Err.Raise vbObjectError + exType, errSource, Message
End Sub

Friend Function GetUniqueToolID() As Long
Static nextToolID As Long

    ' if it has been set, set it
    If nextToolID = 0 Then
        nextToolID = 20000
    End If
    
    ' increment for the next tool id
    nextToolID = nextToolID + 1
    
    ' return the next valid id
    GetUniqueToolID = nextToolID

End Function

Friend Function FindBand(ByVal BandName As String) As Band
        '<EhHeader>
        On Error GoTo FindBand_Err
        '</EhHeader>
       
        Dim eachBand As Band
        Dim bandIndex As Long
    
        ' loop through each band in activebar
100     For bandIndex = 0 To m_ab.Bands.Count - 1
102         Set eachBand = m_ab.Bands(bandIndex)
        
            ' if the band names match, return the object it does
104         If eachBand.Name = BandName Then
106             Set FindBand = eachBand
                Exit Function
            End If
        
        Next
    
        '<EhFooter>
        Exit Function

FindBand_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISClient.AB3Common.FindBand", _
                  "AB3Common component failure Line:" & Erl & Err.Description & Err.number
        '</EhFooter>
End Function

Friend Function IndexOfBand(ByVal thisBand As Band) As Long

    Dim bandIndex As Long
    For bandIndex = 0 To m_ab.Bands.Count - 1
        If m_ab.Bands(bandIndex) Is thisBand Then
            IndexOfBand = bandIndex
            Exit Function
        End If
    Next
    
    IndexOfBand = -1
    
End Function
Friend Function FindSharedTool(ByVal toolName As String) As Tool

    Dim eachTool As Tool
    
    ' loop through all shared tools
    For Each eachTool In m_ab.Tools
    
        ' check each shared tool's name and return if it matches
        If eachTool.Name = toolName Then
            Set FindSharedTool = eachTool
            Exit Function
        End If
        
    Next
        
End Function

Friend Function FindSpecificTool(ByVal toolName As String) As Tool
    
    Dim eachBand As Band
    Dim eachTool As Tool
    
    ' loop through each band
    For Each eachBand In m_ab.Bands
    
        ' in each band, loop through each tool
        For Each eachTool In eachBand.Tools
            
            ' check to see if the name matches
            If eachTool.Name = toolName Then
            
                ' if so, return this tool
                Set FindSpecificTool = eachTool
                Exit Function
                
            End If
        
        Next
    
    Next
        
End Function

Friend Function FindCBItemData(ByVal Combo As Tool, ByVal ItemData As Long) As Integer
    Dim Index As Long
    
    ' check each item data agains the value passed in
    For Index = 0 To Combo.CBListCount - 1
        If Combo.CBItemData(Index) = ItemData Then
            FindCBItemData = Index
            Exit Function
        End If
    Next
    
    ' was not found
    FindCBItemData = -1
    
End Function

Friend Function FindItemData(ByVal comboOrList As Object, ByVal Value As Long) As Integer
    Dim Index As Long
    
    ' ensure we have a listbox or a combo box
    If Not (TypeOf comboOrList Is ComboBox Or TypeOf comboOrList Is ListBox) Then
        ThrowException "Expecting a listbox or combobox", "AB3Common.FindItemData"
    End If
    
    ' find the item data that matches the value passed in
    For Index = 0 To comboOrList.ListCount - 1
        If comboOrList.ItemData(Index) = Value Then
            FindItemData = Index
            Exit Function
        End If
    Next
    
    ' return invalid index
    FindItemData = -1
    
End Function

Friend Function StripName(ByVal caption As String) As String
Dim cleanName As String
    
    ' remove nasty characters
    cleanName = Replace$(caption, "&", "")
    cleanName = Replace$(cleanName, "...", "")
    cleanName = Replace$(cleanName, " ", "")
    
    ' return the cleaned up name
    StripName = cleanName
    
End Function

Friend Sub ABAddFlag(ByVal bandFlag As ActiveBar3LibraryCtl.BandFlags, ByVal Band As ActiveBar3LibraryCtl.Band)
    Band.Flags = Band.Flags Or bandFlag
End Sub

Friend Sub ABRemoveFlag(ByVal bandFlag As ActiveBar3LibraryCtl.BandFlags, ByVal Band As ActiveBar3LibraryCtl.Band)
    Band.Flags = Band.Flags And Not bandFlag
End Sub

Friend Function ABAddFlagEx(ByVal bandFlag As ActiveBar3LibraryCtl.BandFlags, ByVal currentFlags As Long) As Long
    currentFlags = currentFlags Or bandFlag
    ABAddFlagEx = currentFlags
End Function

Friend Function ABRemoveFlagEx(ByVal bandFlag As ActiveBar3LibraryCtl.BandFlags, ByVal currentFlags As Long)
    currentFlags = currentFlags And Not bandFlag
    ABRemoveFlagEx = currentFlags
End Function

