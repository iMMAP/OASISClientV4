Attribute VB_Name = "mdOASISDataImport"
Option Explicit
Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000 ' new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0

Public gbHideGUI As Boolean

Public g_sAppPath As String

Sub Main()
    
    CreateAppPath
    
    CheckCommands
    
    If g_sAppPath = "" Then g_sAppPath = App.Path
    
    If InStr(UCase$(Command$), "ADMIN") > 0 Then
        frmWFPMain.c1TabMain.TabVisible(1) = True
    End If

    If InStr(UCase$(Command$), "HIDDEN") > 0 Then
        gbHideGUI = True
        frmWFPMain.cmdUpdate_Click
        DoEvents
        On Error Resume Next
        Unload frmWFPMain
        End
    Else
        frmWFPMain.Show
    End If


End Sub


Private Sub CheckCommands()
Dim a_strArgs() As String
Dim blnDebug As Boolean
Dim strFileName As String
Dim ofs As FileSystemObject
Dim i As Integer
    
   
    a_strArgs = Split(Command$, " ")

    For i = LBound(a_strArgs) To UBound(a_strArgs)

        Select Case LCase(a_strArgs(i))

            Case "-path"
                Set ofs = New FileSystemObject
                On Error Resume Next
                If IsNumeric(a_strArgs(i + 1)) Then
                    g_sAppPath = (SpecialFolderPathEx(a_strArgs(i + 1)))
                
                    If Not ofs.FolderExists(g_sAppPath) Then
                        g_sAppPath = ""
                    End If
                
                Else
                    If ofs.FolderExists(a_strArgs(i + 1)) Then
                        g_sAppPath = a_strArgs(i + 1)
                    End If
                End If
                
                Set ofs = Nothing
        End Select
      
      
    Next i

End Sub

Private Sub CreateAppPath()
Dim ofs As New FileSystemObject

    If ofs.FolderExists(SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        g_sAppPath = SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client"
    ElseIf ofs.FolderExists(SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        g_sAppPath = SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client"
    ElseIf ofs.FolderExists(SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        g_sAppPath = SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client"
    ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        g_sAppPath = SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client"
    ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        g_sAppPath = SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client"
    End If
    
End Sub

