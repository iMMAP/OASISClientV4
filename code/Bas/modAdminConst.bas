Attribute VB_Name = "AdminConsts"
Option Explicit

Public r1
Public r2
Public r3
Public cindex1
Public cindex2
Public cindex3
Public counter1
Public counter2
Public counter3
Public rounds1
Public rounds2
Public rounds3


Public Function CreateAppPath() As String
Dim ofs As New FileSystemObject

On Error Resume Next

    If ofs.FolderExists(SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        CreateAppPath = SpecialFolderPath(sftUserApplicationData) & "iMMAP - OASIS\OASIS client"
    ElseIf ofs.FolderExists(SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        CreateAppPath = SpecialFolderPath(sftUserLocalApplicationData) & "iMMAP - OASIS\OASIS client"
    ElseIf ofs.FolderExists(SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        CreateAppPath = SpecialFolderPath(sftUserMyDocuments) & "iMMAP - OASIS\OASIS client"
    ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        CreateAppPath = SpecialFolderPath(sftCommonApplicationData) & "iMMAP - OASIS\OASIS client"
    ElseIf ofs.FolderExists(SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client\Data\DB") Then
        CreateAppPath = SpecialFolderPath(sftCommonMyDocuments) & "iMMAP - OASIS\OASIS client"
    End If
    
    If Len(CreateAppPath) < 1 Then CreateAppPath = App.Path
    
    Set ofs = Nothing
    
End Function

