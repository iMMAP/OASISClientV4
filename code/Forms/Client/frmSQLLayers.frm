VERSION 5.00
Begin VB.Form frmSQLLayers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OASIS SQL Layers"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4605
   Icon            =   "frmSQLLayers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox comPath 
      Height          =   315
      Left            =   30
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   540
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.ComboBox comSQLLayers 
      Height          =   315
      Left            =   90
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   420
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   285
      Left            =   2820
      TabIndex        =   0
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "frmSQLLayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public m_bOK As Boolean

Private Sub cmdCancel_Click()
    m_bOK = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    m_bOK = True
    Me.Hide
End Sub

Public Function FileExists(sFullPath As String) As Boolean
        '<EhHeader>
        On Error GoTo FileExists_Err
        '</EhHeader>
        Dim oFile As New Scripting.FileSystemObject
100     FileExists = oFile.FileExists(sFullPath)
        '<EhFooter>
        Exit Function

FileExists_Err:
        MsgBox Err.Description & vbCrLf & "in OASISClient.frmDynamicDataMenu.FileExists " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Public Sub Init(sAdo As String)

    On Error Resume Next
    
    Dim oRS As New ADODB.Recordset
    Dim oCn As New ADODB.Connection
     
    comSQLLayers.Clear
    comSQLLayers.AddItem "--NONE--"
                
    With oCn
        .CursorLocation = adUseClient
        .ConnectionString = sAdo
        .Open
    End With
                
    If oCn.State = adStateOpen Then
                
        Set oRS = oCn.OpenSchema(adSchemaTables)
        oRS.Sort = "ORDINAL_POSITION DESC"
                   
        While Not oRS.EOF
            
            If right(oRS!TABLE_NAME, 4) = "_FEA" Then
                
                If DoesTableExist(sAdo, left(oRS!TABLE_NAME, Len(oRS!TABLE_NAME) - 4) & "_GEO") Then
                    comSQLLayers.AddItem right$(left$(oRS!TABLE_NAME, Len(oRS!TABLE_NAME) - 4), Len(left$(oRS!TABLE_NAME, Len(oRS!TABLE_NAME) - 4)) - 3)
                End If
                
            End If
               
            oRS.MoveNext
               
        Wend
                    
        oRS.Close
        oCn.Close
        
    End If
    
    If comSQLLayers.ListCount > 0 Then comSQLLayers.ListIndex = 0
    Set oCn = Nothing
    Set oRS = Nothing
    
End Sub

