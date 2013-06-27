VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRC4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''''''''''''''''' Start RC4 section
'    ' RC4 Variables
Private sBox(255)
Private Key(255)

''
' This routine called by EnDeCrypt function.
' Initializes the sbox and the key array
Sub RC4Initialize(strPwd)
        '<EhHeader>
        On Error GoTo RC4Initialize_Err
        '</EhHeader>

        Dim tempSwap
        Dim a
        Dim b
        Dim intLength As Integer

100     intLength = Len(strPwd)

102     For a = 0 To 255
104         Key(a) = Asc(Mid(strPwd, (a Mod intLength) + 1, 1))
106         sBox(a) = a
        Next

108     b = 0

110     For a = 0 To 255
112         b = (b + sBox(a) + Key(a)) Mod 256
114         tempSwap = sBox(a)
116         sBox(a) = sBox(b)
118         sBox(b) = tempSwap
        Next

        '<EhFooter>
        Exit Sub

RC4Initialize_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsRC4.RC4Initialize", _
                  "clsRC4 component failure"
        '</EhFooter>
End Sub

''
' This routine does all the work.
' Call it both to ENcrypt and to DEcrypt your data.
Function RC4(plaintxt, psw)
        '<EhHeader>
        On Error GoTo RC4_Err
        '</EhHeader>

        Dim Temp
        Dim a
        Dim i: i = 0
        Dim j: j = 0
        Dim K
        Dim cipherby
        Dim cipher

100     RC4Initialize psw

102     For a = 1 To Len(plaintxt)
104         i = (i + 1) Mod 256
106         j = (j + sBox(i)) Mod 256
108         Temp = sBox(i)
110         sBox(i) = sBox(j)
112         sBox(j) = Temp

114         K = sBox((sBox(i) + sBox(j)) Mod 256)

116         cipherby = Asc(Mid(plaintxt, a, 1)) Xor K
118         cipher = cipher & Chr(cipherby)
        Next

120     RC4 = cipher

        '<EhFooter>
        Exit Function

RC4_Err:
        Err.Raise vbObjectError + 100, _
                  "OASISRemoteAdmin.clsRC4.RC4", _
                  "clsRC4 component failure"
        '</EhFooter>
End Function

