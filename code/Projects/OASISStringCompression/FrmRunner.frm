VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCommand1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCommand1_Click()
        '<EhHeader>
        On Error GoTo cmdCommand1_Click_Err
        '</EhHeader>
        Dim aa As New OASISStringCompression.OASISCompression
        Dim en As String
        Dim dn As String
    
100     dd = "keith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6ukeith doyle 234j32io5hiuo43h6u"
102     dd = dd & "rh43 nmt iu4b tf78g4 fdbc87rq42h rdxn8743 nrxd8t bgn4378dx 5gn7853g ntx87naewrg tx8fbewF Cnmfd vf dsav vgtw4efzjhckhfzdbkj"
104     dd = dd & "fdg kjlhd cshb e9758rc yxz598s7dy zdn8cg5784g cx8gs97  ngcdn987g cs8b7876zs4 nx48gx8s g764 cgiktru7"
106     dd = dd & "gfrdh mc8976h89345nb,54wen4d 7954w nm468954376894738965748056kjhvehkrtwjand  dd"
108     dd = dd & "grtffk654e89fj6n9843 6534879 4e 875 875437  5943 977 9579437 58 7684r 573vt y7h uh5rgv ubsvg khszdrhlktzv iusrhgy"
110     dd = dd & "etjgf98hdgrs79 vhg7 gh rgh78 vzgs 8gvg8g7vg8sg v8yiuved"
112     dd = dd & "76 ne2oi5m 2esmn rh32 b4 5   vsd rytv frgfr ghrte y5 7 sth fd hyd hg fsd hdfsfd hdfdh"
114     dd = dd & "543m,n6tjk fj3k54h fuershg v987xgygfhjkdgxhjrb z4ngrvui43wsi67v9oxzdljh4n5liovs8u8o37645x lz8 c87twz tv v"
116     dd = dd & "rh43 nmt iu4b tf78g4 fdbc87rq42h rdxn8743 nrxd8t bgn4378dx 5gn7853g ntx87naewrg tx8fbewF Cnmfd vf dsav vgtw4efzjhckhfzdbkj"
118     dd = dd & "fdg kjlhd cshb e9758rc yxz598s7dy zdn8cg5784g cx8gs97  ngcdn987g cs8b7876zs4 nx48gx8s g764 cgiktru7"
120     dd = dd & "gfrdh mc8976h89345nb,54wen4d 7954w nm468954376894738965748056kjhvehkrtwjand  dd"
122     dd = dd & "grtffk654e89fj6n9843 6534879 4e 875 875437  5943 977 9579437 58 7684r 573vt y7h uh5rgv ubsvg khszdrhlktzv iusrhgy"
124     dd = dd & "etjgf98hdgrs79 vhg7 gh rgh78 vzgs 8gvg8g7vg8sg v8yiuved"
126     dd = dd & "76 ne2oi5m 2esmn rh32 b4 5   vsd rytv frgfr ghrte y5 7 sth fd hyd hg fsd hdfsfd hdfdh"
128     dd = dd & "543m,n6tjk fj3k54h fuershg v987xgygfhjkdgxhjrb z4ngrvui43wsi67v9oxzdljh4n5liovs8u8o37645x lz8 c87twz tv v+="
    
130     dd = "getservertime="
132     MsgBox Len(dd) & " [" & dd & "]"

134     en = aa.ConvertByteArrayToString(aa.CompressStringToByteArray(dd))
136     MsgBox Len(en) & " [" & en & "]"
    
140     dn = aa.DecompressStringToString(en)
142     MsgBox Len(dn) & " [" & dn & "]"
    
        '<EhFooter>
        Exit Sub

cmdCommand1_Click_Err:
        MsgBox Err.Description & vbCrLf & "in Project1.Form1.cmdCommand1_Click " & "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
