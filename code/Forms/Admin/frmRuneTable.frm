VERSION 5.00
Begin VB.Form RuneTable 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11115
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRuneTable.frx":0000
   ScaleHeight     =   9705
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select any 5 cards by clicking on it"
      ForeColor       =   &H0000FF00&
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   7560
      Width           =   2295
   End
   Begin VB.Image Info 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   10320
      ToolTipText     =   "Legend Of Runes"
      Top             =   0
      Width           =   270
   End
   Begin VB.Image Ext 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   10680
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   360
   End
   Begin VB.Image InfoDown 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   -8400
      Picture         =   "frmRuneTable.frx":16E342
      Top             =   7320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ExtDown 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   -9240
      Picture         =   "frmRuneTable.frx":16E7D0
      Top             =   7320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image InfoUp 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   -8760
      Picture         =   "frmRuneTable.frx":16EDEE
      Top             =   7320
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image ExtUp 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   -9600
      Picture         =   "frmRuneTable.frx":16F27C
      Top             =   7320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   24
      Left            =   7320
      Top             =   7320
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   23
      Left            =   5760
      Top             =   7320
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   22
      Left            =   4200
      Top             =   7320
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   21
      Left            =   2640
      Top             =   7320
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   20
      Left            =   9480
      Top             =   5040
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   19
      Left            =   7920
      Top             =   5040
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   18
      Left            =   6360
      Top             =   5040
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   17
      Left            =   4800
      Top             =   5040
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   16
      Left            =   3240
      Top             =   5040
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   15
      Left            =   1680
      Top             =   5040
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   14
      Left            =   120
      Top             =   5040
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   13
      Left            =   9480
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   12
      Left            =   7920
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   11
      Left            =   6360
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   10
      Left            =   4800
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   9
      Left            =   3240
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   8
      Left            =   1680
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   7
      Left            =   120
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   6
      Left            =   9480
      Top             =   480
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   5
      Left            =   7920
      Top             =   480
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   4
      Left            =   6360
      Top             =   480
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   3
      Left            =   4800
      Top             =   480
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   2
      Left            =   3240
      Top             =   480
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   1
      Left            =   1680
      Top             =   480
      Width           =   1560
   End
   Begin VB.Image CardTable 
      BorderStyle     =   1  'Fixed Single
      Height          =   2310
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   1560
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   25
      Left            =   360
      Picture         =   "frmRuneTable.frx":16F89A
      Top             =   -9960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   24
      Left            =   360
      Picture         =   "frmRuneTable.frx":17155E
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   23
      Left            =   360
      Picture         =   "frmRuneTable.frx":171DE7
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   22
      Left            =   360
      Picture         =   "frmRuneTable.frx":17273C
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   21
      Left            =   360
      Picture         =   "frmRuneTable.frx":173118
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   20
      Left            =   360
      Picture         =   "frmRuneTable.frx":173A2A
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   19
      Left            =   360
      Picture         =   "frmRuneTable.frx":174433
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   18
      Left            =   360
      Picture         =   "frmRuneTable.frx":174DC7
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   17
      Left            =   360
      Picture         =   "frmRuneTable.frx":175717
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   16
      Left            =   360
      Picture         =   "frmRuneTable.frx":176178
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   15
      Left            =   360
      Picture         =   "frmRuneTable.frx":176B33
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   14
      Left            =   360
      Picture         =   "frmRuneTable.frx":1774A5
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   13
      Left            =   360
      Picture         =   "frmRuneTable.frx":177E5B
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   12
      Left            =   360
      Picture         =   "frmRuneTable.frx":178809
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   11
      Left            =   360
      Picture         =   "frmRuneTable.frx":179100
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   10
      Left            =   360
      Picture         =   "frmRuneTable.frx":179A51
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   9
      Left            =   360
      Picture         =   "frmRuneTable.frx":17A445
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   8
      Left            =   360
      Picture         =   "frmRuneTable.frx":17AE44
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   7
      Left            =   360
      Picture         =   "frmRuneTable.frx":17B82C
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   6
      Left            =   360
      Picture         =   "frmRuneTable.frx":17C17D
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   5
      Left            =   360
      Picture         =   "frmRuneTable.frx":17CB1C
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   4
      Left            =   360
      Picture         =   "frmRuneTable.frx":17D441
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   3
      Left            =   360
      Picture         =   "frmRuneTable.frx":17DD80
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   2
      Left            =   360
      Picture         =   "frmRuneTable.frx":17E661
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   1
      Left            =   360
      Picture         =   "frmRuneTable.frx":17F016
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Cards 
      Height          =   2250
      Index           =   0
      Left            =   360
      Picture         =   "frmRuneTable.frx":17F9B5
      Top             =   -8400
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "RuneTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TempInt As Integer, TempArr(24) As Integer, MainArr(24) As Integer
Dim CardNos As Integer
Dim Notes(24, 4) As String
Function Reset()
        '<EhHeader>
        On Error GoTo Reset_Err
        '</EhHeader>
        
        Dim i As Integer

100     For i = 0 To 24
102         CardTable(i).BorderStyle = 0
104         CardTable(i).Picture = Cards(25).Picture
106         TempArr(i) = 0
108     Next i

110     CardNos = -1
112     SetCards
114     ResultTable.Text1.Text = " "
116     AddNotes
118     Info.BorderStyle = 0
120     Ext.BorderStyle = 0
122     Info.Picture = InfoUp.Picture
124     Ext.Picture = ExtUp.Picture
        '<EhFooter>
        Exit Function

Reset_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.Reset " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Function SetCards()
        '<EhHeader>
        On Error GoTo SetCards_Err
        '</EhHeader>
        Dim i As Integer

100     For i = 0 To 24
Repeat:
            Rem Randomize (Second(Now) + Day(Now))
102         TempInt = Int((25) * Rnd() + 0)

104         If TempArr(TempInt) = 1 Then
106             GoTo Repeat
            Else
108             MainArr(i) = TempInt
110             TempArr(TempInt) = 1
            End If

112     Next i

        '<EhFooter>
        Exit Function

SetCards_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.SetCards " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function
Function AddNotes()
        'Aansoos
        '<EhHeader>
        On Error GoTo AddNotes_Err
        '</EhHeader>
100     Notes(0, 0) = " At your curent state you are displaying your Talking abilities(speeches) and your intelligence. You are or will be getting warnings on some matters. Many things are happening unexpectedly for you. A message is waiting you."
102     Notes(0, 1) = " This period is very pleasent for you and won't face big challenges during this time."
104     Notes(0, 2) = " You should display your talking skills,knowledge and mental courage so that you face your challenges bravely."
106     Notes(0, 3) = " You should display your talking skills,knowledge and mental courage so that you face your challenges bravely."
108     Notes(0, 4) = " At the end you may get help from elders and will display your mental courage, knowledge and bravery. Be prepared to hear a good news."
        'Barkana
110     Notes(1, 0) = " You are getting help during various stages of your growth."
112     Notes(1, 1) = " Your actions may lead to some wounded relations. So be careful."
114     Notes(1, 2) = " Recheck your ideals and the factors that helped in your spiritual growth."
116     Notes(1, 3) = " Recheck your ideals and the factors that helped in your spiritual growth."
118     Notes(1, 4) = " At the end there may be wounded relations and prosperity."
        'Dagus
120     Notes(2, 0) = " Your difficulties are about to end and you are waiting for good times.You are or may undergo tranformations and is feeling secure. You will escape from accidents."
122     Notes(2, 1) = " You will find that there is a delay in almost all your activities."
124     Notes(2, 2) = " Wait for a favourable time."
126     Notes(2, 3) = " Wait for a favourable time."
128     Notes(2, 4) = " At the end employment, End to the delay of marriage, Encouragement from others, Freedom and End to your difficulties is predicted."
        'Eesha
130     Notes(3, 0) = " This is period where you won't have neither growth, prosperity ,luck neither well being. You will feel extremely lazy. Your mind will be occupied with hatred, anger, and tension."
132     Notes(3, 1) = " You may face a lot of problems during this period such as laziness,anger,tension,failure,sadness etc. Overall it is a bad time."
134     Notes(3, 2) = " Try to conquer laziness,tension and anger."
136     Notes(3, 3) = " Try to conquer laziness,tension and anger."
138     Notes(3, 4) = " At the end your hard work will become a waste if not careful. You may feel extremely lazy and this will be a period of concern."
        'Elhaas
140     Notes(4, 0) = " You have some desire in your mind and is praying to God for it."
142     Notes(4, 1) = " Pray to God and control your desires."
144     Notes(4, 2) = " Pray to God and control your desires."
146     Notes(4, 3) = " Pray to God and control your desires."
148     Notes(4, 4) = " You will correct someone's path, and will decide to study. You may donate dress, food and will have guests. A change in the profession is also predicted. You will get help during journeys."
        'Fehu
150     Notes(5, 0) = " You are having a very prosperous time with physical wealths, gains, wealth, progress, victory, luck and will succeed in love."
152     Notes(5, 1) = " You are having a very prosperous time with physical wealths, gains, wealth, progress, victory, luck and will succeed in love."
154     Notes(5, 2) = " You are having a very prosperous time with physical wealths, gains, wealth, progress, victory, luck and will succeed in love."
156     Notes(5, 3) = " You are having a very prosperous time with physical wealths, gains, wealth, progress, victory, luck and will succeed in love."
158     Notes(5, 4) = " You are having a very prosperous time with physical wealths, gains, wealth, progress, victory, luck and will succeed in love."
        'Geho
160     Notes(6, 0) = " An aggregate of various characters and emotions is seen. A chance of marriage, reunion, retrieval of a lost object, love, meeting your relations and unexpected help from them are also predicted."
162     Notes(6, 1) = " An aggregate of various characters and emotions is seen. A chance of marriage, reunion, retrieval of a lost object, love, meeting your relations and unexpected help from them are also predicted."
164     Notes(6, 2) = " An aggregate of various characters and emotions is seen. A chance of marriage, reunion, retrieval of a lost object, love, meeting your relations and unexpected help from them are also predicted."
166     Notes(6, 3) = " An aggregate of various characters and emotions is seen. A chance of marriage, reunion, retrieval of a lost object, love, meeting your relations and unexpected help from them are also predicted."
168     Notes(6, 4) = " An aggregate of various characters and emotions is seen. A chance of marriage, reunion, retrieval of a lost object, love, meeting your relations and unexpected help from them are also predicted."
        'Haegalus
170     Notes(7, 0) = " Your paths are filled with difficulties. There is an indication of changes and freedom."
172     Notes(7, 1) = " Your path will be filled with difficulties but there are indications of a change."
174     Notes(7, 2) = " Have a hope of a good tommorow and handle your difficulties with proper care. "
176     Notes(7, 3) = " Have a hope of a good tommorow and handle your difficulties with proper care."
178     Notes(7, 4) = " You will hear the news of a birth. Take special care in your health."
        'Inguvaas
180     Notes(8, 0) = " You will successfully complete any task which you will carry out. Prosperity, growth, progress  and concern about the comming days are also predicted. You will find the energy to face any difficulty in yourself."
182     Notes(8, 1) = " You will successfully complete any task which you will carry out. Prosperity, growth, progress  and concern about the comming days are also predicted. You will find the energy to face any difficulty in yourself.  But still, one of your novel idea won't reach completion."
184     Notes(8, 2) = " You will successfully complete any task which you will carry out. Prosperity, growth, progress  and concern about the comming days are also predicted. You will find the energy to face any difficulty in yourself."
186     Notes(8, 3) = " You will successfully complete any task which you will carry out. Prosperity, growth, progress  and concern about the comming days are also predicted. You will find the energy to face any difficulty in yourself."
188     Notes(8, 4) = " You will successfully complete any task which you will carry out. Prosperity, growth, progress  and concern about the comming days are also predicted. You will find the energy to face any difficulty in yourself."
        'Iwaas
190     Notes(9, 0) = " You need to face certain difficulties today which will prove helpful tommorow. Stability, Longevity, Resistance, Ways to progress and destruction of every evils is predicted."
192     Notes(9, 1) = " You need to face certain difficulties today which will prove helpful tommorow. Stability, Longevity, Resistance, Ways to progress and destruction of every evils is predicted."
194     Notes(9, 2) = " You need to face certain difficulties today which will prove helpful tommorow. Stability, Longevity, Resistance, Ways to progress and destruction of every evils is predicted."
196     Notes(9, 3) = " You need to face certain difficulties today which will prove helpful tommorow. Stability, Longevity, Resistance, Ways to progress and destruction of every evils is predicted."
198     Notes(9, 4) = " You need to face certain difficulties today which will prove helpful tommorow. Stability, Longevity, Resistance, Ways to progress and destruction of every evils is predicted."
        'Jiraa
200     Notes(10, 0) = " Things are happening around you in a very slow ratevand everything is happening according to the laws of nature."
202     Notes(10, 1) = " Your desires will have a long delay to get completed."
204     Notes(10, 2) = " Your desires will have a long delay to get completed."
206     Notes(10, 3) = " Your desires will have a long delay to get completed."
208     Notes(10, 4) = " Things are happening around you in a very slow ratevand everything is happening according to the laws of nature."
        'Kaenaas
210     Notes(11, 0) = " You will get help from your relations, difficulties in your path are vanishing and will have a good time. Successful romance and marriage are also predicted."
212     Notes(11, 1) = " You will get help from your relations, difficulties in your path are vanishing and will have a good time. Successful romance and marriage are also predicted."
214     Notes(11, 2) = " You will get help from your relations, difficulties in your path are vanishing and will have a good time. Successful romance and marriage are also predicted."
216     Notes(11, 3) = " You will get help from your relations, difficulties in your path are vanishing and will have a good time. Successful romance and marriage are also predicted."
218     Notes(11, 4) = " You will get help from your relations, difficulties in your path are vanishing and will have a good time. Successful romance and marriage are also predicted."
        'Lagoos
220     Notes(12, 0) = " The predictions includes emotions, dreams, success in your activities, attaining mysterious knowledge, attention towards spiritual matters, listening to your inner mind and forseeing future, birth of a girl child in the family."
222     Notes(12, 1) = " The predictions includes emotions, dreams, success in your activities, attaining mysterious knowledge, attention towards spiritual matters, listening to your inner mind and forseeing future, birth of a girl child in the family."
224     Notes(12, 2) = " The predictions includes emotions, dreams, success in your activities, attaining mysterious knowledge, attention towards spiritual matters, listening to your inner mind and forseeing future, birth of a girl child in the family."
226     Notes(12, 3) = " The predictions includes emotions, dreams, success in your activities, attaining mysterious knowledge, attention towards spiritual matters, listening to your inner mind and forseeing future, birth of a girl child in the family."
228     Notes(12, 4) = " The predictions includes emotions, dreams, success in your activities, attaining mysterious knowledge, attention towards spiritual matters, listening to your inner mind and forseeing future, birth of a girl child in the family."
        'Mannas
230     Notes(13, 0) = " The study of your own intelligence, identity, personality, difficulties etc. are predicted. Know about your family and factors disturbing your mind. Victory and failure begins from you."
232     Notes(13, 1) = " The study of your own intelligence, identity, personality, difficulties etc. are predicted. Know about your family and factors disturbing your mind. Victory and failure begins from you."
234     Notes(13, 2) = " The study of your own intelligence, identity, personality, difficulties etc. are predicted. Know about your family and factors disturbing your mind. Victory and failure begins from you."
236     Notes(13, 3) = " The study of your own intelligence, identity, personality, difficulties etc. are predicted. Know about your family and factors disturbing your mind. Victory and failure begins from you."
238     Notes(13, 4) = " The study of your own intelligence, identity, personality, difficulties etc. are predicted. Know about your family and factors disturbing your mind. Victory and failure begins from you."
        'Myvoh
240     Notes(14, 0) = " Predictions are friendship and belief, emotional unity, cooperation and progress, changing residence, travel, growth and bride&groom."
242     Notes(14, 1) = " Predictions are friendship and belief, emotional unity, cooperation and progress, changing residence, travel, growth and bride&groom."
244     Notes(14, 2) = " Predictions are friendship and belief, emotional unity, cooperation and progress, changing residence, travel, growth and bride&groom."
246     Notes(14, 3) = " Predictions are friendship and belief, emotional unity, cooperation and progress, changing residence, travel, growth and bride&groom."
248     Notes(14, 4) = " Predictions are friendship and belief, emotional unity, cooperation and progress, changing residence, travel, growth and bride&groom."
        'Noutis
250     Notes(15, 0) = " You will explore into dark sides of your life and wiil have to work hard. Pain, sadness and misfortunes are seen. You will fight ahead the difficulties confronting you."
252     Notes(15, 1) = " You will explore into dark sides of your life and wiil have to work hard. Pain, sadness and misfortunes are seen. You will fight ahead the difficulties confronting you."
254     Notes(15, 2) = " You will explore into dark sides of your life and wiil have to work hard. Pain, sadness and misfortunes are seen. You will fight ahead the difficulties confronting you."
256     Notes(15, 3) = " You will explore into dark sides of your life and wiil have to work hard. Pain, sadness and misfortunes are seen. You will fight ahead the difficulties confronting you."
258     Notes(15, 4) = " You will explore into dark sides of your life and wiil have to work hard. Pain, sadness and misfortunes are seen. You will fight ahead the difficulties confronting you."
        'Othello
260     Notes(16, 0) = " Social life and partition are seen. You may loose something dearest to you."
262     Notes(16, 1) = " Social life and partition are seen. You may loose something dearest to you."
264     Notes(16, 2) = " Social life and partition are seen. You may loose something dearest to you."
266     Notes(16, 3) = " Social life and partition are seen. You may loose something dearest to you."
268     Notes(16, 4) = " Social life and partition are seen. You may loose something dearest to you."
        'Perthru
270     Notes(17, 0) = " You may get unexpected wealth. You may and will have to keep some secret. It is a time without much gains, but you may get a job, if you are searching for it."
272     Notes(17, 1) = " You may get unexpected wealth. You may and will have to keep some secret. It is a time without much gains, but you may get a job, if you are searching for it."
274     Notes(17, 2) = " You may get unexpected wealth. You may and will have to keep some secret. It is a time without much gains, but you may get a job, if you are searching for it."
276     Notes(17, 3) = " You may get unexpected wealth. You may and will have to keep some secret. It is a time without much gains, but you may get a job, if you are searching for it."
278     Notes(17, 4) = " You may get unexpected wealth. You may and will have to keep some secret. It is a time without much gains, but you may get a job, if you are searching for it."
        'Rurisaas
280     Notes(18, 0) = " You may hear a bad news and get strong warnings. You will have a new confidence and will find explore new fields."
282     Notes(18, 1) = " You may hear a bad news and get strong warnings. You will have a new confidence and will find explore new fields."
284     Notes(18, 2) = " You may hear a bad news and get strong warnings. You will have a new confidence and will find explore new fields."
286     Notes(18, 3) = " You may hear a bad news and get strong warnings. You will have a new confidence and will find explore new fields."
288     Notes(18, 4) = " You may hear a bad news and get strong warnings. You will have a new confidence and will find explore new fields."
        'Rydo
290     Notes(19, 0) = " Travel and a thirst for spirituality is forseen. Directions, journey, cooperation, reunion, friendly visits, helps etc. are also seen."
292     Notes(19, 1) = " Travel and a thirst for spirituality is forseen. Directions, journey, cooperation, reunion, friendly visits, helps etc. are also seen."
294     Notes(19, 2) = " Travel and a thirst for spirituality is forseen. Directions, journey, cooperation, reunion, friendly visits, helps etc. are also seen."
296     Notes(19, 3) = " Travel and a thirst for spirituality is forseen. Directions, journey, cooperation, reunion, friendly visits, helps etc. are also seen."
298     Notes(19, 4) = " Travel and a thirst for spirituality is forseen. Directions, journey, cooperation, reunion, friendly visits, helps etc. are also seen."
        'Sovilo
300     Notes(20, 0) = " Your difficulties seems to end and you can continue the journey.Have courage and self confidence.You must need rest and take care of your health."
302     Notes(20, 1) = " Your difficulties seems to end and you can continue the journey.Have courage and self confidence.You must need rest and take care of your health."
304     Notes(20, 2) = " Your difficulties seems to end and you can continue the journey.Have courage and self confidence.You must need rest and take care of your health."
306     Notes(20, 3) = " Your difficulties seems to end and you can continue the journey.Have courage and self confidence.You must need rest and take care of your health."
308     Notes(20, 4) = " Your difficulties seems to end and you can continue the journey.Have courage and self confidence.You must need rest and take care of your health."
        'Thivaas
310     Notes(21, 0) = " Justice, love, courage, sacrificing your wealth and well being for others, helping others, leading others are predicted. Durng this time you may show authority, get a job and may find success in all your activities."
312     Notes(21, 1) = " Justice, love, courage, sacrificing your wealth and well being for others, helping others, leading others are predicted. Durng this time you may show authority, get a job and may find success in all your activities."
314     Notes(21, 2) = " Justice, love, courage, sacrificing your wealth and well being for others, helping others, leading others are predicted. Durng this time you may show authority, get a job and may find success in all your activities."
316     Notes(21, 3) = " Justice, love, courage, sacrificing your wealth and well being for others, helping others, leading others are predicted. Durng this time you may show authority, get a job and may find success in all your activities."
318     Notes(21, 4) = " Justice, love, courage, sacrificing your wealth and well being for others, helping others, leading others are predicted. Durng this time you may show authority, get a job and may find success in all your activities."
        'Uroos
320     Notes(22, 0) = " Creative abilities, health, strength, gains, progress, fiscal gains, introduction to new environments, sexual strength, growth using your own abilities are foreseen."
322     Notes(22, 1) = " Creative abilities, health, strength, gains, progress, fiscal gains, introduction to new environments, sexual strength, growth using your own abilities are foreseen."
324     Notes(22, 2) = " Creative abilities, health, strength, gains, progress, fiscal gains, introduction to new environments, sexual strength, growth using your own abilities are foreseen."
326     Notes(22, 3) = " Creative abilities, health, strength, gains, progress, fiscal gains, introduction to new environments, sexual strength, growth using your own abilities are foreseen."
328     Notes(22, 4) = " Creative abilities, health, strength, gains, progress, fiscal gains, introduction to new environments, sexual strength, growth using your own abilities are foreseen."
        'Voonjo
330     Notes(23, 0) = " This is a time during which misunderstandings vanishes and love replaces of hatred, Prosperity replaces poverty and Victory repalces failure. You have happy good days ahead and your desires will be fulfilled. Visit to other nations and the path to success is also predicted."
332     Notes(23, 1) = " This is a time during which misunderstandings vanishes and love replaces of hatred, Prosperity replaces poverty and Victory repalces failure. You have happy good days ahead and your desires will be fulfilled. Visit to other nations and the path to success is also predicted."
334     Notes(23, 2) = " This is a time during which misunderstandings vanishes and love replaces of hatred, Prosperity replaces poverty and Victory repalces failure. You have happy good days ahead and your desires will be fulfilled. Visit to other nations and the path to success is also predicted."
336     Notes(23, 3) = " This is a time during which misunderstandings vanishes and love replaces of hatred, Prosperity replaces poverty and Victory repalces failure. You have happy good days ahead and your desires will be fulfilled. Visit to other nations and the path to success is also predicted."
338     Notes(23, 4) = " This is a time during which misunderstandings vanishes and love replaces of hatred, Prosperity replaces poverty and Victory repalces failure. You have happy good days ahead and your desires will be fulfilled. Visit to other nations and the path to success is also predicted."
        'Yaid
340     Notes(24, 0) = " You are after something that is unattainable. An answerless mystery is confronting you."
342     Notes(24, 1) = " You are after something that is unattainable. An answerless mystery is confronting you."
344     Notes(24, 2) = " You are after something that is unattainable. An answerless mystery is confronting you."
346     Notes(24, 3) = " You are after something that is unattainable. An answerless mystery is confronting you."
348     Notes(24, 4) = " You are after something that is unattainable. An answerless mystery is confronting you."
        '<EhFooter>
        Exit Function

AddNotes_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.AddNotes " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Function

Private Sub CardTable_Click(Index As Integer)
        '<EhHeader>
        On Error GoTo CardTable_Click_Err
        '</EhHeader>
100     CardNos = CardNos + 1

102     If CardNos <= 4 Then
104         CardTable(Index).Picture = Cards(MainArr(Index)).Picture
106         ResultTable.Image1(CardNos) = Cards(MainArr(Index)).Picture
108         ResultTable.Text1.Text = ResultTable.Text1.Text + Notes(MainArr(Index), CardNos)
        End If

110     If CardNos = 4 Then
112         Unload Me
114         ResultTable.Show
        End If

        '<EhFooter>
        Exit Sub

CardTable_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.CardTable_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Command1_Click()
        '<EhHeader>
        On Error GoTo Command1_Click_Err
        '</EhHeader>
        
        Dim i As Integer

100     For i = 0 To 24
102         CardTable(i).Picture = Cards(MainArr(i)).Picture
104     Next i

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.Command1_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Ext_Click()
        '<EhHeader>
        On Error GoTo Ext_Click_Err
        '</EhHeader>
100     Unload ResultTable
102     Unload Legend
104     Unload Me
        '<EhFooter>
        Exit Sub

Ext_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.Ext_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Ext_MouseDown(Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)
        '<EhHeader>
        On Error GoTo Ext_MouseDown_Err
        '</EhHeader>
100     Ext.Picture = ExtDown.Picture
        '<EhFooter>
        Exit Sub

Ext_MouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.Ext_MouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Ext_MouseMove(Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)
        '<EhHeader>
        On Error GoTo Ext_MouseMove_Err
        '</EhHeader>
100     Ext.Picture = ExtUp.Picture
        '<EhFooter>
        Exit Sub

Ext_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.Ext_MouseMove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Form_Load()
        '<EhHeader>
        On Error GoTo Form_Load_Err
        '</EhHeader>
100     Reset
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.Form_Load " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Info_Click()
        '<EhHeader>
        On Error GoTo Info_Click_Err
        '</EhHeader>
100     Legend.Show
        '<EhFooter>
        Exit Sub

Info_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.Info_Click " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Info_MouseDown(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
        '<EhHeader>
        On Error GoTo Info_MouseDown_Err
        '</EhHeader>
100     Info.Picture = InfoDown.Picture
        '<EhFooter>
        Exit Sub

Info_MouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.Info_MouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Info_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           X As Single, _
                           Y As Single)
        '<EhHeader>
        On Error GoTo Info_MouseMove_Err
        '</EhHeader>
100     Info.Picture = InfoUp.Picture
        '<EhFooter>
        Exit Sub

Info_MouseMove_Err:
        MsgBox Err.Description & vbCrLf & _
               "in OASISRemoteAdmin.RuneTable.Info_MouseMove " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub
