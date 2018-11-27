VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   9300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13635
   BeginProperty Font 
      Name            =   "ZELDA"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   9300
   ScaleWidth      =   13635
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      Caption         =   "FACTS/JOKES LEFT"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10920
      TabIndex        =   9
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Bahnschrift"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   1455
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   5880
      Top             =   8880
   End
   Begin VB.CommandButton BtnInstruction 
      BackColor       =   &H8000000E&
      Height          =   1815
      Left            =   11160
      Picture         =   "CHATBOT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton BtnQuestion 
      BackColor       =   &H8000000E&
      Height          =   1815
      Left            =   11160
      Picture         =   "CHATBOT.frx":25E9
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4800
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox convo 
      Height          =   7095
      Left            =   1200
      TabIndex        =   5
      Top             =   240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   12515
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"CHATBOT.frx":4A20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bahnschrift Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6000
      Top             =   600
   End
   Begin VB.CommandButton enter 
      BackColor       =   &H8000000E&
      Height          =   1815
      Left            =   11160
      Picture         =   "CHATBOT.frx":4AAF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6960
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000E&
      Height          =   1815
      Left            =   5520
      Picture         =   "CHATBOT.frx":65FD
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "HP Simplified"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3840
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   2400
      Width           =   5055
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1200
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   7440
      Width           =   9135
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   375
      URL             =   "App.Path & ""\"" &  Ding"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   661
   End
   Begin VB.Label LabelTimer 
      Caption         =   "20"
      Height          =   255
      Left            =   5400
      TabIndex        =   4
      Top             =   8880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   2760
      Picture         =   "CHATBOT.frx":814B
      Top             =   480
      Width           =   7200
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SOUND EFFECT
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal IpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'STRINGS FOR CHAT
Dim botname As String
Dim nickname As String
Dim beconvo As String
Dim txtconvo As String

Dim partfact As Integer
Dim time(1) As String
'STRING FOR ARRAYS
Dim reaction(2) As String
Dim emotion(6) As String
Dim greeting(7) As String
Dim leave(6) As String
Dim bad(2) As String
Dim fact(1) As String
Dim feeling(1) As String
Dim question(1) As String
Dim questionone(0) As String
Dim questiontwo(0) As String
Dim questionthree(0) As String
Dim questionfour(0) As String
Dim questionfive(0) As String
Dim questionsix(1) As String
Dim questionseven(1) As String
Dim questioneight(1) As String
Dim questionnine(0) As String
Dim questionten(0) As String
Dim questiononeone(0) As String
Dim questiononetwo(0) As String
Dim questiononethree(1) As String
Dim questiononefour(0) As String



Private Sub LoadUser()
Form5.Height = 6510
Form5.Width = 12150
Me.BackColor = RGB(255, 255, 255)
'SHOW THE TEXTBOX WHERE USER WILL WRITE USERNAME
beconvo = convo.Text
Image1.Visible = True
Text2.Visible = True
Text2.Text = ""
Command1.Visible = True
convo.Visible = False
Text.Visible = False
BtnInstruction.Visible = False
BtnQuestion.Visible = False
enter.Visible = False
botname = "XENO: "
Command1.Default = True
convo.Text = txtconvo

End Sub


Private Sub BtnInstruction_Click()
Form3.Show
Form5.Hide
End Sub

Private Sub BtnQuestion_Click()
Me.Hide
QUESTIONS.Show
End Sub

Private Sub Command1_Click()
Call CheckUserBlank
End Sub
Private Sub CheckUserBlank()
If Text2.Text = "" Then
Message = MsgBox("Please don't leave the username box blank", vbOKOnly, "Blank Userbox")
Else: Call LoadConvoForm
End If
End Sub
Private Sub LoadConvoForm()
Form5.Height = 9885
Form5.Width = 13875
'ENTER USERNAME DETAILS DISAPPEAR
    Image1.Visible = False
    
        Text2.Visible = False
        
            Command1.Visible = False
            
                convo.Visible = True
                
                    Text.Visible = True
                    
                        enter.Visible = True
                        
                            Text.Text = ""
                            
                                nickname = Text2.Text
                                        
                                        BtnInstruction.Visible = True
                                        
                                                BtnQuestion.Visible = True
                                                        
                                                        Label1.Visible = True
                                                            
                                                            Frame1.Visible = True
                                    
'ENTER KEY ON KEYBOARD CAN BE USED
        enter.Default = True
'CURSOR GOES TO INPUT TEXTBOX
        Me.Text.SetFocus
    
        Text.Text = imput
        
Call Start
End Sub

Private Sub UserAgain()
Image1.Visible = True
Text2.Visible = True
Text2.Text = ""
Command1.Visible = True
convo.Visible = False
Text.Visible = False
enter.Visible = False
End Sub
Private Sub Start()
'STARTING CONVERSATION
convo.Text = botname & "Hello, " & nickname & ", " & "my name is Xeno. How are you doing?" & vbNewLine
'               Call speech
End Sub

Private Sub Convo_Change()
convo.SelStart = Len(convo.Text)

LabelTimer.Caption = 0
Timer2.Enabled = True
End Sub

Private Sub Enter_Click()
If Text.Text = "" Then
Q = MsgBox("Please type in something in the imput box", vbOKOnly, "Blank")
End If
'WRITE TEXT THAT USER WRITES IN MAIN TEXTBOX (CONVO.TEXT)
beconvo = convo.Text
convo.Text = beconvo & nickname & ": " & imput & vbNewLine
' Call speech
Call Conversation
End Sub


Private Sub Form_Load()

Call LoadUser

time(0) = "DATE"
time(1) = "TIME"

'USER GREETING
greeting(0) = "HI"
greeting(1) = "HELLO"
greeting(2) = "HEY"
greeting(3) = "G'DAY"
greeting(4) = "HOWDY"
greeting(5) = "BONJOUR"
greeting(6) = "MORNING"
greeting(7) = "HEY THERE"

'USER BYE
leave(0) = "BYE"
leave(1) = "AU REVOIR"
leave(2) = "GOODBYE"
leave(3) = "SEE YOU SOON"
leave(4) = "SEE YOU LATER"
leave(5) = "GTG"
leave(6) = "GOT TO GO"

'EMOTIONS 1
emotion(0) = "GOOD"
emotion(1) = "GREAT"
emotion(2) = "AMAZING"
emotion(3) = "WELL"
emotion(4) = "OKAY"
emotion(5) = "OK"
emotion(6) = "NOT BAD"

'EMOTIONS 2
bad(0) = "BAD"
bad(1) = "HORRIBLE"
bad(2) = "SAD"

'USER QUESTIONS
feeling(0) = "HOW ARE YOU"
feeling(1) = "HOW YOU DOING"



'QUESTIONS TO THE BOT
question(0) = "CARNATIC"
question(1) = "MUSIC"

questionone(0) = "INSTRUMENT"
questiontwo(0) = "VEENA"
questionthree(0) = "MIRUDANGAM"
questionfour(0) = "SHRUTI"
questionfive(0) = "TRINITY"

questionsix(0) = "THYAGARAJA"
questionsix(1) = "SWAMIGAL"

questionseven(0) = "SHYAMA"
questionseven(1) = "SASTRI"

questioneight(0) = "MUTHUSWAMI"
questioneight(1) = "DIKSHITAR"

questionnine(0) = "KRITI"
questionten(0) = "THLLAANA"
questiononeone(0) = "SWARA"
questiononetwo(0) = "STHAYI"

questiononethree(0) = "AROHANAM"
questiononethree(1) = "AVAROHANAM"

questiononefour(0) = "JANAKA"

'REACTIONS TO ANSWERS
reaction(0) = "WOW"
reaction(1) = "COOL"
reaction(2) = "I DIDN'T KNOW THAT"


'KEYWORD FACTS
fact(0) = "FACT"
fact(1) = "FUN"

End Sub



Private Sub Conversation()

For num = 0 To UBound(greeting)
    If InStr(1, Text.Text, greeting(num), vbTextCompare) <> 0 Then
    Call GreetingSub
    Text.Text = ""
    Exit Sub
End If
Next num

For num = 0 To UBound(time)
    If InStr(1, Text.Text, time(num), vbTextCompare) <> 0 Then
    Call timesub
    Text.Text = ""
    Exit Sub
End If
Next

'GREETING ARRAY
For num = 0 To UBound(leave)
    If InStr(1, Text.Text, leave(num), vbTextCompare) <> 0 Then
    Call LeaveSub
    Text.Text = ""
    Exit Sub
End If
Next num

'EMOTION ARRAY
For num = 0 To UBound(emotion)
    If InStr(1, Text.Text, emotion(num), vbTextCompare) <> 0 Then
    Call EmotiveSub
    Text.Text = ""
    Exit Sub
End If
Next num

'BAD ARRAY
For num = 0 To UBound(bad)
    If InStr(1, Text.Text, bad(num), vbTextCompare) <> 0 Then
    Call bademotion
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(feeling)
    If InStr(1, Text.Text, feeling(num), vbTextCompare) <> 0 Then
    Call feelingsub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(fact)
    If InStr(1, Text.Text, fact(num), vbTextCompare) <> 0 Then
    Call factsub
    Text.Text = ""
    Exit Sub
End If
Next

'QUESTIONS
For num = 0 To UBound(question)
    If InStr(1, Text.Text, question(num), vbTextCompare) <> 0 Then
    Call questionsub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(questionone)
    If InStr(1, Text.Text, questionone(num), vbTextCompare) <> 0 Then
    Call questiononesub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(questiontwo)
    If InStr(1, Text.Text, questiontwo(num), vbTextCompare) <> 0 Then
    Call questiontwoSub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(questionthree)
    If InStr(1, Text.Text, questionthree(num), vbTextCompare) <> 0 Then
    Call questionthreeSub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(questionfour)
    If InStr(1, Text.Text, questionfour(num), vbTextCompare) <> 0 Then
    Call questionfourSub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(questionfive)
    If InStr(1, Text.Text, questionfive(num), vbTextCompare) <> 0 Then
    Call questionfiveSub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(questionsix)
    If InStr(1, Text.Text, questionsix(num), vbTextCompare) <> 0 Then
    Call questionsixSub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(questionseven)
    If InStr(1, Text.Text, questionseven(num), vbTextCompare) <> 0 Then
    Call questionsevenSub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(questioneight)
    If InStr(1, Text.Text, questioneight(num), vbTextCompare) <> 0 Then
    Call questioneightSub
    Text.Text = ""
    Exit Sub
End If
Next
For num = 0 To UBound(questionnine)
    If InStr(1, Text.Text, questionnine(num), vbTextCompare) <> 0 Then
    Call questionnineSub
    Text.Text = ""
    Exit Sub
End If
Next
For num = 0 To UBound(questionten)
    If InStr(1, Text.Text, questionten(num), vbTextCompare) <> 0 Then
    Call questiontenSub
    Text.Text = ""
    Exit Sub
End If
Next
For num = 0 To UBound(questiononeone)
    If InStr(1, Text.Text, questiononeone(num), vbTextCompare) <> 0 Then
    Call questiononeoneSub
    Text.Text = ""
    Exit Sub
End If
Next
For num = 0 To UBound(questiononetwo)
    If InStr(1, Text.Text, questiononetwo(num), vbTextCompare) <> 0 Then
    Call questiononetwoSub
    Text.Text = ""
    Exit Sub
End If
Next
For num = 0 To UBound(questiononethree)
    If InStr(1, Text.Text, questiononethree(num), vbTextCompare) <> 0 Then
    Call questiononethreeSub
    Text.Text = ""
    Exit Sub
End If
Next
For num = 0 To UBound(questiononefour)
    If InStr(1, Text.Text, questiononefour(num), vbTextCompare) <> 0 Then
    Call questiononefourSub
    Text.Text = ""
    Exit Sub
End If
Next
For num = 0 To UBound(fact)
    If InStr(1, Text.Text, fact(num), vbTextCompare) <> 0 Then
    Call factsub
    Text.Text = ""
    Exit Sub
End If
Next

For num = 0 To UBound(reaction)
    If InStr(1, Text.Text, reaction(num), vbTextCompare) <> 0 Then
    Call reactionsub
    Text.Text = ""
    Exit Sub
End If
Next
' IF USER DOESNT UNDERSTAND
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "I don't quite understand" & vbNewLine
        Text.Text = ""
End Sub

'QUESTION ANSWERS FROM HERE ON
Private Sub GreetingSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "How are you?" & vbNewLine
        
End Sub
Private Sub LeaveSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "don't leave me alone ;-; " & vbNewLine
        Answer = MsgBox("Are you sure you want to quit?", vbYesNo, "Quit")
        If Answer = vbYes Then
        Form5.Hide
        Form1.Show
        End If
        
End Sub
Private Sub EmotiveSub()
       convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "That's great to know!" & vbNewLine & botname & "Do you know what I love? Carnatic music. It's beautiful." & vbNewLine & botname & "You can ask me questions about Carnatic Music. Just ask me a question (you can refer to the question bank), and we can get started!" & vbNewLine
End Sub
Private Sub bademotion()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "It'll get better soon :( Do you want to learn more about Carnatic Music? You can ask me a question (or if you don't know any, then refer to the question bank), or I can tell you a fun fact/joke! Who knows, maybe you'll be :) again" & vbNewLine
End Sub
Private Sub feelingsub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "I'm doing great, thank you for asking!" & vbNewLine
End Sub
Private Sub questionsub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "Indian music is able to be categorised into 2 major components: North Indian and South Indian music." & vbNewLine & botname & "The South Indian music is known as Carnatic music, its name coming from the Tamil word 'Carnatakam', meaning very ancient." & vbNewLine
End Sub
Private Sub questiononesub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "Many instruments are utilised, including the Thambura, Veena, Violin, Flute, Mirudangam and Ghatam. These instruments are used as accompaniments to vocal performances. " & vbNewLine
End Sub
Private Sub questiontwoSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "The veena is considered the 'queen of the music instruments' in India. It is the national instrument of India and is also one of the oldest instruments." & vbNewLine & botname & "It is made of, most of the time, jack wood. However, if the entire veena is made from a single piece of jack wood, then the veena is called an Ekanda Veena" & vbNewLine
End Sub
Private Sub questionthreeSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "The mirudangam is a percussion instrument, used as am accompaniment to vocal and instrumental performances." & vbNewLine & botname & "It is hollowed out of a block of wood, generally Jack, red or Margosa wood. Animal skins are used for the making of this instrument. Fun fact, in North Indian Music, the instrument is known as a Pakkawaj and has a smaller head." & vbNewLine
End Sub
Private Sub questionfourSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "The shruti box is an instrument that works on a system of bellows. It is used as an accompaniment to performances and used to tune the voice in classical singing. " & vbNewLine
End Sub
Private Sub questionfiveSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "The Carnatic Trinity consists of Sri Thyagaraja Swamigal, Sri Shyama Sastri, and Sri Muthuswami Dikshitar. They play a huge part in Indian philosophy, in where many people believe their songs help an individual onto the path of salvation." & vbNewLine & botname & " If you'll like to know in more detail about the individuals in the trinity, ask! " & vbNewLine
End Sub
Private Sub questionsixSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "Thyagaraja Swamigal (B: 1767, D:1847) was one of the three in the music trinity. He learnt many songs from his mother and was a devotee of Lord Naratha" & vbNewLine & botname & "He once sang a song at 8pm and finished at 4am, the next day. The Maharajah of Thanjaavur heard about this and invited Thyagaraja to give him a gift, which he refused. " & vbNewLine
End Sub
Private Sub questionsevenSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "Sri Shyama Sastri (B: 1762, D:1827) was one of the three in the music trinity. He was fluent in both Sanskrit, Tamil and Telugu as a young child, and began to compose keertanais in both languages" & vbNewLine & botname & "None of his family were musicians but he was taught some music by his uncle. He composed over 300 songs however unfortunately most of them are lost. " & vbNewLine
End Sub
Private Sub questioneightSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "Sri Muthuswami Dikshitar (B: 1776, D: 1835) was the youngest of the three in the music trinity. He is famous for his kriti's composed in Sanskrit. Currently, 461 kritis in over 191 ragams are available today" & vbNewLine & botname & "You can ask me about kritis if you don't know what they are." & vbNewLine & botname & "Once, he decided to go visit his only surviving brother, however on his way there he noticed the dry conditions of the land. He composed a song and sang it with emotion, and in no time at all there was a heavy downpour!" & vbNewLine
End Sub
Private Sub questionnineSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "A kriti includes a Pallavi, Anupallavi and Charanam." & vbNewLine & botname & "The Pallavi is the start of the song (generally two lines), the anupallavi has two to three lines, follows the Pallavi and is generally at a higher pitch than the pallavi." & vbNewLine & botname & "The Charanam follows the anupallavi and is the last part to the kriti. The kriti ends after singing all three and then the Pallavi again. " & vbNewLine
End Sub
Private Sub questiontenSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "A thllaana is composed in such a way so that it creates a sense of enthusiasm and joy." & vbNewLine & botname & "It is usually sang in dance concerts, and gives scope to the dancer to display their skill in foot works. They may be composed in Sanskrit, Telugu or Tamil." & vbNewLine
End Sub
Private Sub questiononeoneSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "A swara is a particular type of musical sound. The basic swara (known as saptha swara) is S, R, G, M, P, D, N (pronounced SA, RI, GA, MA, PA, DA, NI)." & vbNewLine & botname & "This is known as the double harmonic scale in Western Music Theory. If you'll like to know how to sing it, watch this video: https://www.youtube.com/watch?v=I7C3E6PJ0CA " & vbNewLine
End Sub
Private Sub questiononetwoSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "A sthayi is an octave in Carnatic Music. There are 3 main sthayis." & vbNewLine & botname & "Mandra Sthayi (Thakkustayi)- a dot is placed under the notes to show that it belongs in the lower octave." & vbNewLine & botname & "Madhya Sthaya (Samasthayi)- this is the most used sthayi. No dots are used to indicate that it belongs in the middle octave." & vbNewLine & botname & "Thaarasthayi (Hechusthayi)- a dot is placed over the notes to show that it belongs in the higher octave. " & vbNewLine
End Sub
Private Sub questiononethreeSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "Every ragam has an arohanam and an avarohanam. The arohanam is the swaram in ascending order (S, R, G, M, P, D, N, S) and the avarohanam is the swaram in descending order (S, N, D, P, M, G, R, S)." & vbNewLine
End Sub
Private Sub questiononefourSub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "A Janaka ragam has seven swarams in both arohanam and avarohanam order. They must occur in regular order and the same type of swaram must occur in both the arohanam and avarohanam order." & vbNewLine
End Sub
Private Sub factsub()
partfact = partfact + 1
Call fsub
End Sub
Private Sub reactionsub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "I know! Do you want to ask another question or hear a fun fact? Ask away!" & vbNewLine
End Sub
Private Sub timesub()
        convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "The current date and time is " & Now & vbNewLine
End Sub

Private Sub fsub()
'IF THE USER WANTS A FACT
Select Case partfact
Case 1
'NUMBER OF FACTS LEFT
Label1.Caption = 4
'SOUND EFFECT
soundfile = App.Path & "/ding.wav"
returnval = PlaySound(soundfile, 0, &H0)
'FACT
convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "What do both cricket and carnatic music have in common? In both, the pitch is important" & vbNewLine & botname & "You have 4 jokes/fun facts left" & vbNewLine
Case 2
Label1.Caption = 3
soundfile = App.Path & "/ding.wav"
returnval = PlaySound(soundfile, 0, &H0)
    
convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "Swara: What the audience tries not to do when listening to bad music." & vbNewLine & botname & "You have 3 jokes/fun facts left" & vbNewLine
Case 3
Label1.Caption = 2
soundfile = App.Path & "/ding.wav"
returnval = PlaySound(soundfile, 0, &H0)
    
convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "It is said that Thyagaraja has composed over 24000 songs!" & vbNewLine & botname & "You have 2 jokes/fun facts left" & vbNewLine
Case 4
Label1.Caption = 1
soundfile = App.Path & "/ding.wav"
returnval = PlaySound(soundfile, 0, &H0)
    
convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "The first documented song came from 1200 CE" & vbNewLine & botname & "You have 1 joke/fun fact left" & vbNewLine
Case 5
Label1.Caption = None
soundfile = App.Path & "/ding.wav"
returnval = PlaySound(soundfile, 0, &H0)
    
convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "A crater on Mercury is named after Thyagaraja" & vbNewLine & botname & "You have no more jokes/fun facts left" & vbNewLine

Call Unavailable

End Select

End Sub
Private Sub Unavailable()
'WHEN CASE RUNS OUT
convo.Text = beconvo & nickname & ": " & Text.Text & vbNewLine & botname & "I have no more jokes/fun facts left. If you want to ask a question about Carnatic Music, ask away. " & vbNewLine
End Sub


Private Sub Timer2_Timer()
'IF USER DOES NOT DO ANYTHING FOR 20 SECONDS
LabelTimer.Caption = Val(LabelTimer) + 1
If LabelTimer.Caption > 20 Then
    convo.Text = beconvo & vbNewLine & botname & "Hey are you still there? I feel lonely" & vbNewLine
LabelTimer.Caption = 0
End If
End Sub
