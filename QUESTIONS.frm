VERSION 5.00
Begin VB.Form QUESTIONS 
   Caption         =   "Form4"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   LinkTopic       =   "Form4"
   ScaleHeight     =   7380
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      BackColor       =   &H80000014&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Bahnschrift Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   7215
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   8280
      Picture         =   "QUESTIONS.frx":0000
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000014&
      Caption         =   "Q U E S T I O N S"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2160
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "QUESTIONS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.BackColor = RGB(255, 255, 255)
Label2.Caption = "What's the time?" & vbNewLine & "What is Carnatic music?" & vbNewLine & "What are some instruments used in Carnatic Music?" & vbNewLine & "What is the veena?" & vbNewLine & "What is the Mirudangam?" & vbNewLine & "What is a Shruti Box?" & vbNewLine & "Who are the Trinity?" & vbNewLine & "Who is Sri Thyagaraja Swamigal?" & vbNewLine & "Who is Sri Shyama Sastri" & vbNewLine & "Who is Sri Muthuswami Dikshitar?" & vbNewLine & "What is a kriti?" & vbNewLine & "What is a Thllaana?" & vbNewLine & "What is a swara?" & vbNewLine & "What is a Sthayi?" & vbNewLine & "What is the arohanam and avarohanam?" & vbNewLine & "What is a Janaka ragam?"
End Sub

Private Sub Image1_Click()
Me.Hide
Form5.Show
End Sub
