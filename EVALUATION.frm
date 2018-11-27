VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   9705
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11610
   LinkTopic       =   "Form2"
   ScaleHeight     =   9705
   ScaleWidth      =   11610
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      BackColor       =   &H80000008&
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
      Left            =   5400
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   8760
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000008&
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
      Left            =   5400
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   7320
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000008&
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
      Left            =   5400
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
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
      Left            =   5400
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   2880
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   1800
      Left            =   9840
      Picture         =   "EVALUATION.frx":0000
      Top             =   3600
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   9840
      Picture         =   "EVALUATION.frx":14F0
      Top             =   5640
      Width           =   1800
   End
   Begin VB.Line Line2 
      BorderWidth     =   6
      X1              =   9600
      X2              =   9600
      Y1              =   2280
      Y2              =   9600
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000012&
      Caption         =   "/10"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   6120
      TabIndex        =   21
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000012&
      Caption         =   "ANSWER:"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3960
      TabIndex        =   20
      Top             =   8880
      Width           =   1455
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000012&
      Caption         =   "ANSWER:"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3960
      TabIndex        =   19
      Top             =   7440
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   "ANSWER:"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3960
      TabIndex        =   18
      Top             =   3000
      Width           =   1455
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000012&
      Caption         =   "ANSWER:"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3960
      TabIndex        =   17
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000012&
      Caption         =   "ANSWER:"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   495
      Left            =   3960
      TabIndex        =   16
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000012&
      Caption         =   "/10"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   6120
      TabIndex        =   11
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Caption         =   "YES/NO: DID YOU HAVE TO REFER TO THE QUESTION BANK FREQUENTLY?"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3000
      TabIndex        =   9
      Top             =   8040
      Width           =   6855
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000012&
      Caption         =   "YES/NO: WOULD YOU RECOMMEND XENO TO A FRIEND?"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3000
      TabIndex        =   8
      Top             =   6600
      Width           =   6855
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "OUT OF 10: HOW WELL WAS XENO ABLE TO ANSWER THE QUESTIONS?"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3000
      TabIndex        =   7
      Top             =   5160
      Width           =   6855
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "YES/NO: WAS THE CONVERSATION FLUENT?"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3000
      TabIndex        =   6
      Top             =   3600
      Width           =   6855
   End
   Begin VB.Line Line1 
      BorderWidth     =   6
      X1              =   2880
      X2              =   2880
      Y1              =   2160
      Y2              =   9480
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
      Caption         =   "OUT OF 10: WAS THE GUI AESTHETICALLY PLEASING?"
      BeginProperty Font 
         Name            =   "Nirmala UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   3120
      TabIndex        =   5
      Top             =   2160
      Width           =   6855
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000008&
      Caption         =   "QUESTION 5"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   960
      TabIndex        =   4
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000008&
      Caption         =   "QUESTION 4"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   960
      TabIndex        =   3
      Top             =   6600
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000008&
      Caption         =   "QUESTION 3"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   960
      TabIndex        =   2
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000008&
      Caption         =   "QUESTION 2"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   960
      TabIndex        =   1
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   2400
      Picture         =   "EVALUATION.frx":284C
      Top             =   360
      Width           =   7200
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000008&
      Caption         =   "QUESTION 1"
      BeginProperty Font 
         Name            =   "Bahnschrift"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numbertotext As String

Private Sub Form_Load()
Me.BackColor = RGB(0, 0, 0)
Me.Height = 11010
Me.Width = 12280
Line1.BorderColor = RGB(255, 209, 80)
Line2.BorderColor = RGB(255, 209, 80)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text1.ForeColor = RGB(255, 209, 80)
Text2.ForeColor = RGB(255, 209, 80)
Text3.ForeColor = RGB(255, 209, 80)
Text4.ForeColor = RGB(255, 209, 80)
Text5.ForeColor = RGB(255, 209, 80)
End Sub

Private Sub Image2_Click()
AnswerTwo = MsgBox("Are you sure you want to leave?", vbYesNo, "Quit")
If AnswerTwo = vbYes Then
Form2.Hide
Form1.Show
End If
End Sub

Private Sub Image3_Click()

If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
Answer = MsgBox("Please fill in all the boxes", vbOKOnly, "Empty")
Else: Answerthree = MsgBox("Thank you for your response", vbOKOnly, "Thank you")
If Answerthree = vbOK Then
Form2.Hide
Form1.Show
End If
End If
End Sub


Private Sub Text2_Change()
  numbertotext = Text2.Text
  If IsNumeric(numbertotext) Then
MsgBox "Integers are not allowed"
Text2.Text = ""
End If
End Sub

Private Sub Text4_Change()
  numbertotext = Text4.Text
  If IsNumeric(numbertotext) Then
MsgBox "Integers are not allowed"
Text4.Text = ""
End If
End Sub

Private Sub Text5_Change()
  numbertotext = Text5.Text
  If IsNumeric(numbertotext) Then
MsgBox "Integers are not allowed"
Text5.Text = ""
End If
End Sub
