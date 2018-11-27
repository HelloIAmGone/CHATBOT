VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6165
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   6360
      Picture         =   "MENU.frx":0000
      Top             =   3120
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   1800
      Picture         =   "MENU.frx":17C1
      Top             =   3120
      Width           =   1800
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   1560
      X2              =   8280
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   1560
      X2              =   8280
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Image Image5 
      Height          =   1800
      Left            =   4080
      Picture         =   "MENU.frx":2F95
      Top             =   3120
      Width           =   1800
   End
   Begin VB.Image Image6 
      Height          =   1800
      Left            =   1320
      Picture         =   "MENU.frx":5443
      Top             =   600
      Width           =   7200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Me.BackColor = RGB(255, 255, 255)
Line1.BorderColor = RGB(255, 209, 80)
Line2.BorderColor = RGB(255, 209, 80)
Form1.Height = 6750
Form1.Width = 10065

End Sub

Private Sub Image1_Click()
Form1.Hide
Form5.Show
End Sub

Private Sub Image2_Click()
End
End Sub

Private Sub Image5_Click()
Form1.Hide
Form2.Show
End Sub
