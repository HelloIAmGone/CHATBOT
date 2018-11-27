VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   8880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7245
   LinkTopic       =   "Form3"
   ScaleHeight     =   8880
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.Image Image2 
      Height          =   1800
      Left            =   2760
      Picture         =   "Form3.frx":0000
      Top             =   7080
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   7050
      Left            =   1440
      Picture         =   "Form3.frx":1B4C
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.BackColor = RGB(255, 255, 255)
End Sub


Private Sub Image2_Click()
Form3.Hide
Form5.Show
End Sub
