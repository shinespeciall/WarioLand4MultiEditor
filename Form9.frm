VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "output"
   ClientHeight    =   14100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   14100
   ScaleWidth      =   5205
   Begin VB.CommandButton Command1 
      Caption         =   "Clear All"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   13320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   12975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form9.Text1.Text = ""
End Sub

Private Sub Form_Load()
Form9.Move Form4.width + 18510, 0, MDIForm1.width - Form4.width - 18510 - 450, MDIForm1.height - 1150
Form9.Text1.FontSize = 20
End Sub
